from pathlib import Path
from unittest.mock import MagicMock, patch

import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook

from Invoice_Data import CustomsDeclaration
from app.core.workbook_staging import WorkbookStagingCancelled, WorkbookStagingTimeout

from app.services.ezarc_source_service import (
    ALLOCATION_FIELDS, HANGZHOU_SOURCE, SHEET_NAME, SourceAllocation,
    ezarc_maintenance_path, plan_ezarc_sources, split_ezarc_rows, write_ezarc_deductions,
)


NANJING = '32019 南京其他'
JIAXING = '33049 嘉兴'


def make_table(tmp_path, existing=False):
    path = tmp_path / 'maintenance.xlsx'
    book = Workbook()
    sheet = book.active
    sheet.title = SHEET_NAME
    sheet.append(['SKU', '供应商名称', '货源地', '结余需报关数量', '26EA1' if existing else 'old'])
    sheet.append([100, '南京供应商', NANJING, 550, 300 if existing else None])
    sheet.append(['200N', '南京供应商', NANJING, 1840, 1840 if existing else None])
    sheet.append(['300N', '嘉兴供应商', JIAXING, 0, None])
    book.save(path)
    book.close()
    return path


def test_plan_allocates_by_sku_and_only_deducts_hangzhou(tmp_path):
    path = make_table(tmp_path)
    allocations, update = plan_ezarc_sources(path, '26EA1', {'100': 300, '200N': 2000, '300N': 40, 'missing': 90})
    assert allocations == {
        '100': SourceAllocation(300, 0, NANJING, '南京供应商'),
        '200N': SourceAllocation(1840, 160, NANJING, '南京供应商'),
        '300N': SourceAllocation(0, 40, JIAXING, '嘉兴供应商'),
    }
    assert update.insert_column == 6
    assert update.deductions == ((2, 550, 300), (3, 1840, 1840))


def test_existing_ticket_replays_saved_numbers_instead_of_current_balances(tmp_path):
    path = make_table(tmp_path, existing=True)
    allocations, update = plan_ezarc_sources(path, '26EA1', {'100': 300, '200N': 2000, '300N': 40})
    assert allocations['200N'].hangzhou == 1840
    assert allocations['300N'].external == 40
    assert not update.deductions
    with patch('app.services.ezarc_source_service.start_isolated_office_application') as start:
        write_ezarc_deductions(update)
        start.assert_not_called()
    with pytest.raises(ValueError, match='超过本次发货'):
        plan_ezarc_sources(path, '26EA1', {'100': 299, '200N': 2000, '300N': 40})


def test_nothing_to_deduct_does_not_create_a_column(tmp_path):
    path = make_table(tmp_path)
    before = path.read_bytes()
    allocations, update = plan_ezarc_sources(path, '26EA1', {'300N': 40, 'missing': 90})
    write_ezarc_deductions(update)
    assert allocations['300N'].external == 40
    assert path.read_bytes() == before


def test_split_keeps_physical_and_financial_totals():
    row = {'SKU': '200N', '发货数量': 300, '境内货源地': HANGZHOU_SOURCE,
           '境内货源地显示': '杭州其他', '供应商名称': '原供应商', '箱数': 2,
           **dict.fromkeys(ALLOCATION_FIELDS, 300)}
    result = split_ezarc_rows(pd.DataFrame([row]), {'200N': SourceAllocation(250, 50, NANJING, '南京供应商')})
    assert list(result['发货数量']) == [250, 50]
    assert list(result['境内货源地']) == [HANGZHOU_SOURCE, NANJING]
    assert list(result['供应商名称']) == ['原供应商', '南京供应商']
    assert list(result['报关箱数']) == [2, 0]
    assert (result[list(ALLOCATION_FIELDS)].sum() == 300).all()
    assert result['箱数'].sum() == 2


def test_changed_workbook_is_not_saved(tmp_path):
    path = make_table(tmp_path)
    _, update = plan_ezarc_sources(path, '26EA1', {'100': 300})
    path.touch()
    path.write_bytes(path.read_bytes() + b'changed')
    with patch('app.services.ezarc_source_service.start_isolated_office_application') as start:
        with pytest.raises(RuntimeError, match='已被修改'):
            write_ezarc_deductions(update)
    start.assert_not_called()


def test_office_owner_file_prevents_opening_workbook_for_write(tmp_path):
    path = make_table(tmp_path)
    _, update = plan_ezarc_sources(path, '26EA1', {'100': 300})
    before = path.read_bytes()
    owner = path.with_name(f'~${path.name}')
    owner.write_bytes(b'Office owner data')
    with patch('app.services.ezarc_source_service.start_isolated_office_application') as start:
        with pytest.raises(RuntimeError, match='Office 占用标记'):
            write_ezarc_deductions(update)
    start.assert_not_called()
    assert path.read_bytes() == before
    assert owner.read_bytes() == b'Office owner data'
    assert set(tmp_path.iterdir()) == {path, owner}


def test_change_during_local_calculation_preserves_newer_shared_file(tmp_path):
    path = make_table(tmp_path)
    _, update = plan_ezarc_sources(path, '26EA1', {'100': 300})
    newer = path.read_bytes() + b'newer workbook'

    def calculate(local_path, _update):
        assert local_path.parent != path.parent
        assert local_path.read_bytes() == path.read_bytes()
        path.write_bytes(newer)

    with patch('app.services.ezarc_source_service._calculate_local_deductions', side_effect=calculate), \
            patch('app.services.ezarc_source_service._verify_saved_deductions'):
        with pytest.raises(RuntimeError, match='已被修改'):
            write_ezarc_deductions(update)
    assert path.read_bytes() == newer
    assert set(tmp_path.iterdir()) == {path}


def test_failed_atomic_replace_preserves_the_original(tmp_path):
    path = make_table(tmp_path)
    _, update = plan_ezarc_sources(path, '26EA1', {'100': 300})
    before = path.read_bytes()
    with patch('app.services.ezarc_source_service._calculate_local_deductions'), \
            patch('app.services.ezarc_source_service._verify_saved_deductions'), \
            patch('app.services.ezarc_source_service.os.replace', side_effect=PermissionError):
        with pytest.raises(RuntimeError, match='原表未更新'):
            write_ezarc_deductions(update)
    assert path.read_bytes() == before
    assert set(tmp_path.iterdir()) == {path}


def test_invalid_balance_formula_prevents_save(tmp_path):
    path = make_table(tmp_path)
    _, update = plan_ezarc_sources(path, '26EA1', {'100': 300})
    office = MagicMock()
    book = office.Workbooks.Open.return_value
    book.ReadOnly = False
    balance = book.Worksheets.return_value.Cells.return_value
    balance.HasFormula = True
    balance.Value2 = 550
    with patch('app.services.ezarc_source_service.start_isolated_office_application', return_value=office):
        with pytest.raises(ValueError, match='未正确包含本票扣减'):
            write_ezarc_deductions(update)
    book.Save.assert_not_called()
    book.Close.assert_called_once_with(False)


def test_pdf_error_does_not_deduct(tmp_path):
    from app.flows import invoice_flow

    path = make_table(tmp_path)
    _, update = plan_ezarc_sources(path, '26EA1', {'100': 300})
    generator = MagicMock(shipping_details={'invoice_number': '26EA1'})

    def process_invoice(*args):
        generator.ezarc_maintenance_update = update

    generator.process_invoice.side_effect = process_invoice
    with patch.multiple(
            invoice_flow,
            render_detailed_list=MagicMock(), render_packing_list=MagicMock(), render_bill=MagicMock(),
            render_contract=MagicMock(), render_customs_declaration=MagicMock(), render_elements_sheet=MagicMock(),
            ensure_default_sheet_removed=MagicMock(), finalize_workbook_page_setup=MagicMock(),
            export_workbook_to_pdf=MagicMock(), export_worksheet_to_pdf=MagicMock(side_effect=RuntimeError('PDF failed')),
    ), patch.object(invoice_flow, 'write_ezarc_deductions') as write:
        with pytest.raises(RuntimeError, match='PDF failed'):
            invoice_flow.InvoiceFlow(generator).generate_all_documents(
                'invoice', 'product', 'names', 'sites', 'elements', str(tmp_path), 'pictures', False, False,
            )
    write.assert_not_called()


def test_register_all_documents_writes_split_quantity_to_both_ezarc_sheets(tmp_path):
    register_path = tmp_path / '1.易达报关明细.xlsx'
    internal_sheet, external_sheet = '2026（内部开票）', '2026（外部开票）'
    book = Workbook()
    book.active.title = internal_sheet
    book.create_sheet(external_sheet)
    for sheet in book:
        sheet.append(['报关渠道', '合同号'])
        sheet.append(['亚马逊', 'OLD'])
    book.save(register_path)
    book.close()

    shared = {'报关中文品名': '锯片', '是否含金刚石': False, 'HS CODE': '8202911000',
              '英文品名': 'Saw Blade', '单位': '件', '采购员': '', '申报单价': 2}
    g = CustomsDeclaration()
    g.lcl_flag = False
    g.company_info_index = 0
    g.shipping_details.update(brand='EZARC', invoice_number='26EA1', currency='USD')
    g.ezarc_source_rows = pd.DataFrame([
        dict(shared, SKU='200N', 发货数量=250, 境内货源地=HANGZHOU_SOURCE, 境内货源地显示='杭州其他',
             外部开票=False, 申报总价=500, 供应商名称='内部供应商'),
        dict(shared, SKU='200N', 发货数量=50, 境内货源地=NANJING, 境内货源地显示='南京其他',
             外部开票=True, 申报总价=100, 供应商名称='南京供应商'),
        dict(shared, SKU='unmatched', 发货数量=25, 境内货源地=HANGZHOU_SOURCE, 境内货源地显示='杭州其他',
             外部开票=False, 申报总价=50, 供应商名称='其他供应商'),
        dict(shared, SKU='with-diamond', 是否含金刚石=True, 发货数量=15, 境内货源地=HANGZHOU_SOURCE,
             境内货源地显示='杭州其他', 外部开票=False, 申报总价=45, 供应商名称='内部供应商'),
    ])
    g.result_df = pd.DataFrame([
        dict(shared, 是否含金刚石=True, 发货数量=15, 申报单价=3, 总价=45,
             境内货源地=HANGZHOU_SOURCE, 境内货源地显示='杭州其他'),
        dict(shared, 发货数量=50, 总价=100, 境内货源地=NANJING, 境内货源地显示='南京其他'),
        dict(shared, 发货数量=275, 总价=550, 境内货源地=HANGZHOU_SOURCE, 境内货源地显示='杭州其他'),
    ])
    with patch('Invoice_Data.append_call_record'), patch.object(Workbook, 'save', autospec=True, side_effect=Workbook.save) as save:
        g.flow.register_all_documents('shipment.xlsx', 'unused', 'products.xlsx', 'unused', str(register_path),
                                      'unused', 'unused', 'unused', register_shipping=False, register_customs=True)
    save.assert_called_once()
    assert set(tmp_path.iterdir()) == {register_path}

    book = load_workbook(register_path, read_only=True)
    try:
        internal = list(book[internal_sheet].values)
        external = list(book[external_sheet].values)
        assert internal[2][1] == external[2][1] == '26EA1'
        assert [(row[8], row[9]) for row in internal[2:]] == [('200N', 250), ('unmatched', 25), ('with-diamond', 15)]
        assert [row[10:18] for row in internal[2:]] == [
            ('锯片', 'Saw Blade', '8202911000', 15, '件', 3, 45, 'USD'),
            ('锯片', 'Saw Blade', '8202911000', 275, '件', 2, 550, 'USD'),
            (None,) * 8,
        ]
        assert [(row[7], row[9], row[11], row[12], row[14]) for row in external[2:]] == [
            ('200N', 50, 2, 100, '南京供应商'),
        ]
    finally:
        book.close()


def test_external_register_uses_the_same_lines_and_amounts_as_the_declaration():
    g = CustomsDeclaration()
    g.company_info_index = 0
    g.lcl_flag = False
    g.shipping_details.update(g.company_info[0])
    g.shipping_details.update(brand='EZARC', invoice_number='26EA304', currency='USD', country_zh='日本',
                              TOTAL_PACKED_IN=6, freight=0, **{'TOTAL_G.W': 12, 'TOTAL_N.W': 6})
    common = {
        '报关中文品名': '锯片', '箱数': 1, '总PCS': 1, '总重量（kg）-箱子': 2, '箱子总毛重（kg）': 2,
        '总体积（m³）-箱子': .02, '箱子长度（cm）': 40, '箱子宽度（cm）': 20, '报关箱数': 1,
    }
    specifications = [
        ('801401', 80, 834.224, JIAXING, '嘉兴', False, '供应商甲', '采购甲'),
        ('801612', 54, 225.844, JIAXING, '嘉兴', False, '供应商甲', '采购甲'),
        ('801905N', 40, 647.242, JIAXING, '嘉兴', False, '供应商乙', None),
        ('802062', 100, 740.22, NANJING, '南京其他', False, '供应商丙', '采购乙'),
        ('D1', 5, 33.33, JIAXING, '嘉兴', True, '供应商甲', '采购甲'),
        ('HZ', 20, 100, HANGZHOU_SOURCE, '杭州其他', False, '内部供应商', ''),
    ]
    g.ezarc_source_rows = pd.DataFrame([
        dict(common, SKU=sku, 发货数量=quantity, 申报总价=amount, 境内货源地=source, 境内货源地显示=display,
             是否含金刚石=diamond, 外部开票=source != HANGZHOU_SOURCE, 供应商名称=supplier, 采购员=buyer)
        for sku, quantity, amount, source, display, diamond, supplier, buyer in specifications
    ])
    products = pd.DataFrame({'*SKU': g.ezarc_source_rows['SKU'], '英文品名': 'Saw Blade',
                             'HS CODE': '8202911000', '单位': '套'})
    g.result_df = g._summarize_non_lcl(g.ezarc_source_rows, products)
    book = Workbook()
    with patch.object(g, 'add_images'):
        g.add_customs_declaration_sheet('unused.xlsx', 'pictures', wb=book)
    declaration = book['报关单']
    external_lines = [
        (row[1], row[4], row[5], row[6], row[7])
        for row in declaration.iter_rows(min_row=18, max_row=17 + len(g.result_df), values_only=True)
        if row[13] != HANGZHOU_SOURCE
    ]
    book.close()

    entries = {sheet: frame for frame, _, sheet, _, _ in g._build_ezarc_register_entries('unused.xlsx')}
    external = entries['2026（外部开票）']
    assert list(external[['HS编码', '数量', '单位', '单价', '总价']].itertuples(index=False, name=None)) == external_lines
    by_sku = external.set_index('SKU')
    merged = by_sku.loc['801401、801612、801905N']
    assert (merged['数量'], merged['单价'], merged['总价']) == (174, 9.81, 1707.31)
    assert (merged['供应商名称'], merged['采购员']) == ('供应商甲、供应商乙', '采购甲')
    assert set(by_sku.index) == {'801401、801612、801905N', '802062', 'D1'}
    assert list(entries['2026（内部开票）']['SKU']) == ['HZ']


@pytest.mark.parametrize('second_sheet, old_contract, error', [
    (False, '', '缺少工作表'),
    (True, '26EA1 & 26EA2', '发票号已登记'),
])
def test_invalid_second_register_sheet_does_not_save_first(tmp_path, second_sheet, old_contract, error):
    path = tmp_path / 'registration.xlsx'
    book = Workbook()
    book.active.title = 'internal'
    book.active.append(['报关渠道', '合同号'])
    if second_sheet:
        external = book.create_sheet('external')
        external.append(['报关渠道', '合同号'])
        external.append(['亚马逊', old_contract])
    book.save(path)
    book.close()
    before = path.read_bytes()

    generator = CustomsDeclaration()
    generator.shipping_details['invoice_number'] = '26EA1'
    rows = pd.DataFrame([['亚马逊', '26EA1']])
    entries = [(rows, str(path), name, 'EZARC', 1) for name in ('internal', 'external')]
    with patch.object(Workbook, 'save', autospec=True, side_effect=Workbook.save) as save:
        with pytest.raises(ValueError, match=error):
            generator._write_ezarc_register_batch(entries)
    save.assert_not_called()
    assert path.read_bytes() == before
    assert set(tmp_path.iterdir()) == {path}


def test_registration_refuses_to_overwrite_concurrent_changes(tmp_path):
    path = tmp_path / 'registration.xlsx'
    book = Workbook()
    book.active.append(['报关渠道', '合同号'])
    book.save(path)
    book.close()
    original = path.read_bytes()
    generator = CustomsDeclaration()
    generator.shipping_details['invoice_number'] = '26EA1'
    entry = (pd.DataFrame([['亚马逊', '26EA1']]), str(path), 'Sheet', 'EZARC', 1)
    original_save = Workbook.save

    def save_and_update_target(book, result):
        original_save(book, result)
        other = load_workbook(path)
        try:
            other.active['A1'] = 'Concurrent update'
            original_save(other, path)
        finally:
            other.close()

    with patch.object(Workbook, 'save', autospec=True, side_effect=save_and_update_target):
        with pytest.raises(RuntimeError, match='登记期间已被修改'):
            generator._write_ezarc_register_batch([entry])
    assert path.read_bytes() != original
    assert set(tmp_path.iterdir()) == {path}
    book = load_workbook(path, read_only=True)
    try:
        assert book.active['A1'].value == 'Concurrent update'
        assert book.active['B2'].value is None
    finally:
        book.close()


@pytest.mark.parametrize('stop_at_handoff', [False, True])
def test_customs_staging_abort_preserves_original_and_reports_completed_shipping(tmp_path, stop_at_handoff):
    path = tmp_path / 'registration.xlsx'
    book = Workbook()
    book.active.append(['报关渠道', '合同号'])
    book.save(path)
    book.close()
    original = path.read_bytes()
    generator = CustomsDeclaration()
    generator.shipping_details.update(brand='EZARC', invoice_number='26EA1')
    generator.ezarc_source_rows = pd.DataFrame()
    entry = (pd.DataFrame([['亚马逊', '26EA1']]), str(path), 'Sheet', 'EZARC', 1)
    stages = []
    staging_states = []

    def stage_workbook(*args, **kwargs):
        if stop_at_handoff:
            return tmp_path / '.ezarc-register-unpublished.tmp'
        raise WorkbookStagingTimeout('开票刷盘超时，原表未更新')

    with patch.object(generator, 'build_invoicing_register_entries', return_value=[entry]), \
            patch.object(generator, 'register_shipping_invoice_number') as shipping, \
            patch('Invoice_Data.stage_customs_workbook', side_effect=stage_workbook), \
            patch('Invoice_Data.workbook_write_lock') as lock, patch('Invoice_Data.os.replace') as replace, \
            patch('Invoice_Data.log_message') as log, patch('Invoice_Data.append_call_record') as metrics:
        error = WorkbookStagingCancelled if stop_at_handoff else WorkbookStagingTimeout
        with pytest.raises(error, match='原表未更新.*出运明细已登记'):
            generator.register_all_documents(
                'invoice', 'shipping', 'product', 'a', str(path), 'c', 'd', 'e', progress=stages.append,
                staging_state=staging_states.append, staging_cancelled=lambda: stop_at_handoff and staging_states[-1] is False,
            )
    assert staging_states == [True, False]
    shipping.assert_called_once()
    lock.assert_not_called()
    replace.assert_not_called()
    metrics.assert_not_called()
    assert path.read_bytes() == original
    assert stages == ['报关开票：检查登记表', '报关开票：读取登记表', '报关开票：本地查重并生成工作簿',
                      '报关开票：校验新增行', '报关开票：写入并刷盘临时文件']
    assert any('写入并刷盘临时文件耗时' in call.args[0] for call in log.call_args_list)


def test_test_mode_points_to_only_the_test_subdirectory():
    assert ezarc_maintenance_path(True) == ezarc_maintenance_path().parent / '测试' / ezarc_maintenance_path().name
    assert isinstance(ezarc_maintenance_path(), Path)
