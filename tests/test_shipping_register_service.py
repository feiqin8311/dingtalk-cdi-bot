from unittest.mock import patch

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

from app.services.shipping_register_service import SHEET_NAME, append_shipping_register


@pytest.fixture
def shipping_file(tmp_path):
    path = tmp_path / 'shipping.xlsx'
    book = Workbook()
    sheet = book.active
    sheet.title = SHEET_NAME
    sheet.append(['发货日期', '发票号', '数量'])
    sheet.append([None, 'OLD & INV-OLD', 20])
    sheet['A25'].font = Font(bold=True)  # Style alone is not a data row.
    book.create_sheet('其他')['A1'] = '=SUM(1,2)'
    book.save(path)
    book.close()
    return path


def test_append_checks_once_and_preserves_contents_and_style(shipping_file):
    rows = [[None, 'INV-NEW', 13], [None, 'INV-NEW', 13]]
    stages = []
    with patch.object(Workbook, 'save', autospec=True, side_effect=Workbook.save) as save:
        assert append_shipping_register(shipping_file, 'INV-NEW', rows, progress=stages.append) == (2, 4)
    save.assert_called_once()
    assert stages == ['出运明细：读取登记表', '出运明细：查重并准备写入', '出运明细：校验新增行', '出运明细：写回登记表']
    book = load_workbook(shipping_file)
    try:
        sheet = book[SHEET_NAME]
        assert sheet['B2'].value == 'OLD & INV-OLD'
        assert [list(row) for row in sheet.iter_rows(min_row=3, max_row=4, values_only=True)] == rows
        assert sheet['B3'].border.bottom.style == 'thin'
        assert sheet['A25'].font.bold
        assert book['其他']['A1'].value == '=SUM(1,2)'
    finally:
        book.close()
    assert list(shipping_file.parent.iterdir()) == [shipping_file]


def test_preserves_each_rows_logistics_number(shipping_file):
    rows = [[None, 'INV-NEW', 13], [None, 'INV-ANOTHER', 5]]
    assert append_shipping_register(shipping_file, 'INV-NEW', rows) == (2, 4)
    book = load_workbook(shipping_file, read_only=True)
    try:
        assert [list(row) for row in book[SHEET_NAME].iter_rows(min_row=3, max_row=4, values_only=True)] == rows
    finally:
        book.close()


def test_duplicate_or_open_workbook_is_not_changed(shipping_file):
    before = shipping_file.read_bytes()
    with pytest.raises(ValueError, match='发票号已登记'):
        append_shipping_register(shipping_file, 'INV-OLD', [[None, 'INV-OLD', 1]])
    owner = shipping_file.with_name(f'~${shipping_file.name}')
    owner.touch()
    with pytest.raises(PermissionError, match='正在被占用'):
        append_shipping_register(shipping_file, 'INV-NEW', [[None, 'INV-NEW', 1]])
    assert shipping_file.read_bytes() == before
    assert set(shipping_file.parent.iterdir()) == {shipping_file, owner}


def test_changed_workbook_is_not_overwritten(shipping_file):
    def save_other_workbook(message):
        if message == '出运明细：写回登记表':
            book = load_workbook(shipping_file)
            try:
                book[SHEET_NAME]['B3'] = 'OTHER-TICKET'
                book.save(shipping_file)
            finally:
                book.close()

    with pytest.raises(RuntimeError, match='登记期间已被修改'):
        append_shipping_register(shipping_file, 'INV-NEW', [[None, 'INV-NEW', 1]], progress=save_other_workbook)
    book = load_workbook(shipping_file, read_only=True)
    try:
        assert book[SHEET_NAME]['B3'].value == 'OTHER-TICKET'
    finally:
        book.close()
    assert list(shipping_file.parent.iterdir()) == [shipping_file]


def test_replace_error_leaves_original_and_cleans_up(shipping_file):
    before = shipping_file.read_bytes()
    with patch('app.services.shipping_register_service.os.replace', side_effect=PermissionError):
        with pytest.raises(RuntimeError, match='原表未更新'):
            append_shipping_register(shipping_file, 'INV-NEW', [[None, 'INV-NEW', 1]])
    assert shipping_file.read_bytes() == before
    assert list(shipping_file.parent.iterdir()) == [shipping_file]
