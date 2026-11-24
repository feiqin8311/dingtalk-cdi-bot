from dataclasses import dataclass
from hashlib import sha256
from io import BytesIO
from numbers import Real
from pathlib import Path
from tempfile import NamedTemporaryFile, TemporaryDirectory
import math
import os

from openpyxl import load_workbook
import pandas as pd
from pywintypes import error as WindowsError
import pythoncom
import win32com.client as win32
import win32con
import win32file

from app.core.excel_io import workbook_write_lock
from app.core.office_automation import start_isolated_office_application
from app.services.lumiya_source_service import source_display


MAINTENANCE_DIR = Path(r"\\192.168.0.45\供应链管理\2 物流发货管理\17.单证数据表维护")
MAINTENANCE_FILENAME = "海联琨之杰  励沣采购报关数量（后续改为易达直采）.xlsx"
SHEET_NAME = "易达直采"
HANGZHOU_SOURCE = "33019 杭州其他"


def ezarc_maintenance_path(test_mode=False):
    directory = MAINTENANCE_DIR / "测试" if test_mode else MAINTENANCE_DIR
    return directory / MAINTENANCE_FILENAME


def _text(value):
    return "" if value is None else str(value).strip()


def sku_key(value):
    if isinstance(value, Real) and math.isfinite(value) and value == int(value):
        return str(int(value))
    key = _text(value)
    integer, dot, fraction = key.partition('.')
    return integer if dot and integer.isdecimal() and fraction and set(fraction) == {'0'} else key


def _quantity(value, label):
    if isinstance(value, bool) or not isinstance(value, Real) or not math.isfinite(value) or value != int(value) or value < 0:
        raise ValueError(f"{label}须为非负整数，实际值：{value!r}")
    return int(value)


@dataclass(frozen=True)
class SourceAllocation:
    hangzhou: int
    external: int
    source: str
    supplier: str


@dataclass(frozen=True)
class MaintenanceUpdate:
    path: Path
    fingerprint: bytes
    invoice_number: str
    insert_column: int
    balance_column: int
    # worksheet row, expected prior balance, deduction
    deductions: tuple[tuple[int, int, int], ...]


def plan_ezarc_sources(path, invoice_number, quantities):
    """Read the maintenance sheet once; existing shipment columns replay their saved split."""
    path = Path(path)
    invoice_number = _text(invoice_number)
    if not invoice_number or invoice_number.startswith(("=", "+", "-", "@")):
        raise ValueError("物流商单号不适合作为海联维护表列标题")
    data = path.read_bytes()
    workbook = load_workbook(BytesIO(data), read_only=True, data_only=True, keep_links=False)
    try:
        sheet = workbook[SHEET_NAME]
        rows = sheet.values
        headers = [_text(value) for value in next(rows)]
        required = ("SKU", "供应商名称", "货源地", "结余需报关数量")
        if any(headers.count(name) != 1 for name in required):
            raise ValueError(f"海联维护表表头需各包含一列：{'、'.join(required)}")
        sku_col, supplier_col, source_col, balance_col = (headers.index(name) for name in required)
        last_header = max(i for i, name in enumerate(headers) if name)
        ticket_columns = [i for i, name in enumerate(headers) if i > balance_col and name == invoice_number]
        if len(ticket_columns) > 1:
            raise ValueError(f"海联维护表有多个 {invoice_number} 列，请先合并为一列")
        ticket_col = ticket_columns[0] if ticket_columns else None
        allocations = {}
        deductions = []
        for row_number, row in enumerate(rows, 2):
            sku = sku_key(row[sku_col])
            saved = row[ticket_col] if ticket_col is not None else None
            if sku not in quantities:
                if saved not in (None, "", 0):
                    raise ValueError(f"海联维护表 {invoice_number} 列有本票不包含的 SKU：{sku}，请核对")
                continue
            if sku in allocations:
                raise ValueError(f"海联维护表 SKU {sku} 重复，无法确定扣减行")
            quantity = _quantity(quantities[sku], f"SKU {sku} 发货数量")
            if ticket_col is not None:
                hangzhou = _quantity(0 if saved in (None, "") else saved, f"SKU {sku} 本票已扣数量")
                if hangzhou > quantity:
                    raise ValueError(f"SKU {sku} 本票已扣 {hangzhou}，超过本次发货 {quantity}，请核对")
            else:
                balance = _quantity(row[balance_col], f"SKU {sku} 结余需报关数量")
                hangzhou = min(quantity, balance)
                if hangzhou:
                    deductions.append((row_number, balance, hangzhou))
            external = quantity - hangzhou
            source, supplier = _text(row[source_col]), _text(row[supplier_col])
            if external and (not source_display(source) or source == HANGZHOU_SOURCE or not supplier):
                raise ValueError(f"SKU {sku} 需要报实际货源地，请核对维护表的货源地和供应商名称")
            allocations[sku] = SourceAllocation(hangzhou, external, source, supplier)
        # Do not overwrite untitled input cells further to the right of the shipment columns.
        insert_column = (sheet.max_column or last_header + 1) + 1
    finally:
        workbook.close()
    update = MaintenanceUpdate(path, sha256(data).digest(), invoice_number, insert_column,
                               balance_col + 1, tuple(deductions))
    return allocations, update


ALLOCATION_FIELDS = ('总重量（kg）-箱子', '总体积（m³）-箱子', '箱子总毛重（kg）', '运费分摊', '总PCS', '申报总价')


def split_ezarc_rows(rows, allocations):
    result = []
    for row in rows.to_dict('records'):
        allocation = allocations.get(row['SKU'])
        row['外部开票'] = False
        row['报关箱数'] = row['箱数']
        if allocation is None or not allocation.external:
            result.append(row)
            continue
        actual = dict(row, 发货数量=allocation.external, 境内货源地=allocation.source,
                      境内货源地显示=source_display(allocation.source), 供应商名称=allocation.supplier, 外部开票=True)
        if allocation.hangzhou:
            ratio = allocation.hangzhou / row['发货数量']
            internal = dict(row, 发货数量=allocation.hangzhou)
            for field in (*ALLOCATION_FIELDS, '箱数'):
                internal[field] = row[field] * ratio
                actual[field] = row[field] - internal[field]
            # Preserve whole cartons on the paperwork without counting a shared carton twice.
            internal['报关箱数'] = math.floor(internal['箱数'] + 0.5)
            actual['报关箱数'] = row['箱数'] - internal['报关箱数']
            result.append(internal)
        result.append(actual)
    return pd.DataFrame(result)


def _checked_workbook_bytes(update):
    # Check both the workbook handle and the Office owner file; some Office versions release
    # their workbook handle between saves but keep the owner file while the window is open.
    path = update.path
    owner_file = path.with_name(f'~${path.name}')
    if owner_file.exists():
        raise RuntimeError('海联维护表仍有 Office 占用标记，请关闭 WPS/Excel 中的此表后重试')
    try:
        handle = win32file.CreateFile(str(path), win32con.GENERIC_READ | win32con.GENERIC_WRITE, 0, None,
                                     win32con.OPEN_EXISTING, 0, None)
        try:
            _, data = win32file.ReadFile(handle, win32file.GetFileSize(handle))
        finally:
            handle.Close()
    except WindowsError as exc:
        raise RuntimeError('海联维护表被占用或不可写，请关闭 WPS/Excel 中的此表并检查权限后重试') from exc
    if sha256(data).digest() != update.fingerprint:
        raise RuntimeError('海联维护表在生成资料期间已被修改，请重新生成以核对最新结余')
    return data


def write_ezarc_deductions(update):
    if not update.deductions:
        return
    with workbook_write_lock(update.path, '海联维护表'), TemporaryDirectory(prefix='ezarc-maintenance-') as tmp:
        local_path = Path(tmp) / update.path.name
        local_path.write_bytes(_checked_workbook_bytes(update))
        _calculate_local_deductions(local_path, update)
        _verify_saved_deductions(local_path, update)
        # Stage the complete workbook on the same filesystem, then replace it in one operation.
        # If the share refuses replacements or the target is opened meanwhile, never overwrite in place.
        staged_path = None
        try:
            with NamedTemporaryFile(dir=update.path.parent, prefix='.ezarc-', suffix='.tmp', delete=False) as staged:
                staged_path = Path(staged.name)
                staged.write(local_path.read_bytes())
                staged.flush()
                os.fsync(staged.fileno())
            _checked_workbook_bytes(update)
            try:
                os.replace(staged_path, update.path)
            except OSError as exc:
                raise RuntimeError('海联维护表无法替换，可能正在使用或缺少共享目录权限；原表未更新') from exc
        finally:
            if staged_path is not None:
                staged_path.unlink(missing_ok=True)


def _verify_saved_deductions(path, update):
    workbook = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        sheet = workbook[SHEET_NAME]
        if sheet.cell(1, update.insert_column).value != update.invoice_number:
            raise ValueError('本地海联维护表保存后未找到本票列，原表未更新')
        for row, prior, deduction in update.deductions:
            if (sheet.cell(row, update.insert_column).value != deduction
                    or sheet.cell(row, update.balance_column).value != prior - deduction):
                raise ValueError(f'本地海联维护表保存后第 {row} 行扣减校验失败，原表未更新')
    finally:
        workbook.close()


def _calculate_local_deductions(path, update):
    # Office recalculates formulas on a local copy only; it never opens the shared source for writing.
    pythoncom.CoInitialize()
    app = book = None
    try:
        app = start_isolated_office_application(win32)
        if app is None:
            raise RuntimeError("无法启动 Excel/WPS 更新海联维护表")
        app.EnableEvents = False
        try:
            app.AutomationSecurity = 3
        except Exception:
            pass
        book = app.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False,
                                  IgnoreReadOnlyRecommended=True, AddToMru=False)
        if book.ReadOnly:
            raise RuntimeError("海联维护表本地副本不可写，原表未更新")
        sheet = book.Worksheets(SHEET_NAME)
        header = sheet.Cells(1, update.insert_column)
        header.NumberFormat = "@"
        header.Value2 = update.invoice_number
        for row, _, deduction in update.deductions:
            sheet.Cells(row, update.insert_column).Value2 = deduction
        sheet.Calculate()
        for row, prior, deduction in update.deductions:
            cell = sheet.Cells(row, update.balance_column)
            if not cell.HasFormula or cell.Value2 != prior - deduction:
                raise ValueError(f"海联维护表第 {row} 行结余公式未正确包含本票扣减，未保存")
        book.Save()
    finally:
        try:
            if book is not None:
                book.Close(False)
        finally:
            try:
                if app is not None:
                    app.Quit()
            finally:
                pythoncom.CoUninitialize()
