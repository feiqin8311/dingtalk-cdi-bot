from io import BytesIO
import os
from pathlib import Path
from tempfile import NamedTemporaryFile

from openpyxl import load_workbook
from openpyxl.styles import Border, Side

from app.core.excel_io import workbook_write_lock


SHEET_NAME = '出运明细'
_INVOICE_COLUMN = 2
_THIN = Side(style='thin')
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def append_shipping_register(path, invoice_number, rows, *, progress=None):
    """Append shipment rows against a snapshot, then publish the verified workbook."""
    path = Path(path)
    invoice = str(invoice_number or '').strip()
    if not rows:
        raise ValueError('出运明细没有待登记的行，原表未更新')
    expected_invoices = [row[_INVOICE_COLUMN - 1] for row in rows]
    if not invoice or not any(invoice in str(value) for value in expected_invoices):
        raise ValueError('出运明细的发票号为空或与待登记数据不一致，未写入')

    owner = path.with_name(f'~${path.name}')

    def check_owner():
        if owner.exists():
            raise PermissionError(f'出运明细正在被占用，原表未更新：{path}')

    def report(message):
        if progress is not None:
            progress(message)

    report('出运明细：读取登记表')
    check_owner()
    snapshot = path.read_bytes()
    report('出运明细：查重并准备写入')
    with BytesIO(snapshot) as source:
        cached = load_workbook(source, read_only=True, data_only=True)
        try:
            if SHEET_NAME not in cached.sheetnames:
                raise ValueError(f'缺少“{SHEET_NAME}”工作表：{path}')
            existing_rows = cached[SHEET_NAME].values
            header = next(existing_rows, ())
            if len(header) < _INVOICE_COLUMN or header[_INVOICE_COLUMN - 1] != '发票号':
                raise ValueError(f'出运明细的 B 列不是“发票号”：{path}')
            last_row = 1
            for row_index, row in enumerate(existing_rows, 2):
                if invoice in str(row[_INVOICE_COLUMN - 1] or ''):
                    raise ValueError(f'发票号已登记（{invoice}）')
                if any(value not in (None, '') for value in row):
                    last_row = row_index
        finally:
            cached.close()

    start_row = last_row + 1
    end_row = last_row + len(rows)
    with BytesIO(snapshot) as source, BytesIO() as updated:
        book = load_workbook(source)
        try:
            sheet = book[SHEET_NAME]
            for row_index, row in enumerate(rows, start_row):
                for column, value in enumerate(row, 1):
                    cell = sheet.cell(row_index, column)
                    if cell.value not in (None, ''):
                        raise ValueError(f'出运明细 {cell.coordinate} 已有内容，原表未更新')
                    cell.value = value
                    cell.border = _BORDER
            book.save(updated)
        finally:
            book.close()

        report('出运明细：校验新增行')
        verified = load_workbook(updated, read_only=True, data_only=True)
        try:
            written = list(verified[SHEET_NAME].iter_rows(
                min_row=start_row, max_row=end_row, min_col=_INVOICE_COLUMN, max_col=_INVOICE_COLUMN,
                values_only=True,
            ))
            if [value for (value,) in written] != expected_invoices:
                raise ValueError('出运明细新增行校验失败，原表未更新')
        finally:
            verified.close()

        report('出运明细：写回登记表')
        staged_path = None
        try:
            with NamedTemporaryFile(dir=path.parent, prefix='.shipping-register-', suffix='.tmp', delete=False) as staged:
                staged_path = Path(staged.name)
                with updated.getbuffer() as data:
                    staged.write(data)
                staged.flush()
                os.fsync(staged.fileno())
            with workbook_write_lock(path, '出运明细'):
                check_owner()
                if path.read_bytes() != snapshot:
                    raise RuntimeError('出运明细在登记期间已被修改，原表未更新，请重试')
                try:
                    os.replace(staged_path, path)
                except OSError as exc:
                    raise RuntimeError(f'无法替换出运明细，原表未更新：{path}') from exc
        finally:
            if staged_path is not None:
                staged_path.unlink(missing_ok=True)

    return last_row, end_row
