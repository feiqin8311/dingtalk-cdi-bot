def check_invoice_number_in_sheet(ws, invoice_number: str, column_name: str) -> int:
    col_idx = None
    for cell in ws[1]:
        if str(cell.value).strip() == column_name:
            col_idx = cell.column
            break
    if col_idx is None:
        raise ValueError(f"登记表未找到列：{column_name}")

    invoice_number = str(invoice_number).strip()
    if not invoice_number:
        return 0

    for row_idx in range(2, ws.max_row + 1):
        value = ws.cell(row=row_idx, column=col_idx).value
        value_str = str(value) if value is not None else ""
        if value_str == invoice_number or invoice_number in value_str:
            raise Exception("发票号已登记")

    return 0


def find_last_data_row_ws(ws) -> int:
    max_row = ws.max_row
    max_col = ws.max_column
    for row_idx in range(max_row, 0, -1):
        row = ws[row_idx]
        for cell in row[:max_col]:
            value = cell.value
            if value is not None and str(value).strip() != "":
                return row_idx
    return 0
