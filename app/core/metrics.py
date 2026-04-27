import datetime
import os

import openpyxl
from openpyxl import load_workbook


def append_call_record(metrics_path: str, sheet_name: str, headers: list[str], module_name: str,
                       action: str, success: bool, duration_seconds: float, invoice_number: str = "",
                       source_file: str = "", error_message: str = "") -> None:
    try:
        metrics_dir = os.path.dirname(metrics_path)
        if metrics_dir and not os.path.exists(metrics_dir):
            os.makedirs(metrics_dir, exist_ok=True)

        if os.path.exists(metrics_path):
            wb = load_workbook(metrics_path)
            ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.create_sheet(sheet_name)
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = sheet_name

        if ws.max_row == 1 and ws.cell(row=1, column=1).value is None:
            for idx, header in enumerate(headers, start=1):
                ws.cell(row=1, column=idx, value=header)
        elif ws.max_row == 1:
            current_headers = [ws.cell(row=1, column=i).value for i in range(1, len(headers) + 1)]
            if current_headers != headers:
                ws.insert_rows(1)
                for idx, header in enumerate(headers, start=1):
                    ws.cell(row=1, column=idx, value=header)

        ws.append([
            datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            module_name,
            action,
            "成功" if success else "失败",
            round(duration_seconds, 3),
            invoice_number or "",
            source_file or "",
            error_message or "",
        ])
        wb.save(metrics_path)
    except Exception as exc:
        print(f"写入调用统计失败: {exc}, 路径: {metrics_path}")
