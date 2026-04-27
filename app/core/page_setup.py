from __future__ import annotations

from openpyxl import load_workbook

try:
    import win32com.client as win32
except Exception:
    win32 = None


def apply_persistent_page_setup(ws) -> None:
    ws.sheet_view.view = "pageBreakPreview"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1


def set_excel_page_view(file_path: str, enable_com: bool = True) -> None:
    workbook = load_workbook(file_path)
    for sheet in workbook.worksheets:
        apply_persistent_page_setup(sheet)
    workbook.save(file_path)
    workbook.close()

    if not enable_com or win32 is None:
        return

    app = None
    for prog_id in ("Ket.Application", "WPS.Application", "Excel.Application"):
        try:
            app = win32.Dispatch(prog_id)
            if app is not None:
                break
        except Exception:
            continue

    if app is None or not hasattr(app, "Workbooks"):
        return

    try:
        workbook = app.Workbooks.Open(file_path)
        for sheet in workbook.Worksheets:
            sheet.Activate()
            page_setup = sheet.PageSetup
            page_setup.Zoom = False
            page_setup.FitToPagesWide = 1
            page_setup.FitToPagesTall = 1
        workbook.Save()
        workbook.Close()
    except Exception:
        return
