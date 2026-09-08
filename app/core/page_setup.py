from __future__ import annotations

from openpyxl import load_workbook

from app.core.office_automation import start_isolated_office_application

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

    app = start_isolated_office_application(win32)

    if app is None or not hasattr(app, "Workbooks"):
        return

    office_workbook = None
    try:
        office_workbook = app.Workbooks.Open(file_path)
        for sheet in office_workbook.Worksheets:
            sheet.Activate()
            page_setup = sheet.PageSetup
            page_setup.Zoom = False
            page_setup.FitToPagesWide = 1
            page_setup.FitToPagesTall = 1
        office_workbook.Save()
    except Exception:
        pass
    finally:
        if office_workbook is not None:
            try:
                office_workbook.Close(False)
            except Exception:
                pass
        try:
            app.Quit()
        except Exception:
            pass
