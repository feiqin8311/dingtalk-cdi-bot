from __future__ import annotations

from openpyxl import load_workbook

from app.core.office_automation import (
    configure_office_local_printer,
    disable_office_print_communication,
    start_isolated_office_application,
)

try:
    import win32com.client as win32
except Exception:
    win32 = None


def apply_persistent_page_setup(ws, fit_to_height: int | None = None) -> None:
    ws.sheet_view.view = "pageBreakPreview"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.paperSize = getattr(ws, "PAPERSIZE_A4", 9)
    ws.page_setup.orientation = getattr(ws, "ORIENTATION_PORTRAIT", "portrait")
    ws.page_setup.fitToWidth = 1
    if fit_to_height is None:
        title = str(getattr(ws, "title", "")).strip()
        fit_to_height = 0 if title in ("要素", "elements") else 1
    ws.page_setup.fitToHeight = fit_to_height



def set_excel_page_view(file_path: str, enable_com: bool = False) -> None:
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
        disable_office_print_communication(app)
        office_workbook = app.Workbooks.Open(file_path)
        configure_office_local_printer(app)
        for sheet in office_workbook.Worksheets:
            sheet.Activate()
            page_setup = sheet.PageSetup
            page_setup.Zoom = False
            page_setup.PaperSize = 9
            page_setup.Orientation = 1
            page_setup.FitToPagesWide = 1
            if str(getattr(sheet, "Name", "")).strip() in ("要素", "elements"):
                page_setup.FitToPagesTall = False
            else:
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

