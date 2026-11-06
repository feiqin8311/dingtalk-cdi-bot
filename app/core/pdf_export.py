from __future__ import annotations

import os

from app.core.office_automation import (
    configure_office_local_printer,
    start_isolated_office_application,
)

try:
    import win32com.client as win32
except Exception:
    win32 = None


def _open_office_workbook(workbook_path: str):
    if win32 is None:
        raise RuntimeError("无法导出 PDF：未安装可用的 Windows Office/WPS 自动化组件。")

    app = start_isolated_office_application(win32)

    if app is None or not hasattr(app, "Workbooks"):
        raise RuntimeError("无法启动 WPS 或 Excel，无法导出 PDF。")

    workbook = app.Workbooks.Open(os.path.abspath(workbook_path))
    configure_office_local_printer(app)
    worksheets = getattr(workbook, "Worksheets", None)
    if worksheets is not None:
        try:
            for sheet in worksheets:
                page_setup = getattr(sheet, "PageSetup", None)
                if page_setup is not None:
                    page_setup.Zoom = False
                    page_setup.PaperSize = 9
                    page_setup.Orientation = 1
                    page_setup.FitToPagesWide = 1
                    if str(getattr(sheet, "Name", "")).strip() in ("要素", "elements"):
                        page_setup.FitToPagesTall = False
                    else:
                        page_setup.FitToPagesTall = 1
        except (TypeError, Exception):
            pass
    return app, workbook



def _export_to_pdf(workbook_path: str, output_path: str, exporter, description: str) -> None:
    if win32 is None:
        raise RuntimeError("无法导出 PDF：未安装可用的 Windows Office/WPS 自动化组件。")

    app = None
    workbook = None
    try:
        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        app, workbook = _open_office_workbook(workbook_path)
        exporter(workbook, os.path.abspath(output_path))
    except RuntimeError:
        raise
    except Exception as exc:
        raise RuntimeError(f"导出{description}PDF失败：{exc}") from exc
    finally:
        if workbook is not None:
            try:
                workbook.Close(False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass


def export_worksheet_to_pdf(workbook_path: str, worksheet_name: str, output_path: str) -> None:
    """Export one worksheet from an Excel workbook as a standalone PDF."""
    _export_to_pdf(
        workbook_path,
        output_path,
        lambda workbook, pdf_path: workbook.Worksheets(worksheet_name).ExportAsFixedFormat(0, pdf_path),
        f"工作表“{worksheet_name}”",
    )


def export_workbook_to_pdf(workbook_path: str, output_path: str) -> None:
    """Export all worksheets from an Excel workbook as one PDF."""
    _export_to_pdf(
        workbook_path,
        output_path,
        lambda workbook, pdf_path: workbook.ExportAsFixedFormat(0, pdf_path),
        "报关资料工作簿",
    )
