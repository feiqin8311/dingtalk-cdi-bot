from app.core.page_setup import apply_persistent_page_setup


def ensure_default_sheet_removed(workbook) -> None:
    if "Sheet" in workbook.sheetnames and len(workbook.sheetnames) > 1:
        del workbook["Sheet"]


def finalize_workbook_page_setup(workbook) -> None:
    for sheet in workbook.worksheets:
        apply_persistent_page_setup(sheet)
