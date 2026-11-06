import os
import tempfile
import unittest
from unittest.mock import Mock, patch

from app.core.pdf_export import export_worksheet_to_pdf


class PdfExportTests(unittest.TestCase):
    def test_exports_entire_workbook_to_pdf(self):
        app = Mock()
        workbook = Mock()
        app.Workbooks.Open.return_value = workbook

        with tempfile.TemporaryDirectory() as tmp_dir:
            workbook_path = os.path.join(tmp_dir, "报关资料.xlsx")
            output_path = os.path.join(tmp_dir, "INV-001 报关资料.pdf")

            with patch("app.core.pdf_export.win32") as win32:
                win32.DispatchEx.return_value = app
                from app.core.pdf_export import export_workbook_to_pdf

                export_workbook_to_pdf(workbook_path, output_path)

            win32.Dispatch.assert_not_called()
            win32.DispatchEx.assert_called_once_with("Ket.Application")
            self.assertIs(app.PrintCommunication, False)
            workbook.ExportAsFixedFormat.assert_called_once_with(0, os.path.abspath(output_path))
            workbook.Close.assert_called_once_with(False)
            app.Quit.assert_called_once_with()

    def test_exports_only_requested_worksheet_to_pdf(self):
        app = Mock()
        app.ActivePrinter = "AnyDesk Printer 在 Ne01:"
        workbook = Mock()
        worksheet = Mock()
        app.Workbooks.Open.return_value = workbook
        workbook.Worksheets.return_value = worksheet

        with tempfile.TemporaryDirectory() as tmp_dir:
            workbook_path = os.path.join(tmp_dir, "报关资料.xlsx")
            output_path = os.path.join(tmp_dir, "INV-001 销售合同.pdf")

            with patch("app.core.pdf_export.win32") as win32, patch(
                "app.core.office_automation._iter_local_virtual_printer_candidates",
                return_value=["Microsoft Print to PDF 在 Ne00:"],
            ):
                win32.DispatchEx.return_value = app
                export_worksheet_to_pdf(workbook_path, "合同", output_path)

            win32.Dispatch.assert_not_called()
            win32.DispatchEx.assert_called_once_with("Ket.Application")
            app.Workbooks.Open.assert_called_once_with(os.path.abspath(workbook_path))
            self.assertEqual(app.ActivePrinter, "Microsoft Print to PDF 在 Ne00:")
            self.assertIs(app.PrintCommunication, False)
            workbook.Worksheets.assert_called_once_with("合同")
            worksheet.ExportAsFixedFormat.assert_called_once_with(0, os.path.abspath(output_path))
            workbook.Close.assert_called_once_with(False)
            app.Quit.assert_called_once_with()

    def test_raises_clear_error_when_no_office_application_is_available(self):
        with patch("app.core.pdf_export.win32") as win32:
            win32.DispatchEx.side_effect = Exception("not installed")

            with self.assertRaisesRegex(RuntimeError, "无法启动 WPS 或 Excel"):
                export_worksheet_to_pdf("report.xlsx", "合同", "contract.pdf")

