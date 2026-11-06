import os
import sys
import tempfile
import types
import unittest
from unittest.mock import Mock, patch

from openpyxl import Workbook, load_workbook


win32com_module = types.ModuleType("win32com")
win32com_client = types.ModuleType("win32com.client")
win32com_client.Dispatch = lambda *args, **kwargs: None
win32com_module.client = win32com_client
sys.modules.setdefault("win32com", win32com_module)
sys.modules.setdefault("win32com.client", win32com_client)

from app.core.page_setup import apply_persistent_page_setup, set_excel_page_view


class CorePageSetupTests(unittest.TestCase):
    def test_apply_persistent_page_setup_survives_save(self):
        wb = Workbook()
        ws = wb.active
        apply_persistent_page_setup(ws)
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            path = tmp.name
        try:
            wb.save(path)
            reloaded = load_workbook(path)
            ws2 = reloaded.active
            self.assertEqual(ws2.sheet_view.view, "pageBreakPreview")
            self.assertTrue(ws2.sheet_properties.pageSetUpPr.fitToPage)
            self.assertEqual(ws2.page_setup.fitToWidth, 1)
            self.assertEqual(ws2.page_setup.fitToHeight, 1)
        finally:
            if "reloaded" in locals():
                reloaded.close()
            os.remove(path)

    def test_set_excel_page_view_is_safe_without_com(self):
        wb = Workbook()
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            path = tmp.name
        try:
            wb.save(path)
            wb.close()
            set_excel_page_view(path)
            reloaded = load_workbook(path)
            ws2 = reloaded.active
            self.assertEqual(ws2.sheet_view.view, "pageBreakPreview")
            self.assertTrue(ws2.sheet_properties.pageSetUpPr.fitToPage)
            self.assertEqual(ws2.page_setup.fitToWidth, 1)
            self.assertEqual(ws2.page_setup.fitToHeight, 1)
        finally:
            if "reloaded" in locals():
                reloaded.close()
            os.remove(path)

    def test_set_excel_page_view_can_skip_com_stage(self):
        wb = Workbook()
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            path = tmp.name
        try:
            wb.save(path)
            wb.close()
            with patch("app.core.page_setup.win32") as win32:
                set_excel_page_view(path)
                win32.DispatchEx.assert_not_called()
            reloaded = load_workbook(path)
            ws2 = reloaded.active
            self.assertEqual(ws2.sheet_view.view, "pageBreakPreview")
            self.assertTrue(ws2.sheet_properties.pageSetUpPr.fitToPage)
            self.assertEqual(ws2.page_setup.fitToWidth, 1)
            self.assertEqual(ws2.page_setup.fitToHeight, 1)
        finally:
            if "reloaded" in locals():
                reloaded.close()
            os.remove(path)

    def test_set_excel_page_view_uses_an_isolated_office_process(self):
        wb = Workbook()
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            path = tmp.name
        app = Mock()
        office_workbook = Mock()
        office_workbook.Worksheets = []
        app.Workbooks.Open.return_value = office_workbook

        try:
            wb.save(path)
            wb.close()

            with patch("app.core.page_setup.win32") as win32:
                win32.DispatchEx.return_value = app
                set_excel_page_view(path, enable_com=True)

            win32.Dispatch.assert_not_called()
            win32.DispatchEx.assert_called_once_with("Ket.Application")
            self.assertIs(app.PrintCommunication, False)
            office_workbook.Save.assert_called_once_with()
            office_workbook.Close.assert_called_once_with(False)
            app.Quit.assert_called_once_with()
        finally:
            os.remove(path)

    def test_apply_persistent_page_setup_elements_sheet_unconstrained_height(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "要素"
        apply_persistent_page_setup(ws)
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            path = tmp.name
        try:
            wb.save(path)
            reloaded = load_workbook(path)
            ws2 = reloaded["要素"]
            self.assertEqual(ws2.sheet_view.view, "pageBreakPreview")
            self.assertTrue(ws2.sheet_properties.pageSetUpPr.fitToPage)
            self.assertEqual(ws2.page_setup.fitToWidth, 1)
            self.assertEqual(ws2.page_setup.fitToHeight, 0)
        finally:
            if "reloaded" in locals():
                reloaded.close()
            os.remove(path)

    def test_set_excel_page_view_elements_uses_unconstrained_height(self):
        wb = Workbook()
        ws_elements = wb.active
        ws_elements.title = "要素"
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            path = tmp.name
        try:
            wb.save(path)
            wb.close()
            set_excel_page_view(path)
            reloaded = load_workbook(path)
            ws2 = reloaded["要素"]
            self.assertEqual(ws2.sheet_view.view, "pageBreakPreview")
            self.assertTrue(ws2.sheet_properties.pageSetUpPr.fitToPage)
            self.assertEqual(ws2.page_setup.fitToWidth, 1)
            self.assertEqual(ws2.page_setup.fitToHeight, 0)
        finally:
            if "reloaded" in locals():
                reloaded.close()
            os.remove(path)


if __name__ == "__main__":
    unittest.main()

