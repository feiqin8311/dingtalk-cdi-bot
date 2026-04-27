import os
import sys
import tempfile
import types
import unittest

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
            set_excel_page_view(path, enable_com=False)
            reloaded = load_workbook(path)
            ws2 = reloaded.active
            self.assertEqual(ws2.sheet_view.view, "pageBreakPreview")
            self.assertEqual(ws2.page_setup.fitToWidth, 1)
            self.assertEqual(ws2.page_setup.fitToHeight, 1)
        finally:
            if "reloaded" in locals():
                reloaded.close()
            os.remove(path)


if __name__ == "__main__":
    unittest.main()
