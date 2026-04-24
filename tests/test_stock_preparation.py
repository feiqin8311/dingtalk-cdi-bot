import os
import sys
import tempfile
import types
import unittest
from unittest.mock import patch

import pandas as pd
from openpyxl import Workbook, load_workbook


sys.modules.setdefault("num2words", types.SimpleNamespace(num2words=lambda *args, **kwargs: "ONE"))

win32com_module = types.ModuleType("win32com")
win32com_client = types.ModuleType("win32com.client")
win32com_client.Dispatch = lambda *args, **kwargs: None
win32com_module.client = win32com_client
sys.modules.setdefault("win32com", win32com_module)
sys.modules.setdefault("win32com.client", win32com_client)

from Stock_Preparation import CustomsDeclaration, _ELEMENTS_GROUP_CACHE


class FindDataGroupOrderTests(unittest.TestCase):
    def setUp(self):
        _ELEMENTS_GROUP_CACHE.clear()

    def test_preserves_result_df_row_order_across_shops(self):
        result_df = pd.DataFrame(
            [
                {
                    "SKU": "SKU-B-1",
                    "报关中文品名": "品名B1",
                    "HS CODE": "HS-B1",
                    "备货数量": 1,
                    "单位": "个",
                    "是否含金刚石": False,
                    "店铺": "SHOP-B",
                },
                {
                    "SKU": "SKU-A-1",
                    "报关中文品名": "品名A1",
                    "HS CODE": "HS-A1",
                    "备货数量": 1,
                    "单位": "个",
                    "是否含金刚石": False,
                    "店铺": "SHOP-A",
                },
                {
                    "SKU": "SKU-B-2",
                    "报关中文品名": "品名B2",
                    "HS CODE": "HS-B2",
                    "备货数量": 1,
                    "单位": "个",
                    "是否含金刚石": False,
                    "店铺": "SHOP-B",
                },
            ]
        )

        elements_by_sheet = {
            "SHOP-A": pd.DataFrame([["品名：品名A1"], ["HS-A1"], [""]]),
            "SHOP-B": pd.DataFrame([["品名：品名B1"], ["HS-B1"], [""], ["品名：品名B2"], ["HS-B2"], [""]]),
        }

        def fake_read_excel(_path, sheet_name=None, header=None):
            return elements_by_sheet[sheet_name].copy()

        with patch("Stock_Preparation.pd.read_excel", side_effect=fake_read_excel):
            groups = CustomsDeclaration.find_data_group("dummy.xlsx", result_df)

        self.assertEqual([group[-1] for group in groups], ["SKU-B-1", "SKU-A-1", "SKU-B-2"])


class PageSetupPersistenceTests(unittest.TestCase):
    def test_apply_persistent_page_setup_survives_save_and_reload(self):
        workbook = Workbook()
        worksheet = workbook.active

        CustomsDeclaration.apply_persistent_page_setup(worksheet)

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            temp_path = tmp.name

        try:
            workbook.save(temp_path)
            reloaded = load_workbook(temp_path)
            reloaded_ws = reloaded.active

            self.assertEqual(reloaded_ws.sheet_view.view, "pageBreakPreview")
            self.assertEqual(reloaded_ws.page_setup.fitToWidth, 1)
            self.assertEqual(reloaded_ws.page_setup.fitToHeight, 1)
        finally:
            if "reloaded" in locals():
                reloaded.close()
            os.remove(temp_path)

    def test_instance_can_call_set_excel_page_view(self):
        workbook = Workbook()
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            temp_path = tmp.name

        try:
            workbook.save(temp_path)
            workbook.close()

            generator = CustomsDeclaration()
            generator.set_excel_page_view(temp_path)

            reloaded = load_workbook(temp_path)
            reloaded_ws = reloaded.active

            self.assertEqual(reloaded_ws.sheet_view.view, "pageBreakPreview")
            self.assertEqual(reloaded_ws.page_setup.fitToWidth, 1)
            self.assertEqual(reloaded_ws.page_setup.fitToHeight, 1)
        finally:
            if "reloaded" in locals():
                reloaded.close()
            os.remove(temp_path)


if __name__ == "__main__":
    unittest.main()
