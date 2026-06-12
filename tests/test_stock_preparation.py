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

from stock_preparation import CustomsDeclaration, _ELEMENTS_GROUP_CACHE


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

        with patch("stock_preparation.pd.read_excel", side_effect=fake_read_excel):
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


class RendererCommonTests(unittest.TestCase):
    def test_ensure_default_sheet_removed(self):
        from app.renderers.common import ensure_default_sheet_removed

        wb = Workbook()
        wb.create_sheet("明细单")
        ensure_default_sheet_removed(wb)
        self.assertNotIn("Sheet", wb.sheetnames)

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


class StockFlowBehaviorTests(unittest.TestCase):
    def test_stock_flow_generates_documents_without_reentering_generator_method(self):
        from app.flows.stock_flow import StockFlow

        class DummyWorkbook:
            def save(self, _path):
                return None

        class DummyGenerator:
            def __init__(self):
                self.shipping_details = {"invoice_number": "INV-002"}
                self.page_view_paths = []
                self.log_messages = []
                self.call_records = []

            def _prepare_declaration_data(self, *_args):
                self.prepare_called = True

            def copy_declaration_template_file(self, output_dir):
                self.template_output_dir = output_dir
                return os.path.join(output_dir, "decl.xlsx")

            def fill_declaration_template_file(self, path):
                self.template_output_path = path

            def set_excel_page_view(self, path):
                self.page_view_paths.append(path)

            def log_message(self, message):
                self.log_messages.append(message)

            def append_call_record(self, **kwargs):
                self.call_records.append(kwargs)

        generator = DummyGenerator()
        flow = StockFlow(generator)

        with patch("app.flows.stock_flow.openpyxl.Workbook", side_effect=[DummyWorkbook(), DummyWorkbook()]), \
                patch("app.flows.stock_flow.render_detailed_list") as render_detailed_list, \
                patch("app.flows.stock_flow.render_packing_list") as render_packing_list, \
                patch("app.flows.stock_flow.render_bill") as render_bill, \
                patch("app.flows.stock_flow.render_contract") as render_contract, \
                patch("app.flows.stock_flow.render_customs_declaration") as render_customs_declaration, \
                patch("app.flows.stock_flow.render_elements_sheet") as render_elements_sheet, \
                patch("app.flows.stock_flow.ensure_default_sheet_removed"), \
                patch("app.flows.stock_flow.finalize_workbook_page_setup"):
            flow.generate_all_documents(
                "invoice.xlsx",
                "product.xlsx",
                "product_name.xlsx",
                "site.xlsx",
                "elements.xlsx",
                "/tmp/out",
                "/tmp/pictures",
                False,
                True,
            )

        self.assertTrue(generator.prepare_called)
        self.assertEqual(generator.template_output_dir, "/tmp/out")
        self.assertEqual(generator.template_output_path, "/tmp/out/decl.xlsx")
        self.assertEqual(
            generator.page_view_paths,
            [
                "/tmp/out/INV-002 报关资料.xlsx",
                "/tmp/out/INV-002 发票合同.xlsx",
            ],
        )
        self.assertEqual(len(generator.call_records), 1)
        self.assertTrue(generator.call_records[0]["success"])
        self.assertEqual(render_detailed_list.call_count, 1)
        self.assertEqual(render_packing_list.call_count, 1)
        self.assertEqual(render_bill.call_count, 2)
        self.assertEqual(render_contract.call_count, 2)
        self.assertEqual(render_customs_declaration.call_count, 1)
        self.assertEqual(render_elements_sheet.call_count, 1)


class StockRegisterTargetSelectionTests(unittest.TestCase):
    def test_register_combined_invoice_marks_non_virtual_warehouse_as_warehouse_shipping(self):
        generator = CustomsDeclaration()

        stock_df = pd.DataFrame(
            [
                {
                    "备货单号": "STOCK-001",
                    "SKU": "SKU-1",
                    "备货数量": 1,
                    "箱数": 1,
                    "总重量（kg）-箱子": 10,
                    "总体积（m³）-箱子": 0.1,
                    "国家": "美国",
                    "单箱数量": 1,
                    "箱子毛重（kg）": 10,
                    "箱子长度（cm）": 10,
                    "箱子宽度（cm）": 10,
                    "箱子高度（cm）": 10,
                    "第三方仓产品名": "Product 1",
                },
                {
                    "备货单号": "STOCK-002",
                    "SKU": "SKU-2",
                    "备货数量": 2,
                    "箱数": 1,
                    "总重量（kg）-箱子": 20,
                    "总体积（m³）-箱子": 0.2,
                    "国家": "美国",
                    "单箱数量": 2,
                    "箱子毛重（kg）": 20,
                    "箱子长度（cm）": 20,
                    "箱子宽度（cm）": 20,
                    "箱子高度（cm）": 20,
                    "第三方仓产品名": "Product 2",
                },
            ]
        )
        stock_details_df = pd.DataFrame(
            [
                {
                    "SKU": "SKU-1",
                    "备货单号": "STOCK-001",
                    "收货仓库": "EZARC NA-US美国仓(AWD)",
                    "物流中心编码": "IUSP",
                    "发货仓库（单据）": "杭州虚拟仓",
                    "物流商": "物流商A",
                },
                {
                    "SKU": "SKU-2",
                    "备货单号": "STOCK-002",
                    "收货仓库": "FBT仓-TIKTOK",
                    "物流中心编码": "FBT",
                    "发货仓库（单据）": "义乌仓",
                    "物流商": "物流商A",
                },
            ]
        )
        brand_df = pd.DataFrame(
            [
                {"*SKU": "SKU-1", "品牌": "EZARC"},
                {"*SKU": "SKU-2", "品牌": "EZARC"},
            ]
        )

        with patch.object(generator, "_read_excel_cached", side_effect=[stock_df, stock_details_df, brand_df]):
            result = generator.register_combined_invoice("stock.xlsx", "product.xlsx")

        self.assertEqual(result["发货仓库（单据）"].tolist(), ["工厂直发", "仓库发货"])

    def test_build_invoicing_register_uses_company_title_instead_of_first_brand(self):
        generator = CustomsDeclaration()
        generator.company_info_index = 0
        generator.lcl_flag = False
        generator.shipping_details = {
            "currency": "USD",
            "invoice_number": "INV-001",
            "is_awd": False,
        }
        generator.result_df = pd.DataFrame(
            [
                {
                    "SKU": "SKU-1",
                    "报关中文品名": "品名1",
                    "英文品名": "Name1",
                    "HS CODE": "HS1",
                    "备货数量": 1,
                    "单位": "PCS",
                    "申报单价": 10,
                    "总价": 10,
                }
            ]
        )

        register_invoice_df = pd.DataFrame(
            [
                {
                    "SKU": "SKU-1",
                    "店铺": "TOLESA",
                    "备货数量": 1,
                    "收货仓库": "仓盛美中仓 Bolingbrook",
                }
            ]
        )
        product_name_df = pd.DataFrame(
            [{"*SKU": "SKU-1", "报关中文品名": "品名1"}]
        )

        with patch.object(generator, "register_combined_invoice", return_value=register_invoice_df), \
                patch.object(generator, "_read_excel_cached", return_value=product_name_df):
            _register_df, register_filepath, _sheet_name, _brand = generator.build_invoicing_register_df(
                "invoice.xlsx",
                "product.xlsx",
                "lumiya.xlsx",
                "easycan.xlsx",
                "sutong.xlsx",
                "liulang.xlsx",
                "jiangsheng.xlsx",
            )

        self.assertEqual(register_filepath, "easycan.xlsx")


if __name__ == "__main__":
    unittest.main()
