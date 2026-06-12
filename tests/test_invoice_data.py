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

from invoice_data import CustomsDeclaration


class InvoicePageSetupTests(unittest.TestCase):
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


class InvoiceFlowImportTests(unittest.TestCase):
    def test_invoice_flow_is_importable(self):
        from app.flows.invoice_flow import InvoiceFlow

        self.assertIsNotNone(InvoiceFlow)

    def test_invoice_flow_generates_documents_without_reentering_generator_method(self):
        from app.flows.invoice_flow import InvoiceFlow

        class DummyWorkbook:
            def save(self, _path):
                return None

        class DummyGenerator:
            def __init__(self):
                self.lcl_flag = None
                self.shipping_details = {"invoice_number": "INV-001"}
                self.process_invoice_calls = 0
                self.page_view_paths = []
                self.log_messages = []
                self.call_records = []

            def process_invoice(self, *_args):
                self.process_invoice_calls += 1

            def lcl_situation(self, *_args):
                raise AssertionError("non-lcl path should not call lcl_situation")

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
        flow = InvoiceFlow(generator)

        with patch("app.flows.invoice_flow.openpyxl.Workbook", return_value=DummyWorkbook()), \
                patch("app.flows.invoice_flow.render_detailed_list") as render_detailed_list, \
                patch("app.flows.invoice_flow.render_packing_list") as render_packing_list, \
                patch("app.flows.invoice_flow.render_bill") as render_bill, \
                patch("app.flows.invoice_flow.render_contract") as render_contract, \
                patch("app.flows.invoice_flow.render_customs_declaration") as render_customs_declaration, \
                patch("app.flows.invoice_flow.render_elements_sheet") as render_elements_sheet, \
                patch("app.flows.invoice_flow.ensure_default_sheet_removed"), \
                patch("app.flows.invoice_flow.finalize_workbook_page_setup"):
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

        self.assertFalse(generator.lcl_flag)
        self.assertEqual(generator.process_invoice_calls, 1)
        self.assertEqual(generator.template_output_dir, "/tmp/out")
        self.assertEqual(generator.template_output_path, "/tmp/out/decl.xlsx")
        self.assertEqual(generator.page_view_paths, ["/tmp/out/INV-001 报关资料.xlsx"])
        self.assertEqual(len(generator.call_records), 1)
        self.assertTrue(generator.call_records[0]["success"])
        self.assertEqual(render_detailed_list.call_count, 1)
        self.assertEqual(render_packing_list.call_count, 1)
        self.assertEqual(render_bill.call_count, 1)
        self.assertEqual(render_contract.call_count, 1)
        self.assertEqual(render_customs_declaration.call_count, 1)
        self.assertEqual(render_elements_sheet.call_count, 1)

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


class InvoiceRegisterTargetSelectionTests(unittest.TestCase):
    def test_register_combined_invoice_marks_non_virtual_warehouse_as_warehouse_shipping(self):
        generator = CustomsDeclaration()

        invoice_df = pd.DataFrame(
            [
                {
                    "发货单号": "SHIP-001",
                    "SKU": "SKU-1",
                    "发货数量": 1,
                    "箱数": 1,
                    "总重量（kg）-箱子": 10,
                    "总体积（m³）-箱子": 0.1,
                    "国家": "美国",
                    "单箱数量": 1,
                    "货件单号": "FBA-1",
                    "箱子毛重（kg）": 10,
                    "箱子长度（cm）": 10,
                    "箱子宽度（cm）": 10,
                    "箱子高度（cm）": 10,
                },
                {
                    "发货单号": "SHIP-002",
                    "SKU": "SKU-2",
                    "发货数量": 2,
                    "箱数": 1,
                    "总重量（kg）-箱子": 20,
                    "总体积（m³）-箱子": 0.2,
                    "国家": "美国",
                    "单箱数量": 2,
                    "货件单号": "FBA-2",
                    "箱子毛重（kg）": 20,
                    "箱子长度（cm）": 20,
                    "箱子宽度（cm）": 20,
                    "箱子高度（cm）": 20,
                },
            ]
        )
        invoice_details_df = pd.DataFrame(
            [
                {
                    "SKU": "SKU-1",
                    "运输方式": "海派",
                    "物流商单号": "INV-001",
                    "店铺": "EZARC",
                    "物流中心编码": "ONT6",
                    "发货单号": "SHIP-001",
                    "发货仓库（单据）": "杭州虚拟仓",
                    "物流商": "物流商A",
                },
                {
                    "SKU": "SKU-2",
                    "运输方式": "海派",
                    "物流商单号": "INV-001",
                    "店铺": "EZARC",
                    "物流中心编码": "ONT6",
                    "发货单号": "SHIP-002",
                    "发货仓库（单据）": "义乌仓",
                    "物流商": "物流商A",
                },
            ]
        )

        with patch.object(generator, "_read_excel_cached", side_effect=[invoice_df, invoice_details_df]):
            result = generator.register_combined_invoice("invoice.xlsx")

        self.assertEqual(result["发货仓库（单据）"].tolist(), ["工厂直发", "仓库发货"])

    def test_register_combined_invoice_raises_when_detail_sku_has_unmatched_suffix(self):
        generator = CustomsDeclaration()

        invoice_df = pd.DataFrame(
            [
                {
                    "发货单号": "SHIP-001",
                    "SKU": "T19104A",
                    "发货数量": 100,
                    "箱数": 2,
                    "总重量（kg）-箱子": 22,
                    "总体积（m³）-箱子": 0.03,
                    "国家": "日本",
                    "单箱数量": 50,
                    "货件单号": "FBA-1",
                    "箱子毛重（kg）": 11,
                    "箱子长度（cm）": 16,
                    "箱子宽度（cm）": 32,
                    "箱子高度（cm）": 29,
                }
            ]
        )
        invoice_details_df = pd.DataFrame(
            [
                {
                    "SKU": "T19104AJP",
                    "运输方式": "海卡",
                    "物流商单号": "26TL038",
                    "店铺": "TOLESA JP-JP",
                    "物流中心编码": "XJE2",
                    "发货单号": "SHIP-001",
                    "发货仓库（单据）": "余杭新仓库",
                    "物流商": "浙江凯西国际货运代理有限公司",
                }
            ]
        )

        with patch.object(generator, "_read_excel_cached", side_effect=[invoice_df, invoice_details_df]):
            with self.assertRaisesRegex(ValueError, "T19104A"):
                generator.register_combined_invoice("invoice.xlsx")

    def test_build_invoicing_register_uses_company_title_instead_of_first_brand(self):
        generator = CustomsDeclaration()
        generator.company_info_index = 4
        generator.lcl_flag = False
        generator.shipping_details = {
            "currency": "USD",
            "invoice_number": "INV-001",
            "transportation_method": "海派",
            "logistics_provider": "物流商A",
            "country_zh": "美国",
        }
        generator.result_df = pd.DataFrame(
            [
                {
                    "SKU": "SKU-1",
                    "报关中文品名": "品名1",
                    "英文品名": "Name1",
                    "HS CODE": "HS1",
                    "发货数量": 1,
                    "单位": "PCS",
                    "申报单价": 10,
                    "总价": 10,
                    "运费分摊": 0,
                    "是否含金刚石": False,
                }
            ]
        )

        invoice_df = pd.DataFrame(
            [
                {
                    "SKU": "SKU-1",
                    "店铺": "EZARC",
                    "发货单号": "SHIP-001",
                    "发货数量": 1,
                    "物流中心编码": "ONT6",
                    "国家": "美国",
                    "货件单号": "FBA123",
                    "物流商单号": "INV-001",
                }
            ]
        )
        product_name_df = pd.DataFrame(
            [{"*SKU": "SKU-1", "报关中文品名": "品名1"}]
        )

        with patch.object(generator, "register_combined_invoice", return_value=invoice_df), \
                patch.object(generator, "_read_excel_cached", return_value=product_name_df):
            _register_df, register_filepath, _sheet_name, _brand, _original_rows = generator.build_invoicing_register_df(
                "invoice.xlsx",
                "product.xlsx",
                "lumiya.xlsx",
                "easycan.xlsx",
                "sutong.xlsx",
                "liulang.xlsx",
                "jiangsheng.xlsx",
            )

        self.assertEqual(register_filepath, "jiangsheng.xlsx")


if __name__ == "__main__":
    unittest.main()
