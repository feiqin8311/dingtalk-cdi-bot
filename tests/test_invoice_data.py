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


class InvoiceBasicInformationTests(unittest.TestCase):
    def test_zero_volume_integer_column_is_normalized_to_float_before_fallback(self):
        generator = CustomsDeclaration()
        invoice_df = pd.DataFrame(
            {
                "发货单号": ["SHIP-001"],
                "货件单号": ["FBA-001"],
                "国家": ["美国"],
                "SKU": [922101],
                "总体积（m³）-箱子": pd.Series([0], dtype="int64"),
            }
        )
        invoice_details_df = pd.DataFrame(
            {
                "SKU": [922101],
                "物流中心编码": ["ONT6"],
                "运输方式": ["海派"],
                "物流商单号": ["INV-001"],
                "店铺": ["EZARC"],
                "物流商": ["物流商A"],
            }
        )
        site_address_df = pd.DataFrame({"目的港": ["ONT6"], "收件人": ["Amazon"]})
        product_df = pd.DataFrame(
            {
                "*SKU": [922101],
                "报关中文品名": ["测试品名"],
                "单品毛重(g)": [100],
                "外箱规格长(cm)": [10],
                "外箱规格宽(cm)": [10],
                "外箱规格高(cm)": [10],
                "单箱数量(pcs)": [1],
                "材质": ["钢"],
                "货源地": ["杭州"],
                "供应商名称": ["供应商A"],
                "报关用采购成本(CNY)": [1],
            }
        )

        with patch.object(
            generator,
            "_read_excel_cached",
            side_effect=[invoice_df, invoice_details_df, site_address_df, product_df],
        ):
            result_df, _, _ = generator.get_basic_information("invoice.xlsx", "site.xlsx", "product.xlsx")

        self.assertEqual(result_df.loc[0, "总体积（m³）-箱子"], 0.01)
        self.assertEqual(result_df["总体积（m³）-箱子"].dtype.kind, "f")


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
                patch("app.flows.invoice_flow.finalize_workbook_page_setup"), \
                patch("app.flows.invoice_flow.os.makedirs"), \
                patch("app.flows.invoice_flow.export_workbook_to_pdf") as export_workbook_to_pdf, \
                patch("app.flows.invoice_flow.export_worksheet_to_pdf") as export_worksheet_to_pdf:
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
        output_dir = "/tmp/out/INV-001 报关资料"
        self.assertEqual(generator.template_output_dir, "/tmp/out")
        self.assertEqual(generator.template_output_path, "/tmp/out/decl.xlsx")
        self.assertEqual(generator.page_view_paths, [f"{output_dir}/INV-001 报关资料.xlsx"])
        export_workbook_to_pdf.assert_called_once_with(
            f"{output_dir}/INV-001 报关资料.xlsx", f"{output_dir}/INV-001 报关资料.pdf"
        )
        export_worksheet_to_pdf.assert_called_once_with(
            f"{output_dir}/INV-001 报关资料.xlsx", "合同", f"{output_dir}/INV-001 销售合同.pdf"
        )
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

    def test_build_invoicing_register_df_for_sutong_deduplicates_sku_rows(self):
        generator = CustomsDeclaration()
        generator.company_info_index = 2
        generator.lcl_flag = True
        generator.result_df = pd.DataFrame(
            [
                {
                    "SKU": "425202",
                    "报关中文品名": "剪刀",
                    "英文品名": "Scissors",
                    "HS CODE": "8213000000",
                    "发货数量": 480,
                    "单位": "套",
                    "申报单价": 2.67,
                    "总价": 1281.6,
                },
                {
                    "SKU": "425205",
                    "报关中文品名": "橡皮",
                    "英文品名": "Eraser",
                    "HS CODE": "4016920000",
                    "发货数量": 150,
                    "单位": "套",
                    "申报单价": 0.6,
                    "总价": 90,
                },
            ]
        )

        register_combined_df = pd.DataFrame(
            [
                {
                    "SKU": "425202",
                    "物流商单号": "26YP065",
                    "店铺": "TRAILFUN",
                    "发货数量": 96,
                    "发货单号": "SP-1",
                    "箱数": 1,
                    "单箱数量": 96,
                },
                {
                    "SKU": "425202",
                    "物流商单号": "26YP065",
                    "店铺": "TRAILFUN",
                    "发货数量": 96,
                    "发货单号": "SP-2",
                    "箱数": 1,
                    "单箱数量": 96,
                },
                {
                    "SKU": "425202",
                    "物流商单号": "26YP065",
                    "店铺": "TRAILFUN",
                    "发货数量": 96,
                    "发货单号": "SP-3",
                    "箱数": 1,
                    "单箱数量": 96,
                },
                {
                    "SKU": "425202",
                    "物流商单号": "26YP065",
                    "店铺": "TRAILFUN",
                    "发货数量": 96,
                    "发货单号": "SP-4",
                    "箱数": 1,
                    "单箱数量": 96,
                },
                {
                    "SKU": "425202",
                    "物流商单号": "26YP065",
                    "店铺": "TRAILFUN",
                    "发货数量": 96,
                    "发货单号": "SP-5",
                    "箱数": 1,
                    "单箱数量": 96,
                },
                {
                    "SKU": "425205",
                    "物流商单号": "26YP065",
                    "店铺": "TRAILFUN",
                    "发货数量": 30,
                    "发货单号": "SP-6",
                    "箱数": 1,
                    "单箱数量": 30,
                },
                {
                    "SKU": "425205",
                    "物流商单号": "26YP065",
                    "店铺": "TRAILFUN",
                    "发货数量": 30,
                    "发货单号": "SP-7",
                    "箱数": 1,
                    "单箱数量": 30,
                },
                {
                    "SKU": "425205",
                    "物流商单号": "26YP065",
                    "店铺": "TRAILFUN",
                    "发货数量": 30,
                    "发货单号": "SP-8",
                    "箱数": 1,
                    "单箱数量": 30,
                },
                {
                    "SKU": "425205",
                    "物流商单号": "26YP065",
                    "店铺": "TRAILFUN",
                    "发货数量": 30,
                    "发货单号": "SP-9",
                    "箱数": 1,
                    "单箱数量": 30,
                },
                {
                    "SKU": "425205",
                    "物流商单号": "26YP065",
                    "店铺": "TRAILFUN",
                    "发货数量": 30,
                    "发货单号": "SP-10",
                    "箱数": 1,
                    "单箱数量": 30,
                },
            ]
        )
        product_df = pd.DataFrame(
            [
                {"*SKU": "425202", "报关中文品名": "剪刀"},
                {"*SKU": "425205", "报关中文品名": "橡皮"},
            ]
        )

        with patch.object(generator, "register_combined_invoice", return_value=register_combined_df), \
                patch.object(generator, "_read_excel_cached", return_value=product_df):
            register_df, register_filepath, sheet_name, brand, original_rows = generator.build_invoicing_register_df(
                "invoice.xlsx",
                "product.xlsx",
                "lumiya.xlsx",
                "easycan.xlsx",
                "sutong.xlsx",
                "liulang.xlsx",
                "jiangsheng.xlsx",
            )

        self.assertEqual(register_filepath, "sutong.xlsx")
        self.assertEqual(sheet_name, "2026")
        self.assertEqual(brand, "TRAILFUN")
        self.assertEqual(original_rows, 10)
        self.assertEqual(len(register_df), 2)
        self.assertEqual(register_df["SKU"].tolist(), ["425202", "425205"])
        self.assertEqual(register_df["数量"].tolist(), [480, 150])

    def test_build_invoicing_register_df_for_easycan_keeps_detail_rows_in_order(self):
        generator = CustomsDeclaration()
        generator.company_info_index = 0
        generator.lcl_flag = False
        generator.result_df = pd.DataFrame(
            [
                {
                    "SKU": "SKU-B", "报关中文品名": "报关品名B", "英文品名": "Name B",
                    "HS CODE": "HS-B", "发货数量": 20, "单位": "套", "申报单价": 2, "总价": 40,
                },
                {
                    "SKU": "SKU-A", "报关中文品名": "报关品名A", "英文品名": "Name A",
                    "HS CODE": "HS-A", "发货数量": 10, "单位": "件", "申报单价": 3, "总价": 30,
                },
            ]
        )
        register_combined_df = pd.DataFrame(
            [
                {"SKU": "SKU-A", "物流商单号": "26EA177", "店铺": "EZARC", "发货数量": 1},
                {"SKU": "SKU-B", "物流商单号": "26EA177", "店铺": "EZARC", "发货数量": 2},
                {"SKU": "SKU-B", "物流商单号": "26EA177", "店铺": "EZARC", "发货数量": 2},
                {"SKU": "SKU-EXTRA", "物流商单号": "26EA177", "店铺": "EZARC", "发货数量": 3},
            ]
        )
        product_df = pd.DataFrame(
            [
                {"*SKU": "SKU-A", "报关中文品名": "源品名A"},
                {"*SKU": "SKU-B", "报关中文品名": "源品名B"},
                {"*SKU": "SKU-EXTRA", "报关中文品名": "源品名C"},
            ]
        )

        with patch.object(generator, "register_combined_invoice", return_value=register_combined_df), \
                patch.object(generator, "_read_excel_cached", return_value=product_df):
            register_df, register_filepath, sheet_name, _brand, _original_rows = generator.build_invoicing_register_df(
                "invoice.xlsx", "product.xlsx", "lumiya.xlsx", "easycan.xlsx", "sutong.xlsx",
                "liulang.xlsx", "jiangsheng.xlsx",
            )

        self.assertEqual(register_filepath, "easycan.xlsx")
        self.assertEqual(sheet_name, "2026（内部开票）")
        self.assertEqual(register_df["SKU"].tolist(), ["SKU-A", "SKU-B", "SKU-B", "SKU-EXTRA"])
        self.assertEqual(register_df["开票/报关品名"].tolist()[:2], ["报关品名B", "报关品名A"])
        self.assertEqual(register_df["数量"].tolist()[:2], [20, 10])
        self.assertTrue(pd.isna(register_df.loc[2, "开票/报关品名"]))


class InvoiceCustomsDeclarationSheetTests(unittest.TestCase):
    def test_customs_declaration_mode_includes_invoice_number(self):
        generator = CustomsDeclaration()
        generator.shipping_details = {
            "domestic_shipper": "杭州素桐电子商务有限公司91330106MA2KC3T74A",
            "production_sales_unit": "杭州素桐电子商务有限公司91330106MA2KC3T74A",
            "invoice_number": "26YP065",
            "country_zh": "美国",
            "TOTAL_PACKED_IN": 1,
            "TOTAL_G.W": 2.5,
            "TOTAL_N.W": 2.0,
            "freight": 12.3,
            "domestic_source_of_goods": "33019 杭州其他",
            "images": ["image0", "image1"],
        }
        generator.result_df = pd.DataFrame(
            [
                {
                    "SKU": "SKU-1",
                    "报关中文品名": "品名1",
                    "是否含金刚石": False,
                    "商品编号": "1234567890",
                    "数量及单位": "1PCS",
                    "箱数": "1",
                    "预留1": "",
                    "预留2": "",
                    "单价": 10.5,
                    "总价": 10.5,
                }
            ]
        )

        wb = Workbook()

        with patch.object(generator, "add_images"):
            generator.add_customs_declaration_sheet("unused.xlsx", "/tmp", wb=wb)

        ws = wb["报关单"]
        cell_value = ws["E8"].value

        self.assertEqual(str(cell_value), "9710(26YP065)")


class InvoiceLumiyaCompanyInfoTests(unittest.TestCase):
    def test_consolidate_lcl_shipping_rows_counts_shared_cartons_once(self):
        generator = CustomsDeclaration()
        source_df = pd.DataFrame(
            [
                {"发货单号": "SP260721036", "SKU": "304002", "发货数量": 40, "箱数": 1,
                 "箱子毛重（kg）": 15.88, "箱子长度（cm）": 35.56, "箱子宽度（cm）": 33.03,
                 "箱子高度（cm）": 30.48, "总重量（kg）-箱子": 15.88, "总体积（m³）-箱子": 0.04},
                {"发货单号": "SP260721036", "SKU": "304004", "发货数量": 84, "箱数": 1,
                 "箱子毛重（kg）": 15.88, "箱子长度（cm）": 35.56, "箱子宽度（cm）": 33.03,
                 "箱子高度（cm）": 30.48, "总重量（kg）-箱子": 15.88, "总体积（m³）-箱子": 0.04},
                {"发货单号": "SP260721036", "SKU": "304002", "发货数量": 40, "箱数": 1,
                 "箱子毛重（kg）": 15.88, "箱子长度（cm）": 35.56, "箱子宽度（cm）": 33.03,
                 "箱子高度（cm）": 30.49, "总重量（kg）-箱子": 15.88, "总体积（m³）-箱子": 0.04},
                {"发货单号": "SP260721036", "SKU": "304004", "发货数量": 84, "箱数": 1,
                 "箱子毛重（kg）": 15.88, "箱子长度（cm）": 35.56, "箱子宽度（cm）": 33.03,
                 "箱子高度（cm）": 30.49, "总重量（kg）-箱子": 15.88, "总体积（m³）-箱子": 0.04},
            ]
        )

        result = generator.consolidate_lcl_shipping_rows(source_df)

        self.assertEqual(result["SKU"].tolist(), ["304002", "304004"])
        self.assertEqual(result["发货数量"].tolist(), [80, 168])
        self.assertEqual(result.loc[0, "箱数"], 2)
        self.assertEqual(result.loc[0, "总重量（kg）-箱子"], 31.76)
        self.assertEqual(result.loc[0, "总体积（m³）-箱子"], 0.08)
        self.assertTrue(pd.isna(result.loc[1, "箱数"]))

    def test_update_shipment_quantity_uses_each_distinct_carton_quantity(self):
        generator = CustomsDeclaration()
        shipment_df = pd.DataFrame(
            [
                {
                    "发货单号": "SP260721036",
                    "SKU": "304002",
                    "发货数量": 80,
                    "单箱数量": 40,
                    "箱数": 1,
                    "箱子毛重（kg）": 15.88,
                    "总重量（kg）-箱子": 15.88,
                    "总体积（m³）-箱子": 0.04,
                    "箱子高度（cm）": 30.48,
                },
                {
                    "发货单号": "SP260721036",
                    "SKU": "304004",
                    "发货数量": 168,
                    "单箱数量": 84,
                    "箱数": 1,
                    "箱子毛重（kg）": 15.88,
                    "总重量（kg）-箱子": 15.88,
                    "总体积（m³）-箱子": 0.04,
                    "箱子高度（cm）": 30.48,
                },
                {
                    "发货单号": "SP260721036",
                    "SKU": "304002",
                    "发货数量": 80,
                    "单箱数量": 40,
                    "箱数": 1,
                    "箱子毛重（kg）": 15.88,
                    "总重量（kg）-箱子": 15.88,
                    "总体积（m³）-箱子": 0.04,
                    "箱子高度（cm）": 30.49,
                },
                {
                    "发货单号": "SP260721036",
                    "SKU": "304004",
                    "发货数量": 168,
                    "单箱数量": 84,
                    "箱数": 1,
                    "箱子毛重（kg）": 15.88,
                    "总重量（kg）-箱子": 15.88,
                    "总体积（m³）-箱子": 0.04,
                    "箱子高度（cm）": 30.49,
                },
            ]
        )

        result = generator.update_shipment_quantity(shipment_df)

        self.assertEqual(len(result), 4)
        self.assertEqual(result["发货数量"].tolist(), [40, 84, 40, 84])
        self.assertEqual(result["箱数"].tolist(), [1, 1, 1, 1])
        self.assertEqual(result["总重量（kg）-箱子"].tolist(), [15.88, 15.88, 15.88, 15.88])
        self.assertEqual(result["总体积（m³）-箱子"].tolist(), [0.04, 0.04, 0.04, 0.04])

    def test_lumiya_company_info_uses_updated_address(self):
        generator = CustomsDeclaration()
        lumiya_info = generator.company_info[1]

        self.assertEqual(lumiya_info["company_name_en"], "Hangzhou Lumiya Technology Co., Ltd.")
        self.assertTrue(lumiya_info["consignor"].startswith("Hangzhou Lumiya Technology Co., Ltd.\n"))
        self.assertIn("BENEFICIARY NAME: Hangzhou Lumiya Technology Co., Ltd.", lumiya_info["bank_informations"])
        self.assertIn("Add: Rm W318-2, Fl.3, Unit 5, Bld. 10 Area E, Aicheng Str.1#, Wuchang Street, Yuhang District, Hangzhou, Zhejiang, 311121, CN", lumiya_info["consignor"])
        self.assertIn("ADDRESS: Rm W318-2, Fl.3, Unit 5, Bld. 10 Area E, Aicheng Str.1#, Wuchang Street, Yuhang District, Hangzhou, Zhejiang, 311121, CN", lumiya_info["bank_informations"])

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

    def test_register_combined_invoice_does_not_expand_rows_for_duplicate_detail_keys(self):
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
                    "发货单号": "SHIP-001",
                    "SKU": "SKU-1",
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
                    "SKU": "SKU-1",
                    "运输方式": "海派",
                    "物流商单号": "INV-001",
                    "店铺": "EZARC",
                    "物流中心编码": "ONT6",
                    "发货单号": "SHIP-001",
                    "发货仓库（单据）": "杭州虚拟仓",
                    "物流商": "物流商A",
                },
            ]
        )

        with patch.object(generator, "_read_excel_cached", side_effect=[invoice_df, invoice_details_df]):
            result = generator.register_combined_invoice("invoice.xlsx")

        self.assertEqual(len(result), 2)
        self.assertEqual(result["发货数量"].tolist(), [1, 2])

    def test_register_combined_invoice_reuses_one_detail_for_multiple_packing_rows(self):
        generator = CustomsDeclaration()
        invoice_df = pd.DataFrame(
            [
                {
                    "发货单号": "SP260714032",
                    "SKU": "T6HW84",
                    "发货数量": 60,
                    "箱数": 3,
                    "总重量（kg）-箱子": 48,
                    "总体积（m³）-箱子": 0.20,
                    "国家": "英国",
                    "单箱数量": 20,
                    "货件单号": "FBA15M1N5HTC",
                    "箱子毛重（kg）": 16,
                    "箱子长度（cm）": 33,
                    "箱子宽度（cm）": 46,
                    "箱子高度（cm）": 44,
                },
                {
                    "发货单号": "SP260714032",
                    "SKU": "T6HW84",
                    "发货数量": 18,
                    "箱数": 1,
                    "总重量（kg）-箱子": 16,
                    "总体积（m³）-箱子": 0.07,
                    "国家": "英国",
                    "单箱数量": 18,
                    "货件单号": "FBA15M1N5HTC",
                    "箱子毛重（kg）": 16,
                    "箱子长度（cm）": 33,
                    "箱子宽度（cm）": 46,
                    "箱子高度（cm）": 44,
                },
            ]
        )
        invoice_details_df = pd.DataFrame(
            [
                {
                    "发货单号": "SP260714032",
                    "SKU": "T6HW84",
                    "运输方式": "海卡",
                    "物流商单号": "26TL066",
                    "店铺": "SIXPRAY EU-UK",
                    "物流中心编码": "LBA8",
                    "发货仓库（单据）": "义乌仓",
                    "物流商": "堡森国际物流有限公司",
                }
            ]
        )

        with patch.object(generator, "_read_excel_cached", side_effect=[invoice_df, invoice_details_df]):
            result = generator.register_combined_invoice("invoice.xlsx")

        self.assertEqual(result["发货数量"].tolist(), [60, 18])
        self.assertEqual(result["物流商单号"].tolist(), ["26TL066", "26TL066"])
        self.assertEqual(result["物流中心编码"].tolist(), ["LBA8", "LBA8"])

    def test_build_invoicing_register_df_drops_duplicate_final_rows(self):
        generator = CustomsDeclaration()
        generator.company_info_index = 2
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
                    "发货数量": 3,
                    "单位": "PCS",
                    "申报单价": 10,
                    "总价": 30,
                }
            ]
        )

        invoice_df = pd.DataFrame(
            [
                {
                    "SKU": "SKU-1",
                    "店铺": "EZARC",
                    "物流商单号": "INV-001",
                    "发货单号": "SHIP-001",
                    "发货数量": 3,
                },
                {
                    "SKU": "SKU-1",
                    "店铺": "EZARC",
                    "物流商单号": "INV-001",
                    "发货单号": "SHIP-001",
                    "发货数量": 3,
                },
            ]
        )
        product_name_df = pd.DataFrame([{"*SKU": "SKU-1", "报关中文品名": "品名1"}])

        with patch.object(generator, "register_combined_invoice", return_value=invoice_df), \
                patch.object(generator, "_read_excel_cached", return_value=product_name_df):
            register_df, _register_filepath, _sheet_name, _brand, _original_rows = generator.build_invoicing_register_df(
                "invoice.xlsx",
                "product.xlsx",
                "lumiya.xlsx",
                "easycan.xlsx",
                "sutong.xlsx",
                "liulang.xlsx",
                "jiangsheng.xlsx",
            )

        self.assertEqual(len(register_df), 1)

    def test_register_shipping_invoice_number_preserves_duplicate_detail_rows(self):
        generator = CustomsDeclaration()
        generator.lcl_flag = False
        generator.shipping_details = {"invoice_number": "INV-001"}

        register_source_df = pd.DataFrame(
            [
                {
                    "物流商单号": "INV-001",
                    "物流中心编码": "ONT6",
                    "国家": "美国",
                    "运输方式": "海派",
                    "货件单号": "FBA-1",
                    "店铺": "EZARC",
                    "SKU": "SKU-1",
                    "发货数量": 3,
                    "箱数": 1,
                    "总重量（kg）-箱子": 10,
                    "总体积（m³）-箱子": 0.1,
                    "发货仓库（单据）": "工厂直发",
                    "物流商": "物流商A",
                },
                {
                    "物流商单号": "INV-001",
                    "物流中心编码": "ONT6",
                    "国家": "美国",
                    "运输方式": "海派",
                    "货件单号": "FBA-1",
                    "店铺": "EZARC",
                    "SKU": "SKU-1",
                    "发货数量": 3,
                    "箱数": 1,
                    "总重量（kg）-箱子": 10,
                    "总体积（m³）-箱子": 0.1,
                    "发货仓库（单据）": "工厂直发",
                    "物流商": "物流商A",
                },
            ]
        )

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            temp_path = tmp.name

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "出运明细"
            ws.append([
                "发货日期", "发票号", "站点", "物流中心编码", "国家", "运输方式", "货件单号", "店铺", "SKU", "发货数量",
                "箱数", "总重量（kg）-箱子", "总体积（m³）-箱子", "关税", "VAT/消费税", "货运费", "发货仓库（单据）", "物流商"
            ])
            wb.save(temp_path)
            wb.close()

            with patch.object(generator, "register_combined_invoice", return_value=register_source_df):
                deleted_rows, actual_rows, written_rows = generator.register_shipping_invoice_number(
                    "invoice.xlsx", temp_path
                )

            written_df = pd.read_excel(temp_path, sheet_name="出运明细")
            self.assertEqual(len(written_df.dropna(how="all")), 2)
            self.assertEqual(deleted_rows, 0)
            self.assertEqual(actual_rows, 3)
            self.assertEqual(written_rows, 2)
        finally:
            os.remove(temp_path)


if __name__ == "__main__":
    unittest.main()
