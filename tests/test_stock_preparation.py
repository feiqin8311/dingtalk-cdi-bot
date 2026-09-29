import os
import sys
import tempfile
import types
import unittest
from unittest.mock import call, patch

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


class StockBasicInformationTests(unittest.TestCase):
    def test_zero_volume_integer_column_is_normalized_to_float_before_fallback(self):
        generator = CustomsDeclaration()
        stock_details_df = pd.DataFrame(
            {
                "备货单号": ["STOCK-001"],
                "发货仓库（单据）": ["杭州虚拟仓"],
                "收货仓库": ["EZARC NA-US美国仓(AWD)"],
                "sku": ["922101"],
            }
        )
        product_df = pd.DataFrame(
            {
                "*SKU": ["922101"],
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
                "品牌": ["EZARC"],
            }
        )
        tracking_df = pd.DataFrame(
            {
                "备货单号": ["STOCK-001"],
                "物流商": ["物流商A"],
                "物流渠道": ["海派"],
                "查询单号": ["INV-001"],
            }
        )
        packing_df = pd.DataFrame(
            {
                "备货单号": ["STOCK-001"],
                "SKU": ["922101"],
                "国家": ["美国"],
                "总体积（m³）-箱子": pd.Series([0], dtype="int64"),
            }
        )

        with patch.object(
            generator,
            "_read_excel_cached",
            side_effect=[stock_details_df, product_df, tracking_df, packing_df],
        ):
            result_df, _, _ = generator.get_basic_information("stock.xlsx", "site.xlsx", "product.xlsx")

        self.assertEqual(result_df.loc[0, "总体积（m³）-箱子"], 0.01)
        self.assertEqual(result_df["总体积（m³）-箱子"].dtype.kind, "f")
        self.assertEqual(generator.shipping_details["recipient"], "Amazon.com Services, Inc.")
        self.assertTrue(generator.shipping_details["is_awd"])
        self.assertEqual(generator.shipping_details["currency"], "USD")
        self.assertEqual(generator.shipping_details["currency_symbol"], "$")
        self.assertEqual(generator.shipping_details["country_zh"], "美国")
        self.assertEqual(generator.shipping_details["country"], "USA")

    def test_non_awd_currency_stays_usd(self):
        generator = CustomsDeclaration()
        stock_details_df = pd.DataFrame(
            {
                "备货单号": ["STOCK-001"],
                "发货仓库（单据）": ["杭州虚拟仓"],
                "收货仓库": ["FBT仓-TIKTOK"],
                "sku": ["922101"],
            }
        )
        product_df = pd.DataFrame(
            {
                "*SKU": ["922101"],
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
                "品牌": ["EZARC"],
            }
        )
        tracking_df = pd.DataFrame(
            {
                "备货单号": ["STOCK-001"],
                "物流商": ["物流商A"],
                "物流渠道": ["海派"],
                "查询单号": ["INV-001"],
            }
        )
        packing_df = pd.DataFrame(
            {
                "备货单号": ["STOCK-001"],
                "SKU": ["922101"],
                "国家": ["美国"],
                "总体积（m³）-箱子": pd.Series([0.1], dtype="float64"),
            }
        )

        with patch.object(
            generator,
            "_read_excel_cached",
            side_effect=[stock_details_df, product_df, tracking_df, packing_df],
        ):
            generator.get_basic_information("stock.xlsx", "site.xlsx", "product.xlsx")

        self.assertFalse(generator.shipping_details["is_awd"])
        self.assertEqual(generator.shipping_details["currency"], "USD")
        self.assertEqual(generator.shipping_details["currency_cn"], "美元")
        self.assertEqual(generator.shipping_details["currency_symbol"], "$")

    def test_country_comes_from_packing_sheet(self):
        generator = CustomsDeclaration()
        stock_details_df = pd.DataFrame(
            {
                "备货单号": ["STOCK-001"],
                "发货仓库（单据）": ["良品仓"],
                "收货仓库": ["EZARC NA-US美国仓(AWD)"],
                "sku": ["922101"],
            }
        )
        product_df = pd.DataFrame(
            {
                "*SKU": ["922101"],
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
                "品牌": ["LIBRATON"],
            }
        )
        tracking_df = pd.DataFrame(
            {
                "备货单号": ["STOCK-001"],
                "物流商": ["物流商A"],
                "物流渠道": ["海派"],
                "查询单号": ["INV-001"],
            }
        )
        packing_df = pd.DataFrame(
            {
                "备货单号": ["STOCK-001"],
                "SKU": ["922101"],
                "国家": ["澳洲"],
                "总体积（m³）-箱子": pd.Series([0.1], dtype="float64"),
            }
        )

        with patch.object(
            generator,
            "_read_excel_cached",
            side_effect=[stock_details_df, product_df, tracking_df, packing_df],
        ):
            generator.get_basic_information("stock.xlsx", "site.xlsx", "product.xlsx")

        self.assertEqual(generator.shipping_details["country_zh"], "澳洲")
        self.assertEqual(generator.shipping_details["country"], "AU")

    def test_country_alias_from_packing_sheet_is_normalized(self):
        generator = CustomsDeclaration()
        stock_details_df = pd.DataFrame(
            {
                "备货单号": ["STOCK-001"],
                "发货仓库（单据）": ["杭州虚拟仓"],
                "收货仓库": ["EZARC NA-US美国仓(AWD)"],
                "sku": ["922101"],
            }
        )
        product_df = pd.DataFrame(
            {
                "*SKU": ["922101"],
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
                "品牌": ["EZARC"],
            }
        )
        tracking_df = pd.DataFrame(
            {
                "备货单号": ["STOCK-001"],
                "物流商": ["物流商A"],
                "物流渠道": ["海派"],
                "查询单号": ["INV-001"],
            }
        )
        packing_df = pd.DataFrame(
            {
                "备货单号": ["STOCK-001"],
                "SKU": ["922101"],
                "国家": ["Us"],
                "总体积（m³）-箱子": pd.Series([0.1], dtype="float64"),
            }
        )

        with patch.object(
            generator,
            "_read_excel_cached",
            side_effect=[stock_details_df, product_df, tracking_df, packing_df],
        ):
            normalized_packing_df, _, _ = generator.get_basic_information("stock.xlsx", "site.xlsx", "product.xlsx")

        self.assertEqual(generator.shipping_details["country_zh"], "美国")
        self.assertEqual(generator.shipping_details["country"], "USA")
        self.assertEqual(normalized_packing_df.loc[0, "国家"], "美国")

    def test_warehouse_recipient_mapping(self):
        cases = [
            ("EZARC NA-US美国仓(AWD)", "Amazon.com Services, Inc."),
            ("topg-德克萨斯州", "TOPG INTERNATIONAL LIMITED\n5115 south shaver st,Houston,TX,77034"),
            ("FBT仓-TIKTOK", "Fulfilled by Tiktok"),
            ("仓盛美中仓 Bolingbrook", "ShipSage\n460 Gibraltar Drive, Bolingbrook IL 60440, US"),
            ("海盟-新泽西", "GCE Logistics East Inc.\n545 Monmouth Rd, Jackson Township, NJ 08527"),
        ]
        for warehouse, recipient in cases:
            with self.subTest(warehouse=warehouse):
                series = pd.Series([warehouse])
                self.assertEqual(CustomsDeclaration.resolve_warehouse_recipient(series), recipient)

    def test_unmatched_warehouse_raises(self):
        with self.assertRaisesRegex(ValueError, "收货仓库未匹配到收件人"):
            CustomsDeclaration.resolve_warehouse_recipient(pd.Series(["未知仓库"]))


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

        with patch("app.flows.stock_flow.openpyxl.Workbook", return_value=DummyWorkbook()), \
                patch("app.flows.stock_flow.render_detailed_list") as render_detailed_list, \
                patch("app.flows.stock_flow.render_packing_list") as render_packing_list, \
                patch("app.flows.stock_flow.render_bill") as render_bill, \
                patch("app.flows.stock_flow.render_contract") as render_contract, \
                patch("app.flows.stock_flow.render_customs_declaration") as render_customs_declaration, \
                patch("app.flows.stock_flow.render_elements_sheet") as render_elements_sheet, \
                patch("app.flows.stock_flow.ensure_default_sheet_removed"), \
                patch("app.flows.stock_flow.finalize_workbook_page_setup"), \
                patch("app.flows.stock_flow.os.makedirs"), \
                patch("app.flows.stock_flow.export_workbook_to_pdf") as export_workbook_to_pdf, \
                patch("app.flows.stock_flow.export_worksheet_to_pdf") as export_worksheet_to_pdf:
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
        output_dir = "/tmp/out/INV-002 报关资料"
        self.assertEqual(generator.template_output_dir, "/tmp/out")
        self.assertEqual(generator.template_output_path, "/tmp/out/decl.xlsx")
        self.assertEqual(
            generator.page_view_paths,
            [
                f"{output_dir}/INV-002 报关资料.xlsx",
            ],
        )
        export_workbook_to_pdf.assert_called_once_with(
            f"{output_dir}/INV-002 报关资料.xlsx", f"{output_dir}/INV-002 报关资料.pdf"
        )
        export_worksheet_to_pdf.assert_called_once_with(
            f"{output_dir}/INV-002 报关资料.xlsx", "合同", f"{output_dir}/INV-002 销售合同.pdf"
        )
        self.assertEqual(len(generator.call_records), 1)
        self.assertTrue(generator.call_records[0]["success"])
        self.assertEqual(render_detailed_list.call_count, 1)
        self.assertEqual(render_packing_list.call_count, 1)
        self.assertEqual(render_bill.call_count, 1)
        self.assertEqual(render_contract.call_count, 1)
        self.assertEqual(render_customs_declaration.call_count, 1)
        self.assertEqual(render_elements_sheet.call_count, 1)


class StockCustomsDeclarationSheetTests(unittest.TestCase):
    def test_awd_customs_declaration_mode_includes_invoice_number(self):
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
            "is_awd": True,
            "recipient": "Amazon.com Services, Inc.",
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
        self.assertEqual(ws["A6"].value, "Amazon.com Services, Inc.")

    def test_non_awd_customs_declaration_mode_also_includes_invoice_number(self):
        generator = CustomsDeclaration()
        generator.shipping_details = {
            "domestic_shipper": "杭州易达工具有限公司3301965G22",
            "production_sales_unit": "杭州易达工具有限公司91330108092054785Y",
            "invoice_number": "INV-002",
            "country_zh": "美国",
            "TOTAL_PACKED_IN": 1,
            "TOTAL_G.W": 2.5,
            "TOTAL_N.W": 2.0,
            "freight": 12.3,
            "domestic_source_of_goods": "33019 杭州其他",
            "images": ["image0", "image1"],
            "is_awd": False,
            "recipient": "Fulfilled by Tiktok",
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
        self.assertEqual(str(ws["E8"].value), "9710(INV-002)")
        self.assertEqual(ws["A6"].value, "Fulfilled by Tiktok")


class StockDeclarationTemplateTests(unittest.TestCase):
    def test_non_liulang_declaration_template_order_type_stays_b(self):
        generator = CustomsDeclaration()
        generator.shipping_details = {
            "brand": "EZARC",
            "invoice_number": "INV-002",
            "TOTAL_PRICE": 10,
            "freight": 1.2,
        }
        generator.result_df = pd.DataFrame(
            [
                {
                    "HS CODE": "8203200000",
                    "报关中文品名": "测试品名",
                    "单位": "个",
                    "备货数量": 4,
                    "申报单价": 99,
                    "总价": 10,
                }
            ]
        )

        wb = Workbook()
        wb.active.title = "申报模板"
        unit_ws = wb.create_sheet("计量单位代码")
        unit_ws.append([])
        unit_ws.append(["计量单位代码", "计量单位名称"])
        unit_ws.append(["007", "个"])

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            template_path = tmp.name
        try:
            wb.save(template_path)

            generator.fill_declaration_template_file(template_path)

            result_wb = load_workbook(template_path)
            try:
                self.assertEqual(result_wb.worksheets[0]["D4"].value, "B")
                self.assertEqual(result_wb.worksheets[0]["R4"].value, 2.5)
                self.assertEqual(result_wb.worksheets[0]["R4"].number_format, "0.0000")
            finally:
                result_wb.close()
        finally:
            os.remove(template_path)


class StockLumiyaCompanyInfoTests(unittest.TestCase):
    def test_lumiya_company_info_uses_updated_address(self):
        generator = CustomsDeclaration()
        lumiya_info = generator.company_info[1]

        self.assertEqual(lumiya_info["company_name_en"], "Hangzhou Lumiya Technology Co., Ltd.")
        self.assertTrue(lumiya_info["consignor"].startswith("Hangzhou Lumiya Technology Co., Ltd.\n"))
        self.assertIn("BENEFICIARY NAME: Hangzhou Lumiya Technology Co., Ltd.", lumiya_info["bank_informations"])
        self.assertIn("Add: Rm W318-2, Fl.3, Unit 5, Bld. 10 Area E, Aicheng Str.1#, Wuchang Street, Yuhang District, Hangzhou, Zhejiang, 311121, CN", lumiya_info["consignor"])
        self.assertIn("ADDRESS: Rm W318-2, Fl.3, Unit 5, Bld. 10 Area E, Aicheng Str.1#, Wuchang Street, Yuhang District, Hangzhou, Zhejiang, 311121, CN", lumiya_info["bank_informations"])


class StockRegisterTargetSelectionTests(unittest.TestCase):
    def test_register_shipping_invoice_number_preserves_packing_country(self):
        generator = CustomsDeclaration()
        generator.lcl_flag = False
        generator.shipping_details = {"invoice_number": "INV-AU"}
        register_source_df = pd.DataFrame(
            [
                {
                    "备货单号": "STOCK-001",
                    "SKU": "SKU-1",
                    "备货数量": 1,
                    "箱数": 1,
                    "总重量（kg）-箱子": 10,
                    "总体积（m³）-箱子": 0.1,
                    "国家": "澳洲",
                    "单箱数量": 1,
                    "箱子毛重（kg）": 10,
                    "箱子长度（cm）": 10,
                    "箱子宽度（cm）": 10,
                    "箱子高度（cm）": 10,
                    "第三方仓产品名": "Product 1",
                    "店铺": "LIBRATON",
                    "物流中心编码": "",
                    "发货仓库（单据）": "仓库发货",
                    "物流商": "物流商A",
                    "收货仓库": "澳大利亚卫星仓AWD",
                }
            ]
        )
        tracking_df = pd.DataFrame(
            [
                {
                    "备货单号": "STOCK-001",
                    "物流商": "物流商A",
                    "物流渠道": "海派",
                    "查询单号": "INV-AU",
                }
            ]
        )

        wb = Workbook()
        ws = wb.active
        ws.title = "出运明细"
        ws.append([
            "发货日期", "发票号", "站点", "物流中心编码", "国家", "出运方式", "FBA编号", "店铺",
            "SKU", "备货数量", "箱数", "总重量（kg）-箱子", "总体积（m³）-箱子", "关税",
            "VAT/消费税", "货运费", "发货仓库（单据）", "物流商",
        ])
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            register_path = tmp.name
        try:
            wb.save(register_path)

            with patch.object(generator, "register_combined_invoice", return_value=register_source_df), \
                    patch.object(generator, "_read_excel_cached", return_value=tracking_df):
                generator.register_shipping_invoice_number("stock.xlsx", register_path, "product.xlsx")

            result_wb = load_workbook(register_path)
            try:
                self.assertEqual(result_wb["出运明细"]["E2"].value, "澳洲")
            finally:
                result_wb.close()
        finally:
            os.remove(register_path)

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
            register_df, register_filepath, _sheet_name, _brand = generator.build_invoicing_register_df(
                "invoice.xlsx",
                "product.xlsx",
                "lumiya.xlsx",
                "easycan.xlsx",
                "sutong.xlsx",
                "liulang.xlsx",
                "jiangsheng.xlsx",
            )

        self.assertEqual(register_filepath, "easycan.xlsx")
        self.assertEqual(register_df.loc[0, "报关渠道"], "海外仓")
        self.assertEqual(register_df.loc[0, "报关方式（9710/一般贸易）"], 9710)
        self.assertEqual(register_df.loc[0, "成交方式(FOB/CIF/CFR)"], "CFR")
        self.assertEqual(register_df.loc[0, "币别"], "USD")

    def test_liulang_invoicing_register_uses_9810(self):
        generator = CustomsDeclaration()
        generator.company_info_index = 3
        generator.lcl_flag = False
        generator.shipping_details = {
            "currency": "USD",
            "invoice_number": "26EBA42",
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
                    "店铺": "SIXPRAY",
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
            register_df, register_filepath, sheet_name, _brand = generator.build_invoicing_register_df(
                "invoice.xlsx",
                "product.xlsx",
                "lumiya.xlsx",
                "easycan.xlsx",
                "sutong.xlsx",
                "liulang.xlsx",
                "jiangsheng.xlsx",
            )

        self.assertEqual(register_filepath, "liulang.xlsx")
        self.assertEqual(sheet_name, "2026（内部开票）")
        self.assertEqual(register_df.loc[0, "报关方式（9710/一般贸易）"], 9810)


if __name__ == "__main__":
    unittest.main()
