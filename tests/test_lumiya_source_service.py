import unittest

import pandas as pd

from app.services.lumiya_source_service import DEYUN_SUPPLIER_NAME
from app.services.lumiya_source_service import EXTERNAL_INVOICE_SHEET
from app.services.lumiya_source_service import INTERNAL_INVOICE_SHEET
from app.services.lumiya_source_service import NINGBO_SOURCE_CODE
from app.services.lumiya_source_service import apply_lumiya_source_columns
from app.services.lumiya_source_service import attach_lumiya_register_source
from app.services.lumiya_source_service import build_lumiya_external_declaration_register
from app.services.lumiya_source_service import split_lumiya_register_entries
from app.services.lumiya_source_service import update_dominant_source_display


class LumiyaSourceServiceTests(unittest.TestCase):
    def test_deyun_rows_use_ningbo_source_for_lumiya_only(self):
        source_df = pd.DataFrame({
            "供应商名称": [DEYUN_SUPPLIER_NAME, "杭州供应商"],
            "总价": [100, 200],
        })

        lumiya_df = apply_lumiya_source_columns(source_df, 1, "33019 杭州其他")
        easycan_df = apply_lumiya_source_columns(source_df, 0, "33019 杭州其他")

        self.assertEqual(lumiya_df.loc[0, "境内货源地"], NINGBO_SOURCE_CODE)
        self.assertEqual(lumiya_df.loc[1, "境内货源地"], "33019 杭州其他")
        self.assertTrue((easycan_df["境内货源地"] == "33019 杭州其他").all())

    def test_dominant_source_display_uses_larger_customs_amount(self):
        shipping_details = {"domestic_source_of_goods": "33019 杭州其他"}
        result_df = pd.DataFrame({
            "境内货源地显示": ["宁波其他", "杭州其他"],
            "总价": [300, 200],
        })

        update_dominant_source_display(shipping_details, result_df)

        self.assertEqual(shipping_details["domestic_source_display"], "宁波其他")

    def test_lumiya_register_entries_split_deyun_to_external_invoice_sheet(self):
        register_df = pd.DataFrame({
            "报关渠道": ["亚马逊", "亚马逊"],
            "物流商单号": ["INV-001", "INV-001"],
            "出口日期": ["", ""],
            "预录单编号": ["", ""],
            "报关方式（9710/一般贸易）": [9710, 9710],
            "成交方式(FOB/CIF/CFR)": ["CFR", "CFR"],
            "店铺": ["LIBRATON", "LIBRATON"],
            "报关中文品名": ["工具", "工具"],
            "SKU": ["A", "B"],
            "数量": [1, 2],
            "开票/报关品名": ["工具", "工具"],
            "英文品名": ["tool", "tool"],
            "HS CODE": ["82000000", "82000000"],
            "单位": ["个", "个"],
            "申报单价": [10, 20],
            "总价": [10, 40],
            "币别": ["USD", "USD"],
        })
        result_df = pd.DataFrame({
            "境内货源地": [NINGBO_SOURCE_CODE, "33019 杭州其他"],
            "境内货源地显示": ["宁波其他", "杭州其他"],
            "供应商名称": [DEYUN_SUPPLIER_NAME, "杭州供应商"],
            "采购员": ["张三", "李四"],
        })

        source_register_df = attach_lumiya_register_source(register_df, result_df, 1)
        entries = split_lumiya_register_entries(source_register_df, 1, "lumiya.xlsx", "物流商单号")

        sheets = {sheet: df for df, _filepath, sheet in entries}
        self.assertEqual(set(sheets), {INTERNAL_INVOICE_SHEET, EXTERNAL_INVOICE_SHEET})
        self.assertEqual(sheets[INTERNAL_INVOICE_SHEET].loc[0, "SKU"], "B")
        self.assertEqual(sheets[EXTERNAL_INVOICE_SHEET].loc[0, "SKU"], "A")
        self.assertEqual(sheets[EXTERNAL_INVOICE_SHEET].loc[0, "供应商名称"], DEYUN_SUPPLIER_NAME)
        self.assertNotIn("境内货源地", sheets[INTERNAL_INVOICE_SHEET].columns)

    def test_lumiya_external_register_uses_declaration_summary_rows(self):
        register_df = pd.DataFrame({
            "报关渠道": ["亚马逊"],
            "物流商单号": ["26LT291"],
            "报关方式（9710/一般贸易）": [9710],
            "店铺": ["LIBRATON"],
        })
        declaration_df = pd.DataFrame({
            "SKU": ["919301", "919014"],
            "报关中文品名": ["打钉枪", "拉铆枪套装"],
            "英文品名": ["Nail gun", "Rivet gun set"],
            "HS CODE": ["8205590000", "8205590000"],
            "单位": ["套", "套"],
            "发货数量": [42, 248],
            "总价": [379.26, 4791.36],
            "境内货源地": [NINGBO_SOURCE_CODE, NINGBO_SOURCE_CODE],
            "供应商名称": [DEYUN_SUPPLIER_NAME, DEYUN_SUPPLIER_NAME],
        })

        result = build_lumiya_external_declaration_register(
            register_df, declaration_df, 1, "物流商单号", "发货数量", "USD"
        )

        self.assertEqual(result["SKU"].tolist(), ["919301", "919014"])
        self.assertEqual(result["数量"].tolist(), [42, 248])
        self.assertEqual(result["总价"].tolist(), [379.26, 4791.36])
        self.assertEqual(result["申报单价"].tolist(), [9.03, 19.32])


if __name__ == "__main__":
    unittest.main()
