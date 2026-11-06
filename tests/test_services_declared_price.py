import unittest

import pandas as pd

from app.services.declared_price_service import select_declared_price_columns


class DeclaredPriceServiceTests(unittest.TestCase):
    def test_easycan_invoice_uses_usd_column(self):
        source = pd.DataFrame({
            "*SKU": ["A1"],
            "易达-欣耀\n申报单价（CNY）": [12.5],
            "易达-亚马逊\n申报单价(USD)": [3.5],
        })
        result = select_declared_price_columns(
            source,
            sheet_key="易达报关单价",
            company_info_index=0,
            is_awd=False,
            document_type="invoice",
        )
        self.assertEqual(result.loc[0, "申报单价"], 3.5)

    def test_easycan_stock_uses_usd_column(self):
        source = pd.DataFrame({
            "*SKU": ["A1"],
            "易达-欣耀\n申报单价（CNY）": [12.5],
            "易达-亚马逊\n申报单价(USD)": [3.5],
        })
        result = select_declared_price_columns(
            source,
            sheet_key="易达报关单价",
            company_info_index=0,
            is_awd=False,
            document_type="stock",
        )
        self.assertEqual(result.loc[0, "申报单价"], 3.5)

    def test_easycan_awd_uses_usd_column(self):
        source = pd.DataFrame({
            "*SKU": ["A1"],
            "易达-欣耀\n申报单价（CNY）": [12.5],
            "易达-亚马逊\n申报单价(USD)": [3.5],
        })
        result = select_declared_price_columns(
            source,
            sheet_key="易达报关单价",
            company_info_index=0,
            is_awd=True,
            document_type="stock",
        )
        self.assertEqual(result.loc[0, "申报单价"], 3.5)


if __name__ == "__main__":
    unittest.main()
