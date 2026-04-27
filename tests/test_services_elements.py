import unittest
from unittest.mock import patch

import pandas as pd

from app.services.elements_service import find_data_group


class ElementsServiceTests(unittest.TestCase):
    def test_preserves_row_order_and_matches_sheet_specific_groups(self):
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
            ]
        )

        elements_by_sheet = {
            "SHOP-A": pd.DataFrame([["品名：品名A1"], ["HS-A1"], [""]]),
            "SHOP-B": pd.DataFrame([["品名：品名B1"], ["HS-B1"], [""]]),
        }

        def fake_read_excel(_path, sheet_name=None, header=None):
            return elements_by_sheet[sheet_name].copy()

        with patch("app.services.elements_service.pd.read_excel", side_effect=fake_read_excel):
            groups = find_data_group("dummy.xlsx", result_df, cache={})

        self.assertEqual([group[-1] for group in groups], ["SKU-B-1", "SKU-A-1"])


if __name__ == "__main__":
    unittest.main()
