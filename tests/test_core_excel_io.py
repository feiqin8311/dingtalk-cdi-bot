import os
import tempfile
import unittest

import pandas as pd

from app.core.excel_io import ExcelCache


class ExcelCacheTests(unittest.TestCase):
    def test_read_excel_cached_returns_copy(self):
        source = pd.DataFrame({"SKU": ["A1"], "数量": [2]})
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            path = tmp.name
        try:
            source.to_excel(path, index=False, sheet_name="数据")
            cache = ExcelCache()
            df1 = cache.read(path, "数据")
            df2 = cache.read(path, "数据")
            df1.loc[0, "数量"] = 99
            self.assertEqual(df2.loc[0, "数量"], 2)
        finally:
            os.remove(path)


if __name__ == "__main__":
    unittest.main()
