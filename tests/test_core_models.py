import unittest

from app.core.models import select_register_filepath


class RegisterTargetSelectionTests(unittest.TestCase):
    def test_select_register_filepath_maps_known_company_indexes(self):
        self.assertEqual(
            select_register_filepath(0, "lumiya.xlsx", "easycan.xlsx", "sutong.xlsx", "liulang.xlsx", "jiangsheng.xlsx"),
            "easycan.xlsx",
        )
        self.assertEqual(
            select_register_filepath(1, "lumiya.xlsx", "easycan.xlsx", "sutong.xlsx", "liulang.xlsx", "jiangsheng.xlsx"),
            "lumiya.xlsx",
        )
        self.assertEqual(
            select_register_filepath(2, "lumiya.xlsx", "easycan.xlsx", "sutong.xlsx", "liulang.xlsx", "jiangsheng.xlsx"),
            "sutong.xlsx",
        )
        self.assertEqual(
            select_register_filepath(3, "lumiya.xlsx", "easycan.xlsx", "sutong.xlsx", "liulang.xlsx", "jiangsheng.xlsx"),
            "liulang.xlsx",
        )
        self.assertEqual(
            select_register_filepath(4, "lumiya.xlsx", "easycan.xlsx", "sutong.xlsx", "liulang.xlsx", "jiangsheng.xlsx"),
            "jiangsheng.xlsx",
        )

    def test_select_register_filepath_falls_back_to_sutong_for_unknown_index(self):
        self.assertEqual(
            select_register_filepath(99, "lumiya.xlsx", "easycan.xlsx", "sutong.xlsx", "liulang.xlsx", "jiangsheng.xlsx"),
            "sutong.xlsx",
        )


if __name__ == "__main__":
    unittest.main()
