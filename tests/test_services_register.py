import unittest

from openpyxl import Workbook

from app.services.register_service import check_invoice_number_in_sheet, find_last_data_row_ws


class RegisterServiceTests(unittest.TestCase):
    def test_check_invoice_number_raises_when_duplicate_exists(self):
        wb = Workbook()
        ws = wb.active
        ws.append(["发票号"])
        ws.append(["INV-001"])

        with self.assertRaises(Exception):
            check_invoice_number_in_sheet(ws, "INV-001", "发票号")

    def test_find_last_data_row_ws_returns_last_non_blank_row(self):
        wb = Workbook()
        ws = wb.active
        ws.append(["头"])
        ws.append(["A"])
        ws.append([None])

        self.assertEqual(find_last_data_row_ws(ws), 2)


if __name__ == "__main__":
    unittest.main()
