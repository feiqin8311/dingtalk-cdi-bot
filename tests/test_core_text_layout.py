import unittest

from app.core.text_layout import calculate_wrapped_text_height


class TextLayoutTests(unittest.TestCase):
    def test_reserves_height_for_multiline_address(self):
        address = "\n".join(["address line"] * 7)

        self.assertEqual(calculate_wrapped_text_height(address, 90, 75), 115)

    def test_counts_long_lines_that_wrap_in_the_cell(self):
        self.assertEqual(calculate_wrapped_text_height("x" * 61, 45, 30), 55)
