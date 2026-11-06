import os
import unittest
from PyQt6.QtWidgets import QApplication

# 确保 QApplication 初始化
app = QApplication.instance() or QApplication([])

from ui_pyqt import QTInterface


class TestUIOption(unittest.TestCase):
    def setUp(self):
        self.window = QTInterface()

    def tearDown(self):
        self.window.close()

    def test_default_test_mode_is_no(self):
        """测试选项默认为否"""
        self.assertTrue(self.window.test_no.isChecked())
        self.assertFalse(self.window.test_yes.isChecked())

    def test_switching_to_test_mode_updates_register_paths(self):
        """切换为测试模式后，登记设置路径自动切换为 files 目录下的测试文件"""
        # 激活测试模式
        self.window.test_yes.setChecked(True)

        self.assertFalse(self.window.test_no.isChecked())
        self.assertTrue(self.window.test_yes.isChecked())

        # 检查登记设置输入框内容
        for i, path_widget in enumerate(self.window.register_paths):
            current_path = path_widget.layout().itemAt(1).widget().text()
            expected_path = self.window.test_register_paths[i]
            self.assertEqual(current_path, expected_path)
            self.assertTrue(expected_path.endswith(
                ("2026出运明细.xlsx", "2.鲁米亚报关明细.xlsx", "1.易达报关明细.xlsx",
                 "5.素桐报关明细.xlsx", "3.六浪报关明细.xlsx", "4.匠晟报关明细.xlsx")
            ))

    def test_switching_back_from_test_mode_restores_prod_paths(self):
        """切换回正式模式后，登记设置路径恢复为正式路径"""
        # 先切到测试模式
        self.window.test_yes.setChecked(True)
        # 再切回正式模式
        self.window.test_no.setChecked(True)

        self.assertTrue(self.window.test_no.isChecked())
        self.assertFalse(self.window.test_yes.isChecked())

        for i, path_widget in enumerate(self.window.register_paths):
            current_path = path_widget.layout().itemAt(1).widget().text()
            expected_path = self.window.prod_register_paths[i]
            self.assertEqual(current_path, expected_path)


if __name__ == "__main__":
    unittest.main()
