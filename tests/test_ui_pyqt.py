import json
import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import QApplication, QDialogButtonBox

from Ui_PyQt import ProcessingThread, QTInterface
from app.services.ezarc_source_service import ezarc_maintenance_path


class TestUIOption(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        with patch.object(QTInterface, "load_default_settings"), patch("Ui_PyQt.QIcon", return_value=QIcon()):
            self.window = QTInterface()
        self.addCleanup(self.window.close)
        self.window.show()
        self.app.processEvents()

    def test_default_test_mode_is_off_and_available_on_main_window(self):
        self.assertIs(self.window.test_mode_checkbox.window(), self.window)
        self.assertTrue(self.window.test_mode_checkbox.isVisible())
        self.assertFalse(self.window.test_mode_checkbox.isChecked())
        self.assertFalse(self.window.test_mode_warning.isVisible())

    def test_packing_options_are_exclusive_and_default_to_no_consolidation(self):
        consolidate = self.window.consolidation_radio
        do_not_consolidate = self.window.non_consolidation_radio
        self.assertTrue(do_not_consolidate.isChecked())
        self.assertFalse(consolidate.isChecked())

        consolidate.click()
        self.assertTrue(consolidate.isChecked())
        self.assertFalse(do_not_consolidate.isChecked())
        do_not_consolidate.click()
        do_not_consolidate.click()
        self.assertTrue(do_not_consolidate.isChecked())
        self.assertFalse(consolidate.isChecked())

    def test_switching_test_mode_updates_paths_and_warning(self):
        edits = [row.layout().itemAt(1).widget() for row in self.window.register_paths]
        edits[0].setText("custom-production-register.xlsx")
        production_paths = [edit.text() for edit in edits]

        self.window.test_mode_checkbox.click()
        self.assertEqual([edit.text() for edit in edits], self.window.test_register_paths)
        self.assertTrue(self.window.test_mode_warning.isVisible())

        self.window.test_mode_checkbox.click()
        self.assertEqual([edit.text() for edit in edits], production_paths)
        self.assertFalse(self.window.test_mode_warning.isVisible())

    def test_settings_dialogs_do_not_resize_main_window(self):
        start_size = self.window.size()
        for button, dialog in (
            (self.window.upload_btn, self.window.upload_dialog),
            (self.window.settings_btn, self.window.settings_dialog),
        ):
            with self.subTest(title=dialog.windowTitle()):
                button.click()
                self.app.processEvents()
                self.assertTrue(dialog.isVisible())
                self.assertEqual(self.window.size(), start_size)
                dialog.findChild(QDialogButtonBox).button(QDialogButtonBox.StandardButton.Close).click()
                self.assertFalse(dialog.isVisible())
                self.assertEqual(self.window.size(), start_size)

    def test_upload_edits_are_used_after_closing_dialog(self):
        self.window.upload_btn.click()
        row = self.window.file_paths[2]
        with patch("Ui_PyQt.QFileDialog.getOpenFileName", return_value=("products.xlsx", "")) as browse:
            row.layout().itemAt(2).widget().click()
        self.assertIs(browse.call_args.args[0], self.window.upload_dialog)
        self.window.upload_dialog.reject()
        self.window.upload_btn.click()
        self.assertEqual(row.layout().itemAt(1).widget().text(), "products.xlsx")
        self.window.upload_dialog.reject()

    def test_test_mode_allows_separate_register_overrides_and_keeps_maintenance_fixed(self):
        self.window.settings_btn.click()
        edit = self.window.ezarc_path_edit
        self.assertEqual(edit.text(), str(ezarc_maintenance_path()))
        self.assertTrue(edit.isVisible())
        with patch("Ui_PyQt.QFileDialog.getOpenFileName", return_value=("hailian.xlsx", "")) as browse:
            edit.parentWidget().layout().itemAt(2).widget().click()
        self.assertIs(browse.call_args.args[0], self.window.settings_dialog)
        self.assertEqual(browse.call_args.args[1], "选择海联 Excel")
        self.assertEqual(edit.text(), "hailian.xlsx")
        self.window.settings_dialog.reject()

        self.window.test_mode_checkbox.click()
        self.window.settings_btn.click()
        self.assertEqual(edit.text(), str(ezarc_maintenance_path(True)))
        self.assertTrue(edit.isReadOnly())
        self.assertFalse(edit.parentWidget().layout().itemAt(2).widget().isEnabled())
        folder = ezarc_maintenance_path(True).parent
        layout = self.window.register_paths[0].layout()
        register_edit = layout.itemAt(1).widget()
        browse_button = layout.itemAt(2).widget()
        self.assertFalse(register_edit.isReadOnly())
        self.assertTrue(browse_button.isEnabled())
        selected = str(folder / "自定义出运明细.xlsx")
        with patch("Ui_PyQt.QFileDialog.getOpenFileName", return_value=(selected, "")) as browse:
            browse_button.click()
        self.assertEqual(browse.call_args.args[2], str(folder))
        self.assertEqual(register_edit.text(), selected)
        with patch("Ui_PyQt.QFileDialog.getOpenFileName", return_value=(str(folder.parent / "正式.xlsx"), "")), patch(
            "Ui_PyQt.QMessageBox.warning",
        ) as warning:
            browse_button.click()
        warning.assert_called_once()
        self.assertEqual(register_edit.text(), selected)
        self.window.settings_dialog.reject()
        self.window.test_mode_checkbox.click()
        self.assertEqual(register_edit.text(), self.window.prod_register_paths[0])
        self.window.test_mode_checkbox.click()
        self.assertEqual(register_edit.text(), selected)
        with TemporaryDirectory() as temp:
            settings_file = Path(temp) / "default_settings.json"
            self.window.settings_file = settings_file
            with patch("Ui_PyQt.QMessageBox.information"):
                self.window.save_default_settings()
            saved = json.loads(settings_file.read_text(encoding="utf-8"))
            self.assertEqual(saved["register1"], self.window.prod_register_paths[0])
            self.assertEqual(saved["test_register1"], selected)
            self.assertEqual(saved["ezarc_maintenance_path"], "hailian.xlsx")
            saved["ezarc_test_maintenance_path"] = "old-override.xlsx"
            settings_file.write_text(json.dumps(saved, ensure_ascii=False), encoding="utf-8")

            with patch.object(QTInterface, "load_default_settings"), patch("Ui_PyQt.QIcon", return_value=QIcon()):
                reloaded = QTInterface()
            self.addCleanup(reloaded.close)
            reloaded.settings_file = settings_file
            reloaded.load_default_settings()
            self.assertTrue(reloaded.test_mode_checkbox.isChecked())
            self.assertEqual(reloaded.register_path_values()[0], selected)
            self.assertEqual(reloaded.ezarc_path_edit.text(), str(ezarc_maintenance_path(True)))
            reloaded.test_mode_checkbox.setChecked(False)
            self.assertEqual(reloaded.ezarc_path_edit.text(), "hailian.xlsx")
            self.assertFalse(reloaded.ezarc_path_edit.isReadOnly())
            self.assertEqual([row.layout().itemAt(1).widget().text() for row in reloaded.register_paths],
                             self.window.prod_register_paths)

    def test_stop_customs_button_only_appears_while_staging(self):
        thread = ProcessingThread([], 'out', 'pictures', False, True, False, False)
        self.window.processing_thread = thread
        thread.staging_active.connect(self.window.set_customs_stop_available)
        button = self.window.stop_customs_btn
        self.assertTrue(button.isHidden())
        thread._set_staging_active(True)
        self.assertFalse(button.isHidden())
        self.assertTrue(button.isEnabled())
        button.click()
        self.assertTrue(thread._stop_staging.is_set())
        self.assertFalse(button.isEnabled())
        self.assertIn('正在停止', self.window.status_bar.currentMessage())
        thread._set_staging_active(False)
        self.assertTrue(button.isHidden())

    def test_action_bar_and_warning_stay_visible_when_form_scrolls(self):
        self.assertEqual(self.window.form_scroll.verticalScrollBar().maximum(), 0)
        self.window.test_mode_checkbox.setChecked(True)
        self.window.resize(560, 190)
        self.app.processEvents()
        scroll_bar = self.window.form_scroll.verticalScrollBar()
        self.assertGreater(scroll_bar.maximum(), 0)
        central = self.window.centralWidget()
        for widget in (self.window.process_btn, self.window.test_mode_warning):
            self.assertIs(widget.parentWidget(), central)
            self.assertTrue(widget.isVisible())
            self.assertTrue(central.rect().contains(widget.geometry()))
            self.assertGreater(widget.geometry().top(), self.window.form_scroll.geometry().bottom())
        original_position = self.window.process_btn.pos()
        scroll_bar.setValue(scroll_bar.maximum())
        self.assertEqual(self.window.process_btn.pos(), original_position)

    def test_existing_default_settings_format_round_trips(self):
        defaults = {
            "product_info": "products.xlsx", "register1": "custom-register.xlsx",
            "consolidation": "是", "register": "否", "declaration": "否", "test_mode": "是",
        }
        with TemporaryDirectory() as temp:
            settings_file = Path(temp) / "default_settings.json"
            settings_file.write_text(json.dumps(defaults, ensure_ascii=False), encoding="utf-8")
            self.window.settings_file = settings_file
            self.window.load_default_settings()
            self.assertTrue(self.window.consolidation_radio.isChecked())
            self.assertFalse(self.window.non_consolidation_radio.isChecked())
            self.assertTrue(self.window.register_types_row.isHidden())
            self.assertFalse(self.window.declaration_checkbox.isChecked())
            self.assertTrue(self.window.test_mode_warning.isVisible())
            self.assertEqual(self.window.ezarc_path_edit.text(), str(ezarc_maintenance_path(True)))
            self.window.test_mode_checkbox.setChecked(False)
            self.assertEqual(self.window.ezarc_path_edit.text(), str(ezarc_maintenance_path()))
            with patch("Ui_PyQt.QMessageBox.information") as message:
                self.window.save_default_settings()
            message.assert_called_once()
            saved = json.loads(settings_file.read_text(encoding="utf-8"))
            self.assertEqual({key: saved[key] for key in defaults}, {**defaults, "test_mode": "否"})

            self.window.non_consolidation_radio.click()
            with patch("Ui_PyQt.QMessageBox.information"):
                self.window.save_default_settings()
            saved = json.loads(settings_file.read_text(encoding="utf-8"))
            self.assertEqual(saved["consolidation"], "否")
            self.window.consolidation_radio.click()
            self.window.load_default_settings()
            self.assertTrue(self.window.non_consolidation_radio.isChecked())
            self.assertFalse(self.window.consolidation_radio.isChecked())


if __name__ == "__main__":
    unittest.main()
