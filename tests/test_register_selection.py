import os
import unittest
from unittest.mock import MagicMock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt6.QtWidgets import QApplication
from Ui_PyQt import ProcessingThread, QTInterface
from Invoice_Data import CustomsDeclaration as InvoiceDeclaration
from Stock_Preparation import CustomsDeclaration as StockDeclaration


class RegisterSelectionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_ui_defaults_and_master_switch(self):
        with patch.object(QTInterface, "load_default_settings"):
            window = QTInterface()
        try:
            self.assertTrue(window.register_shipping.isChecked())
            self.assertTrue(window.register_customs.isChecked())
            window.register_no.setChecked(True)
            self.assertFalse(window.register_shipping.isEnabled())
            self.assertFalse(window.register_customs.isEnabled())
            window.register_yes.setChecked(True)
            self.assertTrue(window.register_shipping.isEnabled())
            self.assertTrue(window.register_customs.isEnabled())
            self.assertTrue(window.register_shipping.isChecked())
            self.assertTrue(window.register_customs.isChecked())
        finally:
            window.close()

    def test_at_least_one_selection_required(self):
        with patch.object(QTInterface, "load_default_settings"):
            window = QTInterface()
        try:
            window.register_shipping.setChecked(False)
            window.register_customs.setChecked(False)
            with patch("Ui_PyQt.QMessageBox.critical") as message:
                window.process_data()
            message.assert_called_once()
            self.assertIn("至少一项", message.call_args.args[2])
        finally:
            window.close()

    def test_ui_requires_only_selected_register_paths(self):
        for shipping, customs in ((True, False), (False, True)):
            with self.subTest(shipping=shipping, customs=customs), patch.object(QTInterface, "load_default_settings"):
                window = QTInterface()
                try:
                    window.register_shipping.setChecked(shipping)
                    window.register_customs.setChecked(customs)
                    for index, widget in enumerate(window.file_paths):
                        widget.layout().itemAt(1).widget().setText("" if index == 1 else f"base{index}")
                    window.output_path_widget.layout().itemAt(1).widget().setText("output")
                    for index, widget in enumerate(window.register_paths):
                        required = shipping if index == 0 else customs
                        widget.layout().itemAt(1).widget().setText(f"register{index}" if required else "")
                    with patch("Ui_PyQt.ProcessingThread") as thread, patch("Ui_PyQt.QMessageBox.critical") as error:
                        window.process_data()
                    error.assert_not_called()
                    self.assertEqual(thread.call_args.kwargs,
                                     {"register_shipping": shipping, "register_customs": customs})
                    thread.return_value.start.assert_called_once()
                finally:
                    window.close()

    def test_worker_checks_only_selected_register_paths(self):
        for use_stock in (False, True):
            for shipping, customs in ((True, False), (False, True), (True, True)):
                with self.subTest(use_stock=use_stock, shipping=shipping, customs=customs):
                    paths = [f"path{i}" for i in range(12)]
                    generator = MagicMock()
                    target = "Ui_PyQt.stock_preparation.CustomsDeclaration" if use_stock else "Ui_PyQt.invoice_data.CustomsDeclaration"
                    thread = ProcessingThread(paths, "out", "pictures", False, True, False, use_stock,
                                              register_shipping=shipping, register_customs=customs)
                    with patch(target, return_value=generator), patch("Ui_PyQt.os.path.exists", return_value=True) as exists:
                        thread.run()
                    checked = [call.args[0] for call in exists.call_args_list]
                    self.assertEqual("path6" in checked, shipping)
                    self.assertEqual("path7" in checked, customs)
                    kwargs = generator.flow.register_all_documents.call_args.kwargs
                    self.assertEqual(kwargs, {"register_shipping": shipping, "register_customs": customs})

    def test_business_registration_types_are_independent(self):
        for cls in (InvoiceDeclaration, StockDeclaration):
            for shipping, customs in ((True, False), (False, True), (True, True), (False, False)):
                with self.subTest(cls=cls.__name__, shipping=shipping, customs=customs):
                    generator = cls.__new__(cls)
                    generator.shipping_details = {"invoice_number": "INV"}
                    generator.result_df = object()
                    entry = ([1], "customs", "sheet", "brand")
                    if cls is InvoiceDeclaration:
                        entry += (1,)
                    generator.build_invoicing_register_entries = MagicMock(return_value=[entry])
                    generator.ensure_file_writable = MagicMock()
                    generator.check_invoice_number = MagicMock(return_value=0)
                    generator.write_invoicing_register = MagicMock(return_value=6)
                    generator.find_last_data_row = MagicMock(return_value=5)
                    generator.register_shipping_invoice_number = MagicMock(
                        return_value=(0, 6) if cls is StockDeclaration else (0, 6, 1)
                    )
                    generator.register_combined_invoice = MagicMock(return_value=[1])
                    with patch(f"{cls.__module__}.append_call_record"):
                        generator.register_all_documents("invoice", "shipping", "product", "a", "b", "c", "d", "e",
                                                         register_shipping=shipping, register_customs=customs)
                    self.assertEqual(generator.register_shipping_invoice_number.called, shipping)
                    self.assertEqual(generator.build_invoicing_register_entries.called, customs)
                    self.assertEqual(generator.ensure_file_writable.called, customs)
                    self.assertEqual(generator.write_invoicing_register.called, customs)
                    checked_sheets = [call.args[1] for call in generator.find_last_data_row.call_args_list]
                    self.assertEqual('出运明细' in checked_sheets, shipping)
                    self.assertEqual('sheet' in checked_sheets, customs)


if __name__ == "__main__":
    unittest.main()
