import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import MagicMock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import QApplication
from Ui_PyQt import ProcessingThread, QTInterface, shared_test_register_paths
from Invoice_Data import CustomsDeclaration as InvoiceDeclaration
from Stock_Preparation import CustomsDeclaration as StockDeclaration
from app.services.ezarc_source_service import ezarc_maintenance_path


class RegisterSelectionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        icon = patch("Ui_PyQt.QIcon", return_value=QIcon())
        icon.start()
        self.addCleanup(icon.stop)

    def test_ui_defaults_and_master_switch(self):
        with patch.object(QTInterface, "load_default_settings"):
            window = QTInterface()
        try:
            self.assertTrue(window.register_shipping.isChecked())
            self.assertTrue(window.register_customs.isChecked())
            window.register_shipping.setChecked(False)
            window.register_checkbox.click()
            self.assertTrue(window.register_types_row.isHidden())
            window.register_checkbox.click()
            self.assertFalse(window.register_types_row.isHidden())
            self.assertFalse(window.register_shipping.isChecked())
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
                    maintenance_path = "hailian-production.xlsx" if shipping else ""
                    window.ezarc_path_edit.setText(maintenance_path)
                    with patch("Ui_PyQt.ProcessingThread") as thread, patch("Ui_PyQt.QMessageBox.critical") as error:
                        window.process_data()
                    error.assert_not_called()
                    self.assertEqual(len(thread.call_args.args[0]), 12)
                    self.assertEqual(thread.call_args.kwargs, {
                        "register_shipping": shipping, "register_customs": customs, "test_mode": False,
                        "maintenance_path": maintenance_path,
                    })
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
                    kwargs = generator.flow.register_all_documents.call_args.kwargs.copy()
                    if not use_stock:
                        reports = []
                        thread.progress.connect(reports.append)
                        kwargs.pop('progress')('登记进度')
                        self.assertEqual(reports, ['登记进度'])
                        self.assertFalse(thread.request_staging_stop())
                        set_active = kwargs.pop('staging_state')
                        is_cancelled = kwargs.pop('staging_cancelled')
                        set_active(True)
                        self.assertFalse(is_cancelled())
                        self.assertTrue(thread.request_staging_stop())
                        self.assertTrue(is_cancelled())
                        set_active(False)
                        self.assertFalse(thread.request_staging_stop())
                    self.assertEqual(kwargs, {"register_shipping": shipping, "register_customs": customs})

    def test_worker_routes_to_shared_tests_without_requiring_other_company_files(self):
        shared_paths = shared_test_register_paths()
        self.assertEqual(Path(shared_paths[3]).name, "5.素桐报关明细.xlsx")
        shared_paths[2] = str(ezarc_maintenance_path(True).parent / "易达-另一个测试表.xlsx")
        paths = [f"path{i}" for i in range(6)] + shared_paths
        for use_stock in (False, True):
            with self.subTest(use_stock=use_stock):
                target = "stock_preparation" if use_stock else "invoice_data"
                thread = ProcessingThread(paths, "out", "pictures", False, True, False, use_stock,
                                          test_mode=True, maintenance_path="formal-maintenance.xlsx")
                with patch(f"Ui_PyQt.{target}.CustomsDeclaration") as constructor, patch(
                    "Ui_PyQt.os.path.exists", side_effect=lambda path: path in paths[:6] + shared_paths[:1],
                ) as exists:
                    thread.run()
                self.assertEqual([call.args[0] for call in exists.call_args_list], paths[:6] + shared_paths[:1])
                flow = constructor.return_value.flow
                flow.generate_all_documents.assert_called_once()
                self.assertEqual(flow.generate_all_documents.call_args.args[5], "out")
                args = flow.register_all_documents.call_args.args
                self.assertEqual([args[1], *args[3:8]], shared_paths)

    def test_worker_stops_on_missing_test_register_without_using_formal_paths(self):
        shared_paths = shared_test_register_paths()
        shared_shipping = shared_paths[0]
        paths = ["invoice.xlsx"] + [""] * 5 + shared_paths
        thread = ProcessingThread(paths, "out", "pictures", False, True, False, False,
                                  register_customs=False, test_mode=True)
        results = []
        thread.finished.connect(lambda success, message: results.append((success, message)))
        with patch("Ui_PyQt.invoice_data.CustomsDeclaration") as invoice, patch(
            "Ui_PyQt.os.path.exists", side_effect=lambda path: path != shared_shipping,
        ) as exists:
            thread.run()
        invoice.assert_not_called()
        self.assertEqual([call.args[0] for call in exists.call_args_list], ["invoice.xlsx", shared_shipping])
        self.assertEqual(results, [(False, f"文件不存在或无法访问：{shared_shipping}")])

    def test_worker_refuses_outside_test_paths_before_processing(self):
        for path in (r"C:\本机\易达.xlsx", str(ezarc_maintenance_path()),
                     str(ezarc_maintenance_path(True).parent / ".." / "正式.xlsx")):
            with self.subTest(path=path):
                inputs = ["invoice.xlsx"] + [""] * 5 + shared_test_register_paths()
                inputs[8] = path
                thread = ProcessingThread(inputs, "out", "pictures", False, True, False, False,
                                          register_shipping=False, test_mode=True)
                results = []
                thread.finished.connect(lambda success, message: results.append((success, message)))
                with patch("Ui_PyQt.invoice_data.CustomsDeclaration") as invoice:
                    thread.run()
                invoice.assert_not_called()
                self.assertEqual(len(results), 1)
                self.assertFalse(results[0][0])
                self.assertIn("测试登记文件必须是以下目录内", results[0][1])
                self.assertIn(path, results[0][1])

    def test_worker_passes_maintenance_path_only_to_invoice_flow(self):
        for test_mode, custom_path in ((False, "hailian.xlsx"), (True, "test-copy.xlsx"), (True, "")):
            with self.subTest(test_mode=test_mode, custom_path=custom_path):
                thread = ProcessingThread(
                    ["invoice.xlsx"] + [""] * 11, "out", "pictures", False, False, False, False,
                    test_mode=test_mode, maintenance_path=custom_path,
                )
                with patch("Ui_PyQt.invoice_data.CustomsDeclaration") as invoice, patch(
                    "Ui_PyQt.os.path.exists", side_effect=lambda path: path == "invoice.xlsx",
                ):
                    thread.run()
                expected = ezarc_maintenance_path(True) if test_mode else custom_path
                invoice.assert_called_once_with(maintenance_path=expected)
                invoice.return_value.flow.generate_all_documents.assert_called_once()
                invoice.return_value.flow.register_all_documents.assert_not_called()

        stock_thread = ProcessingThread(
            ["", "stock.xlsx"] + [""] * 10, "out", "pictures", False, False, False, True,
            maintenance_path="unused.xlsx",
        )
        with patch("Ui_PyQt.stock_preparation.CustomsDeclaration") as stock, patch(
            "Ui_PyQt.os.path.exists", side_effect=lambda path: path == "stock.xlsx",
        ):
            stock_thread.run()
        stock.assert_called_once_with()
        stock.return_value.flow.generate_all_documents.assert_called_once()

    def test_registration_still_requires_the_actual_customs_workbook(self):
        for cls in (InvoiceDeclaration, StockDeclaration):
            with self.subTest(cls=cls.__name__), TemporaryDirectory() as temp:
                target = str(Path(temp) / "missing-customs.xlsx")
                generator = cls()
                entry = ([1], target, "sheet", "EZARC")
                if cls is InvoiceDeclaration:
                    entry += (1,)
                with patch.object(generator, "build_invoicing_register_entries", return_value=[entry]), patch.object(
                    generator, "register_shipping_invoice_number",
                ) as register_shipping, patch(f"{cls.__module__}.append_call_record"):
                    with self.assertRaisesRegex(Exception, "报关开票登记表不存在") as caught:
                        generator.register_all_documents("invoice", "shipping", "product", "a", "b", "c", "d", "e")
                self.assertIn(target, str(caught.exception))
                register_shipping.assert_not_called()

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
                    self.assertEqual('出运明细' in checked_sheets, shipping and cls is StockDeclaration)
                    self.assertEqual('sheet' in checked_sheets, customs)


if __name__ == "__main__":
    unittest.main()
