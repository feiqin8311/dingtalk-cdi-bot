import os
import sys

from app.core.workbook_staging import STAGING_FLAG, run_staging_worker

if __name__ == '__main__' and sys.argv[1:2] == [STAGING_FLAG]:
    sys.exit(run_staging_worker(sys.argv[2:]))

import json
from threading import Event, Lock
import pythoncom
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QFileDialog, QDialog, QDialogButtonBox,
    QFrame, QScrollArea, QStatusBar, QMessageBox, QCheckBox, QRadioButton, QButtonGroup
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QIcon
import Invoice_Data as invoice_data
import Stock_Preparation as stock_preparation
from app.services.ezarc_source_service import ezarc_maintenance_path


EZARC_MAINTENANCE_LABEL = "海联 Excel"
EZARC_PATH_SETTING = "ezarc_maintenance_path"
REGISTER_FILENAMES = (
    "2026出运明细.xlsx", "2.鲁米亚报关明细.xlsx", "1.易达报关明细.xlsx",
    "5.素桐报关明细.xlsx", "3.六浪报关明细.xlsx", "4.匠晟报关明细.xlsx",
)


def shared_test_register_paths():
    folder = ezarc_maintenance_path(True).parent
    return [str(folder / name) for name in REGISTER_FILENAMES]


def normalize_test_register_path(path):
    folder = str(ezarc_maintenance_path(True).parent)
    parent, filename = os.path.split(os.path.normpath(path.strip()))
    if (os.path.normcase(parent) != os.path.normcase(folder)
            or ':' in filename or os.path.splitext(filename)[1].lower() != '.xlsx'):
        raise ValueError(f"测试登记文件必须是以下目录内的 .xlsx 文件：\n{folder}\n所选文件：{path}")
    return os.path.join(folder, filename)


class ProcessingThread(QThread):
    """处理线程，避免界面冻结"""
    finished = pyqtSignal(bool, str)  # 信号：处理完成(成功, 消息)
    progress = pyqtSignal(str)
    staging_active = pyqtSignal(bool)

    def __init__(self, input_paths, output_path, pictures_path, is_consolidation, is_register,
                 is_declaration, use_stock, register_shipping=True, register_customs=True, test_mode=False,
                 *, maintenance_path=None):
        super().__init__()
        self.input_paths = list(input_paths)
        self.output_path = output_path
        self.pictures_path = pictures_path
        self.is_consolidation = is_consolidation
        self.is_register = is_register
        self.is_declaration = is_declaration
        self.use_stock = use_stock
        self.register_shipping = register_shipping
        self.register_customs = register_customs
        self.test_mode = test_mode
        self.maintenance_path = (
            ezarc_maintenance_path(True) if test_mode else maintenance_path or ezarc_maintenance_path()
        )
        self._staging_lock = Lock()
        self._staging_running = False
        self._stop_staging = Event()

    def _set_staging_active(self, active):
        with self._staging_lock:
            self._staging_running = active
        self.staging_active.emit(active)

    def request_staging_stop(self):
        with self._staging_lock:
            if not self._staging_running:
                return False
            self._stop_staging.set()
            return True

    def run(self):
        """执行处理任务"""
        pythoncom.CoInitialize()
        try:
            if self.test_mode and self.is_register:
                for index in range(6, 12):
                    selected = self.register_shipping if index == 6 else self.register_customs
                    if selected:
                        self.input_paths[index] = normalize_test_register_path(self.input_paths[index])
            # 检查文件是否存在
            paths_to_check = self.input_paths[:6]
            if self.is_register and self.register_shipping:
                paths_to_check += self.input_paths[6:7]
            # 测试目录不必备齐所有公司表，由登记流程检查本票实际使用的表。
            if self.is_register and self.register_customs and not self.test_mode:
                paths_to_check += self.input_paths[7:12]
            for path in filter(None, paths_to_check):
                if not os.path.exists(path):
                    detail = f"文件不存在或无法访问：{path}" if self.test_mode else f"文件不存在: {os.path.basename(path)}"
                    self.finished.emit(False, detail)
                    return

            # 初始化CustomsDeclaration()类
            if self.use_stock:
                generator = stock_preparation.CustomsDeclaration()
                generator.flow.generate_all_documents(
                    self.input_paths[1],  # 备货单
                    self.input_paths[2],  # 产品信息单证专用
                    self.input_paths[3],  # 品名
                    self.input_paths[4],  # 全站点地址
                    self.input_paths[5],  # 要素表
                    self.output_path,
                    self.pictures_path, self.is_consolidation, self.is_declaration
                )
                if self.is_register:
                    generator.flow.register_all_documents(
                        self.input_paths[1],   # 备货单
                        self.input_paths[6],   # 出运明细
                        self.input_paths[2],   # 产品信息单证专用
                        self.input_paths[7],   # 鲁米亚报关明细
                        self.input_paths[8],   # 易达报关明细
                        self.input_paths[9],   # 素桐报关开票明细
                        self.input_paths[10],  # 六浪报关明细
                        self.input_paths[11],  # 匠晟报关明细
                        self.input_paths[3],   # 品名
                        self.input_paths[4],   # 全站点地址
                        self.is_consolidation,
                        register_shipping=self.register_shipping,
                        register_customs=self.register_customs,
                    )
            else:
                generator = invoice_data.CustomsDeclaration(maintenance_path=self.maintenance_path)
                generator.flow.generate_all_documents(
                    self.input_paths[0],  # 发货单
                    self.input_paths[2],  # 产品信息单证专用
                    self.input_paths[3],  # 品名
                    self.input_paths[4],  # 全站点地址
                    self.input_paths[5],  # 要素表
                    self.output_path,
                    self.pictures_path, self.is_consolidation, self.is_declaration
                )

                # 登记所有文档（可选）
                if self.is_register:
                    generator.flow.register_all_documents(
                        self.input_paths[0],   # 发货单
                        self.input_paths[6],   # 出运明细
                        self.input_paths[2],   # 产品信息单证专用
                        self.input_paths[7],   # 鲁米亚报关明细
                        self.input_paths[8],   # 易达报关明细
                        self.input_paths[9],   # 素桐报关开票明细
                        self.input_paths[10],  # 六浪报关明细
                        self.input_paths[11],  # 匠晟报关明细
                        register_shipping=self.register_shipping,
                        register_customs=self.register_customs,
                        progress=self.progress.emit,
                        staging_state=self._set_staging_active,
                        staging_cancelled=self._stop_staging.is_set,
                    )

            # 处理成功
            self.finished.emit(True, "报关资料生成成功！")

        except Exception as e:
            self.finished.emit(False, f"处理过程中发生错误: {str(e)}")
        finally:
            pythoncom.CoUninitialize()


class ModernButton(QPushButton):
    """现代化样式按钮"""

    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setMinimumHeight(35)

    def setprimary(self):
        """设置主要按钮样式"""
        self.setStyleSheet("""
            QPushButton {
                background-color: #6200ee;
                color: white;
                border: 1px solid transparent;
                border-radius: 5px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:focus {
                border-color: #38106a;
            }
            QPushButton:hover {
                background-color: #7c3aed;
            }
            QPushButton:pressed {
                background-color: #512da8;
            }
            QPushButton:disabled {
                background-color: #9e9e9e;
            }
        """)

    def setSecondary(self):
        """设置次要按钮样式"""
        self.setStyleSheet("""
            QPushButton {
                background-color: #f5f5f5;
                color: #333;
                border: 1px solid #ddd;
                border-radius: 5px;
                padding: 8px 16px;
            }
            QPushButton:focus {
                border-color: #38106a;
            }
            QPushButton:hover {
                background-color: #e8e8e8;
            }
            QPushButton:pressed {
                background-color: #dcdcdc;
            }
        """)


class QTInterface(QMainWindow):
    """报关资料生成工具主窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("报关资料生成工具")
        self.resize(760, 360)
        self.setMinimumSize(560, 190)

        # 设置窗口图标
        self.setWindowIcon(QIcon(r'\\192.168.0.45\供应链管理\2 物流发货管理\17.单证数据表维护\pictures\tool.ico'))

        # 初始化默认设置文件路径
        # self.settings_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "default_settings.json")
        self.settings_file = os.path.join(os.path.dirname(sys.executable), "default_settings.json")
        # 路径标签与配置键
        self.base_file_labels = ["发货单路径", "备货单路径", "产品信息单证专用路径", "品名", "全站点地址", "要素表"]
        self.register_file_labels = ["出运明细", "鲁米亚报关明细", "易达报关明细", "素桐报关开票明细", "六浪报关明细", "匠晟报关明细"]
        self.file_dialog_labels = (
            [label.replace("路径", "") for label in self.base_file_labels]
            + self.register_file_labels + [EZARC_MAINTENANCE_LABEL]
        )
        self.file_keys = ["delivery", "stocking", "product_info", "product_name", "full_site_address", "elements"]
        self.register_keys = ["register1", "register2", "register3", "register4", "register5", "register6"]
        # 初始化默认路径
        self.init_paths()

        # 设置应用样式
        self.setup_styles()

        # 创建中央部件和主布局
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setSpacing(10)
        self.main_layout.setContentsMargins(18, 14, 18, 10)

        # 设置UI
        self.setup_ui()

        # 加载默认设置
        self.load_default_settings()

        # 初始化处理线程
        self.processing_thread = None

    def init_paths(self):
        """初始化默认路径"""
        # 设置默认输出路径为桌面
        self.default_output = os.path.join(os.path.expanduser("~"), "Desktop")

        # 获取所在文件夹路径
        self.folder_path = r"\\192.168.0.45\供应链管理\2 物流发货管理\17.单证数据表维护"
        self.product_info_filepath = os.path.join(self.folder_path, "产品信息单证专用.xlsx")
        self.product_name_filepath = os.path.join(self.folder_path, "品名.xlsx")
        self.full_site_address_filepath = os.path.join(self.folder_path, "全站点地址.xlsx")
        self.elements_filepath = os.path.join(self.folder_path, "要素.xlsx")
        self.stocking_filepath = ""
        self.pictures_path = os.path.join(self.folder_path, "pictures")

        # 测试专用
        # self.invoice_number_register_filepath = r"\\192.168.0.45\供应链管理\2 物流发货管理\17.单证数据表维护\测试\2025出运明细2025.9.19.xlsx"
        # self.register_folder_path = r"\\192.168.0.45\供应链管理\2 物流发货管理\17.单证数据表维护\测试"

        # 登记设置相关文件路径
        self.invoice_number_register_filepath = r"\\192.168.0.45\供应链管理\2 物流发货管理\2026出运明细.xlsx"
        self.register_folder_path = r"\\192.168.0.45\供应链管理\2 物流发货管理\1 报关开票登记"
        self.lumiya_register_filepath = os.path.join(self.register_folder_path, "2.鲁米亚报关明细.xlsx")
        self.easycan_register_filepath = os.path.join(self.register_folder_path, "1.易达报关明细.xlsx")
        self.sutong_register_filepath = os.path.join(self.register_folder_path, "5.素桐报关明细.xlsx")
        self.liulang_register_filepath = os.path.join(self.register_folder_path, "3.六浪报关明细.xlsx")
        self.jiangsheng_register_filepath = os.path.join(self.register_folder_path, "4.匠晟报关明细.xlsx")

        self.prod_register_paths = [
            self.invoice_number_register_filepath,
            self.lumiya_register_filepath,
            self.easycan_register_filepath,
            self.sutong_register_filepath,
            self.liulang_register_filepath,
            self.jiangsheng_register_filepath
        ]
        self.last_prod_register_paths = list(self.prod_register_paths)
        self.ezarc_prod_path = str(ezarc_maintenance_path())
        self.test_register_paths = shared_test_register_paths()

    def setup_styles(self):
        """设置应用样式"""
        self.setStyleSheet("""
            QMainWindow, QDialog, QWidget#formContent {
                background-color: #fafafa;
            }
            QLabel[sectionHeading="true"] {
                color: #424242;
                font-weight: bold;
            }
            QLabel#testModeWarning {
                color: #78350f;
                background-color: #fef3c7;
                border-radius: 4px;
                padding: 5px 8px;
                font-weight: bold;
            }
            QFrame#actionSeparator {
                color: #ddd;
            }
            QLineEdit {
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 8px;
                background-color: white;
                selection-background-color: #bb86fc;
            }
            QLineEdit:focus {
                border: 2px solid #6200ee;
            }
            QLabel {
                color: #424242;
            }
            QStatusBar {
                background-color: #f0f0f0;
                color: #666666;
            }
            QCheckBox, QRadioButton {
                spacing: 5px;
            }
        """)

    def setup_ui(self):
        """主窗口只展示生成资料所需的常用操作。"""
        form = QWidget()
        form_layout = QVBoxLayout(form)
        form_layout.setContentsMargins(0, 0, 4, 0)
        form_layout.setSpacing(6)
        form_layout.addWidget(self.section_heading("文件路径"))

        self.file_paths = [
            self.create_file_widget(self.file_dialog_labels[0], "", 0),
            self.create_file_widget(self.file_dialog_labels[1], self.stocking_filepath, 1),
        ]
        for path_widget in self.file_paths:
            form_layout.addWidget(path_widget)
        self.output_path_widget = self.create_file_widget("输出目录", self.default_output, -1, is_output=True)
        form_layout.addWidget(self.output_path_widget)

        form_layout.addSpacing(6)
        form_layout.addWidget(self.section_heading("生成选项"))
        options_layout = QHBoxLayout()
        options_layout.setSpacing(20)
        packing_layout = QHBoxLayout()
        packing_layout.setSpacing(10)
        self.consolidation_radio = QRadioButton("拼箱")
        self.non_consolidation_radio = QRadioButton("不拼箱")
        packing_group = QButtonGroup(self)
        for option in (self.consolidation_radio, self.non_consolidation_radio):
            packing_group.addButton(option)
            packing_layout.addWidget(option)
        self.non_consolidation_radio.setChecked(True)
        options_layout.addLayout(packing_layout)

        self.declaration_checkbox = QCheckBox("生成申报文件")
        self.declaration_checkbox.setChecked(True)
        options_layout.addWidget(self.declaration_checkbox)
        self.register_checkbox = QCheckBox("登记")
        self.register_checkbox.setChecked(True)
        options_layout.addWidget(self.register_checkbox)
        options_layout.addStretch()
        self.test_mode_checkbox = QCheckBox("测试模式")
        options_layout.addWidget(self.test_mode_checkbox)
        form_layout.addLayout(options_layout)

        self.register_types_row = QWidget()
        register_layout = QHBoxLayout(self.register_types_row)
        register_layout.setContentsMargins(0, 0, 0, 0)
        register_layout.setSpacing(12)
        register_layout.addWidget(QLabel("登记类型："))
        self.register_shipping = QCheckBox("出运明细")
        self.register_customs = QCheckBox("报关明细")
        for checkbox in (self.register_shipping, self.register_customs):
            checkbox.setChecked(True)
            register_layout.addWidget(checkbox)
        register_layout.addStretch()
        form_layout.addWidget(self.register_types_row)
        self.register_checkbox.toggled.connect(self.register_types_row.setVisible)

        form_layout.addStretch()
        self.form_scroll = self.scroll_area_for(form)
        self.main_layout.addWidget(self.form_scroll, 1)

        self.test_mode_warning = QLabel("测试模式已开启")
        self.test_mode_warning.setObjectName("testModeWarning")
        self.test_mode_warning.hide()
        self.main_layout.addWidget(self.test_mode_warning)
        self.setup_settings_dialogs()
        self.setup_action_bar()

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就绪")
        self.test_mode_checkbox.toggled.connect(self.on_test_mode_toggled)

    @staticmethod
    def section_heading(text):
        heading = QLabel(text)
        heading.setProperty("sectionHeading", True)
        return heading

    @staticmethod
    def scroll_area_for(content):
        content.setObjectName("formContent")
        scroll_area = QScrollArea()
        scroll_area.setFrameShape(QFrame.Shape.NoFrame)
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(content)
        return scroll_area

    def create_settings_dialog(self, title):
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(16, 14, 16, 12)

        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(0, 0, 4, 0)
        content_layout.setSpacing(6)
        layout.addWidget(self.scroll_area_for(content), 1)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.button(QDialogButtonBox.StandardButton.Close).setText("关闭")
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        return dialog, content_layout

    def setup_settings_dialogs(self):
        self.upload_dialog, upload_layout = self.create_settings_dialog("文件上传")
        self.upload_dialog.resize(720, 260)
        upload_layout.addWidget(self.section_heading("资料文件路径"))
        default_paths = (
            self.product_info_filepath, self.product_name_filepath,
            self.full_site_address_filepath, self.elements_filepath,
        )
        for i, (label, default_path) in enumerate(zip(self.base_file_labels[2:], default_paths), start=2):
            path_widget = self.create_file_widget(label, default_path, i)
            self.file_paths.append(path_widget)
            upload_layout.addWidget(path_widget)
        upload_layout.addStretch()

        self.settings_dialog, settings_layout = self.create_settings_dialog("更多设置")
        self.settings_dialog.resize(720, 390)
        settings_layout.addWidget(self.section_heading("登记设置"))
        self.register_paths = []
        for i, (label, default_path) in enumerate(zip(self.register_file_labels, self.prod_register_paths), start=6):
            path_widget = self.create_file_widget(label, default_path, i)
            self.register_paths.append(path_widget)
            settings_layout.addWidget(path_widget)

        ezarc_row = self.create_file_widget(
            EZARC_MAINTENANCE_LABEL, self.ezarc_prod_path, len(self.file_dialog_labels) - 1,
        )
        self.ezarc_path_edit = ezarc_row.layout().itemAt(1).widget()
        self.ezarc_path_edit.setToolTip("用于 EZARC 发货单的货源地拆分和扣减回写。测试模式固定使用共享测试目录。")
        settings_layout.addWidget(ezarc_row)
        settings_layout.addStretch()

    def setup_action_bar(self):
        separator = QFrame()
        separator.setObjectName("actionSeparator")
        separator.setFrameShape(QFrame.Shape.HLine)
        self.main_layout.addWidget(separator)

        button_layout = QHBoxLayout()
        button_layout.setSpacing(8)
        self.upload_btn = ModernButton("文件上传…")
        self.upload_btn.setSecondary()
        self.upload_btn.clicked.connect(self.upload_dialog.open)
        button_layout.addWidget(self.upload_btn)

        self.settings_btn = ModernButton("更多设置…")
        self.settings_btn.setSecondary()
        self.settings_btn.clicked.connect(self.settings_dialog.open)
        button_layout.addWidget(self.settings_btn)
        button_layout.addStretch()

        self.save_defaults_btn = ModernButton("保存默认设置")
        self.save_defaults_btn.setSecondary()
        self.save_defaults_btn.clicked.connect(self.save_default_settings)
        button_layout.addWidget(self.save_defaults_btn)

        self.stop_customs_btn = ModernButton("停止开票")
        self.stop_customs_btn.setSecondary()
        self.stop_customs_btn.hide()
        self.stop_customs_btn.clicked.connect(self.stop_customs_staging)
        button_layout.addWidget(self.stop_customs_btn)

        self.process_btn = ModernButton("生成报关资料")
        self.process_btn.setprimary()
        self.process_btn.clicked.connect(self.process_data)
        button_layout.addWidget(self.process_btn)
        self.main_layout.addLayout(button_layout)

    def on_test_mode_toggled(self, checked):
        """切换登记路径，并在主窗口持续标记测试模式。"""
        if checked:
            self.remember_production_paths()
        else:
            self.test_register_paths = self.register_path_values()
        self.apply_mode_paths(checked)
        self.status_bar.showMessage("已开启测试模式：登记使用共享测试目录" if checked else "已关闭测试模式：登记使用正式路径")

    def register_path_values(self):
        return [row.layout().itemAt(1).widget().text().strip() for row in self.register_paths]

    def remember_production_paths(self):
        self.last_prod_register_paths = self.register_path_values()
        self.ezarc_prod_path = self.ezarc_path_edit.text().strip()

    def apply_mode_paths(self, test_mode):
        self.test_mode_warning.setVisible(test_mode)
        paths = self.test_register_paths if test_mode else self.last_prod_register_paths
        for row, path in zip(self.register_paths, paths):
            row.layout().itemAt(1).widget().setText(path)
        self.ezarc_path_edit.setText(str(ezarc_maintenance_path(True)) if test_mode else self.ezarc_prod_path)
        self.ezarc_path_edit.setReadOnly(test_mode)
        self.ezarc_path_edit.parentWidget().layout().itemAt(2).widget().setEnabled(not test_mode)

    def create_file_widget(self, label_text, default_path, index, is_output=False):
        """创建文件路径选择控件"""
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        label = QLabel(f"{label_text}:")
        label.setFont(QFont("Segoe UI", 10))
        label.setFixedWidth(100 if index < 2 else 150)
        layout.addWidget(label)

        path_edit = QLineEdit()
        path_edit.setFont(QFont("Segoe UI", 10))
        path_edit.setText(default_path)
        label.setBuddy(path_edit)
        layout.addWidget(path_edit)

        browse_btn = ModernButton("浏览")
        browse_btn.setSecondary()
        browse_btn.setFixedWidth(80)

        if is_output:
            browse_btn.clicked.connect(self.browse_output)
        else:
            browse_btn.clicked.connect(lambda: self.browse_file(index, path_edit))

        layout.addWidget(browse_btn)

        return widget

    def browse_file(self, index, path_edit):
        """浏览文件"""
        file_types = "所有文件 (*.*);;Excel文件 (*.xlsx *.xls);;PDF文件 (*.pdf);;文本文件 (*.txt);;Word文件 (*.doc *.docx)"
        title = f"选择{self.file_dialog_labels[index]}"

        test_register = self.test_mode_checkbox.isChecked() and 6 <= index < 6 + len(self.register_paths)
        directory = str(ezarc_maintenance_path(True).parent) if test_register else ""
        if test_register:
            file_types = "Excel文件 (*.xlsx)"
        path, _ = QFileDialog.getOpenFileName(path_edit.window(), title, directory, file_types)
        if path:
            if test_register:
                try:
                    path = normalize_test_register_path(path)
                except ValueError as exc:
                    QMessageBox.warning(path_edit.window(), "测试登记路径", str(exc))
                    return
            path_edit.setText(path)

    def browse_output(self):
        """浏览输出文件夹"""
        output_edit = self.output_path_widget.layout().itemAt(1).widget()
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")

        path = QFileDialog.getExistingDirectory(
            self,
            "选择输出文件夹",
            desktop_path
        )

        if path:
            output_edit.setText(path)

    def save_default_settings(self):
        """保存当前设置为默认设置"""
        try:
            # 收集所有文件路径
            settings = {}

            for i, key in enumerate(self.file_keys):
                path_edit = self.file_paths[i].layout().itemAt(1).widget()
                settings[key] = path_edit.text().strip()

            # 正式和测试路径独立保存。
            test_mode = self.test_mode_checkbox.isChecked()
            if test_mode:
                self.test_register_paths = [
                    normalize_test_register_path(path) if path else "" for path in self.register_path_values()
                ]
            else:
                self.remember_production_paths()
            settings.update(zip(self.register_keys, self.last_prod_register_paths))
            settings.update(zip((f"test_{key}" for key in self.register_keys), self.test_register_paths))
            settings[EZARC_PATH_SETTING] = self.ezarc_prod_path

            # 保存输出路径
            output_edit = self.output_path_widget.layout().itemAt(1).widget()
            settings["output"] = output_edit.text().strip()

            # 保留已有默认设置的键和值格式。
            settings["consolidation"] = "是" if self.consolidation_radio.isChecked() else "否"
            settings["register"] = "是" if self.register_checkbox.isChecked() else "否"
            settings["declaration"] = "是" if self.declaration_checkbox.isChecked() else "否"
            settings["test_mode"] = "是" if test_mode else "否"

            # 保存到文件
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=4)

            QMessageBox.information(self, "成功", "默认设置已保存！")
            self.status_bar.showMessage("默认设置已保存")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存默认设置时出错: {str(e)}")
            self.status_bar.showMessage("保存默认设置失败")

    def load_default_settings(self):
        """加载默认设置"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)

                # 加载文件路径
                for i, key in enumerate(self.file_keys):
                    if key in settings and settings[key]:
                        path_edit = self.file_paths[i].layout().itemAt(1).widget()
                        path_edit.setText(settings[key])

                # 加载输出路径
                if "output" in settings and settings["output"]:
                    output_edit = self.output_path_widget.layout().itemAt(1).widget()
                    output_edit.setText(settings["output"])

                packing_option = (
                    self.consolidation_radio if settings.get("consolidation") == "是"
                    else self.non_consolidation_radio
                )
                packing_option.setChecked(True)
                for key, checkbox in (
                    ("register", self.register_checkbox),
                    ("declaration", self.declaration_checkbox),
                ):
                    if key in settings:
                        checkbox.setChecked(settings[key] == "是")
                test_mode = settings.get("test_mode") == "是"
                self.test_mode_checkbox.setChecked(test_mode)
                self.last_prod_register_paths = [
                    settings.get(key) or default for key, default in zip(self.register_keys, self.prod_register_paths)
                ]
                self.test_register_paths = [
                    settings.get(f"test_{key}") or default
                    for key, default in zip(self.register_keys, shared_test_register_paths())
                ]
                production_path = settings.get(EZARC_PATH_SETTING)
                self.ezarc_prod_path = production_path.strip() if isinstance(production_path, str) else str(ezarc_maintenance_path())
                self.apply_mode_paths(test_mode)

                self.status_bar.showMessage("默认设置已加载")

        except Exception as e:
            # 如果加载失败，使用默认设置
            print(f"加载默认设置时出错: {str(e)}")
            self.status_bar.showMessage("使用默认设置")

    def process_data(self):
        """处理数据"""
        # 检查必需文件是否都已选择
        missing_files = []

        # 是否登记
        is_register = self.register_checkbox.isChecked()
        register_shipping = is_register and self.register_shipping.isChecked()
        register_customs = is_register and self.register_customs.isChecked()
        if is_register and not (register_shipping or register_customs):
            QMessageBox.critical(self, "登记选项", "请选择至少一项登记内容，或取消勾选“登记”。")
            return
        # 是否生成申报文件
        is_declaration = self.declaration_checkbox.isChecked()

        input_paths = []
        # 收集文件上传区域的文件路径
        for i, path_widget in enumerate(self.file_paths):
            path_edit = path_widget.layout().itemAt(1).widget()
            path = path_edit.text().strip()
            input_paths.append(path)

        delivery_path = input_paths[0]
        stocking_path = input_paths[1]
        use_stock = bool(stocking_path)
        use_delivery = bool(delivery_path)

        if use_stock and use_delivery:
            QMessageBox.critical(self, "文件冲突", "请只填写发货单或备货单路径，不要同时填写。")
            return

        if not use_stock and not use_delivery:
            QMessageBox.critical(self, "文件缺失", "请填写发货单路径或备货单路径。")
            return

        required_base_indices = [2, 3, 4, 5] + ([1] if use_stock else [0])
        for idx in required_base_indices:
            if not input_paths[idx]:
                missing_files.append(self.base_file_labels[idx])

        # 收集登记设置区域的文件路径（仅在选择登记时必填）
        is_test_mode = self.test_mode_checkbox.isChecked()
        for i, path in enumerate(self.register_path_values()):
            input_paths.append(path)

            required = register_shipping if i == 0 else register_customs
            if required and not path:
                missing_files.append(self.register_file_labels[i])

        # 如果有必需文件未选择，弹出提示框
        if missing_files:
            missing_files_str = "、".join(missing_files)
            QMessageBox.critical(self, "文件缺失", f"以下文件为必需项，请上传：\n{missing_files_str}")
            return

        # 获取输出路径
        output_edit = self.output_path_widget.layout().itemAt(1).widget()
        output_path = output_edit.text().strip()

        if not output_path:
            QMessageBox.critical(self, "错误", "请指定输出路径！")
            return

        # 获取选项
        is_consolidation = self.consolidation_radio.isChecked()

        # 禁用处理按钮，防止重复点击
        self.process_btn.setEnabled(False)
        self.process_btn.setText("处理中...")
        self.status_bar.showMessage("正在处理数据...")

        # 创建并启动处理线程
        self.processing_thread = ProcessingThread(
            input_paths, output_path, self.pictures_path, is_consolidation, is_register, is_declaration, use_stock,
            register_shipping=register_shipping, register_customs=register_customs, test_mode=is_test_mode,
            maintenance_path=self.ezarc_path_edit.text().strip(),
        )
        self.processing_thread.progress.connect(self.status_bar.showMessage)
        self.processing_thread.staging_active.connect(self.set_customs_stop_available)
        self.processing_thread.finished.connect(self.on_processing_finished)
        self.processing_thread.start()

    def set_customs_stop_available(self, available):
        self.stop_customs_btn.setEnabled(available)
        self.stop_customs_btn.setVisible(available)

    def stop_customs_staging(self):
        if self.processing_thread.request_staging_stop():
            self.stop_customs_btn.setEnabled(False)
            self.status_bar.showMessage("正在停止本次开票……")

    def on_processing_finished(self, success, message):
        """处理完成回调"""
        self.set_customs_stop_available(False)
        # 启用按钮
        self.process_btn.setEnabled(True)
        self.process_btn.setText("生成报关资料")

        # 显示结果
        if success:
            self.status_bar.showMessage("处理完成")
            QMessageBox.information(self, "完成", message)
        else:
            self.status_bar.showMessage("处理失败")
            QMessageBox.critical(self, "错误", message)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = QTInterface()
    window.show()
    sys.exit(app.exec())
