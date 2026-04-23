import os
import sys
import json
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGroupBox, QLabel, QLineEdit, QPushButton, QFileDialog,
    QStatusBar, QMessageBox, QCheckBox
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QPropertyAnimation, QEasingCurve
from PyQt6.QtGui import QFont, QIcon
import Invoice_Data
import Stock_Preparation


class ToggleButton(QPushButton):
    """自定义切换按钮"""

    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setCheckable(True)
        self.setChecked(False)
        self.setStyleSheet("""
            QPushButton {
                background-color: #f0f0f0;
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 4px 8px;
                text-align: left;
            }
            QPushButton:checked {
                background-color: #e0e0e0;
            }
            QPushButton:hover {
                background-color: #e8e8e8;
            }
        """)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

        # 添加箭头图标
        self.update_arrow()
        self.toggled.connect(self.update_arrow)

    def update_arrow(self):
        if self.isChecked():
            self.setText("▼ " + self.text()[2:])
        else:
            self.setText("▶ " + self.text()[2:])


class CollapsibleWidget(QWidget):
    """可折叠的部件容器"""

    def __init__(self, title, parent=None):
        super().__init__(parent)

        self.content_height = 0

        # 主布局
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # 切换按钮
        self.toggle_button = ToggleButton("▶ " + title)
        self.toggle_button.toggled.connect(self.toggle_content)
        layout.addWidget(self.toggle_button)

        # 内容框架
        self.content_frame = QWidget()
        self.content_frame.setMaximumHeight(0)
        self.content_layout = QVBoxLayout(self.content_frame)
        self.content_layout.setContentsMargins(10, 10, 10, 10)
        self.content_layout.setSpacing(10)

        layout.addWidget(self.content_frame)

        # 动画效果
        self.animation = QPropertyAnimation(self.content_frame, b"maximumHeight")
        self.animation.setDuration(300)
        self.animation.setEasingCurve(QEasingCurve.Type.InOutQuad)
        self.animation.finished.connect(self.on_animation_finished)

    def addWidget(self, widget):
        """添加部件到容器中"""
        self.content_layout.addWidget(widget)

        # 更新内容高度
        self.content_frame.adjustSize()
        self.content_height = self.content_frame.sizeHint().height()

    def toggle_content(self, checked):
        """切换内容的显示/隐藏"""
        if checked:
            self.animation.setStartValue(0)
            self.animation.setEndValue(self.content_height)
        else:
            self.animation.setStartValue(self.content_frame.height())
            self.animation.setEndValue(0)

        self.animation.start()

    def on_animation_finished(self):
        """动画完成后更新布局"""
        # 获取父部件并更新其布局
        parent = self.parent()
        while parent:
            if hasattr(parent, 'layout') and parent.layout():
                parent.layout().activate()
                parent.adjustSize()
            parent = parent.parent()

        # 更新主窗口布局
        main_window = self.window()
        if main_window:
            main_window.adjustSize()


class ProcessingThread(QThread):
    """处理线程，避免界面冻结"""
    finished = pyqtSignal(bool, str)  # 信号：处理完成(成功, 消息)

    def __init__(self, input_paths, output_path, pictures_path, is_consolidation, is_register,
                 is_declaration, use_stock):
        super().__init__()
        self.input_paths = input_paths
        self.output_path = output_path
        self.pictures_path = pictures_path
        self.is_consolidation = is_consolidation
        self.is_register = is_register
        self.is_declaration = is_declaration
        self.use_stock = use_stock

    def run(self):
        """执行处理任务"""
        try:
            # 检查文件是否存在
            for path in [p for p in self.input_paths if p]:
                if not os.path.exists(path):
                    self.finished.emit(False, f"文件不存在: {os.path.basename(path)}")
                    return

            # 初始化CustomsDeclaration()类
            if self.use_stock:
                generator = Stock_Preparation.CustomsDeclaration()
                generator.generate_all_documents(
                    self.input_paths[1],  # 备货单
                    self.input_paths[2],  # 产品信息单证专用
                    self.input_paths[3],  # 品名
                    self.input_paths[4],  # 全站点地址
                    self.input_paths[5],  # 要素表
                    self.output_path,
                    self.pictures_path, self.is_consolidation, self.is_declaration
                )
                if self.is_register:
                    generator.register_all_documents(
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
                        self.is_consolidation
                    )
            else:
                generator = Invoice_Data.CustomsDeclaration()
                generator.generate_all_documents(
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
                    generator.register_all_documents(
                        self.input_paths[0],   # 发货单
                        self.input_paths[6],   # 出运明细
                        self.input_paths[2],   # 产品信息单证专用
                        self.input_paths[7],   # 鲁米亚报关明细
                        self.input_paths[8],   # 易达报关明细
                        self.input_paths[9],   # 素桐报关开票明细
                        self.input_paths[10],  # 六浪报关明细
                        self.input_paths[11]   # 匠晟报关明细
                    )

            # 处理成功
            self.finished.emit(True, "报关资料生成成功！")

        except Exception as e:
            self.finished.emit(False, f"处理过程中发生错误: {str(e)}")


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
                border: none;
                border-radius: 5px;
                padding: 8px 16px;
                font-weight: bold;
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
        self.setGeometry(500, 200, 900, 600)
        # 设置最小宽度，确保折叠后界面不会变窄
        self.setMinimumWidth(900)

        # 设置窗口图标
        self.setWindowIcon(QIcon(r'\\192.168.0.45\供应链管理\2 物流发货管理\17.单证数据表维护\pictures\tool.ico'))

        # 初始化默认设置文件路径
        # self.settings_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "default_settings.json")
        self.settings_file = os.path.join(os.path.dirname(sys.executable), "default_settings.json")
        # 路径标签与配置键
        self.base_file_labels = ["发货单路径", "备货单路径", "产品信息单证专用路径", "品名", "全站点地址", "要素表"]
        self.register_file_labels = ["出运明细", "鲁米亚报关明细", "易达报关明细", "素桐报关开票明细", "六浪报关明细", "匠晟报关明细"]
        self.file_dialog_labels = [label.replace("路径", "") for label in self.base_file_labels] + self.register_file_labels
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
        self.main_layout.setSpacing(20)
        self.main_layout.setContentsMargins(30, 30, 30, 30)

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
        self.lumiya_register_filepath = os.path.join(self.register_folder_path, "2025.9-鲁米亚报关明细.xlsx")
        self.easycan_register_filepath = os.path.join(self.register_folder_path, "2025.9-易达报关明细 修复.xlsx")
        self.sutong_register_filepath = os.path.join(self.register_folder_path, "2025-素桐报关开票明细.xlsx")
        self.liulang_register_filepath = os.path.join(self.register_folder_path, "2025-六浪报关明细.xlsx")
        self.jiangsheng_register_filepath = os.path.join(self.register_folder_path, "2025-匠晟报关明细.xlsx")

    def setup_styles(self):
        """设置应用样式"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #fafafa;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #e0e0e0;
                border-radius: 8px;
                margin-top: 1ex;
                padding-top: 10px;
                background-color: white;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
                color: #6200ee;
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
            QCheckBox {
                spacing: 5px;
            }
            QCheckBox::indicator {
                width: 15px;
                height: 15px;
            }
            QCheckBox::indicator:unchecked {
                border: 1px solid #ddd;
                border-radius: 3px;
                background: white;
            }
            QCheckBox::indicator:checked {
                border: 1px solid #6200ee;
                border-radius: 3px;
                background: #6200ee;
            }
        """)

    def setup_ui(self):
        """设置用户界面"""
        # 标题
        title_label = QLabel("报关资料生成工具")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        title_label.setStyleSheet("color: #6200ee; margin: 10px 0;")
        self.main_layout.addWidget(title_label)

        # 发货单区域
        delivery_group = QGroupBox("发货单")
        delivery_group.setFont(QFont("Segoe UI", 11))
        delivery_layout = QVBoxLayout(delivery_group)
        delivery_layout.setSpacing(15)
        delivery_layout.setContentsMargins(15, 20, 15, 20)

        # 创建发货单路径输入控件
        self.file_paths = []
        delivery_widget = self.create_file_widget(self.base_file_labels[0], "", 0)
        self.file_paths.append(delivery_widget)
        delivery_layout.addWidget(delivery_widget)

        # 创建备货单路径输入控件
        stocking_widget = self.create_file_widget(self.base_file_labels[1], self.stocking_filepath, 1)
        self.file_paths.append(stocking_widget)
        delivery_layout.addWidget(stocking_widget)

        # 创建输出路径输入控件（放在备货单路径下方，不单独分组）
        self.output_path_widget = self.create_file_widget("输出路径", self.default_output, -1, is_output=True)
        delivery_layout.addWidget(self.output_path_widget)

        self.main_layout.addWidget(delivery_group)

        # 拼箱选项区域
        consolidation_group = QGroupBox("拼箱选项")
        consolidation_group.setFont(QFont("Segoe UI", 11))
        consolidation_layout = QHBoxLayout(consolidation_group)
        consolidation_layout.setContentsMargins(15, 20, 15, 20)

        # 创建是否拼箱的单选按钮
        self.consolidation_no = QCheckBox("否")
        self.consolidation_no.setChecked(True)
        self.consolidation_no.toggled.connect(self.on_consolidation_toggled)

        self.consolidation_yes = QCheckBox("是")
        self.consolidation_yes.toggled.connect(self.on_consolidation_toggled)

        consolidation_layout.addWidget(QLabel("是否拼箱:"))
        consolidation_layout.addWidget(self.consolidation_no)
        consolidation_layout.addWidget(self.consolidation_yes)
        consolidation_layout.addStretch()

        self.main_layout.addWidget(consolidation_group)

        # 登记选项区域（默认是）
        register_group = QGroupBox("登记选项")
        register_group.setFont(QFont("Segoe UI", 11))
        register_layout = QHBoxLayout(register_group)
        register_layout.setContentsMargins(15, 20, 15, 20)

        self.register_no = QCheckBox("否")
        self.register_yes = QCheckBox("是")
        self.register_yes.setChecked(True)
        self.register_no.toggled.connect(self.on_register_toggled)
        self.register_yes.toggled.connect(self.on_register_toggled)

        register_layout.addWidget(QLabel("是否登记:"))
        register_layout.addWidget(self.register_no)
        register_layout.addWidget(self.register_yes)
        register_layout.addStretch()

        self.main_layout.addWidget(register_group)

        # 申报选项区域（默认是）
        declaration_group = QGroupBox("申报选项")
        declaration_group.setFont(QFont("Segoe UI", 11))
        declaration_layout = QHBoxLayout(declaration_group)
        declaration_layout.setContentsMargins(15, 20, 15, 20)

        self.declaration_no = QCheckBox("否")
        self.declaration_yes = QCheckBox("是")
        self.declaration_yes.setChecked(True)
        self.declaration_no.toggled.connect(self.on_declaration_toggled)
        self.declaration_yes.toggled.connect(self.on_declaration_toggled)

        declaration_layout.addWidget(QLabel("是否生成申报文件:"))
        declaration_layout.addWidget(self.declaration_no)
        declaration_layout.addWidget(self.declaration_yes)
        declaration_layout.addStretch()

        self.main_layout.addWidget(declaration_group)

        # 文件上传区域（其他文件，可折叠，默认折叠）
        self.upload_section = CollapsibleWidget("文件上传")

        # 创建其他文件路径输入控件
        file_labels = self.base_file_labels[2:]
        default_paths = [
            self.product_info_filepath,
            self.product_name_filepath,
            self.full_site_address_filepath,
            self.elements_filepath
        ]

        for i, (label, default_path) in enumerate(zip(file_labels, default_paths)):
            path_widget = self.create_file_widget(label, default_path, i + 2)
            self.file_paths.append(path_widget)
            self.upload_section.addWidget(path_widget)

        self.main_layout.addWidget(self.upload_section)

        # 登记设置区域
        self.register_section = CollapsibleWidget("登记设置")

        # 创建登记文件路径输入控件
        self.register_paths = []
        register_labels = self.register_file_labels
        register_default_paths = [
            self.invoice_number_register_filepath,
            self.lumiya_register_filepath,
            self.easycan_register_filepath,
            self.sutong_register_filepath,
            self.liulang_register_filepath,
            self.jiangsheng_register_filepath
        ]

        for i, (label, default_path) in enumerate(zip(register_labels, register_default_paths)):
            path_widget = self.create_file_widget(label, default_path, 6 + i)
            self.register_paths.append(path_widget)
            self.register_section.addWidget(path_widget)

        self.main_layout.addWidget(self.register_section)

        # 按钮区域
        button_layout = QHBoxLayout()
        button_layout.setSpacing(15)
        button_layout.addStretch()

        # 保存默认设置按钮
        self.save_defaults_btn = ModernButton("保持默认设置")
        self.save_defaults_btn.setprimary()
        self.save_defaults_btn.clicked.connect(self.save_default_settings)
        button_layout.addWidget(self.save_defaults_btn)

        # 处理按钮
        self.process_btn = ModernButton("生成报关资料")
        self.process_btn.setprimary()
        self.process_btn.clicked.connect(self.process_data)
        button_layout.addWidget(self.process_btn)

        button_layout.addStretch()
        self.main_layout.addLayout(button_layout)

        # 添加底部版权信息
        copyright_label = QLabel("© 2026 报关资料生成工具")
        copyright_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        copyright_label.setFont(QFont("Segoe UI", 8))
        copyright_label.setStyleSheet("color: #9e9e9e; margin-top: 20px;")
        self.main_layout.addWidget(copyright_label)

        # 状态栏
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就绪")

    def on_consolidation_toggled(self, checked):
        """处理拼箱选项切换"""
        if checked:
            sender = self.sender()
            if sender == self.consolidation_yes:
                self.consolidation_no.setChecked(False)
            elif sender == self.consolidation_no:
                self.consolidation_yes.setChecked(False)

    def on_register_toggled(self, checked):
        """处理登记选项切换"""
        if checked:
            sender = self.sender()
            if sender == self.register_yes:
                self.register_no.setChecked(False)
            elif sender == self.register_no:
                self.register_yes.setChecked(False)

    def on_declaration_toggled(self, checked):
        """处理申报选项切换"""
        if checked:
            sender = self.sender()
            if sender == self.declaration_yes:
                self.declaration_no.setChecked(False)
            elif sender == self.declaration_no:
                self.declaration_yes.setChecked(False)

    def create_file_widget(self, label_text, default_path, index, is_output=False):
        """创建文件路径选择控件"""
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        label = QLabel(f"{label_text}:")
        label.setFont(QFont("Segoe UI", 10))
        label.setFixedWidth(150)
        layout.addWidget(label)

        path_edit = QLineEdit()
        path_edit.setFont(QFont("Segoe UI", 10))
        path_edit.setText(default_path)
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

        path, _ = QFileDialog.getOpenFileName(self, title, "", file_types)
        if path:
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

            # 保存登记设置文件路径
            for i, key in enumerate(self.register_keys):
                path_edit = self.register_paths[i].layout().itemAt(1).widget()
                settings[key] = path_edit.text().strip()

            # 保存输出路径
            output_edit = self.output_path_widget.layout().itemAt(1).widget()
            settings["output"] = output_edit.text().strip()

            # 保存拼箱选项
            settings["consolidation"] = "是" if self.consolidation_yes.isChecked() else "否"
            # 保存登记选项
            settings["register"] = "是" if self.register_yes.isChecked() else "否"
            # 保存申报选项
            settings["declaration"] = "是" if self.declaration_yes.isChecked() else "否"

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

                # 加载登记设置文件路径
                for i, key in enumerate(self.register_keys):
                    if key in settings and settings[key]:
                        path_edit = self.register_paths[i].layout().itemAt(1).widget()
                        path_edit.setText(settings[key])

                # 加载输出路径
                if "output" in settings and settings["output"]:
                    output_edit = self.output_path_widget.layout().itemAt(1).widget()
                    output_edit.setText(settings["output"])

                # 加载拼箱选项
                if "consolidation" in settings:
                    if settings["consolidation"] == "是":
                        self.consolidation_yes.setChecked(True)
                        self.consolidation_no.setChecked(False)
                    else:
                        self.consolidation_yes.setChecked(False)
                        self.consolidation_no.setChecked(True)

                # 加载登记选项
                if "register" in settings:
                    if settings["register"] == "是":
                        self.register_yes.setChecked(True)
                        self.register_no.setChecked(False)
                    else:
                        self.register_yes.setChecked(False)
                        self.register_no.setChecked(True)

                # 加载申报选项
                if "declaration" in settings:
                    if settings["declaration"] == "是":
                        self.declaration_yes.setChecked(True)
                        self.declaration_no.setChecked(False)
                    else:
                        self.declaration_yes.setChecked(False)
                        self.declaration_no.setChecked(True)

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
        is_register = self.register_yes.isChecked()
        # 是否生成申报文件
        is_declaration = self.declaration_yes.isChecked()

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
        for i, path_widget in enumerate(self.register_paths):
            path_edit = path_widget.layout().itemAt(1).widget()
            path = path_edit.text().strip()
            input_paths.append(path)

            if is_register and not path:
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
        is_consolidation = self.consolidation_yes.isChecked()

        # 禁用处理按钮，防止重复点击
        self.process_btn.setEnabled(False)
        self.process_btn.setText("处理中...")
        self.status_bar.showMessage("正在处理数据...")

        # 创建并启动处理线程
        self.processing_thread = ProcessingThread(
            input_paths, output_path, self.pictures_path, is_consolidation, is_register, is_declaration, use_stock
        )
        self.processing_thread.finished.connect(self.on_processing_finished)
        self.processing_thread.start()

    def on_processing_finished(self, success, message):
        """处理完成回调"""
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
