# ui/zentao_export_page.py - 修改后的禅道导出页面

import os
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QTextEdit, QFileDialog, QMessageBox, QProgressDialog, QGroupBox, QCheckBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QTextCursor

from core.selenium_worker import SeleniumWorker
from core.settings_manager import SettingsManager
from config.settings import DOWNLOAD_DIR, HEADLESS_MODE_DEFAULT, TEST_REPORT_ID_DEFAULT


class ZentaoExportPage(QWidget):
    # 新增信号：用户登录成功
    user_logged_in = pyqtSignal(object)  # 传递用户信息对象

    def __init__(self, parent=None):
        super().__init__(parent)
        self.worker_thread = None
        self.progress_dialog = None
        self.settings_manager = SettingsManager("zentao_export")
        self.current_user_info = None  # 存储当前用户信息

        self.init_ui()
        self.load_settings()

    def init_ui(self):
        main_layout = QVBoxLayout(self)

        # 禅道登录信息 GroupBox
        login_group_box = QGroupBox("禅道登录信息")
        login_layout = QVBoxLayout()

        self.account_input = self._create_input_field(login_layout, "账号:", "")
        self.account_input.setPlaceholderText("请输入禅道账号")

        self.password_input = self._create_input_field(login_layout, "密码:", "", is_password=True)
        self.password_input.setPlaceholderText("请输入禅道密码")

        # 新增登录测试按钮
        login_test_layout = QHBoxLayout()
        self.test_login_btn = QPushButton("登录并获取用户信息")
        self.test_login_btn.clicked.connect(self._test_login)
        self.test_login_btn.setFixedHeight(30)
        login_test_layout.addWidget(self.test_login_btn)
        login_test_layout.addStretch()
        login_layout.addLayout(login_test_layout)

        # 无头模式复选框
        self.headless_checkbox = QCheckBox("无头模式 (不显示浏览器界面)")
        self.headless_checkbox.setChecked(HEADLESS_MODE_DEFAULT)
        login_layout.addWidget(self.headless_checkbox)

        login_group_box.setLayout(login_layout)
        main_layout.addWidget(login_group_box)

        # 全局参数设置 GroupBox
        global_settings_group_box = QGroupBox("全局参数设置")
        global_settings_layout = QVBoxLayout()

        self.product_name_input = self._create_input_field(global_settings_layout, "产品名称:", "")
        self.product_name_input.setPlaceholderText("请输入产品名称关键字，例如\"2600F\"")

        # 测试单号输入框和保存按钮
        test_report_id_layout = QHBoxLayout()
        self.test_report_id_label = QLabel("测试单号:")
        self.test_report_id_input = QLineEdit(TEST_REPORT_ID_DEFAULT)
        self.test_report_id_input.setPlaceholderText("请输入测试单号，例如\"11111111111\"")
        self.save_test_report_id_button = QPushButton("保存测试单号")
        self.save_test_report_id_button.clicked.connect(self._save_test_report_id)
        test_report_id_layout.addWidget(self.test_report_id_label)
        test_report_id_layout.addWidget(self.test_report_id_input)
        test_report_id_layout.addWidget(self.save_test_report_id_button)
        global_settings_layout.addLayout(test_report_id_layout)

        # 下载目录
        download_dir_layout = QHBoxLayout()
        self.download_dir_label = QLabel("下载目录:")
        self.download_dir_display = QLineEdit(DOWNLOAD_DIR)
        self.download_dir_display.setReadOnly(True)
        self.browse_button = QPushButton("浏览...")
        self.browse_button.clicked.connect(self._browse_download_dir)
        download_dir_layout.addWidget(self.download_dir_label)
        download_dir_layout.addWidget(self.download_dir_display)
        download_dir_layout.addWidget(self.browse_button)
        global_settings_layout.addLayout(download_dir_layout)

        global_settings_group_box.setLayout(global_settings_layout)
        main_layout.addWidget(global_settings_group_box)

        # 操作按钮区域
        button_layout = QHBoxLayout()

        # 开始导出按钮
        self.export_button = QPushButton("开始导出")
        self.export_button.setFixedHeight(40)
        self.export_button.clicked.connect(self._start_export)
        button_layout.addWidget(self.export_button)

        # 刷新用户信息按钮
        self.refresh_user_btn = QPushButton("刷新用户信息")
        self.refresh_user_btn.setFixedHeight(40)
        self.refresh_user_btn.clicked.connect(self.refresh_user_info)
        self.refresh_user_btn.setEnabled(False)  # 默认禁用
        button_layout.addWidget(self.refresh_user_btn)

        main_layout.addLayout(button_layout)

        # 日志输出区域
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setStyleSheet("background-color: #f0f0f0; color: #333; font-family: 'Consolas', 'Monospace';")
        main_layout.addWidget(self.log_output)

    def _create_input_field(self, layout, label_text, default_text="", is_password=False):
        """Helper to create labeled input fields."""
        h_layout = QHBoxLayout()
        label = QLabel(label_text)
        line_edit = QLineEdit(default_text)
        if is_password:
            line_edit.setEchoMode(QLineEdit.Password)
        h_layout.addWidget(label)
        h_layout.addWidget(line_edit)
        layout.addLayout(h_layout)
        return line_edit

    def _browse_download_dir(self):
        """Opens a dialog to select the download directory."""
        initial_dir = self.download_dir_display.text()
        if not os.path.isdir(initial_dir):
            initial_dir = os.path.expanduser("~")

        directory = QFileDialog.getExistingDirectory(self, "选择下载目录", initial_dir)
        if directory:
            self.download_dir_display.setText(directory)
            self.save_settings()
            self.update_log(f"已选择下载目录: {directory}", False)

    def _save_test_report_id(self):
        """Saves the test report ID when the button is clicked."""
        self.save_settings()
        QMessageBox.information(self, "保存成功", "测试单号已保存。")
        self.update_log("测试单号已保存。", False)

    def _test_login(self):
        """测试登录并获取用户信息"""
        account = self.account_input.text().strip()
        password = self.password_input.text().strip()

        if not account or not password:
            QMessageBox.warning(self, "输入错误", "请填写账号和密码")
            return

        self.log_output.clear()
        self.update_log("--- 开始测试登录 ---", False)
        self.test_login_btn.setEnabled(False)

        # 创建进度对话框
        self.progress_dialog = QProgressDialog("正在登录并获取用户信息...", "取消", 0, 100, self)
        self.progress_dialog.setWindowTitle("登录测试")
        self.progress_dialog.setWindowModality(Qt.WindowModal)
        self.progress_dialog.setMinimumDuration(0)
        self.progress_dialog.setValue(0)
        self.progress_dialog.show()

        # 创建只用于登录的工作线程
        self.worker_thread = SeleniumWorker(
            account, password, "", "", "",
            self.headless_checkbox.isChecked(),
            task_type="login_only"  # 只登录，不执行导出
        )

        # 连接信号
        self.worker_thread.log_signal.connect(self.update_log)
        self.worker_thread.finished_signal.connect(self._login_test_finished)
        self.worker_thread.progress_signal.connect(self.progress_dialog.setValue)
        self.worker_thread.user_info_signal.connect(self._on_user_info_received)
        self.progress_dialog.canceled.connect(self._cancel_login_test)

        self.worker_thread.start()

    def _login_test_finished(self, success, message):
        """登录测试完成处理"""
        self.test_login_btn.setEnabled(True)
        if self.progress_dialog:
            self.progress_dialog.hide()

        self.update_log(f"\n--- 登录测试完成: {'成功' if success else '失败'} ---", False)
        self.update_log(message, not success)

        if success:
            QMessageBox.information(self, "登录成功", "登录成功，用户信息已获取")
            self.refresh_user_btn.setEnabled(True)
        else:
            QMessageBox.critical(self, "登录失败", message)

        self.worker_thread = None

    def _on_user_info_received(self, user_info):
        """接收到用户信息"""
        self.current_user_info = user_info
        self.update_log(f"用户信息已获取: {user_info.real_name} ({user_info.account})", False)

        # 发射用户登录信号，通知主窗口
        self.user_logged_in.emit(user_info)

    def _cancel_login_test(self):
        """取消登录测试"""
        if self.worker_thread and self.worker_thread.isRunning():
            self.update_log("用户取消登录测试...", True)
            if self.progress_dialog:
                self.progress_dialog.hide()
            self.test_login_btn.setEnabled(True)

    def refresh_user_info(self):
        """刷新用户信息"""
        if not self.current_user_info:
            QMessageBox.information(self, "提示", "请先测试登录")
            return

        # 重新执行登录测试来刷新用户信息
        self._test_login()

    def _start_export(self):
        """Initiates the data export process in a separate thread."""
        account = self.account_input.text().strip()
        password = self.password_input.text().strip()
        product_name = self.product_name_input.text().strip()
        test_report_id = self.test_report_id_input.text().strip()
        download_dir = self.download_dir_display.text().strip()
        headless_mode = self.headless_checkbox.isChecked()

        if not account or not password or not product_name or not download_dir:
            QMessageBox.warning(self, "输入错误", "账号、密码、产品名称和下载目录都不能为空，请填写完整。")
            return

        # 检查下载目录
        if not os.path.exists(download_dir):
            try:
                os.makedirs(download_dir, exist_ok=True)
                self.update_log(f"已创建下载目录: {download_dir}", False)
            except Exception as e:
                self.update_log(f"错误: 无法创建下载目录 '{download_dir}': {e}", True)
                QMessageBox.critical(self, "目录创建失败", f"无法创建下载目录，请检查权限或路径是否合法。\n错误: {e}")
                return
        elif not os.path.isdir(download_dir):
            self.update_log(f"错误: 下载目录 '{download_dir}' 存在但不是一个目录。", True)
            QMessageBox.critical(self, "路径错误", f"下载目录 '{download_dir}' 存在但不是一个目录，请重新选择。")
            return

        self.save_settings()

        self.log_output.clear()
        self.update_log("--- 开始执行自动化任务 ---", False)
        self.export_button.setEnabled(False)

        self.progress_dialog = QProgressDialog("导出进度", "取消", 0, 100, self)
        self.progress_dialog.setWindowTitle("导出进度")
        self.progress_dialog.setWindowModality(Qt.WindowModal)
        self.progress_dialog.setAutoClose(True)
        self.progress_dialog.setAutoReset(True)
        self.progress_dialog.setMinimumDuration(0)
        self.progress_dialog.setValue(0)
        self.progress_dialog.setLabelText("正在准备...")
        self.progress_dialog.show()

        # 创建导出工作线程
        self.worker_thread = SeleniumWorker(
            account, password, product_name, test_report_id, download_dir, headless_mode, "export"
        )
        self.worker_thread.log_signal.connect(self.update_log)
        self.worker_thread.status_signal.connect(self.progress_dialog.setLabelText)
        self.worker_thread.finished_signal.connect(self._export_finished)
        self.worker_thread.progress_signal.connect(self.progress_dialog.setValue)
        self.worker_thread.user_info_signal.connect(self._on_user_info_received)  # 也监听用户信息
        self.progress_dialog.canceled.connect(self._cancel_export)

        self.worker_thread.start()

    def _export_finished(self, success, message):
        """Handles the completion of the export process."""
        self.export_button.setEnabled(True)
        if self.progress_dialog:
            self.progress_dialog.hide()

        self.update_log(f"\n--- 任务完成: {'成功' if success else '失败'} ---", False)
        self.update_log(message, not success)
        if success:
            QMessageBox.information(self, "任务完成", message)
        else:
            QMessageBox.critical(self, "任务失败", message)
        self.worker_thread = None

    def _cancel_export(self):
        """Handles cancellation of the export process."""
        if self.worker_thread and self.worker_thread.isRunning():
            self.update_log("用户请求取消任务 (此版本暂不支持中断正在进行的Selenium操作)...", True)

            if self.progress_dialog:
                self.progress_dialog.hide()

            self.export_button.setEnabled(True)
            QMessageBox.information(self, "任务取消", "导出任务已请求取消。请等待当前Selenium操作结束。")
        else:
            self.update_log("没有正在运行的任务可以取消。", False)

    def update_log(self, message, is_error=False):
        """Appends a message to the log QTextEdit."""
        cursor = self.log_output.textCursor()
        format = cursor.charFormat()
        if is_error:
            format.setForeground(Qt.red)
        else:
            format.setForeground(Qt.black)
        cursor.setCharFormat(format)
        self.log_output.append(message)
        cursor = self.log_output.textCursor()
        cursor.movePosition(QTextCursor.End)
        self.log_output.setTextCursor(cursor)

    def save_settings(self):
        """Saves settings specific to this tab."""
        settings = {
            "account": self.account_input.text(),
            "password": self.password_input.text(),
            "product_name": self.product_name_input.text(),
            "test_report_id": self.test_report_id_input.text(),
            "download_dir": self.download_dir_display.text(),
            "headless_mode": self.headless_checkbox.isChecked()
        }
        self.settings_manager.save_settings("zentao_export", settings, self.update_log)

    def load_settings(self):
        """Loads settings specific to this tab."""
        default_settings = {
            "account": "",
            "password": "",
            "product_name": "",
            "test_report_id": TEST_REPORT_ID_DEFAULT,
            "download_dir": DOWNLOAD_DIR,
            "headless_mode": HEADLESS_MODE_DEFAULT
        }
        loaded_settings = self.settings_manager.load_settings(
            "zentao_export",
            default_settings=default_settings,
            log_callback=self.update_log
        )
        self.product_name_input.setText(loaded_settings.get("product_name", ""))
        self.test_report_id_input.setText(loaded_settings.get("test_report_id", TEST_REPORT_ID_DEFAULT))
        self.download_dir_display.setText(loaded_settings.get("download_dir", DOWNLOAD_DIR))
        self.headless_checkbox.setChecked(loaded_settings.get("headless_mode", HEADLESS_MODE_DEFAULT))

        # 不自动加载账号密码，保证安全性
        self.account_input.setText("")
        self.password_input.setText("")

        # 确保默认下载目录存在
        initial_download_dir = self.download_dir_display.text()
        if initial_download_dir and not os.path.exists(initial_download_dir):
            try:
                os.makedirs(initial_download_dir, exist_ok=True)
                self.update_log(f"已创建默认下载目录: {initial_download_dir}", False)
            except Exception as e:
                self.update_log(f"警告: 无法创建默认下载目录 '{initial_download_dir}': {e}", True)
                self.download_dir_display.setText(os.path.expanduser("~"))
