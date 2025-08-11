# ui/bug_query_page.py - 历史BUG查询页面

import os
import json
from datetime import datetime, timedelta
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QComboBox, QDateEdit, QTextEdit, QTableWidget, QTableWidgetItem,
    QGroupBox, QGridLayout, QHeaderView, QMessageBox, QFileDialog,
    QCheckBox, QProgressBar, QSplitter, QTabWidget
)
from PyQt5.QtCore import Qt, QDate, pyqtSignal
from PyQt5.QtGui import QTextCursor

from core.settings_manager import SettingsManager
from config.settings import BUG_QUERY_STATUS_OPTIONS, BUG_SEVERITY_OPTIONS
import pandas as pd


class BugQueryPage(QWidget):
    """历史BUG查询页面"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.settings_manager = SettingsManager("bug_query")
        self.bug_query_worker = None
        self.bug_data = []
        self.user_info = None  # 当前登录用户信息

        self.init_ui()
        self.load_settings()

    def init_ui(self):
        main_layout = QHBoxLayout(self)

        # 创建分割器，左侧查询条件，右侧结果显示
        splitter = QSplitter(Qt.Horizontal)

        # 左侧查询面板
        left_panel = self._create_query_panel()
        splitter.addWidget(left_panel)

        # 右侧结果面板
        right_panel = self._create_result_panel()
        splitter.addWidget(right_panel)

        # 设置分割器比例
        splitter.setStretchFactor(0, 1)  # 左侧占1份
        splitter.setStretchFactor(1, 3)  # 右侧占3份

        main_layout.addWidget(splitter)

    def _create_query_panel(self):
        """创建查询条件面板"""
        panel = QWidget()
        layout = QVBoxLayout(panel)

        # 管理员账号配置
        manager_group = QGroupBox("管理员账号配置")
        manager_layout = QGridLayout()

        manager_layout.addWidget(QLabel("管理员账号:"), 0, 0)
        self.manager_account_input = QLineEdit()
        self.manager_account_input.setPlaceholderText("请输入管理员账号")
        manager_layout.addWidget(self.manager_account_input, 0, 1)

        manager_layout.addWidget(QLabel("管理员密码:"), 1, 0)
        self.manager_password_input = QLineEdit()
        self.manager_password_input.setEchoMode(QLineEdit.Password)
        self.manager_password_input.setPlaceholderText("请输入管理员密码")
        manager_layout.addWidget(self.manager_password_input, 1, 1)

        self.save_manager_btn = QPushButton("保存管理员信息")
        self.save_manager_btn.clicked.connect(self.save_manager_config)
        manager_layout.addWidget(self.save_manager_btn, 2, 0, 1, 2)

        manager_group.setLayout(manager_layout)
        layout.addWidget(manager_group)

        # 当前操作人信息
        operator_group = QGroupBox("操作人信息")
        operator_layout = QVBoxLayout()

        self.operator_label = QLabel("当前操作人: 未登录")
        self.operator_label.setStyleSheet("font-weight: bold; color: #333;")
        operator_layout.addWidget(self.operator_label)

        self.login_status_label = QLabel("状态: 请先在主页面登录")
        self.login_status_label.setStyleSheet("color: #999;")
        operator_layout.addWidget(self.login_status_label)

        operator_group.setLayout(operator_layout)
        layout.addWidget(operator_group)

        # 查询条件
        query_group = QGroupBox("查询条件")
        query_layout = QGridLayout()

        query_layout.addWidget(QLabel("产品名称:"), 0, 0)
        self.product_name_input = QLineEdit()
        self.product_name_input.setPlaceholderText("例如: 2600F")
        query_layout.addWidget(self.product_name_input, 0, 1)

        query_layout.addWidget(QLabel("BUG状态:"), 1, 0)
        self.status_combo = QComboBox()
        self.status_combo.addItems(["全部"] + BUG_QUERY_STATUS_OPTIONS)
        query_layout.addWidget(self.status_combo, 1, 1)

        query_layout.addWidget(QLabel("严重程度:"), 2, 0)
        self.severity_combo = QComboBox()
        self.severity_combo.addItems(["全部", "1-严重", "2-主要", "3-次要", "4-建议"])
        query_layout.addWidget(self.severity_combo, 2, 1)

        query_layout.addWidget(QLabel("开始日期:"), 3, 0)
        self.start_date = QDateEdit()
        self.start_date.setDate(QDate.currentDate().addDays(-30))  # 默认30天前
        self.start_date.setCalendarPopup(True)
        query_layout.addWidget(self.start_date, 3, 1)

        query_layout.addWidget(QLabel("结束日期:"), 4, 0)
        self.end_date = QDateEdit()
        self.end_date.setDate(QDate.currentDate())
        self.end_date.setCalendarPopup(True)
        query_layout.addWidget(self.end_date, 4, 1)

        # 高级选项
        self.include_resolved_cb = QCheckBox("包含已解决的BUG")
        self.include_resolved_cb.setChecked(True)
        query_layout.addWidget(self.include_resolved_cb, 5, 0, 1, 2)

        self.include_closed_cb = QCheckBox("包含已关闭的BUG")
        self.include_closed_cb.setChecked(False)
        query_layout.addWidget(self.include_closed_cb, 6, 0, 1, 2)

        query_group.setLayout(query_layout)
        layout.addWidget(query_group)

        # 操作按钮
        button_layout = QVBoxLayout()

        self.query_btn = QPushButton("开始查询")
        self.query_btn.setFixedHeight(35)
        self.query_btn.clicked.connect(self.start_query)
        self.query_btn.setEnabled(False)  # 默认禁用，需要先登录
        button_layout.addWidget(self.query_btn)

        self.export_btn = QPushButton("导出结果")
        self.export_btn.setFixedHeight(35)
        self.export_btn.clicked.connect(self.export_results)
        self.export_btn.setEnabled(False)
        button_layout.addWidget(self.export_btn)

        self.clear_btn = QPushButton("清空结果")
        self.clear_btn.setFixedHeight(35)
        self.clear_btn.clicked.connect(self.clear_results)
        button_layout.addWidget(self.clear_btn)

        layout.addLayout(button_layout)
        layout.addStretch()

        return panel

    def _create_result_panel(self):
        """创建结果显示面板"""
        panel = QWidget()
        layout = QVBoxLayout(panel)

        # 标签页组件
        tab_widget = QTabWidget()

        # BUG列表标签页
        bug_list_tab = QWidget()
        bug_list_layout = QVBoxLayout(bug_list_tab)

        # 结果统计
        self.result_label = QLabel("查询结果: 0 条记录")
        self.result_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        bug_list_layout.addWidget(self.result_label)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        bug_list_layout.addWidget(self.progress_bar)

        # BUG列表表格
        self.bug_table = QTableWidget()
        self.bug_table.setColumnCount(8)
        self.bug_table.setHorizontalHeaderLabels([
            "BUG ID", "标题", "状态", "创建人", "创建时间",
            "严重程度", "指派给", "操作"
        ])

        # 设置表格属性
        header = self.bug_table.horizontalHeader()
        header.setStretchLastSection(True)
        header.setSectionResizeMode(1, QHeaderView.Stretch)  # 标题列自适应
        self.bug_table.setAlternatingRowColors(True)
        self.bug_table.setSelectionBehavior(QTableWidget.SelectRows)

        bug_list_layout.addWidget(self.bug_table)
        tab_widget.addTab(bug_list_tab, "BUG列表")

        # 日志标签页
        log_tab = QWidget()
        log_layout = QVBoxLayout(log_tab)

        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setMaximumHeight(200)
        self.log_output.setStyleSheet("""
            QTextEdit {
                background-color: #f5f5f5;
                font-family: 'Consolas', 'Monaco', monospace;
                font-size: 12px;
            }
        """)
        log_layout.addWidget(self.log_output)
        tab_widget.addTab(log_tab, "操作日志")

        layout.addWidget(tab_widget)

        return panel

    def set_user_info(self, user_info):
        """设置当前登录用户信息"""
        self.user_info = user_info
        if user_info:
            self.operator_label.setText(f"当前操作人: {user_info.real_name} ({user_info.account})")
            self.login_status_label.setText("状态: 已登录，可以进行BUG查询")
            self.login_status_label.setStyleSheet("color: #4CAF50;")

            # 启用查询按钮
            if self.manager_account_input.text() and self.manager_password_input.text():
                self.query_btn.setEnabled(True)
        else:
            self.operator_label.setText("当前操作人: 未登录")
            self.login_status_label.setText("状态: 请先在主页面登录")
            self.login_status_label.setStyleSheet("color: #999;")
            self.query_btn.setEnabled(False)

    def save_manager_config(self):
        """保存管理员配置"""
        if not self.manager_account_input.text() or not self.manager_password_input.text():
            QMessageBox.warning(self, "输入错误", "请填写完整的管理员账号和密码")
            return

        # 这里可以加入密码加密存储
        reply = QMessageBox.question(
            self, "确认保存",
            "管理员密码将被保存到本地配置文件中。\n为了安全起见，建议定期更换密码。\n确定要保存吗？",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.save_settings()
            QMessageBox.information(self, "保存成功", "管理员配置已保存")
            self.log("管理员配置已保存")

            # 如果用户已登录，启用查询按钮
            if self.user_info:
                self.query_btn.setEnabled(True)

    def start_query(self):
        """开始查询历史BUG"""
        if not self.user_info:
            QMessageBox.warning(self, "未登录", "请先在主页面登录禅道系统")
            return

        if not self.manager_account_input.text() or not self.manager_password_input.text():
            QMessageBox.warning(self, "配置错误", "请先配置管理员账号")
            return

        # 构建查询参数
        query_params = {
            'status': self.status_combo.currentText() if self.status_combo.currentIndex() > 0 else None,
            'severity': self.severity_combo.currentIndex() if self.severity_combo.currentIndex() > 0 else None,
            'date_from': self.start_date.date().toString("yyyy-MM-dd"),
            'date_to': self.end_date.date().toString("yyyy-MM-dd"),
            'include_resolved': self.include_resolved_cb.isChecked(),
            'include_closed': self.include_closed_cb.isChecked()
        }


        self.log("开始查询历史BUG...", clear=True)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # 不确定进度条
        self.query_btn.setEnabled(False)

        # 创建查询工作线程
        from core.selenium_worker import BugQueryWorker
        self.bug_query_worker = BugQueryWorker(
            manager_account=self.manager_account_input.text(),
            manager_password=self.manager_password_input.text(),
            operator_name=self.user_info.real_name,
            product_name=self.product_name_input.text(),
            query_params=query_params
        )

        # 连接信号
        self.bug_query_worker.log_signal.connect(self.log)
        self.bug_query_worker.finished_signal.connect(self.query_finished)
        self.bug_query_worker.progress_signal.connect(self.progress_bar.setValue)
        self.bug_query_worker.bug_data_signal.connect(self.display_bug_data)

        # 启动查询
        self.bug_query_worker.start()

    def query_finished(self, success, message):
        """查询完成处理"""
        self.progress_bar.setVisible(False)
        self.query_btn.setEnabled(True)

        if success:
            self.export_btn.setEnabled(len(self.bug_data) > 0)
            self.log(f"查询完成: {message}")
            QMessageBox.information(self, "查询完成", message)
        else:
            self.log(f"查询失败: {message}", is_error=True)
            QMessageBox.critical(self, "查询失败", message)

    def display_bug_data(self, bug_list):
        """显示BUG数据"""
        self.bug_data = bug_list
        self.result_label.setText(f"查询结果: {len(bug_list)} 条记录")

        # 清空表格
        self.bug_table.setRowCount(0)

        # 填充数据
        for i, bug in enumerate(bug_list):
            self.bug_table.insertRow(i)
            self.bug_table.setItem(i, 0, QTableWidgetItem(str(bug.get('id', ''))))
            self.bug_table.setItem(i, 1, QTableWidgetItem(bug.get('title', '')))
            self.bug_table.setItem(i, 2, QTableWidgetItem(bug.get('status', '')))
            self.bug_table.setItem(i, 3, QTableWidgetItem(bug.get('opened_by', '')))
            self.bug_table.setItem(i, 4, QTableWidgetItem(bug.get('opened_date', '')))
            self.bug_table.setItem(i, 5, QTableWidgetItem(bug.get('severity', '')))
            self.bug_table.setItem(i, 6, QTableWidgetItem(bug.get('assigned_to', '')))

            # 操作按钮
            action_btn = QPushButton("详情")
            action_btn.clicked.connect(lambda checked, bug_id=bug.get('id'): self.show_bug_detail(bug_id))
            self.bug_table.setCellWidget(i, 7, action_btn)

    def show_bug_detail(self, bug_id):
        """显示BUG详情"""
        # 这里可以实现BUG详情显示功能
        QMessageBox.information(self, "BUG详情", f"BUG ID: {bug_id}\n详情功能待实现")

    def export_results(self):
        """导出查询结果"""
        if not self.bug_data:
            QMessageBox.warning(self, "无数据", "没有可导出的数据")
            return

        file_name, _ = QFileDialog.getSaveFileName(
            self, "导出BUG查询结果",
            f"BUG查询结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel Files (*.xlsx);;CSV Files (*.csv)"
        )

        if file_name:
            try:
                # 准备导出数据
                export_data = []
                for bug in self.bug_data:
                    export_data.append({
                        'BUG ID': bug.get('id', ''),
                        '标题': bug.get('title', ''),
                        '状态': bug.get('status', ''),
                        '创建人': bug.get('opened_by', ''),
                        '创建时间': bug.get('opened_date', ''),
                        '严重程度': bug.get('severity', ''),
                        '指派给': bug.get('assigned_to', ''),
                    })

                df = pd.DataFrame(export_data)

                if file_name.endswith('.xlsx'):
                    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
                        df.to_excel(writer, sheet_name='BUG查询结果', index=False)

                        # 添加查询信息到第二个sheet
                        query_info = {
                            '查询时间': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
                            '操作人': [self.user_info.real_name if self.user_info else ''],
                            '产品名称': [self.product_name_input.text()],
                            '查询状态': [self.status_combo.currentText()],
                            '严重程度': [self.severity_combo.currentText()],
                            '开始日期': [self.start_date.date().toString("yyyy-MM-dd")],
                            '结束日期': [self.end_date.date().toString("yyyy-MM-dd")],
                            '结果数量': [len(self.bug_data)]
                        }
                        query_df = pd.DataFrame(query_info)
                        query_df.to_excel(writer, sheet_name='查询信息', index=False)
                else:
                    df.to_csv(file_name, index=False, encoding='utf-8-sig')

                self.log(f"导出成功: {file_name}")
                QMessageBox.information(self, "导出成功", f"数据已导出到: {file_name}")

            except Exception as e:
                self.log(f"导出失败: {e}", is_error=True)
                QMessageBox.critical(self, "导出失败", f"导出过程中发生错误: {e}")

    def clear_results(self):
        """清空查询结果"""
        self.bug_data = []
        self.bug_table.setRowCount(0)
        self.result_label.setText("查询结果: 0 条记录")
        self.export_btn.setEnabled(False)
        self.log("查询结果已清空")

    def log(self, message, is_error=False, clear=False):
        """添加日志"""
        if clear:
            self.log_output.clear()

        timestamp = datetime.now().strftime("[%H:%M:%S]")
        formatted_message = f"{timestamp} {message}"

        cursor = self.log_output.textCursor()
        format = cursor.charFormat()
        if is_error:
            format.setForeground(Qt.red)
        else:
            format.setForeground(Qt.black)
        cursor.setCharFormat(format)

        self.log_output.append(formatted_message)
        self.log_output.ensureCursorVisible()

    def save_settings(self):
        """保存设置"""
        settings = {
            "manager_account": self.manager_account_input.text(),
            "manager_password": self.manager_password_input.text(),  # 注意：实际应用中应该加密存储
            "product_name": self.product_name_input.text(),
            "status_index": self.status_combo.currentIndex(),
            "severity_index": self.severity_combo.currentIndex(),
            "include_resolved": self.include_resolved_cb.isChecked(),
            "include_closed": self.include_closed_cb.isChecked()
        }
        self.settings_manager.save_settings("bug_query", settings, self.log)

    def load_settings(self):
        """加载设置"""
        default_settings = {
            "manager_account": "",
            "manager_password": "",
            "product_name": "",
            "status_index": 0,
            "severity_index": 0,
            "include_resolved": True,
            "include_closed": False
        }

        loaded_settings = self.settings_manager.load_settings(
            "bug_query", default_settings, self.log
        )

        self.manager_account_input.setText(loaded_settings.get("manager_account", ""))
        self.manager_password_input.setText(loaded_settings.get("manager_password", ""))
        self.product_name_input.setText(loaded_settings.get("product_name", ""))
        self.status_combo.setCurrentIndex(loaded_settings.get("status_index", 0))
        self.severity_combo.setCurrentIndex(loaded_settings.get("severity_index", 0))
        self.include_resolved_cb.setChecked(loaded_settings.get("include_resolved", True))
        self.include_closed_cb.setChecked(loaded_settings.get("include_closed", False))
