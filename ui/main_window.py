# ui/main_window.py - 修改后的主窗口

import os
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QTabWidget, QMessageBox, QSplitter
from PyQt5.QtCore import Qt
from ui.zentao_export_page import ZentaoExportPage
from ui.acceptance_filling_page import AcceptanceTestFillingPage
from ui.ExcelTool import ExcelTool
from ui.data_chart_page import ZentaoDataChartPage
from ui.user_info_widget import UserInfoWidget
from ui.bug_query_page import BugQueryPage
from core.settings_manager import SettingsManager


class MainApplication(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("XD_自动化报告生成_V2.0")
        self.setGeometry(100, 100, 1200, 950)  # 增加宽度以适应用户信息面板

        self.settings_manager = SettingsManager()
        self.current_user_info = None  # 当前登录用户信息

        self._init_ui_components()
        self._setup_layout()
        self._connect_signals()
        self._load_all_settings()

    def _init_ui_components(self):
        """初始化UI组件"""
        # 创建标签页组件
        self.tabs = QTabWidget()

        # 创建用户信息面板
        self.user_info_widget = UserInfoWidget(self)

        # 初始化各个页面
        self.zentao_export_page = ZentaoExportPage(self)
        self.data_chart_page = ZentaoDataChartPage(self)
        self.excel_tool = ExcelTool()
        self.bug_query_page = BugQueryPage(self)  # 新增历史BUG查询页面

        # 添加标签页
        self.tabs.addTab(self.zentao_export_page, "禅道自动化导出")
        self.tabs.addTab(self.data_chart_page, "禅道数据表单与验收图插入")
        self.tabs.addTab(self.excel_tool, "验收测试结果填充")
        # 历史BUG查询页面默认隐藏，登录后显示
        self.bug_query_tab_index = self.tabs.addTab(self.bug_query_page, "历史BUG查询")
        self.tabs.setTabEnabled(self.bug_query_tab_index, False)  # 默认禁用

    def _setup_layout(self):
        """设置布局"""
        main_layout = QHBoxLayout()

        # 创建水平分割器
        splitter = QSplitter(Qt.Horizontal)

        # 左侧主要内容区域
        main_content = QWidget()
        main_content_layout = QVBoxLayout(main_content)
        main_content_layout.addWidget(self.tabs)

        # 右侧用户信息面板
        self.user_info_widget.setFixedWidth(280)  # 固定宽度

        # 添加到分割器
        splitter.addWidget(main_content)
        splitter.addWidget(self.user_info_widget)

        # 设置分割器比例
        splitter.setStretchFactor(0, 4)  # 主内容区占4份
        splitter.setStretchFactor(1, 1)  # 用户信息区占1份

        main_layout.addWidget(splitter)
        self.setLayout(main_layout)

    def _connect_signals(self):
        """连接信号槽"""
        # 连接禅道导出页面的用户信息信号
        self.zentao_export_page.user_logged_in.connect(self._on_user_logged_in)

        # 连接用户信息面板的刷新信号
        self.user_info_widget.refresh_requested.connect(self._refresh_user_info)

    def _on_user_logged_in(self, user_info):
        """用户登录成功处理"""
        self.current_user_info = user_info

        # 更新用户信息显示
        self.user_info_widget.update_user_info(user_info)

        # 启用历史BUG查询页面
        self.tabs.setTabEnabled(self.bug_query_tab_index, True)

        # 将用户信息传递给BUG查询页面
        self.bug_query_page.set_user_info(user_info)

        # 更新窗口标题
        self.setWindowTitle(f"XD_自动化报告生成_V2.0 - {user_info.real_name} ({user_info.account})")

    def _refresh_user_info(self):
        """刷新用户信息"""
        if self.current_user_info:
            # 重新获取用户信息
            self.zentao_export_page.refresh_user_info()
        else:
            QMessageBox.information(self, "提示", "请先登录禅道系统")

    def _load_all_settings(self):
        """加载所有页面的设置"""
        self.zentao_export_page.load_settings()
        self.data_chart_page.load_settings()
        self.bug_query_page.load_settings()

    def closeEvent(self, event):
        """处理窗口关闭事件"""
        is_zentao_running = self.zentao_export_page and \
                            self.zentao_export_page.worker_thread and \
                            self.zentao_export_page.worker_thread.isRunning()

        is_excel_running = self.data_chart_page and \
                           self.data_chart_page.excel_worker_thread and \
                           self.data_chart_page.excel_worker_thread.isRunning()

        is_bug_query_running = self.bug_query_page and \
                               self.bug_query_page.bug_query_worker and \
                               self.bug_query_page.bug_query_worker.isRunning()

        if is_zentao_running or is_excel_running or is_bug_query_running:
            running_tasks = []
            if is_zentao_running:
                running_tasks.append("禅道自动化导出")
            if is_excel_running:
                running_tasks.append("Excel 处理")
            if is_bug_query_running:
                running_tasks.append("BUG查询")

            task_name = "、".join(running_tasks)

            reply = QMessageBox.question(self, '退出确认',
                                         f"{task_name} 任务正在运行，确定要退出并停止任务吗？",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                if is_zentao_running:
                    self.zentao_export_page._cancel_export()
                if is_excel_running:
                    # Excel处理任务强制结束
                    pass
                if is_bug_query_running:
                    # BUG查询任务强制结束
                    pass
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()
