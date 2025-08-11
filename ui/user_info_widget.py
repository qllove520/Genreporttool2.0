# ui/user_info_widget.py - 更新后的用户信息显示组件

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QGroupBox, QFrame, QPushButton
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QPixmap, QPalette
from datetime import datetime


class UserInfoWidget(QWidget):
    """用户信息显示组件 - 支持6个字段"""
    refresh_requested = pyqtSignal()  # 刷新用户信息信号

    def __init__(self, parent=None):
        super().__init__(parent)
        self.user_info = None
        self.is_logged_in = False
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)

        # 用户信息组框
        self.user_group = QGroupBox("用户信息")
        self.user_group.setVisible(False)  # 默认隐藏，登录后显示
        user_layout = QVBoxLayout()

        # 用户头像和基本信息区域
        header_layout = QHBoxLayout()

        # 头像区域
        self.avatar_label = QLabel("👤")
        self.avatar_label.setFixedSize(60, 60)
        self.avatar_label.setAlignment(Qt.AlignCenter)
        self.avatar_label.setStyleSheet("""
            QLabel {
                border: 2px solid #ccc;
                border-radius: 30px;
                background-color: #f0f0f0;
                font-size: 24px;
            }
        """)
        header_layout.addWidget(self.avatar_label)

        # 基本信息区域
        info_layout = QVBoxLayout()

        self.name_label = QLabel("未登录")
        self.name_label.setFont(QFont("Arial", 12, QFont.Bold))
        info_layout.addWidget(self.name_label)

        self.account_label = QLabel("")
        self.account_label.setStyleSheet("color: #666; font-size: 10px;")
        info_layout.addWidget(self.account_label)

        self.status_label = QLabel("离线")
        self.status_label.setStyleSheet("color: #999; font-size: 10px;")
        info_layout.addWidget(self.status_label)

        header_layout.addLayout(info_layout)
        header_layout.addStretch()

        user_layout.addLayout(header_layout)

        # 分隔线
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        user_layout.addWidget(line)

        # 详细信息区域 - 6个必要字段
        details_layout = QVBoxLayout()

        self.dept_label = QLabel("部门: --")
        self.dept_label.setWordWrap(True)  # 支持换行，适应长部门名
        details_layout.addWidget(self.dept_label)

        self.position_label = QLabel("职位: --")
        details_layout.addWidget(self.position_label)

        self.role_label = QLabel("权限: --")
        self.role_label.setWordWrap(True)  # 权限可能很长
        details_layout.addWidget(self.role_label)

        self.last_login_label = QLabel("最后登录: --")
        self.last_login_label.setWordWrap(True)
        details_layout.addWidget(self.last_login_label)

        user_layout.addLayout(details_layout)

        # 操作按钮区域
        button_layout = QHBoxLayout()

        self.refresh_btn = QPushButton("刷新")
        self.refresh_btn.clicked.connect(self.refresh_requested.emit)
        self.refresh_btn.setFixedHeight(25)
        self.refresh_btn.setFixedWidth(50)
        button_layout.addWidget(self.refresh_btn)

        self.logout_btn = QPushButton("登出")
        self.logout_btn.clicked.connect(self.clear_user_info)
        self.logout_btn.setFixedHeight(25)
        self.logout_btn.setFixedWidth(50)
        button_layout.addWidget(self.logout_btn)

        button_layout.addStretch()
        user_layout.addLayout(button_layout)

        self.user_group.setLayout(user_layout)
        main_layout.addWidget(self.user_group)

        # 未登录提示
        self.login_prompt = QLabel("请先登录禅道系统")
        self.login_prompt.setAlignment(Qt.AlignCenter)
        self.login_prompt.setStyleSheet("""
            QLabel {
                padding: 20px;
                color: #999;
                font-size: 14px;
            }
        """)
        main_layout.addWidget(self.login_prompt)

        main_layout.addStretch()

    def update_user_info(self, user_info):
        """更新用户信息显示 - 支持新的6个字段"""
        self.user_info = user_info
        self.is_logged_in = True

        # 显示用户信息组框，隐藏登录提示
        self.user_group.setVisible(True)
        self.login_prompt.setVisible(False)

        # 更新头部信息
        display_name = user_info.real_name or user_info.account
        self.name_label.setText(display_name)
        self.account_label.setText(f"@{user_info.account}")
        self.status_label.setText("在线")
        self.status_label.setStyleSheet("color: #4CAF50; font-size: 10px;")

        # 更新详细信息 - 使用新的字段结构
        # 处理部门信息（可能包含层级结构）
        dept_text = user_info.department or '--'
        if len(dept_text) > 20:  # 如果部门名称太长，进行换行显示
            dept_text = dept_text.replace(' > ', '\n  > ')
        self.dept_label.setText(f"部门: {dept_text}")

        self.position_label.setText(f"职位: {user_info.position or '--'}")

        # 处理权限信息（可能很长）
        role_text = user_info.role or '--'
        if len(role_text) > 15:  # 如果权限文本太长，进行适当处理
            role_text = role_text.replace(' ', '\n')
        self.role_label.setText(f"权限: {role_text}")

        self.last_login_label.setText(f"最后登录:\n{user_info.last_login or '--'}")

        # 根据权限设置不同的头像和样式
        self._update_avatar_style(user_info.role or "")

        # 设置工具提示，显示完整信息
        self._set_tooltips(user_info)

    def _update_avatar_style(self, role):
        """根据用户权限更新头像样式"""
        if any(keyword in role for keyword in ["管理", "admin", "Admin"]):
            self.avatar_label.setText("👨‍💼")
            self.avatar_label.setStyleSheet("""
                QLabel {
                    border: 2px solid #FF9800;
                    border-radius: 30px;
                    background-color: #FFF3E0;
                    font-size: 24px;
                }
            """)
        elif any(keyword in role for keyword in ["工程师", "测试", "开发"]):
            self.avatar_label.setText("👨‍💻")
            self.avatar_label.setStyleSheet("""
                QLabel {
                    border: 2px solid #2196F3;
                    border-radius: 30px;
                    background-color: #E3F2FD;
                    font-size: 24px;
                }
            """)
        elif any(keyword in role for keyword in ["经理", "主管"]):
            self.avatar_label.setText("👔")
            self.avatar_label.setStyleSheet("""
                QLabel {
                    border: 2px solid #4CAF50;
                    border-radius: 30px;
                    background-color: #E8F5E8;
                    font-size: 24px;
                }
            """)
        else:
            self.avatar_label.setText("👤")
            self.avatar_label.setStyleSheet("""
                QLabel {
                    border: 2px solid #9E9E9E;
                    border-radius: 30px;
                    background-color: #F5F5F5;
                    font-size: 24px;
                }
            """)

    def _set_tooltips(self, user_info):
        """设置工具提示显示完整信息"""
        tooltip_text = f"""用户详细信息:
用户名: {user_info.account}
真实姓名: {user_info.real_name or '--'}
所属部门: {user_info.department or '--'}
职位: {user_info.position or '--'}
权限: {user_info.role or '--'}
最后登录: {user_info.last_login or '--'}"""

        self.user_group.setToolTip(tooltip_text)

    def clear_user_info(self):
        """清除用户信息"""
        self.user_info = None
        self.is_logged_in = False

        # 隐藏用户信息组框，显示登录提示
        self.user_group.setVisible(False)
        self.login_prompt.setVisible(True)

        # 重置显示内容
        self.name_label.setText("未登录")
        self.account_label.setText("")
        self.status_label.setText("离线")
        self.status_label.setStyleSheet("color: #999; font-size: 10px;")

        # 重置头像
        self.avatar_label.setText("👤")
        self.avatar_label.setStyleSheet("""
            QLabel {
                border: 2px solid #ccc;
                border-radius: 30px;
                background-color: #f0f0f0;
                font-size: 24px;
            }
        """)

    def get_user_info(self):
        """获取当前用户信息"""
        return self.user_info

    def is_user_logged_in(self):
        """检查用户是否已登录"""
        return self.is_logged_in

    def get_user_summary(self):
        """获取用户信息摘要文本"""
        if not self.user_info:
            return "未登录"

        return f"{self.user_info.real_name} ({self.user_info.account}) - {self.user_info.position}"
