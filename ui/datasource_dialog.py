"""
数据源管理对话框
"""
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QLineEdit, QSpinBox, QTextEdit, QPushButton,
                             QLabel, QMessageBox, QGroupBox)
from PyQt5.QtCore import Qt
from models import DataSource
from database import DatabaseManager


class DataSourceDialog(QDialog):
    """数据源添加/编辑对话框"""
    
    def __init__(self, parent=None, datasource: DataSource = None, db_manager: DatabaseManager = None):
        super().__init__(parent)
        self.datasource = datasource or DataSource()
        self.db_manager = db_manager
        self.is_edit_mode = datasource is not None and datasource.id is not None
        
        self.init_ui()
        self.load_data()
    
    def init_ui(self):
        """初始化UI"""
        self.setWindowTitle("编辑数据源" if self.is_edit_mode else "添加数据源")
        self.setMinimumWidth(500)
        
        layout = QVBoxLayout()
        
        # 基本信息组
        basic_group = QGroupBox("基本信息")
        basic_layout = QFormLayout()
        
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("数据源名称（唯一标识）")
        basic_layout.addRow("名称*:", self.name_edit)
        
        self.host_edit = QLineEdit()
        self.host_edit.setPlaceholderText("localhost 或 IP地址")
        basic_layout.addRow("主机*:", self.host_edit)
        
        self.port_spin = QSpinBox()
        self.port_spin.setRange(1, 65535)
        self.port_spin.setValue(3306)
        basic_layout.addRow("端口*:", self.port_spin)
        
        self.database_edit = QLineEdit()
        self.database_edit.setPlaceholderText("数据库名称")
        basic_layout.addRow("数据库*:", self.database_edit)
        
        basic_group.setLayout(basic_layout)
        layout.addWidget(basic_group)
        
        # 认证信息组
        auth_group = QGroupBox("认证信息")
        auth_layout = QFormLayout()
        
        self.username_edit = QLineEdit()
        self.username_edit.setPlaceholderText("MySQL用户名")
        auth_layout.addRow("用户名*:", self.username_edit)
        
        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.Password)
        self.password_edit.setPlaceholderText("MySQL密码")
        auth_layout.addRow("密码*:", self.password_edit)
        
        auth_group.setLayout(auth_layout)
        layout.addWidget(auth_group)
        
        # 高级选项组
        advanced_group = QGroupBox("高级选项")
        advanced_layout = QFormLayout()
        
        self.charset_edit = QLineEdit()
        self.charset_edit.setText("utf8mb4")
        self.charset_edit.setPlaceholderText("字符集")
        advanced_layout.addRow("字符集:", self.charset_edit)
        
        self.description_edit = QTextEdit()
        self.description_edit.setPlaceholderText("数据源描述信息（可选）")
        self.description_edit.setMaximumHeight(80)
        advanced_layout.addRow("描述:", self.description_edit)
        
        advanced_group.setLayout(advanced_layout)
        layout.addWidget(advanced_group)
        
        # 按钮组
        button_layout = QHBoxLayout()
        
        self.test_btn = QPushButton("测试连接")
        self.test_btn.clicked.connect(self.test_connection)
        button_layout.addWidget(self.test_btn)
        
        button_layout.addStretch()
        
        self.save_btn = QPushButton("保存")
        self.save_btn.clicked.connect(self.save_datasource)
        self.save_btn.setDefault(True)
        button_layout.addWidget(self.save_btn)
        
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(self.cancel_btn)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
    
    def load_data(self):
        """加载数据到表单"""
        if self.datasource:
            self.name_edit.setText(self.datasource.name)
            self.host_edit.setText(self.datasource.host)
            self.port_spin.setValue(self.datasource.port)
            self.username_edit.setText(self.datasource.username)
            self.password_edit.setText(self.datasource.password)
            self.database_edit.setText(self.datasource.database)
            self.charset_edit.setText(self.datasource.charset)
            self.description_edit.setPlainText(self.datasource.description)
    
    def get_datasource(self) -> DataSource:
        """从表单获取数据源对象"""
        datasource = DataSource(
            id=self.datasource.id if self.is_edit_mode else None,
            name=self.name_edit.text().strip(),
            host=self.host_edit.text().strip(),
            port=self.port_spin.value(),
            username=self.username_edit.text().strip(),
            password=self.password_edit.text(),
            database=self.database_edit.text().strip(),
            charset=self.charset_edit.text().strip() or "utf8mb4",
            description=self.description_edit.toPlainText().strip()
        )
        return datasource
    
    def validate_form(self) -> tuple[bool, str]:
        """验证表单"""
        if not self.name_edit.text().strip():
            return False, "请输入数据源名称"
        if not self.host_edit.text().strip():
            return False, "请输入主机地址"
        if not self.username_edit.text().strip():
            return False, "请输入用户名"
        if not self.password_edit.text():
            return False, "请输入密码"
        if not self.database_edit.text().strip():
            return False, "请输入数据库名称"
        return True, ""
    
    def test_connection(self):
        """测试数据库连接"""
        valid, msg = self.validate_form()
        if not valid:
            QMessageBox.warning(self, "验证失败", msg)
            return
        
        datasource = self.get_datasource()
        
        self.test_btn.setEnabled(False)
        self.test_btn.setText("测试中...")
        
        success, message = self.db_manager.test_connection(datasource)
        
        self.test_btn.setEnabled(True)
        self.test_btn.setText("测试连接")
        
        if success:
            QMessageBox.information(self, "连接成功", message)
        else:
            QMessageBox.critical(self, "连接失败", message)
    
    def save_datasource(self):
        """保存数据源"""
        valid, msg = self.validate_form()
        if not valid:
            QMessageBox.warning(self, "验证失败", msg)
            return
        
        datasource = self.get_datasource()
        
        try:
            if self.is_edit_mode:
                success = self.db_manager.update_datasource(datasource)
                if success:
                    QMessageBox.information(self, "成功", "数据源更新成功！")
                    self.accept()
                else:
                    QMessageBox.warning(self, "失败", "数据源更新失败！")
            else:
                datasource_id = self.db_manager.add_datasource(datasource)
                if datasource_id:
                    QMessageBox.information(self, "成功", "数据源添加成功！")
                    self.accept()
                else:
                    QMessageBox.warning(self, "失败", "数据源添加失败！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存失败：{str(e)}")
