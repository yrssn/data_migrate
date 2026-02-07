"""
主体编号补全功能界面
"""
import re
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QCheckBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from database import DatabaseManager
import pymysql


class ShopIndexCompleteWorker(QThread):
    """主体编号补全工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)  # 返回统计结果
    error = pyqtSignal(str)
    
    def __init__(self, datasource, db_manager, process_customer, process_legal_info, process_legal_part2):
        super().__init__()
        self.datasource = datasource
        self.db_manager = db_manager
        self.process_customer = process_customer
        self.process_legal_info = process_legal_info
        self.process_legal_part2 = process_legal_part2
        self.results = {
            'customer_p_added': 0,
            'customer_s_added': 0,
            'legal_info_added': 0,
            'legal_part2_added': 0,
            'customer_p_skipped': 0,
            'customer_s_skipped': 0,
            'legal_info_skipped': 0,
            'legal_part2_skipped': 0,
            'failed': 0
        }
    
    def run(self):
        try:
            self.log_message.emit("开始连接数据库...")
            self.progress.emit(5)
            
            # 连接数据库
            connection = pymysql.connect(
                host=self.datasource.host,
                port=self.datasource.port,
                user=self.datasource.username,
                password=self.datasource.password,
                database=self.datasource.database,
                charset=self.datasource.charset
            )
            
            cursor = connection.cursor()
            self.log_message.emit("数据库连接成功")
            self.progress.emit(10)
            
            current_time = int(datetime.now().timestamp())
            
            # 处理ba_rlb_customer表
            if self.process_customer:
                self.log_message.emit("开始处理ba_rlb_customer表...")
                self.process_customer_table(cursor, current_time)
                self.progress.emit(40)
            
            # 处理ba_rlb_legal_information表
            if self.process_legal_info:
                self.log_message.emit("开始处理ba_rlb_legal_information表...")
                self.process_legal_information_table(cursor, current_time)
                self.progress.emit(70)
            
            # 处理ba_rlb_legal_information_part2表
            if self.process_legal_part2:
                self.log_message.emit("开始处理ba_rlb_legal_information_part2表...")
                self.process_legal_information_part2_table(cursor, current_time)
                self.progress.emit(90)
            
            # 提交事务
            connection.commit()
            cursor.close()
            connection.close()
            
            self.progress.emit(100)
            self.finished.emit(self.results)
            
        except Exception as e:
            self.error.emit(str(e))
    
    def get_next_shopindex_id(self, cursor, prefix):
        """获取下一个编号"""
        cursor.execute("""
            SELECT shopindex_id FROM ba_shopindex 
            WHERE shopindex_id LIKE %s 
            ORDER BY shopindex_id DESC LIMIT 1
        """, (f"{prefix}%",))
        
        result = cursor.fetchone()
        if result:
            # 提取数字部分
            current_id = result[0]
            match = re.search(r'(\d+)$', current_id)
            if match:
                current_num = int(match.group(1))
                next_num = current_num + 1
                return f"{prefix}{next_num:05d}"
        
        # 如果没有找到，从00001开始
        return f"{prefix}00001"
    
    def process_customer_table(self, cursor, current_time):
        """处理ba_rlb_customer表"""
        # 查找所有客户记录（只处理is_customer=1的记录）
        cursor.execute("""
            SELECT id, admin_id, admin_dept_id 
            FROM ba_rlb_customer 
            WHERE (delete_time IS NULL OR delete_time = 0) AND is_customer = 1
        """)
        customers = cursor.fetchall()
        
        self.log_message.emit(f"找到 {len(customers)} 个客户记录")
        
        for customer in customers:
            customer_id, admin_id, admin_dept_id = customer
            
            try:
                # 确保参数不为None
                if admin_id is None:
                    admin_id = 1
                if admin_dept_id is None:
                    admin_dept_id = 1
                
                # 检查是否已存在P开头的记录
                cursor.execute("""
                    SELECT id FROM ba_shopindex 
                    WHERE legal_id = %s AND (belong_information IS NULL OR belong_information = '') AND shopindex_id LIKE 'P%%'
                """, (customer_id,))
                
                if not cursor.fetchone():
                    # 获取下一个P编号
                    p_shopindex_id = self.get_next_shopindex_id(cursor, "P")
                    
                    # 插入P记录
                    cursor.execute("""
                        INSERT INTO ba_shopindex 
                        (shopindex_id, `select`, admin_id, create_time, update_time, remark, status, legal_id, admin_dept_id, origin_customer_id)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """, (p_shopindex_id, 'opt0', admin_id, current_time, current_time, '脚本补全', 1, customer_id, admin_dept_id, customer_id))
                    
                    self.results['customer_p_added'] += 1
                    self.log_message.emit(f"为客户ID {customer_id} 添加P编号: {p_shopindex_id}")
                else:
                    self.results['customer_p_skipped'] += 1
                
                # 检查是否已存在S开头的记录
                cursor.execute("""
                    SELECT id FROM ba_shopindex 
                    WHERE legal_id = %s AND (belong_information IS NULL OR belong_information = '') AND shopindex_id LIKE 'S%%'
                """, (customer_id,))
                
                if not cursor.fetchone():
                    # 获取下一个S编号
                    s_shopindex_id = self.get_next_shopindex_id(cursor, "S")
                    
                    # 插入S记录
                    cursor.execute("""
                        INSERT INTO ba_shopindex 
                        (shopindex_id, `select`, admin_id, create_time, update_time, remark, status, legal_id, admin_dept_id, origin_customer_id)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """, (s_shopindex_id, 'opt1', admin_id, current_time, current_time, '脚本补全', 1, customer_id, admin_dept_id, customer_id))
                    
                    self.results['customer_s_added'] += 1
                    self.log_message.emit(f"为客户ID {customer_id} 添加S编号: {s_shopindex_id}")
                else:
                    self.results['customer_s_skipped'] += 1
                    
            except Exception as e:
                self.results['failed'] += 1
                self.log_message.emit(f"处理客户ID {customer_id} 时出错: {str(e)}")
    
    def process_legal_information_table(self, cursor, current_time):
        """处理ba_rlb_legal_information表"""
        # 查找所有法人信息记录
        cursor.execute("""
            SELECT id, country, admin_id 
            FROM ba_rlb_legal_information 
            WHERE delete_time IS NULL OR delete_time = 0
        """)
        legal_infos = cursor.fetchall()
        
        self.log_message.emit(f"找到 {len(legal_infos)} 个法人信息记录")
        
        for legal_info in legal_infos:
            legal_info_id, country, admin_id = legal_info
            
            try:
                # 确保参数不为None
                if admin_id is None:
                    admin_id = 1
                if country is None:
                    country = ''
                
                # 检查是否已存在belong_information=1的记录
                cursor.execute("""
                    SELECT id FROM ba_shopindex 
                    WHERE legal_id = %s AND belong_information = 1
                """, (legal_info_id,))
                
                if not cursor.fetchone():
                    # 获取下一个C编号
                    c_shopindex_id = self.get_next_shopindex_id(cursor, "C")
                    
                    # 获取admin_dept_id (如果有的话)
                    admin_dept_id = 1  # 默认值
                    
                    # 插入C记录
                    cursor.execute("""
                        INSERT INTO ba_shopindex 
                        (shopindex_id, country, `select`, admin_id, create_time, update_time, remark, status, legal_id, belong_information, admin_dept_id)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """, (c_shopindex_id, country, 'opt2', admin_id, current_time, current_time, '脚本补全', 1, legal_info_id, 1, admin_dept_id))
                    
                    self.results['legal_info_added'] += 1
                    self.log_message.emit(f"为法人信息ID {legal_info_id} 添加C编号: {c_shopindex_id}")
                else:
                    self.results['legal_info_skipped'] += 1
                    
            except Exception as e:
                self.results['failed'] += 1
                self.log_message.emit(f"处理法人信息ID {legal_info_id} 时出错: {str(e)}")
    
    def process_legal_information_part2_table(self, cursor, current_time):
        """处理ba_rlb_legal_information_part2表"""
        # 查找所有法人信息part2记录
        cursor.execute("""
            SELECT id, country, admin_id 
            FROM ba_rlb_legal_information_part2 
            WHERE delete_time IS NULL OR delete_time = 0
        """)
        legal_part2s = cursor.fetchall()
        
        self.log_message.emit(f"找到 {len(legal_part2s)} 个法人信息part2记录")
        
        for legal_part2 in legal_part2s:
            legal_part2_id, country, admin_id = legal_part2
            
            try:
                # 确保参数不为None
                if admin_id is None:
                    admin_id = 1
                if country is None:
                    country = ''
                
                # 检查是否已存在belong_information=2的记录
                cursor.execute("""
                    SELECT id FROM ba_shopindex 
                    WHERE legal_id = %s AND belong_information = 2
                """, (legal_part2_id,))
                
                if not cursor.fetchone():
                    # 获取下一个C编号
                    c_shopindex_id = self.get_next_shopindex_id(cursor, "C")
                    
                    # 获取admin_dept_id (如果有的话)
                    admin_dept_id = 1  # 默认值
                    
                    # 插入C记录
                    cursor.execute("""
                        INSERT INTO ba_shopindex 
                        (shopindex_id, country, `select`, admin_id, create_time, update_time, remark, status, legal_id, belong_information, admin_dept_id)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """, (c_shopindex_id, country, 'opt2', admin_id, current_time, current_time, '脚本补全', 1, legal_part2_id, 2, admin_dept_id))
                    
                    self.results['legal_part2_added'] += 1
                    self.log_message.emit(f"为法人信息part2 ID {legal_part2_id} 添加C编号: {c_shopindex_id}")
                else:
                    self.results['legal_part2_skipped'] += 1
                    
            except Exception as e:
                self.results['failed'] += 1
                self.log_message.emit(f"处理法人信息part2 ID {legal_part2_id} 时出错: {str(e)}")


class ShopIndexCompleteWidget(QWidget):
    """主体编号补全功能组件"""
    
    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.init_ui()
        self.load_datasources()
    
    def init_ui(self):
        """初始化UI"""
        layout = QVBoxLayout()
        
        # 数据源选择区域
        datasource_group = QGroupBox("选择目标数据库")
        datasource_layout = QFormLayout()
        
        self.datasource_combo = QComboBox()
        self.datasource_combo.currentTextChanged.connect(self.on_datasource_changed)
        datasource_layout.addRow("数据源:", self.datasource_combo)
        
        self.test_connection_btn = QPushButton("测试连接")
        self.test_connection_btn.clicked.connect(self.test_connection)
        self.test_connection_btn.setEnabled(False)
        datasource_layout.addRow("", self.test_connection_btn)
        
        datasource_group.setLayout(datasource_layout)
        layout.addWidget(datasource_group)
        
        # 处理选项区域
        options_group = QGroupBox("选择要处理的表")
        options_layout = QVBoxLayout()
        
        # 说明文字
        info_label = QLabel("""
<b>主体编号补全说明:</b><br>
• <b>ba_rlb_customer表</b>: 为每个客户补全P和S开头的编号<br>
  &nbsp;&nbsp;- P编号: 个人类型 (select='opt0')<br>
  &nbsp;&nbsp;- S编号: 个人事业主类型 (select='opt1')<br>
• <b>ba_rlb_legal_information表</b>: 补全C开头的编号 (belong_information=1)<br>
• <b>ba_rlb_legal_information_part2表</b>: 补全C开头的编号 (belong_information=2)<br><br>
<b>编号规则:</b><br>
• 自动获取当前最大编号，递增生成新编号<br>
• 格式: P00001, S00001, C00001 等<br>
• 已存在的记录会被跳过<br>
• 备注字段统一填写为"脚本补全"
        """)
        info_label.setWordWrap(True)
        options_layout.addWidget(info_label)
        
        # 复选框选项
        self.customer_checkbox = QCheckBox("处理 ba_rlb_customer 表")
        self.customer_checkbox.setChecked(True)
        options_layout.addWidget(self.customer_checkbox)
        
        self.legal_info_checkbox = QCheckBox("处理 ba_rlb_legal_information 表")
        self.legal_info_checkbox.setChecked(True)
        options_layout.addWidget(self.legal_info_checkbox)
        
        self.legal_part2_checkbox = QCheckBox("处理 ba_rlb_legal_information_part2 表")
        self.legal_part2_checkbox.setChecked(True)
        options_layout.addWidget(self.legal_part2_checkbox)
        
        options_group.setLayout(options_layout)
        layout.addWidget(options_group)
        
        # 操作按钮区域
        button_layout = QHBoxLayout()
        
        self.start_btn = QPushButton("开始补全")
        self.start_btn.clicked.connect(self.start_complete)
        self.start_btn.setEnabled(False)
        button_layout.addWidget(self.start_btn)
        
        button_layout.addStretch()
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        button_layout.addWidget(self.progress_bar)
        
        layout.addLayout(button_layout)
        
        # 日志区域
        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout()
        
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(200)
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        
        log_button_layout = QHBoxLayout()
        self.clear_log_btn = QPushButton("清空日志")
        self.clear_log_btn.clicked.connect(self.clear_log)
        log_button_layout.addWidget(self.clear_log_btn)
        
        log_button_layout.addStretch()
        log_layout.addLayout(log_button_layout)
        
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)
        
        # 结果显示区域
        self.result_label = QLabel("")
        self.result_label.setAlignment(Qt.AlignCenter)
        font = QFont()
        font.setBold(True)
        self.result_label.setFont(font)
        layout.addWidget(self.result_label)
        
        self.setLayout(layout)
    
    def load_datasources(self):
        """加载数据源列表"""
        self.datasource_combo.clear()
        self.datasource_combo.addItem("请选择数据源", None)
        
        datasources = self.db_manager.get_all_datasources()
        for ds in datasources:
            self.datasource_combo.addItem(f"{ds.name} ({ds.host}:{ds.port})", ds)
    
    def on_datasource_changed(self):
        """数据源选择变化"""
        current_data = self.datasource_combo.currentData()
        self.test_connection_btn.setEnabled(current_data is not None)
        self.update_start_button_state()
    
    def test_connection(self):
        """测试数据库连接"""
        datasource = self.datasource_combo.currentData()
        if not datasource:
            return
        
        success, message = self.db_manager.test_connection(datasource)
        if success:
            QMessageBox.information(self, "连接成功", message)
        else:
            QMessageBox.critical(self, "连接失败", message)
    
    def update_start_button_state(self):
        """更新开始按钮状态"""
        has_datasource = self.datasource_combo.currentData() is not None
        has_selection = (self.customer_checkbox.isChecked() or 
                        self.legal_info_checkbox.isChecked() or 
                        self.legal_part2_checkbox.isChecked())
        self.start_btn.setEnabled(has_datasource and has_selection)
    
    def start_complete(self):
        """开始补全"""
        datasource = self.datasource_combo.currentData()
        if not datasource:
            QMessageBox.warning(self, "警告", "请选择数据源！")
            return
        
        if not (self.customer_checkbox.isChecked() or 
                self.legal_info_checkbox.isChecked() or 
                self.legal_part2_checkbox.isChecked()):
            QMessageBox.warning(self, "警告", "请至少选择一个要处理的表！")
            return
        
        # 确认对话框
        tables = []
        if self.customer_checkbox.isChecked():
            tables.append("ba_rlb_customer")
        if self.legal_info_checkbox.isChecked():
            tables.append("ba_rlb_legal_information")
        if self.legal_part2_checkbox.isChecked():
            tables.append("ba_rlb_legal_information_part2")
        
        reply = QMessageBox.question(
            self,
            "确认补全",
            f"确定要为以下表补全主体编号吗？\n\n"
            f"数据库: {datasource.name}\n"
            f"表: {', '.join(tables)}\n\n"
            f"此操作将自动为缺失的记录生成编号。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        # 禁用按钮，显示进度条
        self.start_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在补全中...")
        self.log_text.clear()
        
        # 启动工作线程
        self.worker = ShopIndexCompleteWorker(
            datasource, 
            self.db_manager,
            self.customer_checkbox.isChecked(),
            self.legal_info_checkbox.isChecked(),
            self.legal_part2_checkbox.isChecked()
        )
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log_message.connect(self.append_log)
        self.worker.finished.connect(self.on_complete_finished)
        self.worker.error.connect(self.on_complete_error)
        self.worker.start()
    
    def append_log(self, message):
        """添加日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
    
    def on_complete_finished(self, results):
        """补全完成"""
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        result_text = (
            f"补全完成！\n"
            f"客户P编号: 新增 {results['customer_p_added']} 个，跳过 {results['customer_p_skipped']} 个\n"
            f"客户S编号: 新增 {results['customer_s_added']} 个，跳过 {results['customer_s_skipped']} 个\n"
            f"法人信息编号: 新增 {results['legal_info_added']} 个，跳过 {results['legal_info_skipped']} 个\n"
            f"法人信息part2编号: 新增 {results['legal_part2_added']} 个，跳过 {results['legal_part2_skipped']} 个\n"
            f"失败: {results['failed']} 个"
        )
        
        self.result_label.setText(result_text)
        self.result_label.setStyleSheet("color: green;")
        
        QMessageBox.information(self, "补全完成", result_text)
    
    def on_complete_error(self, error_msg):
        """补全错误"""
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.result_label.setText("补全失败！")
        self.result_label.setStyleSheet("color: red;")
        
        QMessageBox.critical(self, "补全错误", f"补全失败：{error_msg}")
    
    def clear_log(self):
        """清空日志"""
        self.log_text.clear()
