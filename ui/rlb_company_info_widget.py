"""
日联部企业基础信息表补全功能界面
"""
import os
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QTableWidget,
                             QTableWidgetItem, QHeaderView, QCheckBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from database import DatabaseManager
import pymysql


class RlbCompanyInfoWorker(QThread):
    """日联部企业基础信息补全工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    
    def __init__(self, datasource, db_manager, selected_records, fill_mode='company_info'):
        super().__init__()
        self.datasource = datasource
        self.db_manager = db_manager
        self.selected_records = selected_records  # 用户选择的记录ID列表
        self.fill_mode = fill_mode  # 'company_info' 或 'address_other'
        self.results = {
            'total': 0,
            'success': 0,
            'failed': 0,
            'error_records': []
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
            
            self.results['total'] = len(self.selected_records)
            
            # 处理每个选中的记录
            for index, record_id in enumerate(self.selected_records):
                try:
                    if self.fill_mode == 'company_info':
                        self.update_company_info(cursor, record_id, index + 1)
                    else:
                        self.update_address_other(cursor, record_id, index + 1)
                    self.results['success'] += 1
                    
                    # 更新进度
                    progress = 10 + int((index + 1) / len(self.selected_records) * 85)
                    self.progress.emit(progress)
                    
                    # 每10条提交一次
                    if (index + 1) % 10 == 0:
                        connection.commit()
                        self.log_message.emit(f"已处理 {index + 1} 条记录，提交事务")
                        
                except Exception as e:
                    self.results['failed'] += 1
                    error_msg = f"记录ID {record_id} 处理失败: {str(e)}"
                    self.log_message.emit(error_msg)
                    self.results['error_records'].append({
                        'id': record_id,
                        'error': str(e)
                    })
            
            # 最终提交
            connection.commit()
            cursor.close()
            connection.close()
            
            self.progress.emit(100)
            self.finished.emit(self.results)
            
        except Exception as e:
            self.error.emit(str(e))
    
    def update_company_info(self, cursor, record_id, sequence):
        """更新单条记录的公司信息"""
        try:
            # 先查询当前记录信息
            cursor.execute("""
                SELECT id, company_name, country, company_information
                FROM ba_rlb_legal_information 
                WHERE id = %s AND create_staff = '系统导入'
            """, (record_id,))
            
            record = cursor.fetchone()
            if not record:
                raise Exception(f"未找到ID为 {record_id} 且create_staff为'系统导入'的记录")
            
            record_id, company_name, country, current_information = record
            
            # 生成标准的公司信息模板
            company_template = """会社法人番号: 
会社名(漢字)： 
会社名(平仮名)： 
会社名(力タ力ナ)： 
会社名(英語)："""
            
            # 只更新company_information字段
            cursor.execute("""
                UPDATE ba_rlb_legal_information 
                SET company_information = %s,
                    update_time = %s
                WHERE id = %s
            """, (
                company_template,
                int(datetime.now().timestamp()),
                record_id
            ))
            
            self.log_message.emit(f"第{sequence}条: 成功更新公司 '{company_name}' (ID: {record_id}) 的company_information")
            
        except Exception as e:
            raise Exception(f"更新记录失败: {str(e)}")
    
    def update_address_other(self, cursor, record_id, sequence):
        """更新单条记录的公司地址信息"""
        try:
            # 先查询当前记录信息
            cursor.execute("""
                SELECT id, company_name, country, company_address_other
                FROM ba_rlb_legal_information 
                WHERE id = %s AND create_staff = '系统导入'
            """, (record_id,))
            
            record = cursor.fetchone()
            if not record:
                raise Exception(f"未找到ID为 {record_id} 且create_staff为'系统导入'的记录")
            
            record_id, company_name, country, current_address = record
            
            # 生成标准的公司地址模板
            address_template = """会社住所郵便番号：
会社住所(日本語)：
会社住所(平仮名)：
会社住所(力タ力ナ)：
会社住所(英語)："""
            
            # 只更新company_address_other字段
            cursor.execute("""
                UPDATE ba_rlb_legal_information 
                SET company_address_other = %s,
                    update_time = %s
                WHERE id = %s
            """, (
                address_template,
                int(datetime.now().timestamp()),
                record_id
            ))
            
            self.log_message.emit(f"第{sequence}条: 成功更新公司 '{company_name}' (ID: {record_id}) 的company_address_other")
            
        except Exception as e:
            raise Exception(f"更新记录失败: {str(e)}")


class RlbCompanyInfoWidget(QWidget):
    """日联部企业基础信息表补全功能组件"""
    
    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.records_data = []  # 存储查询到的记录
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
        
        # 功能说明区域
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        
        info_label = QLabel("""
<b>日联部企业基础信息表补全说明:</b><br>
<b>功能描述:</b><br>
• 检测ba_rlb_legal_information表中create_staff字段为"系统导入"的记录<br>
• 为这些记录补充标准的日本公司信息模板<br><br>
<b>补充内容（分开运行）:</b><br>
• <b>company_information字段</b>: 公司名称信息模板<br>
• <b>company_address_other字段</b>: 公司地址信息模板<br><br>
<b>company_information模板:</b><br>
会社法人番号: / 会社名(漢字)： / 会社名(平仮名)： / 会社名(力タ力ナ)： / 会社名(英語)：<br><br>
<b>company_address_other模板:</b><br>
会社住所郵便番号： / 会社住所(日本語)： / 会社住所(平仮名)： / 会社住所(力タ力ナ)： / 会社住所(英語)：<br><br>
<b>操作流程:</b><br>
1. 选择数据源并测试连接<br>
2. 点击"查询待补全记录"加载数据<br>
3. 在表格中选择需要补全的记录（支持多选）<br>
4. 点击对应按钮分别补全不同字段
        """)
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # 查询按钮区域
        query_layout = QHBoxLayout()
        
        self.query_btn = QPushButton("查询待补全记录")
        self.query_btn.clicked.connect(self.query_records)
        self.query_btn.setEnabled(False)
        query_layout.addWidget(self.query_btn)
        
        query_layout.addStretch()
        
        # 记录统计标签
        self.record_count_label = QLabel("")
        self.record_count_label.setStyleSheet("color: blue; font-weight: bold;")
        query_layout.addWidget(self.record_count_label)
        
        layout.addLayout(query_layout)
        
        # 记录列表区域
        records_group = QGroupBox("待补全记录列表")
        records_layout = QVBoxLayout()
        
        # 全选/取消全选按钮
        select_layout = QHBoxLayout()
        
        self.select_all_btn = QPushButton("全选")
        self.select_all_btn.clicked.connect(self.select_all_records)
        self.select_all_btn.setEnabled(False)
        select_layout.addWidget(self.select_all_btn)
        
        self.deselect_all_btn = QPushButton("取消全选")
        self.deselect_all_btn.clicked.connect(self.deselect_all_records)
        self.deselect_all_btn.setEnabled(False)
        select_layout.addWidget(self.deselect_all_btn)
        
        select_layout.addStretch()
        records_layout.addLayout(select_layout)
        
        # 记录表格
        self.records_table = QTableWidget()
        self.records_table.setSelectionBehavior(QTableWidget.SelectRows)
        records_layout.addWidget(self.records_table)
        
        records_group.setLayout(records_layout)
        layout.addWidget(records_group)
        
        # 操作按钮区域
        button_layout = QHBoxLayout()
        
        self.start_company_info_btn = QPushButton("补全 company_information")
        self.start_company_info_btn.clicked.connect(lambda: self.start_update('company_info'))
        self.start_company_info_btn.setEnabled(False)
        self.start_company_info_btn.setStyleSheet("background-color: #4CAF50; color: white; padding: 8px 16px;")
        button_layout.addWidget(self.start_company_info_btn)
        
        self.start_address_other_btn = QPushButton("补全 company_address_other")
        self.start_address_other_btn.clicked.connect(lambda: self.start_update('address_other'))
        self.start_address_other_btn.setEnabled(False)
        self.start_address_other_btn.setStyleSheet("background-color: #2196F3; color: white; padding: 8px 16px;")
        button_layout.addWidget(self.start_address_other_btn)
        
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
        self.query_btn.setEnabled(current_data is not None)
        
        # 清空之前的查询结果
        self.records_data = []
        self.records_table.setRowCount(0)
        self.record_count_label.setText("")
        self.update_button_states()
    
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
    
    def query_records(self):
        """查询待补全记录"""
        datasource = self.datasource_combo.currentData()
        if not datasource:
            QMessageBox.warning(self, "警告", "请先选择数据源！")
            return
        
        try:
            self.log_text.append("开始查询待补全记录...")
            
            # 连接数据库
            connection = pymysql.connect(
                host=datasource.host,
                port=datasource.port,
                user=datasource.username,
                password=datasource.password,
                database=datasource.database,
                charset=datasource.charset
            )
            cursor = connection.cursor()
            
            # 查询create_staff为"系统导入"的记录
            cursor.execute("""
                SELECT id, company_name, country, legal_id, admin_id, 
                       create_time, update_time, company_information
                FROM ba_rlb_legal_information 
                WHERE create_staff = '系统导入'
                ORDER BY create_time DESC
            """)
            
            records = cursor.fetchall()
            cursor.close()
            connection.close()
            
            self.records_data = records
            self.display_records()
            
            count = len(records)
            self.record_count_label.setText(f"找到 {count} 条待补全记录")
            self.log_text.append(f"查询完成，找到 {count} 条记录")
            
            self.update_button_states()
            
        except Exception as e:
            QMessageBox.critical(self, "查询失败", f"查询记录失败：{str(e)}")
            self.log_text.append(f"查询失败：{str(e)}")
    
    def display_records(self):
        """显示记录列表"""
        if not self.records_data:
            self.records_table.setRowCount(0)
            return
        
        # 设置表格
        headers = ["选择", "ID", "公司名称", "国家", "法人ID", "管理员ID", 
                  "创建时间", "更新时间", "公司信息状态"]
        self.records_table.setColumnCount(len(headers))
        self.records_table.setHorizontalHeaderLabels(headers)
        self.records_table.setRowCount(len(self.records_data))
        
        for row, record in enumerate(self.records_data):
            record_id, company_name, country, legal_id, admin_id, create_time, update_time, company_info = record
            
            # 选择框
            checkbox = QCheckBox()
            self.records_table.setCellWidget(row, 0, checkbox)
            
            # 其他数据
            self.records_table.setItem(row, 1, QTableWidgetItem(str(record_id)))
            self.records_table.setItem(row, 2, QTableWidgetItem(str(company_name) if company_name else ""))
            self.records_table.setItem(row, 3, QTableWidgetItem(str(country) if country else ""))
            self.records_table.setItem(row, 4, QTableWidgetItem(str(legal_id) if legal_id else ""))
            self.records_table.setItem(row, 5, QTableWidgetItem(str(admin_id) if admin_id else ""))
            
            # 时间格式化
            create_time_str = datetime.fromtimestamp(create_time).strftime('%Y-%m-%d %H:%M:%S') if create_time else ""
            update_time_str = datetime.fromtimestamp(update_time).strftime('%Y-%m-%d %H:%M:%S') if update_time else ""
            self.records_table.setItem(row, 6, QTableWidgetItem(create_time_str))
            self.records_table.setItem(row, 7, QTableWidgetItem(update_time_str))
            
            # 公司信息状态
            info_status = "已有内容" if company_info and company_info.strip() else "空白"
            self.records_table.setItem(row, 8, QTableWidgetItem(info_status))
        
        # 调整列宽
        self.records_table.resizeColumnsToContents()
        header = self.records_table.horizontalHeader()
        header.setSectionResizeMode(2, QHeaderView.Stretch)  # 公司名称列自适应
    
    def select_all_records(self):
        """全选记录"""
        for row in range(self.records_table.rowCount()):
            checkbox = self.records_table.cellWidget(row, 0)
            if checkbox:
                checkbox.setChecked(True)
        self.update_button_states()
    
    def deselect_all_records(self):
        """取消全选记录"""
        for row in range(self.records_table.rowCount()):
            checkbox = self.records_table.cellWidget(row, 0)
            if checkbox:
                checkbox.setChecked(False)
        self.update_button_states()
    
    def get_selected_record_ids(self):
        """获取选中的记录ID列表"""
        selected_ids = []
        for row in range(self.records_table.rowCount()):
            checkbox = self.records_table.cellWidget(row, 0)
            if checkbox and checkbox.isChecked():
                record_id = int(self.records_table.item(row, 1).text())
                selected_ids.append(record_id)
        return selected_ids
    
    def update_button_states(self):
        """更新按钮状态"""
        has_records = len(self.records_data) > 0
        has_selection = len(self.get_selected_record_ids()) > 0
        
        self.select_all_btn.setEnabled(has_records)
        self.deselect_all_btn.setEnabled(has_records)
        self.start_company_info_btn.setEnabled(has_selection)
        self.start_address_other_btn.setEnabled(has_selection)
    
    def start_update(self):
        """开始补全"""
        datasource = self.datasource_combo.currentData()
        if not datasource:
            QMessageBox.warning(self, "警告", "请先选择数据源！")
            return
        
        selected_ids = self.get_selected_record_ids()
        if not selected_ids:
            QMessageBox.warning(self, "警告", "请先选择要补全的记录！")
            return
        
        # 确认对话框
        reply = QMessageBox.question(
            self,
            "确认补全",
            f"确定要补全选中的 {len(selected_ids)} 条记录吗？\n\n"
            f"将为这些记录的company_information字段\n"
            f"补充标准的日本公司信息模板。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        # 禁用按钮，显示进度条
        self.start_btn.setEnabled(False)
        self.query_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在补全中...")
        self.log_text.clear()
        
        # 启动工作线程
        self.worker = RlbCompanyInfoWorker(datasource, self.db_manager, selected_ids)
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log_message.connect(self.append_log)
        self.worker.finished.connect(self.on_update_finished)
        self.worker.error.connect(self.on_update_error)
        self.worker.start()
    
    def append_log(self, message):
        """添加日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
    
    def on_update_finished(self, results):
        """补全完成"""
        self.start_btn.setEnabled(True)
        self.query_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        result_text = (
            f"补全完成！\n"
            f"总计：{results['total']} 条\n"
            f"成功：{results['success']} 条\n"
            f"失败：{results['failed']} 条"
        )
        
        self.result_label.setText(result_text)
        self.result_label.setStyleSheet("color: green;")
        
        # 显示详细结果
        detail_msg = result_text
        if results['error_records']:
            detail_msg += "\n\n失败记录：\n"
            for error_record in results['error_records'][:5]:  # 只显示前5个错误
                detail_msg += f"ID {error_record['id']}: {error_record['error']}\n"
            if len(results['error_records']) > 5:
                detail_msg += f"...还有 {len(results['error_records']) - 5} 个错误"
        
        QMessageBox.information(self, "补全完成", detail_msg)
        
        # 重新查询记录以更新显示
        self.query_records()
    
    def on_update_error(self, error_msg):
        """补全错误"""
        self.start_btn.setEnabled(True)
        self.query_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.result_label.setText("补全失败！")
        self.result_label.setStyleSheet("color: red;")
        
        QMessageBox.critical(self, "补全错误", f"补全失败：{error_msg}")
    
    def clear_log(self):
        """清空日志"""
        self.log_text.clear()
