"""
补全客户信息字段初始值功能界面
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


class CustomerInfoCompleteWorker(QThread):
    """客户信息补全工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    
    def __init__(self, datasource, db_manager, selected_records, fill_mode='address_info'):
        super().__init__()
        self.datasource = datasource
        self.db_manager = db_manager
        self.selected_records = selected_records  # 用户选择的记录ID列表
        self.fill_mode = fill_mode  # 'address_info' 或 'name_info'
        self.results = {
            'total': 0,
            'success': 0,
            'skipped': 0,
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
                    if self.fill_mode == 'address_info':
                        self.update_address_info(cursor, record_id, index + 1)
                    else:
                        self.update_name_info(cursor, record_id, index + 1)
                    
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
    
    def update_address_info(self, cursor, record_id, sequence):
        """更新单条记录的住所信息"""
        try:
            # 先查询当前记录信息
            cursor.execute("""
                SELECT id, legal_name, address_information
                FROM ba_rlb_customer 
                WHERE id = %s AND remark = '系统导入'
            """, (record_id,))
            
            record = cursor.fetchone()
            if not record:
                raise Exception(f"未找到ID为 {record_id} 且remark为'系统导入'的记录")
            
            record_id, legal_name, current_address_info = record
            
            # 检查是否已有内容，有内容则跳过
            if current_address_info and current_address_info.strip():
                self.log_message.emit(f"第{sequence}条: 跳过客户 '{legal_name}' (ID: {record_id})，address_information已有内容")
                self.results['skipped'] += 1
                return
            
            # 生成标准的住所信息模板
            address_template = """本人住所(日本語)：
本人住所(平仮名)：
本人住所(力タ力ナ)：
本人住所(英語)：
本人住所邮编番号："""
            
            # 只更新address_information字段
            cursor.execute("""
                UPDATE ba_rlb_customer 
                SET address_information = %s,
                    update_time = %s
                WHERE id = %s
            """, (
                address_template,
                int(datetime.now().timestamp()),
                record_id
            ))
            
            self.results['success'] += 1
            self.log_message.emit(f"第{sequence}条: 成功更新客户 '{legal_name}' (ID: {record_id}) 的address_information")
            
        except Exception as e:
            raise Exception(f"更新记录失败: {str(e)}")
    
    def update_name_info(self, cursor, record_id, sequence):
        """更新单条记录的姓名信息"""
        try:
            # 先查询当前记录信息
            cursor.execute("""
                SELECT id, legal_name, name_information
                FROM ba_rlb_customer 
                WHERE id = %s AND remark = '系统导入'
            """, (record_id,))
            
            record = cursor.fetchone()
            if not record:
                raise Exception(f"未找到ID为 {record_id} 且remark为'系统导入'的记录")
            
            record_id, legal_name, current_name_info = record
            
            # 检查是否已有内容，有内容则跳过
            if current_name_info and current_name_info.strip():
                self.log_message.emit(f"第{sequence}条: 跳过客户 '{legal_name}' (ID: {record_id})，name_information已有内容")
                self.results['skipped'] += 1
                return
            
            # 生成标准的姓名信息模板
            name_template = """名前(漢字)：
名前(平仮名)：
名前(力タ力ナ)：
名前(ロ-マ字)："""
            
            # 只更新name_information字段
            cursor.execute("""
                UPDATE ba_rlb_customer 
                SET name_information = %s,
                    update_time = %s
                WHERE id = %s
            """, (
                name_template,
                int(datetime.now().timestamp()),
                record_id
            ))
            
            self.results['success'] += 1
            self.log_message.emit(f"第{sequence}条: 成功更新客户 '{legal_name}' (ID: {record_id}) 的name_information")
            
        except Exception as e:
            raise Exception(f"更新记录失败: {str(e)}")


class CustomerInfoCompleteWidget(QWidget):
    """补全客户信息字段初始值功能组件"""
    
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
<b>补全客户信息字段初始值说明:</b><br>
<b>功能描述:</b><br>
• 检测ba_rlb_customer表中remark字段为"系统导入"的记录<br>
• 为这些记录补充标准的客户信息字段模板<br><br>
<b>补充内容（分开运行）:</b><br>
• <b>address_information字段</b>: 住所信息模板<br>
• <b>name_information字段</b>: 姓名信息模板<br><br>
<b>住所信息模板:</b><br>
本人住所(日本語)： / 本人住所(平仮名)： / 本人住所(力タ力ナ)： / 本人住所(英語)： / 本人住所邮编番号：<br><br>
<b>姓名信息模板:</b><br>
名前(漢字)： / 名前(平仮名)： / 名前(力タ力ナ)： / 名前(ロ-マ字)：<br><br>
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
        
        self.start_address_btn = QPushButton("补全 address_information（住所）")
        self.start_address_btn.clicked.connect(lambda: self.start_update('address_info'))
        self.start_address_btn.setEnabled(False)
        self.start_address_btn.setStyleSheet("background-color: #4CAF50; color: white; padding: 8px 16px;")
        button_layout.addWidget(self.start_address_btn)
        
        self.start_name_btn = QPushButton("补全 name_information（姓名）")
        self.start_name_btn.clicked.connect(lambda: self.start_update('name_info'))
        self.start_name_btn.setEnabled(False)
        self.start_name_btn.setStyleSheet("background-color: #2196F3; color: white; padding: 8px 16px;")
        button_layout.addWidget(self.start_name_btn)
        
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
            
            # 查询remark为"系统导入"的记录
            cursor.execute("""
                SELECT id, legal_name, create_time, update_time, address_information, name_information
                FROM ba_rlb_customer 
                WHERE remark = '系统导入'
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
        headers = ["选择", "ID", "法人姓名", "创建时间", "更新时间", "住所信息状态", "姓名信息状态"]
        self.records_table.setColumnCount(len(headers))
        self.records_table.setHorizontalHeaderLabels(headers)
        self.records_table.setRowCount(len(self.records_data))
        
        for row, record in enumerate(self.records_data):
            record_id, legal_name, create_time, update_time, address_information, name_information = record
            
            # 选择框
            checkbox = QCheckBox()
            self.records_table.setCellWidget(row, 0, checkbox)
            
            # 其他数据
            self.records_table.setItem(row, 1, QTableWidgetItem(str(record_id)))
            self.records_table.setItem(row, 2, QTableWidgetItem(str(legal_name) if legal_name else ""))
            
            # 时间格式化
            create_time_str = datetime.fromtimestamp(create_time).strftime('%Y-%m-%d %H:%M:%S') if create_time else ""
            update_time_str = datetime.fromtimestamp(update_time).strftime('%Y-%m-%d %H:%M:%S') if update_time else ""
            self.records_table.setItem(row, 3, QTableWidgetItem(create_time_str))
            self.records_table.setItem(row, 4, QTableWidgetItem(update_time_str))
            
            # 住所信息状态
            address_status = "已有内容" if address_information and address_information.strip() else "空白"
            self.records_table.setItem(row, 5, QTableWidgetItem(address_status))
            
            # 姓名信息状态
            name_status = "已有内容" if name_information and name_information.strip() else "空白"
            self.records_table.setItem(row, 6, QTableWidgetItem(name_status))
        
        # 调整列宽
        self.records_table.resizeColumnsToContents()
        header = self.records_table.horizontalHeader()
        header.setSectionResizeMode(2, QHeaderView.Stretch)  # 法人姓名列自适应
    
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
        self.start_address_btn.setEnabled(has_selection)
        self.start_name_btn.setEnabled(has_selection)
    
    def start_update(self, fill_mode='address_info'):
        """开始补全"""
        datasource = self.datasource_combo.currentData()
        if not datasource:
            QMessageBox.warning(self, "警告", "请先选择数据源！")
            return
        
        selected_ids = self.get_selected_record_ids()
        if not selected_ids:
            QMessageBox.warning(self, "警告", "请先选择要补全的记录！")
            return
        
        # 根据模式确定提示信息
        if fill_mode == 'address_info':
            field_name = "address_information"
            template_desc = "住所信息模板（本人住所等）"
        else:
            field_name = "name_information"
            template_desc = "姓名信息模板（名前等）"
        
        # 确认对话框
        reply = QMessageBox.question(
            self,
            "确认补全",
            f"确定要补全选中的 {len(selected_ids)} 条记录吗？\n\n"
            f"将为这些记录的 {field_name} 字段\n"
            f"补充{template_desc}。\n\n"
            f"已有内容的记录会跳过，不会覆盖。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        # 禁用按钮，显示进度条
        self.start_address_btn.setEnabled(False)
        self.start_name_btn.setEnabled(False)
        self.query_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText(f"正在补全 {field_name} 中...")
        self.log_text.clear()
        
        # 启动工作线程
        self.worker = CustomerInfoCompleteWorker(datasource, self.db_manager, selected_ids, fill_mode)
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
        self.start_address_btn.setEnabled(True)
        self.start_name_btn.setEnabled(True)
        self.query_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        skipped = results.get('skipped', 0)
        result_text = (
            f"补全完成！\n"
            f"总计：{results['total']} 条\n"
            f"成功：{results['success']} 条\n"
            f"跳过（已有内容）：{skipped} 条\n"
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
        self.start_address_btn.setEnabled(True)
        self.start_name_btn.setEnabled(True)
        self.query_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.result_label.setText("补全失败！")
        self.result_label.setStyleSheet("color: red;")
        
        QMessageBox.critical(self, "补全错误", f"补全失败：{error_msg}")
    
    def clear_log(self):
        """清空日志"""
        self.log_text.clear()
