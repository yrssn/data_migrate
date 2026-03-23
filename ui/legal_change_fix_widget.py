"""
修复变更法人功能界面
"""
import os
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog,
                             QTableWidget, QTableWidgetItem, QHeaderView, QCheckBox,
                             QListWidget, QListWidgetItem, QAbstractItemView)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from database import DatabaseManager
import pandas as pd
import pymysql


class LegalChangeFixWorker(QThread):
    """修复变更法人工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    
    def __init__(self, datasource, db_manager, selected_records):
        super().__init__()
        self.datasource = datasource
        self.db_manager = db_manager
        self.selected_records = selected_records
        self.results = {
            'total_selected': len(selected_records),
            'success_count': 0,
            'skipped_no_shopindex': 0,
            'skipped_no_legal_id': 0,
            'skipped_no_legal_info': 0,
            'failed_count': 0,
            'failed_records': []
        }
    
    def run(self):
        try:
            self.log_message.emit("开始连接数据库...")
            self.progress.emit(5)
            
            connection = pymysql.connect(
                host=self.datasource.host,
                port=self.datasource.port,
                user=self.datasource.username,
                password=self.datasource.password,
                database=self.datasource.database,
                charset=self.datasource.charset
            )
            cursor = connection.cursor(pymysql.cursors.DictCursor)
            self.log_message.emit("数据库连接成功")
            self.progress.emit(10)
            
            total = len(self.selected_records)
            for i, record in enumerate(self.selected_records):
                try:
                    self.process_record(cursor, record)
                    
                    progress = 10 + int((i + 1) / total * 85)
                    self.progress.emit(progress)
                    
                except Exception as e:
                    self.results['failed_count'] += 1
                    error_msg = f"记录ID {record['id']} 处理失败: {str(e)}"
                    self.log_message.emit(error_msg)
                    self.results['failed_records'].append({
                        'id': record['id'],
                        'error': str(e)
                    })
            
            connection.commit()
            cursor.close()
            connection.close()
            
            self.progress.emit(100)
            self.finished.emit(self.results)
            
        except Exception as e:
            self.error.emit(str(e))
    
    def process_record(self, cursor, record):
        """处理单条记录"""
        register_id = record['id']
        shopindex_id = record['shopindex_id']
        
        self.log_message.emit(f"处理记录 ID={register_id}, shopindex_id={shopindex_id}")
        
        if not shopindex_id:
            self.results['skipped_no_shopindex'] += 1
            self.log_message.emit(f"  记录ID {register_id}: shopindex_id为空，跳过")
            return
        
        cursor.execute("""
            SELECT id, legal_id, belong_information 
            FROM ba_shopindex 
            WHERE id = %s AND (delete_time IS NULL OR delete_time = 0)
        """, (shopindex_id,))
        shopindex_record = cursor.fetchone()
        
        if not shopindex_record:
            self.results['skipped_no_shopindex'] += 1
            self.log_message.emit(f"  记录ID {register_id}: 未找到shopindex_id={shopindex_id}的主体编号记录，跳过")
            return
        
        legal_id = shopindex_record['legal_id']
        belong_information = shopindex_record['belong_information']
        
        self.log_message.emit(f"  找到主体编号: legal_id={legal_id}, belong_information={belong_information}")
        
        if not legal_id:
            self.results['skipped_no_legal_id'] += 1
            self.log_message.emit(f"  记录ID {register_id}: legal_id为空，跳过")
            return
        
        customer_id = None
        information_id = None
        information_part2_id = None
        
        current_time = int(datetime.now().timestamp())
        
        if belong_information is None:
            customer_id = legal_id
            self.log_message.emit(f"  belong_information为NULL，直接使用legal_id={legal_id}作为customer_id")
            
        elif belong_information == 1:
            cursor.execute("""
                SELECT id, legal_id 
                FROM ba_rlb_legal_information 
                WHERE id = %s AND (delete_time IS NULL OR delete_time = 0)
            """, (legal_id,))
            legal_info = cursor.fetchone()
            
            if not legal_info:
                self.results['skipped_no_legal_info'] += 1
                self.log_message.emit(f"  记录ID {register_id}: 在ba_rlb_legal_information表中未找到id={legal_id}的记录，跳过")
                return
            
            customer_id = legal_info['legal_id']
            information_id = legal_info['id']
            self.log_message.emit(f"  从ba_rlb_legal_information获取: customer_id={customer_id}, information_id={information_id}")
            
        elif belong_information == 2:
            cursor.execute("""
                SELECT id, legal_id 
                FROM ba_rlb_legal_information_part2 
                WHERE id = %s AND (delete_time IS NULL OR delete_time = 0)
            """, (legal_id,))
            legal_info = cursor.fetchone()
            
            if not legal_info:
                self.results['skipped_no_legal_info'] += 1
                self.log_message.emit(f"  记录ID {register_id}: 在ba_rlb_legal_information_part2表中未找到id={legal_id}的记录，跳过")
                return
            
            customer_id = legal_info['legal_id']
            information_part2_id = legal_info['id']
            self.log_message.emit(f"  从ba_rlb_legal_information_part2获取: customer_id={customer_id}, information_part2_id={information_part2_id}")
        
        else:
            self.results['skipped_no_legal_info'] += 1
            self.log_message.emit(f"  记录ID {register_id}: belong_information值异常({belong_information})，跳过")
            return
        
        update_sql = """
            UPDATE ba_ptzcb_register 
            SET customer_id = %s, 
                information_id = %s, 
                information_part2_id = %s,
                update_time = %s
            WHERE id = %s
        """
        cursor.execute(update_sql, (customer_id, information_id, information_part2_id, current_time, register_id))
        
        self.results['success_count'] += 1
        self.log_message.emit(f"  记录ID {register_id}: 更新成功 - customer_id={customer_id}, information_id={information_id}, information_part2_id={information_part2_id}")


class LegalChangeFixWidget(QWidget):
    """修复变更法人功能组件"""
    
    STATUS_OPTIONS = {
        0: '待指派',
        1: '进行中',
        2: '审核中',
        3: '银行验证',
        4: '审核失败',
        5: '银行卡验证失败',
        6: '完成',
        7: '失败',
        8: '取消',
        9: '暂停',
        10: '店铺验证失败',
        11: '店铺验证中'
    }
    
    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.loaded_records = []
        self.init_ui()
        self.load_datasources()
    
    def init_ui(self):
        """初始化UI"""
        layout = QVBoxLayout()
        
        datasource_group = QGroupBox("选择数据库")
        datasource_layout = QFormLayout()
        
        self.datasource_combo = QComboBox()
        self.datasource_combo.currentTextChanged.connect(self.on_datasource_changed)
        datasource_layout.addRow("数据库:", self.datasource_combo)
        
        self.test_connection_btn = QPushButton("测试连接")
        self.test_connection_btn.clicked.connect(self.test_connection)
        self.test_connection_btn.setEnabled(False)
        datasource_layout.addRow("", self.test_connection_btn)
        
        datasource_group.setLayout(datasource_layout)
        layout.addWidget(datasource_group)
        
        status_group = QGroupBox("状态筛选")
        status_layout = QVBoxLayout()
        
        status_hint = QLabel("选择要处理的状态（可多选）：")
        status_layout.addWidget(status_hint)
        
        self.status_list = QListWidget()
        self.status_list.setSelectionMode(QAbstractItemView.MultiSelection)
        self.status_list.setMaximumHeight(150)
        
        for status_code, status_name in self.STATUS_OPTIONS.items():
            item = QListWidgetItem(f"{status_code} - {status_name}")
            item.setData(Qt.UserRole, status_code)
            self.status_list.addItem(item)
            if status_code == 6:
                item.setSelected(True)
        
        status_layout.addWidget(self.status_list)
        
        status_btn_layout = QHBoxLayout()
        self.select_all_status_btn = QPushButton("全选状态")
        self.select_all_status_btn.clicked.connect(self.select_all_status)
        status_btn_layout.addWidget(self.select_all_status_btn)
        
        self.deselect_all_status_btn = QPushButton("取消全选")
        self.deselect_all_status_btn.clicked.connect(self.deselect_all_status)
        status_btn_layout.addWidget(self.deselect_all_status_btn)
        
        status_btn_layout.addStretch()
        status_layout.addLayout(status_btn_layout)
        
        status_group.setLayout(status_layout)
        layout.addWidget(status_group)
        
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        
        info_label = QLabel("""
<b>修复变更法人说明:</b><br>
<b>处理逻辑:</b><br>
• 1. 查找ba_ptzcb_register表中，选定状态且customer_id为空的记录<br>
• 2. 根据shopindex_id查找ba_shopindex表<br>
• 3. 根据belong_information字段判断:<br>
&nbsp;&nbsp;- NULL: 将legal_id直接存入customer_id<br>
&nbsp;&nbsp;- 1: 根据legal_id在ba_rlb_legal_information表查找，获取其legal_id存入customer_id，并将该表的id存入information_id<br>
&nbsp;&nbsp;- 2: 根据legal_id在ba_rlb_legal_information_part2表查找，获取其legal_id存入customer_id，并将该表的id存入information_part2_id<br>
• 4. information_id和information_part2_id只能有一个有值<br><br>
<b>字段映射:</b><br>
• ba_ptzcb_register.customer_id ← ba_rlb_customer.id (或直接从legal_id)<br>
• ba_ptzcb_register.information_id ← ba_rlb_legal_information.id (belong_information=1时)<br>
• ba_ptzcb_register.information_part2_id ← ba_rlb_legal_information_part2.id (belong_information=2时)
        """)
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        preview_group = QGroupBox("数据预览")
        preview_layout = QVBoxLayout()
        
        preview_button_layout = QHBoxLayout()
        self.load_data_btn = QPushButton("查询数据")
        self.load_data_btn.clicked.connect(self.load_data)
        self.load_data_btn.setEnabled(False)
        preview_button_layout.addWidget(self.load_data_btn)
        
        self.select_all_btn = QPushButton("全选")
        self.select_all_btn.clicked.connect(self.select_all_records)
        self.select_all_btn.setEnabled(False)
        preview_button_layout.addWidget(self.select_all_btn)
        
        self.deselect_all_btn = QPushButton("全不选")
        self.deselect_all_btn.clicked.connect(self.deselect_all_records)
        self.deselect_all_btn.setEnabled(False)
        preview_button_layout.addWidget(self.deselect_all_btn)
        
        self.record_count_label = QLabel("")
        preview_button_layout.addWidget(self.record_count_label)
        
        preview_button_layout.addStretch()
        preview_layout.addLayout(preview_button_layout)
        
        self.records_table = QTableWidget()
        self.records_table.setMaximumHeight(300)
        preview_layout.addWidget(self.records_table)
        
        preview_group.setLayout(preview_layout)
        layout.addWidget(preview_group)
        
        button_layout = QHBoxLayout()
        
        self.start_btn = QPushButton("开始修复")
        self.start_btn.clicked.connect(self.start_fix)
        self.start_btn.setEnabled(False)
        button_layout.addWidget(self.start_btn)
        
        button_layout.addStretch()
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        button_layout.addWidget(self.progress_bar)
        
        layout.addLayout(button_layout)
        
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
        
        self.export_log_btn = QPushButton("导出日志")
        self.export_log_btn.clicked.connect(self.export_log)
        log_button_layout.addWidget(self.export_log_btn)
        
        log_button_layout.addStretch()
        log_layout.addLayout(log_button_layout)
        
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)
        
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
            display_name = f"{ds.name} ({ds.host}:{ds.port})"
            self.datasource_combo.addItem(display_name, ds)
    
    def on_datasource_changed(self):
        """数据源选择变化"""
        datasource = self.datasource_combo.currentData()
        self.test_connection_btn.setEnabled(datasource is not None)
        self.load_data_btn.setEnabled(datasource is not None)
    
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
    
    def select_all_status(self):
        """全选状态"""
        for i in range(self.status_list.count()):
            self.status_list.item(i).setSelected(True)
    
    def deselect_all_status(self):
        """取消全选状态"""
        for i in range(self.status_list.count()):
            self.status_list.item(i).setSelected(False)
    
    def get_selected_status(self):
        """获取选中的状态列表"""
        selected = []
        for item in self.status_list.selectedItems():
            selected.append(item.data(Qt.UserRole))
        return selected
    
    def load_data(self):
        """加载数据预览"""
        datasource = self.datasource_combo.currentData()
        if not datasource:
            QMessageBox.warning(self, "警告", "请选择数据源！")
            return
        
        selected_status = self.get_selected_status()
        if not selected_status:
            QMessageBox.warning(self, "警告", "请至少选择一个状态！")
            return
        
        try:
            connection = pymysql.connect(
                host=datasource.host,
                port=datasource.port,
                user=datasource.username,
                password=datasource.password,
                database=datasource.database,
                charset=datasource.charset
            )
            cursor = connection.cursor(pymysql.cursors.DictCursor)
            
            status_placeholders = ','.join(['%s'] * len(selected_status))
            query = f"""
                SELECT id, project_id, shopindex_id, status, customer_id, 
                       information_id, information_part2_id, company_name, email
                FROM ba_ptzcb_register 
                WHERE status IN ({status_placeholders}) 
                  AND (customer_id IS NULL OR customer_id = 0)
                  AND (delete_time IS NULL OR delete_time = 0)
                ORDER BY id DESC
            """
            
            cursor.execute(query, selected_status)
            records = cursor.fetchall()
            
            cursor.close()
            connection.close()
            
            self.loaded_records = records
            self.display_records(records)
            
            self.record_count_label.setText(f"共 {len(records)} 条待处理记录")
            
            if records:
                self.select_all_btn.setEnabled(True)
                self.deselect_all_btn.setEnabled(True)
                self.start_btn.setEnabled(True)
            else:
                self.select_all_btn.setEnabled(False)
                self.deselect_all_btn.setEnabled(False)
                self.start_btn.setEnabled(False)
            
        except Exception as e:
            QMessageBox.critical(self, "查询失败", f"查询数据失败：{str(e)}")
    
    def display_records(self, records):
        """显示记录到表格"""
        headers = ["选择", "ID", "项目编号", "主体编号ID", "状态", "公司名称", "邮箱"]
        self.records_table.setRowCount(len(records))
        self.records_table.setColumnCount(len(headers))
        self.records_table.setHorizontalHeaderLabels(headers)
        
        for row_idx, record in enumerate(records):
            checkbox = QCheckBox()
            checkbox.setChecked(True)
            self.records_table.setCellWidget(row_idx, 0, checkbox)
            
            self.records_table.setItem(row_idx, 1, QTableWidgetItem(str(record['id'])))
            self.records_table.setItem(row_idx, 2, QTableWidgetItem(str(record['project_id'] or '')))
            self.records_table.setItem(row_idx, 3, QTableWidgetItem(str(record['shopindex_id'] or '')))
            
            status_code = record['status']
            status_text = f"{status_code} - {self.STATUS_OPTIONS.get(status_code, '未知')}"
            self.records_table.setItem(row_idx, 4, QTableWidgetItem(status_text))
            
            self.records_table.setItem(row_idx, 5, QTableWidgetItem(str(record['company_name'] or '')))
            self.records_table.setItem(row_idx, 6, QTableWidgetItem(str(record['email'] or '')))
        
        header = self.records_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.Stretch)
        header.setSectionResizeMode(6, QHeaderView.Stretch)
    
    def select_all_records(self):
        """全选记录"""
        for row in range(self.records_table.rowCount()):
            checkbox = self.records_table.cellWidget(row, 0)
            if checkbox:
                checkbox.setChecked(True)
    
    def deselect_all_records(self):
        """全不选记录"""
        for row in range(self.records_table.rowCount()):
            checkbox = self.records_table.cellWidget(row, 0)
            if checkbox:
                checkbox.setChecked(False)
    
    def get_selected_records(self):
        """获取选中的记录"""
        selected = []
        for row in range(self.records_table.rowCount()):
            checkbox = self.records_table.cellWidget(row, 0)
            if checkbox and checkbox.isChecked():
                selected.append(self.loaded_records[row])
        return selected
    
    def start_fix(self):
        """开始修复"""
        datasource = self.datasource_combo.currentData()
        if not datasource:
            QMessageBox.warning(self, "警告", "请选择数据源！")
            return
        
        selected_records = self.get_selected_records()
        if not selected_records:
            QMessageBox.warning(self, "警告", "请选择要处理的记录！")
            return
        
        reply = QMessageBox.question(
            self,
            "确认修复",
            f"确定要修复变更法人吗？\n\n"
            f"数据库: {datasource.name}\n"
            f"选中记录: {len(selected_records)} 条\n\n"
            f"此操作将更新ba_ptzcb_register表的\n"
            f"customer_id、information_id、information_part2_id字段。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        self.start_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在修复中...")
        self.log_text.clear()
        
        self.worker = LegalChangeFixWorker(datasource, self.db_manager, selected_records)
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log_message.connect(self.append_log)
        self.worker.finished.connect(self.on_fix_finished)
        self.worker.error.connect(self.on_fix_error)
        self.worker.start()
    
    def append_log(self, message):
        """添加日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
    
    def on_fix_finished(self, results):
        """修复完成"""
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        result_text = (
            f"修复完成！\n"
            f"选中记录：{results['total_selected']} 条\n"
            f"成功修复：{results['success_count']} 条\n"
            f"跳过(无主体编号)：{results['skipped_no_shopindex']} 条\n"
            f"跳过(无legal_id)：{results['skipped_no_legal_id']} 条\n"
            f"跳过(无法人信息)：{results['skipped_no_legal_info']} 条\n"
            f"处理失败：{results['failed_count']} 条"
        )
        
        self.result_label.setText(result_text)
        if results['failed_count'] > 0:
            self.result_label.setStyleSheet("color: orange;")
        else:
            self.result_label.setStyleSheet("color: green;")
        
        QMessageBox.information(self, "修复完成", result_text)
    
    def on_fix_error(self, error_msg):
        """修复错误"""
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.result_label.setText("修复失败！")
        self.result_label.setStyleSheet("color: red;")
        
        QMessageBox.critical(self, "修复错误", f"修复失败：{error_msg}")
    
    def clear_log(self):
        """清空日志"""
        self.log_text.clear()
    
    def export_log(self):
        """导出日志"""
        if not self.log_text.toPlainText():
            QMessageBox.information(self, "提示", "没有日志可以导出")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存日志文件",
            f"修复变更法人日志_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            "文本文件 (*.txt);;所有文件 (*)"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.log_text.toPlainText())
                QMessageBox.information(self, "成功", "日志导出成功！")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出失败：{str(e)}")
