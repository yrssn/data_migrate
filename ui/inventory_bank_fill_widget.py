"""
根据库存ID补全银行卡功能界面
"""
import os
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog,
                             QTableWidget, QTableWidgetItem, QHeaderView)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from database import DatabaseManager
import pandas as pd
import pymysql


class InventoryBankFillWorker(QThread):
    """根据库存ID补全银行卡工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    
    def __init__(self, datasource, excel_file, legal_column, bank_name_column, bank_card_column, inventory_column):
        super().__init__()
        self.datasource = datasource
        self.excel_file = excel_file
        self.legal_column = legal_column
        self.bank_name_column = bank_name_column
        self.bank_card_column = bank_card_column
        self.inventory_column = inventory_column
        self.results = {
            'total_rows': 0,
            'success_count': 0,
            'skipped_no_bank': 0,
            'skipped_no_inventory': 0,
            'skipped_already_has': 0,
            'failed_count': 0,
            'completed_data': [],
            'failed_records': []
        }
    
    def run(self):
        try:
            self.log_message.emit("开始读取Excel文件...")
            self.progress.emit(5)
            
            df = pd.read_excel(self.excel_file, engine='openpyxl')
            self.log_message.emit(f"成功读取Excel文件，共 {len(df)} 行数据")
            self.results['total_rows'] = len(df)
            self.progress.emit(10)
            
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
            self.progress.emit(15)
            
            total = len(df)
            for index, row in df.iterrows():
                try:
                    result = self.process_row(cursor, index + 1, row, total)
                    
                    completed_row = row.to_dict()
                    completed_row['处理结果'] = result['status']
                    completed_row['说明'] = result['message']
                    completed_row['找到的bank_id'] = result.get('bank_id', '')
                    completed_row['处理时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    self.results['completed_data'].append(completed_row)
                    
                    progress = 15 + int((index + 1) / total * 80)
                    self.progress.emit(progress)
                    
                    if (index + 1) % 50 == 0:
                        connection.commit()
                        self.log_message.emit(f"已处理 {index + 1}/{total} 行，提交事务")
                        
                except Exception as e:
                    self.results['failed_count'] += 1
                    error_msg = f"第{index + 1}行处理失败: {str(e)}"
                    self.log_message.emit(error_msg)
                    
                    failed_row = row.to_dict()
                    failed_row['处理结果'] = '失败'
                    failed_row['失败原因'] = str(e)
                    failed_row['处理时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    self.results['completed_data'].append(failed_row)
                    self.results['failed_records'].append({
                        'row': index + 1,
                        'error': str(e),
                        'data': row.to_dict()
                    })
            
            connection.commit()
            cursor.close()
            connection.close()
            
            self.progress.emit(100)
            self.finished.emit(self.results)
            
        except Exception as e:
            self.error.emit(str(e))
    
    def get_cell_value(self, row, column_name):
        """获取单元格值"""
        if column_name not in row.index:
            return None
        value = row[column_name]
        if pd.isna(value) or str(value).strip() == '' or str(value).strip().lower() == 'nan':
            return None
        return str(value).strip()
    
    def process_row(self, cursor, row_num, row, total):
        """处理单行数据"""
        legal_name = self.get_cell_value(row, self.legal_column)
        bank_name = self.get_cell_value(row, self.bank_name_column)
        bank_card_no = self.get_cell_value(row, self.bank_card_column)
        inventory_id = self.get_cell_value(row, self.inventory_column)
        
        if not inventory_id:
            self.results['skipped_no_inventory'] += 1
            return {'status': '跳过', 'message': '库存ID为空'}
        
        # 检查库存记录是否存在
        cursor.execute("""
            SELECT id, bank_id FROM ba_account_mag 
            WHERE id = %s
        """, (inventory_id,))
        inventory_record = cursor.fetchone()
        
        if not inventory_record:
            self.results['skipped_no_inventory'] += 1
            return {'status': '跳过', 'message': f'未找到库存ID={inventory_id}的记录'}
        
        # 检查是否已有bank_id
        if inventory_record['bank_id'] and inventory_record['bank_id'] > 0:
            self.results['skipped_already_has'] += 1
            return {'status': '跳过', 'message': f'库存ID={inventory_id}已有bank_id={inventory_record["bank_id"]}'}
        
        # 查找银行ID
        bank_id = self.find_bank_id(cursor, legal_name, bank_name, bank_card_no, row_num)
        
        if not bank_id:
            self.results['skipped_no_bank'] += 1
            return {'status': '跳过', 'message': '未找到对应的银行记录'}
        
        # 更新库存记录的bank_id
        current_time = int(datetime.now().timestamp())
        cursor.execute("""
            UPDATE ba_account_mag 
            SET bank_id = %s, update_time = %s 
            WHERE id = %s
        """, (bank_id, current_time, inventory_id))
        
        self.results['success_count'] += 1
        self.log_message.emit(f"第{row_num}行: 库存ID={inventory_id} 更新bank_id={bank_id}")
        return {'status': '成功', 'message': f'更新bank_id={bank_id}', 'bank_id': bank_id}
    
    def find_bank_id(self, cursor, legal_name, bank_name, bank_card_no, row_num):
        """查找银行ID"""
        # 优先用银行卡号查找
        if bank_card_no:
            cursor.execute("""
                SELECT id FROM ba_zhb_bank 
                WHERE bank_id = %s AND (delete_time IS NULL OR delete_time = 0)
                ORDER BY update_time DESC LIMIT 1
            """, (bank_card_no,))
            result = cursor.fetchone()
            if result:
                self.log_message.emit(f"  第{row_num}行: 通过银行卡号 {bank_card_no} 找到bank_id={result['id']}")
                return result['id']
        
        # 如果银行卡号没找到，用法人+银行名称查找
        if not legal_name or not bank_name:
            return None
        
        # 查找客户ID
        cursor.execute("""
            SELECT id FROM ba_rlb_customer 
            WHERE legal_name = %s AND (delete_time IS NULL OR delete_time = 0)
        """, (legal_name,))
        customer_record = cursor.fetchone()
        if not customer_record:
            return None
        
        customer_id = customer_record['id']
        
        # 查找银行名称ID
        cursor.execute("""
            SELECT id FROM ba_zhb_bank_name 
            WHERE bank_name = %s AND (delete_time IS NULL OR delete_time = 0)
        """, (bank_name,))
        bank_name_record = cursor.fetchone()
        if not bank_name_record:
            return None
        
        bank_name_id = bank_name_record['id']
        
        # 查找银行记录
        cursor.execute("""
            SELECT id FROM ba_zhb_bank 
            WHERE rlb_customer_id = %s AND bank_name_id = %s AND (delete_time IS NULL OR delete_time = 0)
            ORDER BY update_time DESC LIMIT 1
        """, (customer_id, bank_name_id))
        result = cursor.fetchone()
        if result:
            self.log_message.emit(f"  第{row_num}行: 通过法人+银行名称找到bank_id={result['id']}")
            return result['id']
        
        return None


class InventoryBankFillWidget(QWidget):
    """根据库存ID补全银行卡功能组件"""
    
    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.excel_file = None
        self.excel_columns = []
        self.completed_data = []
        self.failed_records = []
        self.init_ui()
        self.load_datasources()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # 数据源选择
        ds_group = QGroupBox("选择目标数据库")
        ds_layout = QFormLayout()
        
        self.datasource_combo = QComboBox()
        self.datasource_combo.currentTextChanged.connect(self.on_datasource_changed)
        ds_layout.addRow("数据源:", self.datasource_combo)
        
        self.test_connection_btn = QPushButton("测试连接")
        self.test_connection_btn.clicked.connect(self.test_connection)
        self.test_connection_btn.setEnabled(False)
        ds_layout.addRow("", self.test_connection_btn)
        
        ds_group.setLayout(ds_layout)
        layout.addWidget(ds_group)
        
        # Excel文件选择
        file_group = QGroupBox("Excel文件选择")
        file_layout = QFormLayout()
        
        file_row = QHBoxLayout()
        self.file_label = QLabel("未选择文件")
        self.file_label.setStyleSheet("color: gray;")
        file_row.addWidget(self.file_label)
        self.select_file_btn = QPushButton("选择Excel文件")
        self.select_file_btn.clicked.connect(self.select_excel_file)
        file_row.addWidget(self.select_file_btn)
        file_layout.addRow("Excel文件:", file_row)
        
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # 列选择
        column_group = QGroupBox("列选择")
        column_layout = QFormLayout()
        
        self.legal_column_combo = QComboBox()
        self.legal_column_combo.setEnabled(False)
        column_layout.addRow("法人名称列:", self.legal_column_combo)
        
        self.bank_name_column_combo = QComboBox()
        self.bank_name_column_combo.setEnabled(False)
        column_layout.addRow("银行名称列:", self.bank_name_column_combo)
        
        self.bank_card_column_combo = QComboBox()
        self.bank_card_column_combo.setEnabled(False)
        column_layout.addRow("银行卡号列:", self.bank_card_column_combo)
        
        self.inventory_column_combo = QComboBox()
        self.inventory_column_combo.setEnabled(False)
        column_layout.addRow("库存ID列:", self.inventory_column_combo)
        
        column_group.setLayout(column_layout)
        layout.addWidget(column_group)
        
        # 功能说明
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        info_label = QLabel("""
<b>根据库存ID补全银行卡说明:</b><br>
<b>处理逻辑:</b><br>
• 1. 读取Excel中的法人名称、银行名称、银行卡号、库存ID<br>
• 2. 优先用银行卡号在ba_zhb_bank表查找（字段: bank_id）<br>
• 3. 如果没找到，用法人名称+银行名称组合查找:<br>
&nbsp;&nbsp;- 法人名称 → ba_rlb_customer.legal_name → 获取customer_id<br>
&nbsp;&nbsp;- 银行名称 → ba_zhb_bank_name.bank_name → 获取bank_name_id<br>
&nbsp;&nbsp;- 用customer_id + bank_name_id查找ba_zhb_bank表<br>
• 4. 找到后更新ba_account_mag表的bank_id字段<br><br>
<b>跳过情况:</b><br>
• 库存ID为空或不存在<br>
• 库存记录已有bank_id<br>
• 未找到对应的银行记录（说明没绑定银行）
        """)
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # 操作按钮
        btn_layout = QHBoxLayout()
        
        self.start_btn = QPushButton("开始补全")
        self.start_btn.clicked.connect(self.start_process)
        self.start_btn.setEnabled(False)
        btn_layout.addWidget(self.start_btn)
        
        self.export_btn = QPushButton("导出完整结果")
        self.export_btn.clicked.connect(self.export_completed)
        self.export_btn.setEnabled(False)
        btn_layout.addWidget(self.export_btn)
        
        self.export_failed_btn = QPushButton("导出失败记录")
        self.export_failed_btn.clicked.connect(self.export_failed)
        self.export_failed_btn.setEnabled(False)
        btn_layout.addWidget(self.export_failed_btn)
        
        btn_layout.addStretch()
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        btn_layout.addWidget(self.progress_bar)
        
        layout.addLayout(btn_layout)
        
        # 日志
        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout()
        
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(180)
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        
        log_btn_layout = QHBoxLayout()
        self.clear_log_btn = QPushButton("清空日志")
        self.clear_log_btn.clicked.connect(self.log_text.clear)
        log_btn_layout.addWidget(self.clear_log_btn)
        log_btn_layout.addStretch()
        log_layout.addLayout(log_btn_layout)
        
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)
        
        # 结果
        self.result_label = QLabel("")
        self.result_label.setAlignment(Qt.AlignCenter)
        font = QFont()
        font.setBold(True)
        self.result_label.setFont(font)
        layout.addWidget(self.result_label)
        
        self.setLayout(layout)
    
    def load_datasources(self):
        self.datasource_combo.clear()
        self.datasource_combo.addItem("请选择数据源", None)
        for ds in self.db_manager.get_all_datasources():
            self.datasource_combo.addItem(f"{ds.name} ({ds.host}:{ds.port})", ds)
    
    def on_datasource_changed(self):
        ds = self.datasource_combo.currentData()
        self.test_connection_btn.setEnabled(ds is not None)
        self.update_start_btn_state()
    
    def test_connection(self):
        ds = self.datasource_combo.currentData()
        if not ds:
            return
        ok, msg = self.db_manager.test_connection(ds)
        if ok:
            QMessageBox.information(self, "连接成功", msg)
        else:
            QMessageBox.critical(self, "连接失败", msg)
    
    def select_excel_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", "", "Excel文件 (*.xlsx *.xls)")
        if not path:
            return
        try:
            df = pd.read_excel(path, engine='openpyxl', nrows=0)
            self.excel_columns = list(df.columns)
            self.excel_file = path
            self.file_label.setText(os.path.basename(path))
            self.file_label.setStyleSheet("color: green;")
            
            for combo in [self.legal_column_combo, self.bank_name_column_combo, 
                         self.bank_card_column_combo, self.inventory_column_combo]:
                combo.clear()
                combo.addItem("请选择列")
                combo.addItems(self.excel_columns)
                combo.setEnabled(True)
            
            self.log_text.append(f"已加载Excel文件，共 {len(self.excel_columns)} 列")
            self.update_start_btn_state()
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取Excel失败: {str(e)}")
            self.excel_file = None
            self.file_label.setText("未选择文件")
            self.file_label.setStyleSheet("color: gray;")
    
    def update_start_btn_state(self):
        ds_ok = self.datasource_combo.currentData() is not None
        file_ok = self.excel_file is not None
        self.start_btn.setEnabled(ds_ok and file_ok)
    
    def start_process(self):
        ds = self.datasource_combo.currentData()
        if not ds or not self.excel_file:
            QMessageBox.warning(self, "警告", "请选择数据源和Excel文件！")
            return
        
        legal_col = self.legal_column_combo.currentText()
        bank_name_col = self.bank_name_column_combo.currentText()
        bank_card_col = self.bank_card_column_combo.currentText()
        inventory_col = self.inventory_column_combo.currentText()
        
        if legal_col == "请选择列" or bank_name_col == "请选择列" or bank_card_col == "请选择列" or inventory_col == "请选择列":
            QMessageBox.warning(self, "警告", "请选择所有必要的列！")
            return
        
        reply = QMessageBox.question(
            self,
            "确认补全",
            f"确定要根据库存ID补全银行卡吗？\n\n"
            f"数据库: {ds.name}\n"
            f"Excel文件: {os.path.basename(self.excel_file)}\n"
            f"法人名称列: {legal_col}\n"
            f"银行名称列: {bank_name_col}\n"
            f"银行卡号列: {bank_card_col}\n"
            f"库存ID列: {inventory_col}\n\n"
            f"此操作将更新ba_account_mag表的bank_id字段。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        self.start_btn.setEnabled(False)
        self.export_btn.setEnabled(False)
        self.export_failed_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在处理中...")
        self.log_text.clear()
        self.completed_data = []
        self.failed_records = []
        
        self.worker = InventoryBankFillWorker(
            ds, self.excel_file, legal_col, bank_name_col, bank_card_col, inventory_col
        )
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log_message.connect(self.append_log)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()
    
    def append_log(self, msg):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {msg}")
    
    def on_finished(self, results):
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.completed_data = results['completed_data']
        self.failed_records = results['failed_records']
        
        result_text = (
            f"处理完成！\n"
            f"总计: {results['total_rows']} 行\n"
            f"成功: {results['success_count']} 行\n"
            f"跳过(无银行记录): {results['skipped_no_bank']} 行\n"
            f"跳过(无库存记录): {results['skipped_no_inventory']} 行\n"
            f"跳过(已有bank_id): {results['skipped_already_has']} 行\n"
            f"失败: {results['failed_count']} 行"
        )
        
        self.result_label.setText(result_text)
        self.result_label.setStyleSheet("color: green;" if results['failed_count'] == 0 else "color: orange;")
        
        self.export_btn.setEnabled(len(self.completed_data) > 0)
        self.export_failed_btn.setEnabled(len(self.failed_records) > 0)
        
        QMessageBox.information(self, "处理完成", result_text)
    
    def on_error(self, error_msg):
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.result_label.setText("处理失败！")
        self.result_label.setStyleSheet("color: red;")
        QMessageBox.critical(self, "错误", f"处理失败: {error_msg}")
    
    def export_completed(self):
        if not self.completed_data:
            QMessageBox.information(self, "提示", "没有数据可导出")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "保存完整结果",
            f"库存银行卡补全_完整结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel文件 (*.xlsx)"
        )
        if path:
            try:
                df = pd.DataFrame(self.completed_data)
                df.to_excel(path, index=False, engine='openpyxl')
                QMessageBox.information(self, "成功", f"已导出到: {path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出失败: {str(e)}")
    
    def export_failed(self):
        if not self.failed_records:
            QMessageBox.information(self, "提示", "没有失败记录")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "保存失败记录",
            f"库存银行卡补全_失败记录_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel文件 (*.xlsx)"
        )
        if path:
            try:
                data = []
                for r in self.failed_records:
                    row = r['data'].copy()
                    row['失败行号'] = r['row']
                    row['失败原因'] = r['error']
                    data.append(row)
                df = pd.DataFrame(data)
                df.to_excel(path, index=False, engine='openpyxl')
                QMessageBox.information(self, "成功", f"已导出到: {path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出失败: {str(e)}")
