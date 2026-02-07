"""
修正is_customer字段功能界面
"""
import os
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog,
                             QTableWidget, QTableWidgetItem, QHeaderView, QCheckBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from database import DatabaseManager
import pandas as pd
import pymysql


class CustomerStatusFixWorker(QThread):
    """修正is_customer字段工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)  # 返回统计结果
    error = pyqtSignal(str)
    
    def __init__(self, datasource, db_manager, excel_file, compare_column, selected_records):
        super().__init__()
        self.datasource = datasource
        self.db_manager = db_manager
        self.excel_file = excel_file
        self.compare_column = compare_column
        self.selected_records = selected_records
        self.results = {
            'total_selected': len(selected_records),
            'success_count': 0,
            'failed_count': 0,
            'skipped_count': 0,
            'failed_records': []
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
            
            # 读取Excel文件
            self.log_message.emit("开始读取Excel文件...")
            try:
                df = pd.read_excel(self.excel_file, engine='openpyxl')
                self.log_message.emit(f"成功读取Excel文件，共 {len(df)} 行数据")
            except Exception as e:
                self.error.emit(f"读取Excel文件失败: {str(e)}")
                return
            
            self.progress.emit(15)
            
            # 处理选中的记录
            for i, record_index in enumerate(self.selected_records):
                try:
                    if record_index >= len(df):
                        continue
                        
                    row = df.iloc[record_index]
                    self.process_row(cursor, record_index + 1, row)
                    
                    # 更新进度
                    progress = 15 + int((i + 1) / len(self.selected_records) * 80)
                    self.progress.emit(progress)
                    
                except Exception as e:
                    self.results['failed_count'] += 1
                    error_msg = f"第{record_index + 1}行处理失败: {str(e)}"
                    self.log_message.emit(error_msg)
                    
                    self.results['failed_records'].append({
                        'row': record_index + 1,
                        'error': str(e),
                        'data': row.to_dict() if 'row' in locals() else {}
                    })
            
            # 提交事务
            connection.commit()
            cursor.close()
            connection.close()
            
            self.progress.emit(100)
            self.finished.emit(self.results)
            
        except Exception as e:
            self.error.emit(str(e))
    
    def process_row(self, cursor, row_num, row):
        """处理单行数据"""
        try:
            # 清理单元格值函数
            def clean_cell_value(cell_value):
                if cell_value is None or str(cell_value) == 'nan':
                    return ""
                try:
                    if isinstance(cell_value, (int, float)) and cell_value == int(cell_value):
                        return str(int(cell_value))
                    else:
                        return str(cell_value).replace('.0', '') if str(cell_value).endswith('.0') else str(cell_value)
                except:
                    return str(cell_value)
            
            # 获取对比列的值（法人姓名）
            legal_name = clean_cell_value(row.iloc[self.compare_column]) if len(row) > self.compare_column else ""
            
            # 获取G列的值（属性判断）
            col_g = clean_cell_value(row.iloc[6]) if len(row) > 6 else ""  # G列索引为6
            
            if not legal_name:
                self.results['skipped_count'] += 1
                self.log_message.emit(f"第{row_num}行: 法人姓名为空，跳过")
                return
            
            # 如果G列是"企业法人"，跳过处理
            if col_g == "企业法人":
                self.results['skipped_count'] += 1
                self.log_message.emit(f"第{row_num}行: {legal_name} 是企业法人，跳过")
                return
            
            self.log_message.emit(f"第{row_num}行: 处理法人={legal_name}, 属性={col_g}")
            
            # 在ba_rlb_customer表中查找对应的记录
            cursor.execute("""
                SELECT id, legal_name, is_customer 
                FROM ba_rlb_customer 
                WHERE legal_name = %s AND (delete_time IS NULL OR delete_time = 0)
            """, (legal_name,))
            
            customer_results = cursor.fetchall()
            
            if not customer_results:
                self.results['skipped_count'] += 1
                self.log_message.emit(f"第{row_num}行: 未找到法人 {legal_name} 的客户记录")
                return
            
            # 更新找到的所有匹配记录
            updated_count = 0
            current_time = int(datetime.now().timestamp())
            
            for customer_record in customer_results:
                customer_id, db_legal_name, current_is_customer = customer_record
                
                # 如果is_customer已经是1，跳过
                if current_is_customer == 1:
                    self.log_message.emit(f"第{row_num}行: 客户ID {customer_id} ({db_legal_name}) is_customer已经是1，跳过")
                    continue
                
                # 更新is_customer为1
                cursor.execute("""
                    UPDATE ba_rlb_customer 
                    SET is_customer = 1, update_time = %s 
                    WHERE id = %s
                """, (current_time, customer_id))
                
                updated_count += 1
                self.log_message.emit(f"第{row_num}行: 更新客户ID {customer_id} ({db_legal_name}) is_customer = 1")
            
            if updated_count > 0:
                self.results['success_count'] += 1
                self.log_message.emit(f"第{row_num}行: 成功更新 {updated_count} 条记录")
            else:
                self.results['skipped_count'] += 1
                self.log_message.emit(f"第{row_num}行: 所有匹配记录的is_customer已经是1，无需更新")
                
        except Exception as e:
            raise Exception(f"处理第{row_num}行数据失败: {str(e)}")


class CustomerStatusFixWidget(QWidget):
    """修正is_customer字段功能组件"""
    
    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.excel_file = None
        self.excel_data = None
        self.init_ui()
        self.load_datasources()
    
    def init_ui(self):
        """初始化UI"""
        layout = QVBoxLayout()
        
        # 数据源选择区域
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
        
        # Excel文件选择区域
        file_group = QGroupBox("Excel文件选择")
        file_layout = QFormLayout()
        
        file_select_layout = QHBoxLayout()
        self.file_label = QLabel("未选择文件")
        self.file_label.setStyleSheet("color: gray;")
        file_select_layout.addWidget(self.file_label)
        
        self.select_file_btn = QPushButton("选择Excel文件")
        self.select_file_btn.clicked.connect(self.select_excel_file)
        file_select_layout.addWidget(self.select_file_btn)
        
        file_layout.addRow("Excel文件:", file_select_layout)
        
        # 列选择
        self.column_combo = QComboBox()
        self.column_combo.setEnabled(False)
        file_layout.addRow("对比列(法人姓名):", self.column_combo)
        
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # 功能说明区域
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        
        info_label = QLabel("""
<b>修正is_customer字段说明:</b><br>
<b>Excel列要求:</b><br>
• 可选择任意列作为法人姓名对比列<br>
• G列: 属性判断 (如果是"企业法人"将跳过处理)<br><br>
<b>处理逻辑:</b><br>
• 1. 读取Excel文件，获取表头供用户选择对比列<br>
• 2. 根据选择的列与ba_rlb_customer表的legal_name字段进行匹配<br>
• 3. 如果G列是"企业法人"，跳过该记录<br>
• 4. 找到匹配记录后，将is_customer字段设置为1<br>
• 5. 如果is_customer已经是1，则跳过更新<br><br>
<b>更新条件:</b><br>
• 法人姓名匹配成功<br>
• G列不是"企业法人"<br>
• 当前is_customer不等于1<br>
• 记录未被删除(delete_time为空或0)<br><br>
<b>操作流程:</b><br>
1. 选择数据库并测试连接<br>
2. 选择Excel文件，系统自动读取表头<br>
3. 选择用于对比的法人姓名列<br>
4. 点击"加载数据"预览待处理记录<br>
5. 选择需要处理的记录<br>
6. 点击"开始修正"执行更新操作
        """)
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # 数据预览区域
        preview_group = QGroupBox("数据预览")
        preview_layout = QVBoxLayout()
        
        preview_button_layout = QHBoxLayout()
        self.load_data_btn = QPushButton("加载数据")
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
        
        preview_button_layout.addStretch()
        preview_layout.addLayout(preview_button_layout)
        
        # 数据表格
        self.records_table = QTableWidget()
        self.records_table.setMaximumHeight(300)
        preview_layout.addWidget(self.records_table)
        
        preview_group.setLayout(preview_layout)
        layout.addWidget(preview_group)
        
        # 操作按钮区域
        button_layout = QHBoxLayout()
        
        self.start_btn = QPushButton("开始修正")
        self.start_btn.clicked.connect(self.start_fix)
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
            display_name = f"{ds.name} ({ds.host}:{ds.port})"
            self.datasource_combo.addItem(display_name, ds)
    
    def on_datasource_changed(self):
        """数据源选择变化"""
        datasource = self.datasource_combo.currentData()
        self.test_connection_btn.setEnabled(datasource is not None)
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
    
    def select_excel_file(self):
        """选择Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Excel文件",
            "",
            "Excel文件 (*.xlsx *.xls)"
        )
        
        if file_path:
            self.excel_file = file_path
            self.file_label.setText(os.path.basename(file_path))
            self.file_label.setStyleSheet("color: green;")
            
            # 读取Excel表头
            try:
                df = pd.read_excel(file_path, engine='openpyxl')
                self.excel_data = df
                
                # 填充列选择下拉框
                self.column_combo.clear()
                for i, col in enumerate(df.columns):
                    self.column_combo.addItem(f"{chr(65+i)}列: {col}", i)
                
                self.column_combo.setEnabled(True)
                self.update_button_states()
                
            except Exception as e:
                QMessageBox.critical(self, "文件读取失败", f"读取Excel文件失败：{str(e)}")
                self.excel_file = None
                self.file_label.setText("未选择文件")
                self.file_label.setStyleSheet("color: gray;")
    
    def update_button_states(self):
        """更新按钮状态"""
        datasource_selected = self.datasource_combo.currentData() is not None
        file_selected = self.excel_file is not None
        column_selected = self.column_combo.currentData() is not None
        
        self.load_data_btn.setEnabled(datasource_selected and file_selected and column_selected)
    
    def load_data(self):
        """加载数据预览"""
        if not self.excel_data is not None:
            return
        
        try:
            compare_column = self.column_combo.currentData()
            
            # 设置表格
            self.records_table.setRowCount(len(self.excel_data))
            headers = ["选择", "行号", "法人姓名", "G列(属性)", "处理状态"]
            self.records_table.setColumnCount(len(headers))
            self.records_table.setHorizontalHeaderLabels(headers)
            
            # 填充数据
            for index, row in self.excel_data.iterrows():
                # 选择框
                checkbox = QCheckBox()
                self.records_table.setCellWidget(index, 0, checkbox)
                
                # 行号
                self.records_table.setItem(index, 1, QTableWidgetItem(str(index + 1)))
                
                # 法人姓名
                legal_name = str(row.iloc[compare_column]) if len(row) > compare_column else ""
                if legal_name == 'nan':
                    legal_name = ""
                self.records_table.setItem(index, 2, QTableWidgetItem(legal_name))
                
                # G列属性
                col_g = str(row.iloc[6]) if len(row) > 6 else ""
                if col_g == 'nan':
                    col_g = ""
                self.records_table.setItem(index, 3, QTableWidgetItem(col_g))
                
                # 处理状态
                if not legal_name:
                    status = "法人姓名为空"
                    checkbox.setEnabled(False)
                elif col_g == "企业法人":
                    status = "企业法人(跳过)"
                    checkbox.setEnabled(False)
                else:
                    status = "待处理"
                    checkbox.setChecked(True)
                
                self.records_table.setItem(index, 4, QTableWidgetItem(status))
            
            # 调整列宽
            header = self.records_table.horizontalHeader()
            header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(2, QHeaderView.Stretch)
            header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
            
            # 启用按钮
            self.select_all_btn.setEnabled(True)
            self.deselect_all_btn.setEnabled(True)
            self.start_btn.setEnabled(True)
            
        except Exception as e:
            QMessageBox.critical(self, "加载失败", f"加载数据失败：{str(e)}")
    
    def select_all_records(self):
        """全选记录"""
        for row in range(self.records_table.rowCount()):
            checkbox = self.records_table.cellWidget(row, 0)
            if checkbox and checkbox.isEnabled():
                checkbox.setChecked(True)
    
    def deselect_all_records(self):
        """全不选记录"""
        for row in range(self.records_table.rowCount()):
            checkbox = self.records_table.cellWidget(row, 0)
            if checkbox:
                checkbox.setChecked(False)
    
    def get_selected_records(self):
        """获取选中的记录索引"""
        selected = []
        for row in range(self.records_table.rowCount()):
            checkbox = self.records_table.cellWidget(row, 0)
            if checkbox and checkbox.isChecked():
                selected.append(row)
        return selected
    
    def start_fix(self):
        """开始修正"""
        datasource = self.datasource_combo.currentData()
        if not datasource or not self.excel_file:
            QMessageBox.warning(self, "警告", "请选择数据源和Excel文件！")
            return
        
        compare_column = self.column_combo.currentData()
        if compare_column is None:
            QMessageBox.warning(self, "警告", "请选择对比列！")
            return
        
        selected_records = self.get_selected_records()
        if not selected_records:
            QMessageBox.warning(self, "警告", "请选择要处理的记录！")
            return
        
        # 确认对话框
        reply = QMessageBox.question(
            self,
            "确认修正",
            f"确定要修正is_customer字段吗？\n\n"
            f"数据库: {datasource.name}\n"
            f"Excel文件: {os.path.basename(self.excel_file)}\n"
            f"对比列: {self.column_combo.currentText()}\n"
            f"选中记录: {len(selected_records)} 条\n\n"
            f"此操作将根据Excel中的法人姓名\n"
            f"更新ba_rlb_customer表的is_customer字段为1。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        # 禁用按钮，显示进度条
        self.start_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在修正中...")
        self.log_text.clear()
        
        # 启动工作线程
        self.worker = CustomerStatusFixWorker(
            datasource, self.db_manager, self.excel_file, 
            compare_column, selected_records
        )
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
        """修正完成"""
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        result_text = (
            f"修正完成！\n"
            f"选中记录：{results['total_selected']} 条\n"
            f"成功修正：{results['success_count']} 条\n"
            f"跳过处理：{results['skipped_count']} 条\n"
            f"处理失败：{results['failed_count']} 条"
        )
        
        self.result_label.setText(result_text)
        if results['failed_count'] > 0:
            self.result_label.setStyleSheet("color: orange;")
        else:
            self.result_label.setStyleSheet("color: green;")
        
        QMessageBox.information(self, "修正完成", result_text)
    
    def on_fix_error(self, error_msg):
        """修正错误"""
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.result_label.setText("修正失败！")
        self.result_label.setStyleSheet("color: red;")
        
        QMessageBox.critical(self, "修正错误", f"修正失败：{error_msg}")
    
    def clear_log(self):
        """清空日志"""
        self.log_text.clear()
