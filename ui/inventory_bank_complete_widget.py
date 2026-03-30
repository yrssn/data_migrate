"""
补全现有库存关联银行卡信息功能界面
"""
import os
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog,
                             QTableWidget, QTableWidgetItem)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from database import DatabaseManager
import pandas as pd
import pymysql


class InventoryBankCompleteWorker(QThread):
    """库存银行卡信息补全工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    
    def __init__(self, source_datasource, db_manager, excel_file, column_mapping):
        super().__init__()
        self.source_datasource = source_datasource
        self.db_manager = db_manager
        self.excel_file = excel_file
        self.column_mapping = column_mapping
        self.results = {
            'total_rows': 0,
            'success_count': 0,
            'failed_count': 0,
            'completed_data': [],
            'failed_records': []
        }
    
    def run(self):
        try:
            self.log_message.emit("开始读取Excel文件...")
            self.progress.emit(5)
            
            try:
                df = pd.read_excel(self.excel_file, engine='openpyxl')
                self.log_message.emit(f"成功读取Excel文件，共 {len(df)} 行数据")
            except Exception as e:
                self.error.emit(f"读取Excel文件失败: {str(e)}")
                return
            
            self.results['total_rows'] = len(df)
            self.progress.emit(10)
            
            self.log_message.emit("开始连接源数据库...")
            source_connection = pymysql.connect(
                host=self.source_datasource.host,
                port=self.source_datasource.port,
                user=self.source_datasource.username,
                password=self.source_datasource.password,
                database=self.source_datasource.database,
                charset=self.source_datasource.charset
            )
            source_cursor = source_connection.cursor()
            self.log_message.emit("源数据库连接成功")
            self.progress.emit(15)
            
            for index, row in df.iterrows():
                try:
                    completed_row = self.process_row(source_cursor, index + 1, row)
                    self.results['completed_data'].append(completed_row)
                    self.results['success_count'] += 1
                    
                    progress = 15 + int((index + 1) / len(df) * 80)
                    self.progress.emit(progress)
                    
                except Exception as e:
                    self.results['failed_count'] += 1
                    error_msg = f"第{index + 1}行处理失败: {str(e)}"
                    self.log_message.emit(error_msg)
                    
                    failed_row = row.to_dict()
                    failed_row['处理状态'] = '失败'
                    failed_row['失败原因'] = str(e)
                    failed_row['银行卡号'] = ''
                    failed_row['银行名称'] = ''
                    failed_row['银行类型'] = ''
                    failed_row['货币代码'] = ''
                    
                    self.results['completed_data'].append(failed_row)
                    self.results['failed_records'].append({
                        'row': index + 1,
                        'error': str(e),
                        'data': row.to_dict()
                    })
            
            source_cursor.close()
            source_connection.close()
            
            self.progress.emit(100)
            self.finished.emit(self.results)
            
        except Exception as e:
            self.error.emit(str(e))
    
    def get_cell_value(self, row, col_index):
        """获取单元格值"""
        if col_index is None or col_index < 0 or col_index >= len(row):
            return None
        value = row.iloc[col_index]
        if pd.isna(value) or str(value).strip() == '' or str(value).strip().lower() == 'nan':
            return None
        return str(value).strip()
    
    def process_row(self, source_cursor, row_num, row):
        """处理单行数据"""
        try:
            # 获取法人姓名
            legal_name_col = self.column_mapping.get('legal_name')
            legal_name = self.get_cell_value(row, legal_name_col)
            
            if not legal_name:
                raise Exception("法人姓名为空")
            
            # 获取店铺类型名称
            site_type_col = self.column_mapping.get('site_type')
            site_type_name = self.get_cell_value(row, site_type_col)
            
            if not site_type_name:
                raise Exception("店铺类型名称为空")
            
            self.log_message.emit(f"第{row_num}行: 处理法人={legal_name}, 店铺类型={site_type_name}")
            
            # 第一步：根据法人姓名找到legal_id
            source_cursor.execute("""
                SELECT id FROM ea_dy_legal 
                WHERE name = %s AND status = 1 AND (delete_time IS NULL OR delete_time = 0)
            """, (legal_name,))
            legal_result = source_cursor.fetchone()
            
            if not legal_result:
                raise Exception(f"未找到法人: {legal_name}")
            
            legal_id = legal_result[0]
            self.log_message.emit(f"第{row_num}行: 找到法人ID={legal_id}")
            
            # 第二步：根据店铺类型名称找到site_type_id
            source_cursor.execute("""
                SELECT id FROM ea_dy_site_type 
                WHERE name = %s AND status = 1 AND (delete_time IS NULL OR delete_time = 0)
            """, (site_type_name,))
            site_type_result = source_cursor.fetchone()
            
            if not site_type_result:
                raise Exception(f"未找到店铺类型: {site_type_name}")
            
            site_type_id = site_type_result[0]
            self.log_message.emit(f"第{row_num}行: 找到店铺类型ID={site_type_id}")
            
            # 第三步：根据legal_id和siteType找到ea_dy_shop记录
            source_cursor.execute("""
                SELECT id, card_id, name, currency_id FROM ea_dy_shop 
                WHERE legal_id = %s AND siteType = %s
            """, (legal_id, site_type_id))
            shop_result = source_cursor.fetchone()
            
            if not shop_result:
                raise Exception(f"未找到对应的店铺记录 (法人ID:{legal_id}, 店铺类型ID:{site_type_id})")
            
            shop_id, card_id, shop_name, shop_currency_id = shop_result
            self.log_message.emit(f"第{row_num}行: 找到店铺ID={shop_id}, 银行卡ID={card_id}")
            
            # 第四步：根据card_id获取银行卡详细信息
            bank_info = self.get_bank_card_info(source_cursor, card_id, row_num)
            
            # 构建补全后的行数据
            completed_row = row.to_dict()
            completed_row['处理状态'] = '成功'
            completed_row['失败原因'] = ''
            completed_row['店铺ID'] = shop_id
            completed_row['店铺名称'] = shop_name
            completed_row['银行卡ID'] = card_id
            completed_row['银行卡号'] = bank_info['account']
            completed_row['银行名称'] = bank_info['bank_name']
            completed_row['银行类型'] = bank_info['bank_type']
            completed_row['货币代码'] = bank_info['currency_code']
            completed_row['银行卡余额'] = bank_info['balance']
            
            self.log_message.emit(f"第{row_num}行: 成功补全银行卡信息 - {bank_info['bank_name']} {bank_info['account']}")
            
            return completed_row
            
        except Exception as e:
            raise e
    
    def get_bank_card_info(self, source_cursor, card_id, row_num):
        """获取银行卡详细信息"""
        try:
            source_cursor.execute("""
                SELECT 
                    lc.account,
                    lc.balance,
                    lc.bank_card_type,
                    bc.name as bank_name,
                    bc.type as bank_type,
                    cur.code as currency_code,
                    cur.name as currency_name
                FROM ea_dy_legal_cards lc
                LEFT JOIN ea_dy_bankcard bc ON lc.bankcard_id = bc.id
                LEFT JOIN ea_dy_currency cur ON lc.currency_id = cur.id
                WHERE lc.id = %s AND (lc.delete_time IS NULL OR lc.delete_time = 0)
            """, (card_id,))
            
            card_result = source_cursor.fetchone()
            
            if not card_result:
                raise Exception(f"未找到银行卡信息 (card_id: {card_id})")
            
            (account, balance, bank_card_type, bank_name, bank_type, 
             currency_code, currency_name) = card_result
            
            card_type_display = ''
            if bank_card_type == 1:
                card_type_display = '个人卡'
            elif bank_card_type == 2:
                card_type_display = '企业卡'
            elif bank_card_type == 0:
                card_type_display = '第三方卡'
            else:
                card_type_display = f'未知类型({bank_card_type})'
            
            bank_type_display = ''
            if bank_type == 0:
                bank_type_display = '第三方'
            elif bank_type == 1:
                bank_type_display = '个人'
            elif bank_type == 2:
                bank_type_display = '企业'
            else:
                bank_type_display = f'未知类型({bank_type})'
            
            return {
                'account': account or '',
                'balance': balance or 0,
                'bank_card_type': card_type_display,
                'bank_name': bank_name or '',
                'bank_type': bank_type_display,
                'currency_code': currency_code or '',
                'currency_name': currency_name or ''
            }
            
        except Exception as e:
            raise Exception(f"获取银行卡信息失败: {str(e)}")


class InventoryBankCompleteWidget(QWidget):
    """补全现有库存关联银行卡信息功能组件"""
    
    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.excel_file = None
        self.excel_data = None
        self.completed_data = []
        self.failed_records = []
        self.init_ui()
        self.load_datasources()
    
    def init_ui(self):
        """初始化UI"""
        layout = QVBoxLayout()
        
        # 数据源选择区域
        datasource_group = QGroupBox("选择源数据库")
        datasource_layout = QFormLayout()
        
        self.source_datasource_combo = QComboBox()
        self.source_datasource_combo.currentTextChanged.connect(self.on_datasource_changed)
        datasource_layout.addRow("源数据库:", self.source_datasource_combo)
        
        self.test_source_btn = QPushButton("测试连接")
        self.test_source_btn.clicked.connect(self.test_source_connection)
        self.test_source_btn.setEnabled(False)
        datasource_layout.addRow("", self.test_source_btn)
        
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
        
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # 列映射配置区域
        mapping_group = QGroupBox("列映射配置（选择Excel文件后可配置）")
        mapping_layout = QFormLayout()
        
        # 法人姓名列（必选）
        self.legal_name_col_combo = QComboBox()
        self.legal_name_col_combo.setEnabled(False)
        self.legal_name_col_combo.currentIndexChanged.connect(self.update_start_button_state)
        mapping_layout.addRow("法人姓名列 (用于查找ea_dy_legal表):", self.legal_name_col_combo)
        
        # 店铺类型列（必选）
        self.site_type_col_combo = QComboBox()
        self.site_type_col_combo.setEnabled(False)
        self.site_type_col_combo.currentIndexChanged.connect(self.update_start_button_state)
        mapping_layout.addRow("店铺类型列 (用于查找ea_dy_site_type表):", self.site_type_col_combo)
        
        mapping_group.setLayout(mapping_layout)
        layout.addWidget(mapping_group)
        
        # 数据预览区域
        preview_group = QGroupBox("数据预览（前5行）")
        preview_layout = QVBoxLayout()
        
        self.preview_table = QTableWidget()
        self.preview_table.setMaximumHeight(120)
        preview_layout.addWidget(self.preview_table)
        
        preview_group.setLayout(preview_layout)
        layout.addWidget(preview_group)
        
        # 功能说明区域
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        
        info_label = QLabel("""
<b>补全现有库存关联银行卡信息说明:</b><br>
<b>查找逻辑:</b><br>
• 1. 根据法人姓名在ea_dy_legal表中找到legal_id<br>
• 2. 根据店铺类型名称在ea_dy_site_type表中找到site_type_id<br>
• 3. 根据legal_id和siteType在ea_dy_shop表中找到对应店铺记录<br>
• 4. 获取ea_dy_shop.card_id，查询ea_dy_legal_cards获取银行卡信息<br>
<b>补全字段:</b> 店铺ID、店铺名称、银行卡ID、银行卡号、银行卡余额、银行名称、银行类型、货币代码
        """)
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # 操作按钮区域
        button_layout = QHBoxLayout()
        
        self.start_btn = QPushButton("开始补全")
        self.start_btn.clicked.connect(self.start_complete)
        self.start_btn.setEnabled(False)
        button_layout.addWidget(self.start_btn)
        
        self.export_completed_btn = QPushButton("导出补全结果")
        self.export_completed_btn.clicked.connect(self.export_completed_data)
        self.export_completed_btn.setEnabled(False)
        button_layout.addWidget(self.export_completed_btn)
        
        self.export_failed_btn = QPushButton("导出失败记录")
        self.export_failed_btn.clicked.connect(self.export_failed_records)
        self.export_failed_btn.setEnabled(False)
        button_layout.addWidget(self.export_failed_btn)
        
        button_layout.addStretch()
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        button_layout.addWidget(self.progress_bar)
        
        layout.addLayout(button_layout)
        
        # 日志区域
        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout()
        
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(150)
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
        self.source_datasource_combo.clear()
        self.source_datasource_combo.addItem("请选择源数据源", None)
        
        datasources = self.db_manager.get_all_datasources()
        for ds in datasources:
            display_name = f"{ds.name} ({ds.host}:{ds.port})"
            self.source_datasource_combo.addItem(display_name, ds)
    
    def on_datasource_changed(self):
        """数据源选择变化"""
        source_data = self.source_datasource_combo.currentData()
        self.test_source_btn.setEnabled(source_data is not None)
        self.update_start_button_state()
    
    def test_source_connection(self):
        """测试源数据库连接"""
        datasource = self.source_datasource_combo.currentData()
        if not datasource:
            return
        
        success, message = self.db_manager.test_connection(datasource)
        if success:
            QMessageBox.information(self, "源连接成功", message)
        else:
            QMessageBox.critical(self, "源连接失败", message)
    
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
            
            try:
                self.excel_data = pd.read_excel(file_path, engine='openpyxl')
                self.populate_column_combos()
                self.show_preview()
                self.update_start_button_state()
                
            except Exception as e:
                QMessageBox.critical(self, "文件读取失败", f"读取Excel文件失败：{str(e)}")
                self.excel_file = None
                self.excel_data = None
                self.file_label.setText("未选择文件")
                self.file_label.setStyleSheet("color: gray;")
    
    def populate_column_combos(self):
        """填充列选择下拉框"""
        if self.excel_data is None:
            return
        
        columns = self.excel_data.columns.tolist()
        
        col_options = []
        for i, col in enumerate(columns):
            col_letter = chr(65 + i) if i < 26 else f"A{chr(65 + i - 26)}" if i < 52 else f"Col{i+1}"
            col_options.append((f"{col_letter}列: {col}", i))
        
        # 法人姓名列
        self.legal_name_col_combo.clear()
        self.legal_name_col_combo.addItem("请选择列", -1)
        for text, idx in col_options:
            self.legal_name_col_combo.addItem(text, idx)
        self.legal_name_col_combo.setEnabled(True)
        
        # 店铺类型列
        self.site_type_col_combo.clear()
        self.site_type_col_combo.addItem("请选择列", -1)
        for text, idx in col_options:
            self.site_type_col_combo.addItem(text, idx)
        self.site_type_col_combo.setEnabled(True)
        
        # 不自动匹配，让用户手动选择
        # self.auto_match_columns(columns)
    
    def auto_match_columns(self, columns):
        """尝试自动匹配常见列名"""
        for i, col in enumerate(columns):
            col_lower = str(col).lower()
            
            if '法人' in col_lower and ('姓名' in col_lower or '名' in col_lower):
                self.legal_name_col_combo.setCurrentIndex(i + 1)
            elif '店铺' in col_lower and '类型' in col_lower:
                self.site_type_col_combo.setCurrentIndex(i + 1)
            elif '类型' in col_lower and '店铺' not in col_lower:
                if self.site_type_col_combo.currentData() == -1:
                    self.site_type_col_combo.setCurrentIndex(i + 1)
    
    def show_preview(self):
        """显示数据预览"""
        if self.excel_data is None:
            return
        
        preview_df = self.excel_data.head(5)
        
        self.preview_table.setRowCount(len(preview_df))
        self.preview_table.setColumnCount(len(preview_df.columns))
        self.preview_table.setHorizontalHeaderLabels([str(col) for col in preview_df.columns])
        
        for i in range(len(preview_df)):
            for j, col in enumerate(preview_df.columns):
                value = preview_df.iloc[i, j]
                item = QTableWidgetItem(str(value) if pd.notna(value) else "")
                self.preview_table.setItem(i, j, item)
        
        self.preview_table.resizeColumnsToContents()
    
    def update_start_button_state(self):
        """更新开始按钮状态"""
        datasource_selected = self.source_datasource_combo.currentData() is not None
        file_selected = self.excel_file is not None
        legal_name_selected = (self.legal_name_col_combo.currentData() is not None and 
                               self.legal_name_col_combo.currentData() >= 0)
        site_type_selected = (self.site_type_col_combo.currentData() is not None and 
                              self.site_type_col_combo.currentData() >= 0)
        
        self.start_btn.setEnabled(datasource_selected and file_selected and 
                                  legal_name_selected and site_type_selected)
    
    def get_column_mapping(self):
        """获取列映射配置"""
        return {
            'legal_name': self.legal_name_col_combo.currentData(),
            'site_type': self.site_type_col_combo.currentData()
        }
    
    def start_complete(self):
        """开始补全"""
        source_datasource = self.source_datasource_combo.currentData()
        if not source_datasource or not self.excel_file:
            QMessageBox.warning(self, "警告", "请选择源数据源和Excel文件！")
            return
        
        column_mapping = self.get_column_mapping()
        
        mapping_info = []
        mapping_info.append(f"法人姓名: {self.legal_name_col_combo.currentText()}")
        mapping_info.append(f"店铺类型: {self.site_type_col_combo.currentText()}")
        
        reply = QMessageBox.question(
            self,
            "确认补全",
            f"确定要补全库存银行卡信息吗？\n\n"
            f"源数据库: {source_datasource.name}\n"
            f"Excel文件: {os.path.basename(self.excel_file)}\n\n"
            f"列映射配置:\n" + "\n".join(mapping_info) + "\n\n"
            f"此操作将根据Excel中的法人姓名和店铺类型\n"
            f"从源数据库中查找对应的银行卡信息进行补全。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        self.start_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在补全中...")
        self.log_text.clear()
        self.completed_data = []
        self.failed_records = []
        
        self.worker = InventoryBankCompleteWorker(source_datasource, self.db_manager, 
                                                   self.excel_file, column_mapping)
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
        self.completed_data = results['completed_data']
        self.failed_records = results['failed_records']
        
        result_text = (
            f"补全完成！\n"
            f"总行数：{results['total_rows']} 行\n"
            f"成功：{results['success_count']} 行\n"
            f"失败：{results['failed_count']} 行"
        )
        
        self.result_label.setText(result_text)
        if results['failed_count'] > 0:
            self.result_label.setStyleSheet("color: orange;")
            self.export_failed_btn.setEnabled(True)
        else:
            self.result_label.setStyleSheet("color: green;")
        
        self.export_completed_btn.setEnabled(True)
        
        QMessageBox.information(self, "补全完成", result_text)
    
    def on_complete_error(self, error_msg):
        """补全错误"""
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.result_label.setText("补全失败！")
        self.result_label.setStyleSheet("color: red;")
        
        QMessageBox.critical(self, "补全错误", f"补全失败：{error_msg}")
    
    def export_completed_data(self):
        """导出补全结果"""
        if not self.completed_data:
            QMessageBox.information(self, "提示", "没有补全数据可导出")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存补全结果",
            f"inventory_bank_completed_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel文件 (*.xlsx)"
        )
        
        if file_path:
            try:
                df = pd.DataFrame(self.completed_data)
                df.to_excel(file_path, index=False, engine='openpyxl')
                
                QMessageBox.information(self, "导出成功", f"补全结果已导出到：\n{file_path}")
                
            except Exception as e:
                QMessageBox.critical(self, "导出失败", f"导出补全结果时出错：{str(e)}")
    
    def export_failed_records(self):
        """导出失败记录"""
        if not self.failed_records:
            QMessageBox.information(self, "提示", "没有失败记录可导出")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存失败记录",
            f"inventory_bank_failed_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel文件 (*.xlsx)"
        )
        
        if file_path:
            try:
                export_data = []
                for record in self.failed_records:
                    row_data = record['data'].copy()
                    row_data['错误行号'] = record['row']
                    row_data['错误原因'] = record['error']
                    export_data.append(row_data)
                
                df = pd.DataFrame(export_data)
                df.to_excel(file_path, index=False, engine='openpyxl')
                
                QMessageBox.information(self, "导出成功", f"失败记录已导出到：\n{file_path}")
                
            except Exception as e:
                QMessageBox.critical(self, "导出失败", f"导出失败记录时出错：{str(e)}")
    
    def clear_log(self):
        """清空日志"""
        self.log_text.clear()
