"""
补全现有库存关联银行卡信息功能界面
"""
import os
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from database import DatabaseManager
import pandas as pd
import pymysql


class InventoryBankCompleteWorker(QThread):
    """库存银行卡信息补全工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)  # 返回统计结果和补全后的数据
    error = pyqtSignal(str)
    
    def __init__(self, source_datasource, db_manager, excel_file):
        super().__init__()
        self.source_datasource = source_datasource
        self.db_manager = db_manager
        self.excel_file = excel_file
        self.results = {
            'total_rows': 0,
            'success_count': 0,
            'failed_count': 0,
            'completed_data': [],  # 补全后的数据
            'failed_records': []
        }
    
    def run(self):
        try:
            self.log_message.emit("开始读取Excel文件...")
            self.progress.emit(5)
            
            # 读取Excel文件
            try:
                df = pd.read_excel(self.excel_file, engine='openpyxl')
                self.log_message.emit(f"成功读取Excel文件，共 {len(df)} 行数据")
            except Exception as e:
                self.error.emit(f"读取Excel文件失败: {str(e)}")
                return
            
            self.results['total_rows'] = len(df)
            self.progress.emit(10)
            
            # 连接源数据库
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
            
            # 处理每一行数据
            for index, row in df.iterrows():
                try:
                    completed_row = self.process_row(source_cursor, index + 1, row)
                    self.results['completed_data'].append(completed_row)
                    self.results['success_count'] += 1
                    
                    # 更新进度
                    progress = 15 + int((index + 1) / len(df) * 80)
                    self.progress.emit(progress)
                    
                except Exception as e:
                    self.results['failed_count'] += 1
                    error_msg = f"第{index + 1}行处理失败: {str(e)}"
                    self.log_message.emit(error_msg)
                    
                    # 记录失败的行，保留原始数据
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
    
    def process_row(self, source_cursor, row_num, row):
        """处理单行数据"""
        try:
            # 获取Excel列数据
            legal_name = str(row.iloc[1]) if len(row) > 1 else ""  # B列：法人姓名
            site_type_name = str(row.iloc[3]) if len(row) > 3 else ""  # D列：店铺类型名称
            
            if not legal_name or legal_name == 'nan':
                raise Exception("B列(法人姓名)为空")
            if not site_type_name or site_type_name == 'nan':
                raise Exception("D列(店铺类型名称)为空")
            
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
            # 查询银行卡信息，关联银行名称和货币信息
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
            
            # 转换银行卡类型显示
            card_type_display = ''
            if bank_card_type == 1:
                card_type_display = '个人卡'
            elif bank_card_type == 2:
                card_type_display = '企业卡'
            elif bank_card_type == 0:
                card_type_display = '第三方卡'
            else:
                card_type_display = f'未知类型({bank_card_type})'
            
            # 转换银行类型显示
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
        
        # 功能说明区域
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        
        info_label = QLabel("""
<b>补全现有库存关联银行卡信息说明:</b><br>
<b>Excel列要求:</b><br>
• B列: 法人姓名 (用于查找ea_dy_legal表)<br>
• D列: 店铺类型名称 (用于查找ea_dy_site_type表)<br><br>
<b>查找逻辑:</b><br>
• 1. 根据B列法人姓名在ea_dy_legal表中找到legal_id<br>
• 2. 根据D列店铺类型名称在ea_dy_site_type表中找到site_type_id<br>
• 3. 根据legal_id和siteType在ea_dy_shop表中找到对应店铺记录<br>
• 4. 获取ea_dy_shop.card_id，查询ea_dy_legal_cards获取银行卡信息<br>
• 5. 关联ea_dy_bankcard和ea_dy_currency获取完整银行信息<br><br>
<b>补全字段:</b><br>
• 店铺ID、店铺名称<br>
• 银行卡ID、银行卡号、银行卡余额<br>
• 银行名称、银行类型 (0未知(第三方)/1个人/2企业)<br>
• 货币代码、银行卡类型 (个人卡/企业卡/第三方卡)<br>
• 处理状态、失败原因<br><br>
<b>导出结果:</b><br>
• 原Excel数据 + 补全的银行卡信息<br>
• 支持导出处理失败的记录进行分析
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
            self.update_start_button_state()
    
    def update_start_button_state(self):
        """更新开始按钮状态"""
        datasource_selected = self.source_datasource_combo.currentData() is not None
        file_selected = self.excel_file is not None
        self.start_btn.setEnabled(datasource_selected and file_selected)
    
    def start_complete(self):
        """开始补全"""
        source_datasource = self.source_datasource_combo.currentData()
        if not source_datasource or not self.excel_file:
            QMessageBox.warning(self, "警告", "请选择源数据源和Excel文件！")
            return
        
        # 确认对话框
        reply = QMessageBox.question(
            self,
            "确认补全",
            f"确定要补全库存银行卡信息吗？\n\n"
            f"源数据库: {source_datasource.name}\n"
            f"Excel文件: {os.path.basename(self.excel_file)}\n\n"
            f"此操作将根据Excel中的法人姓名和店铺类型\n"
            f"从源数据库中查找对应的银行卡信息进行补全。",
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
        self.completed_data = []
        self.failed_records = []
        
        # 启动工作线程
        self.worker = InventoryBankCompleteWorker(source_datasource, self.db_manager, self.excel_file)
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
                # 创建DataFrame并导出
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
                # 准备导出数据
                export_data = []
                for record in self.failed_records:
                    row_data = record['data'].copy()
                    row_data['错误行号'] = record['row']
                    row_data['错误原因'] = record['error']
                    export_data.append(row_data)
                
                # 创建DataFrame并导出
                df = pd.DataFrame(export_data)
                df.to_excel(file_path, index=False, engine='openpyxl')
                
                QMessageBox.information(self, "导出成功", f"失败记录已导出到：\n{file_path}")
                
            except Exception as e:
                QMessageBox.critical(self, "导出失败", f"导出失败记录时出错：{str(e)}")
    
    def clear_log(self):
        """清空日志"""
        self.log_text.clear()
