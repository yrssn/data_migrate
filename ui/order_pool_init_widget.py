"""
导入平台注册部Excel表初始化订单池功能界面
"""
import os
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog,
                             QTableWidget, QTableWidgetItem, QHeaderView, QScrollArea)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from database import DatabaseManager
import pandas as pd
import pymysql


class OrderPoolInitWorker(QThread):
    """订单池初始化工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    
    def __init__(self, datasource, db_manager, excel_file, column_mapping):
        super().__init__()
        self.datasource = datasource
        self.db_manager = db_manager
        self.excel_file = excel_file
        self.column_mapping = column_mapping
        self.results = {
            'total_rows': 0,
            'success_count': 0,
            'failed_count': 0,
            'failed_records': [],
            'completed_data': []
        }
        self.current_date = datetime.now().strftime('%Y%m%d')
        self.daily_counter = 1
    
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
            
            self.log_message.emit("开始连接数据库...")
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
            self.progress.emit(15)
            
            self.init_daily_counter(cursor)
            
            for index, row in df.iterrows():
                try:
                    order_id, primary_key_id = self.process_row(cursor, index + 1, row)
                    self.results['success_count'] += 1
                    
                    completed_row = row.to_dict()
                    completed_row['生成的订单ID'] = order_id
                    completed_row['数据库主键ID'] = primary_key_id
                    completed_row['处理状态'] = '成功'
                    completed_row['处理时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    self.results['completed_data'].append(completed_row)
                    
                    progress = 15 + int((index + 1) / len(df) * 80)
                    self.progress.emit(progress)
                    
                    if (index + 1) % 100 == 0:
                        connection.commit()
                        self.log_message.emit(f"已处理 {index + 1} 行，提交事务")
                        
                except Exception as e:
                    self.results['failed_count'] += 1
                    error_msg = f"第{index + 1}行处理失败: {str(e)}"
                    self.log_message.emit(error_msg)
                    
                    failed_row = row.to_dict()
                    failed_row['生成的订单ID'] = ''
                    failed_row['数据库主键ID'] = ''
                    failed_row['处理状态'] = '失败'
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
    
    def init_daily_counter(self, cursor):
        """初始化当日订单计数器"""
        try:
            today_prefix = f"OD{self.current_date}"
            cursor.execute("""
                SELECT order_id FROM ba_order 
                WHERE order_id LIKE %s 
                ORDER BY order_id DESC 
                LIMIT 1
            """, (f"{today_prefix}%",))
            
            result = cursor.fetchone()
            if result:
                last_order_id = result[0]
                last_number = int(last_order_id[-3:])
                self.daily_counter = last_number + 1
                self.log_message.emit(f"当日已有订单，从 {self.daily_counter:03d} 开始计数")
            else:
                self.daily_counter = 1
                self.log_message.emit(f"当日首次导入，从 001 开始计数")
                
        except Exception as e:
            self.log_message.emit(f"初始化订单计数器失败: {str(e)}，使用默认值001")
            self.daily_counter = 1
    
    def generate_order_id(self):
        """生成订单号"""
        if self.daily_counter > 999:
            current_datetime = datetime.now()
            next_day = current_datetime.replace(day=current_datetime.day + 1)
            self.current_date = next_day.strftime('%Y%m%d')
            self.daily_counter = 1
            self.log_message.emit(f"当日订单号已满，切换到下一天: {self.current_date}")
        
        order_id = f"OD{self.current_date}{self.daily_counter:03d}"
        self.daily_counter += 1
        return order_id
    
    def get_cell_value(self, row, field_name):
        """根据列映射获取单元格值"""
        col_index = self.column_mapping.get(field_name)
        if col_index is None or col_index < 0:
            return None
        if col_index >= len(row):
            return None
        value = row.iloc[col_index]
        if pd.isna(value) or str(value).strip() == '' or str(value).strip().lower() == 'nan':
            return None
        return str(value).strip()
    
    def process_row(self, cursor, row_num, row):
        """处理单行数据"""
        try:
            platform_name = self.get_cell_value(row, 'platform_name')
            company_country = self.get_cell_value(row, 'company_country')
            attribute_value = self.get_cell_value(row, 'attribute')
            territory_abbr = self.get_cell_value(row, 'territory_abbr')
            status_value = self.get_cell_value(row, 'status')
            
            if not platform_name:
                raise Exception("平台名称为空")
            if not attribute_value:
                raise Exception("属性为空")
            if not territory_abbr:
                raise Exception("属地简称为空")
            
            # 根据状态值决定 regist_status 和 status
            # 完成 → 4, 其他 → 2
            if status_value and status_value == "完成":
                regist_status = 4
                order_status = 4
            else:
                regist_status = 2
                order_status = 2
            
            country_info = f", 企业国别={company_country}" if company_country else ""
            status_info = f", 状态={status_value}→regist_status={regist_status}"
            self.log_message.emit(f"第{row_num}行: 处理平台={platform_name}, 属性={attribute_value}, 属地={territory_abbr}{country_info}{status_info}")
            
            platform_info = self.get_platform_info(cursor, platform_name)
            if not platform_info:
                raise Exception(f"未找到平台: {platform_name}")
            
            platform_id, platform_type = platform_info
            
            shudi_id = self.get_shudi_id(cursor, territory_abbr)
            if not shudi_id:
                raise Exception(f"未找到属地: {territory_abbr}")
            
            attribute_enum = self.convert_attribute(attribute_value)
            if not attribute_enum:
                raise Exception(f"无效的属性值: {attribute_value}")
            
            order_id = self.generate_order_id()
            
            primary_key_id = self.insert_order(cursor, order_id, platform_id, platform_type, shudi_id, 
                                              attribute_enum, attribute_value, company_country, row_num,
                                              regist_status, order_status)
            
            self.log_message.emit(f"第{row_num}行: 成功创建订单 {order_id}，主键ID: {primary_key_id}")
            
            return order_id, primary_key_id
            
        except Exception as e:
            raise e
    
    def get_platform_info(self, cursor, platform_name):
        """根据平台名称获取平台信息"""
        cursor.execute("""
            SELECT id, platform_type FROM ba_platform 
            WHERE platform = %s AND status = 1 AND (delete_time IS NULL OR delete_time = 0)
        """, (platform_name,))
        result = cursor.fetchone()
        return result if result else None
    
    def get_shudi_id(self, cursor, territory_abbr):
        """根据属地简称获取shudi_id"""
        cursor.execute("""
            SELECT id FROM ba_shudi 
            WHERE territory_abbreviation = %s AND status = 1
        """, (territory_abbr,))
        result = cursor.fetchone()
        return result[0] if result else None
    
    def convert_attribute(self, attribute_value):
        """转换属性值到枚举"""
        attribute_mapping = {
            '个人': 'opt0',
            '个人事业主': 'opt1', 
            '企业法人': 'opt2'
        }
        return attribute_mapping.get(attribute_value)
    
    def insert_order(self, cursor, order_id, platform_id, platform_type, shudi_id, 
                    attribute_enum, attribute_value, company_country, row_num,
                    regist_status, order_status):
        """插入订单记录"""
        current_time = int(datetime.now().timestamp())
        
        # done_number: 完成状态为1，其他为0
        done_number = 1 if order_status == 4 else 0
        
        insert_data = {
            'order_id': order_id,
            'shop_type': platform_type,
            'regist_department': 'opt0',
            'platform_id': platform_id,
            'shudi_id': shudi_id,
            'attribute': attribute_enum,
            'order_number': 1,
            'rejected_number': 0,
            'done_number': done_number,
            'status': order_status,
            'demand': '系统初始化数据',
            'feedback': '',
            'admin_id': 1,
            'admin_dept_id': 1,
            'create_time': current_time,
            'update_time': current_time,
            'regesit_status': regist_status,
            'need_data_type': 1,
            'rlb_status': 1
        }
        
        if company_country:
            insert_data['company_country'] = company_country
            self.log_message.emit(f"第{row_num}行: 设置企业国别为 {company_country}")
        else:
            if attribute_value == '企业法人':
                self.log_message.emit(f"第{row_num}行: 企业法人但企业国别为空")
        
        fields = list(insert_data.keys())
        placeholders = ['%s'] * len(fields)
        values = list(insert_data.values())
        
        insert_sql = f"""
            INSERT INTO ba_order ({', '.join(fields)}) 
            VALUES ({', '.join(placeholders)})
        """
        
        cursor.execute(insert_sql, values)
        
        primary_key_id = cursor.lastrowid
        
        self.log_message.emit(f"第{row_num}行: 插入订单数据 - 平台ID:{platform_id}, 属地ID:{shudi_id}, 属性:{attribute_enum}")
        
        return primary_key_id


class OrderPoolInitWidget(QWidget):
    """导入平台注册部Excel表初始化订单池功能组件"""
    
    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.excel_file = None
        self.excel_data = None
        self.failed_records = []
        self.completed_data = []
        self.column_mapping = {}
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
        
        # 平台名称列
        self.platform_col_combo = QComboBox()
        self.platform_col_combo.setEnabled(False)
        self.platform_col_combo.currentIndexChanged.connect(self.update_start_button_state)
        mapping_layout.addRow("平台名称列 (用于查找ba_platform):", self.platform_col_combo)
        
        # 企业国别列
        self.country_col_combo = QComboBox()
        self.country_col_combo.setEnabled(False)
        mapping_layout.addRow("企业国别列 (企业法人时使用，可选):", self.country_col_combo)
        
        # 属性列
        self.attribute_col_combo = QComboBox()
        self.attribute_col_combo.setEnabled(False)
        self.attribute_col_combo.currentIndexChanged.connect(self.update_start_button_state)
        mapping_layout.addRow("属性列 (个人/个人事业主/企业法人):", self.attribute_col_combo)
        
        # 属地简称列
        self.territory_col_combo = QComboBox()
        self.territory_col_combo.setEnabled(False)
        self.territory_col_combo.currentIndexChanged.connect(self.update_start_button_state)
        mapping_layout.addRow("属地简称列 (用于查找ba_shudi):", self.territory_col_combo)
        
        # 状态列（用于决定regist_status和status）
        self.status_col_combo = QComboBox()
        self.status_col_combo.setEnabled(False)
        self.status_col_combo.currentIndexChanged.connect(self.update_start_button_state)
        mapping_layout.addRow("状态列 (完成→4, 其他→2):", self.status_col_combo)
        
        mapping_group.setLayout(mapping_layout)
        layout.addWidget(mapping_group)
        
        # 数据预览区域
        preview_group = QGroupBox("数据预览（前5行）")
        preview_layout = QVBoxLayout()
        
        self.preview_table = QTableWidget()
        self.preview_table.setMaximumHeight(150)
        preview_layout.addWidget(self.preview_table)
        
        preview_group.setLayout(preview_layout)
        layout.addWidget(preview_group)
        
        # 功能说明区域
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        
        info_label = QLabel("""
<b>导入平台注册部Excel表初始化订单池说明:</b><br>
<b>订单号生成规则:</b> OD + 日期(YYYYMMDD) + 3位序号，示例: OD20251001001<br>
<b>数据映射规则:</b><br>
• shop_type: 从ba_platform.platform_type获取 | regist_department: 固定为opt0(平台注册部)<br>
• platform_id: 从ba_platform.id获取 | shudi_id: 从ba_shudi.id获取<br>
• attribute: 个人→opt0, 个人事业主→opt1, 企业法人→opt2 | company_country: 企业法人时从对应列获取<br>
<b>状态逻辑:</b> 状态列值为「完成」→ status=4, regesit_status=4, done_number=1; 其他 → status=2, regesit_status=2, done_number=0<br>
<b>固定字段值:</b> order_number: 1, admin_id: 1, admin_dept_id: 1
        """)
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # 操作按钮区域
        button_layout = QHBoxLayout()
        
        self.start_btn = QPushButton("开始导入")
        self.start_btn.clicked.connect(self.start_import)
        self.start_btn.setEnabled(False)
        button_layout.addWidget(self.start_btn)
        
        self.export_completed_btn = QPushButton("导出完整结果")
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
        
        # 生成列选项
        col_options = []
        for i, col in enumerate(columns):
            col_letter = chr(65 + i) if i < 26 else f"A{chr(65 + i - 26)}" if i < 52 else f"Col{i+1}"
            col_options.append((f"{col_letter}列: {col}", i))
        
        # 平台名称列
        self.platform_col_combo.clear()
        self.platform_col_combo.addItem("请选择列", -1)
        for text, idx in col_options:
            self.platform_col_combo.addItem(text, idx)
        self.platform_col_combo.setEnabled(True)
        
        # 企业国别列（可选）
        self.country_col_combo.clear()
        self.country_col_combo.addItem("不使用（留空）", -1)
        for text, idx in col_options:
            self.country_col_combo.addItem(text, idx)
        self.country_col_combo.setEnabled(True)
        
        # 属性列
        self.attribute_col_combo.clear()
        self.attribute_col_combo.addItem("请选择列", -1)
        for text, idx in col_options:
            self.attribute_col_combo.addItem(text, idx)
        self.attribute_col_combo.setEnabled(True)
        
        # 属地简称列
        self.territory_col_combo.clear()
        self.territory_col_combo.addItem("请选择列", -1)
        for text, idx in col_options:
            self.territory_col_combo.addItem(text, idx)
        self.territory_col_combo.setEnabled(True)
        
        # 状态列
        self.status_col_combo.clear()
        self.status_col_combo.addItem("请选择列", -1)
        for text, idx in col_options:
            self.status_col_combo.addItem(text, idx)
        self.status_col_combo.setEnabled(True)
        
        # 不自动匹配，让用户手动选择
        # self.auto_match_columns(columns)
    
    def auto_match_columns(self, columns):
        """尝试自动匹配常见列名"""
        for i, col in enumerate(columns):
            col_lower = str(col).lower()
            
            if '平台' in col_lower and '名' in col_lower:
                self.platform_col_combo.setCurrentIndex(i + 1)
            elif '国别' in col_lower or '国家' in col_lower:
                self.country_col_combo.setCurrentIndex(i + 1)
            elif '属性' in col_lower:
                self.attribute_col_combo.setCurrentIndex(i + 1)
            elif '属地' in col_lower or '简称' in col_lower:
                self.territory_col_combo.setCurrentIndex(i + 1)
            elif '状态' in col_lower or '进度' in col_lower:
                self.status_col_combo.setCurrentIndex(i + 1)
    
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
        datasource_selected = self.datasource_combo.currentData() is not None
        file_selected = self.excel_file is not None
        
        platform_selected = self.platform_col_combo.currentData() is not None and self.platform_col_combo.currentData() >= 0
        attribute_selected = self.attribute_col_combo.currentData() is not None and self.attribute_col_combo.currentData() >= 0
        territory_selected = self.territory_col_combo.currentData() is not None and self.territory_col_combo.currentData() >= 0
        status_selected = self.status_col_combo.currentData() is not None and self.status_col_combo.currentData() >= 0
        
        self.start_btn.setEnabled(
            datasource_selected and file_selected and 
            platform_selected and attribute_selected and territory_selected and status_selected
        )
    
    def get_column_mapping(self):
        """获取列映射配置"""
        return {
            'platform_name': self.platform_col_combo.currentData(),
            'company_country': self.country_col_combo.currentData() if self.country_col_combo.currentData() >= 0 else None,
            'attribute': self.attribute_col_combo.currentData(),
            'territory_abbr': self.territory_col_combo.currentData(),
            'status': self.status_col_combo.currentData()
        }
    
    def start_import(self):
        """开始导入"""
        datasource = self.datasource_combo.currentData()
        if not datasource or not self.excel_file:
            QMessageBox.warning(self, "警告", "请选择数据源和Excel文件！")
            return
        
        column_mapping = self.get_column_mapping()
        
        # 构建映射说明
        mapping_info = []
        mapping_info.append(f"平台名称: {self.platform_col_combo.currentText()}")
        mapping_info.append(f"企业国别: {self.country_col_combo.currentText()}")
        mapping_info.append(f"属性: {self.attribute_col_combo.currentText()}")
        mapping_info.append(f"属地简称: {self.territory_col_combo.currentText()}")
        mapping_info.append(f"状态: {self.status_col_combo.currentText()}")
        
        reply = QMessageBox.question(
            self,
            "确认导入",
            f"确定要导入订单数据吗？\n\n"
            f"数据库: {datasource.name}\n"
            f"Excel文件: {os.path.basename(self.excel_file)}\n\n"
            f"列映射配置:\n" + "\n".join(mapping_info) + "\n\n"
            f"订单号格式: OD{datetime.now().strftime('%Y%m%d')}XXX\n"
            f"操作不可撤销，请确认配置正确！",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        self.start_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在导入中...")
        self.log_text.clear()
        self.failed_records = []
        self.completed_data = []
        
        self.worker = OrderPoolInitWorker(datasource, self.db_manager, self.excel_file, column_mapping)
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log_message.connect(self.append_log)
        self.worker.finished.connect(self.on_import_finished)
        self.worker.error.connect(self.on_import_error)
        self.worker.start()
    
    def append_log(self, message):
        """添加日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
    
    def on_import_finished(self, results):
        """导入完成"""
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.failed_records = results['failed_records']
        self.completed_data = results['completed_data']
        
        result_text = (
            f"导入完成！\n"
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
        
        QMessageBox.information(self, "导入完成", result_text)
    
    def on_import_error(self, error_msg):
        """导入错误"""
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.result_label.setText("导入失败！")
        self.result_label.setStyleSheet("color: red;")
        
        QMessageBox.critical(self, "导入错误", f"导入失败：{error_msg}")
    
    def export_completed_data(self):
        """导出完整结果"""
        if not self.completed_data:
            QMessageBox.information(self, "提示", "没有完整结果可导出")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存完整结果",
            f"order_import_completed_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel文件 (*.xlsx)"
        )
        
        if file_path:
            try:
                df = pd.DataFrame(self.completed_data)
                df.to_excel(file_path, index=False, engine='openpyxl')
                
                QMessageBox.information(self, "导出成功", f"完整结果已导出到：\n{file_path}")
                
            except Exception as e:
                QMessageBox.critical(self, "导出失败", f"导出完整结果时出错：{str(e)}")
    
    def export_failed_records(self):
        """导出失败记录"""
        if not self.failed_records:
            QMessageBox.information(self, "提示", "没有失败记录可导出")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存失败记录",
            f"order_import_failed_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
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
