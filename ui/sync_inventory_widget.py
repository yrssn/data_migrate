"""
同步库存账号功能界面
"""
import os
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog, QLineEdit)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from database import DatabaseManager
import pandas as pd
import pymysql


class SyncInventoryWorker(QThread):
    """同步库存账号工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    
    def __init__(self, datasource, excel_file, ptzcb_column, legal_column, bank_column, status_column, register_department):
        super().__init__()
        self.datasource = datasource
        self.excel_file = excel_file
        self.ptzcb_column = ptzcb_column  # ba_ptzcb_register查找列
        self.legal_column = legal_column  # 法人姓名列
        self.bank_column = bank_column    # 银行名称列
        self.status_column = status_column # 状态列
        self.register_department = register_department
        self.results = {
            'total_rows': 0,
            'success_count': 0,
            'failed_count': 0,
            'failed_records': [],
            'completed_data': []
        }
        # 序号计数器，按平台-属地分组
        self.inventory_counters = {}
    
    def run(self):
        try:
            self.log_message.emit("开始读取Excel文件...")
            self.progress.emit(5)
            
            # 读取Excel文件
            df = pd.read_excel(self.excel_file, engine='openpyxl')
            self.log_message.emit(f"成功读取Excel文件，共 {len(df)} 行数据")
            self.results['total_rows'] = len(df)
            self.progress.emit(10)
            
            # 连接数据库
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
            
            # 处理每一行数据
            for index, row in df.iterrows():
                try:
                    inventory_num, account_mag_id = self.process_row(cursor, index + 1, row)
                    self.results['success_count'] += 1
                    
                    # 记录成功数据
                    completed_row = row.to_dict()
                    completed_row['生成的库存编号'] = inventory_num
                    completed_row['库存账号主键ID'] = account_mag_id
                    completed_row['处理状态'] = '成功'
                    completed_row['处理时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    self.results['completed_data'].append(completed_row)
                    
                    # 更新进度
                    progress = 15 + int((index + 1) / len(df) * 80)
                    self.progress.emit(progress)
                    
                    # 每50行提交一次
                    if (index + 1) % 50 == 0:
                        connection.commit()
                        self.log_message.emit(f"已处理 {index + 1} 行，提交事务")
                        
                except Exception as e:
                    self.results['failed_count'] += 1
                    error_msg = f"第{index + 1}行处理失败: {str(e)}"
                    self.log_message.emit(error_msg)
                    
                    # 记录失败数据
                    failed_row = row.to_dict()
                    failed_row['生成的库存编号'] = ''
                    failed_row['库存账号主键ID'] = ''
                    failed_row['处理状态'] = '失败'
                    failed_row['失败原因'] = str(e)
                    failed_row['处理时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    self.results['completed_data'].append(failed_row)
                    
                    self.results['failed_records'].append({
                        'row': index + 1,
                        'error': str(e),
                        'data': row.to_dict()
                    })
            
            # 最终提交
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
            # 获取Excel列数据
            ptzcb_value = str(row[self.ptzcb_column]).strip() if self.ptzcb_column in row.index else ""
            legal_name = str(row[self.legal_column]).strip() if self.legal_column in row.index else ""
            bank_name = str(row[self.bank_column]).strip() if self.bank_column in row.index else ""
            status_value = str(row[self.status_column]).strip() if self.status_column in row.index else ""
            
            if not ptzcb_value or ptzcb_value == 'nan':
                raise Exception(f"PTZCB查找列为空")
            
            self.log_message.emit(f"第{row_num}行: 处理数据 - PTZCB值={ptzcb_value}, 法人={legal_name}, 银行={bank_name}, 状态={status_value}")
            
            # 第一步：根据ptzcb_value查找ba_ptzcb_register记录
            ptzcb_record = self.find_ptzcb_record(cursor, ptzcb_value)
            if not ptzcb_record:
                raise Exception(f"未找到对应的ba_ptzcb_register记录")
            
            # 第二步：获取平台和属地信息
            platform_info = self.get_platform_info(cursor, ptzcb_record['platform_id'])
            shudi_info = self.get_shudi_info(cursor, ptzcb_record['shudi_id'])
            
            # 第三步：生成库存编号
            inventory_num = self.generate_inventory_num(platform_info['code'], shudi_info['territory_abbreviation'])
            
            # 第四步：查找银行信息
            bank_id = self.find_bank_id(cursor, legal_name, bank_name, row_num)
            
            # 第五步：获取货币ID
            currency_id = self.get_currency_id(cursor, bank_id) if bank_id else None
            
            # 第六步：转换状态
            account_status = self.convert_status(status_value)
            
            # 第七步：插入ba_account_mag记录
            account_mag_id = self.insert_account_mag(cursor, {
                'inventory_num': inventory_num,
                'shudi_id': ptzcb_record['shudi_id'],
                'platform_id': ptzcb_record['platform_id'],
                'register_id': ptzcb_record['id'],
                'account_status': account_status,
                'shopindex_id': ptzcb_record['shopindex_id'],
                'bank_id': bank_id,
                'order_id': ptzcb_record['order_id'],
                'currency_id': currency_id,
                'legal_id': ptzcb_record['customer_id'],
                'information_id': ptzcb_record['information_id'],
                'information_part2_id': ptzcb_record['information_part2_id'],
                'register_department': self.register_department,
                'source_reg_id': 1
            })
            
            self.log_message.emit(f"第{row_num}行: 成功创建库存账号 {inventory_num} (ID: {account_mag_id})")
            return inventory_num, account_mag_id
            
        except Exception as e:
            raise e
    
    def find_ptzcb_record(self, cursor, ptzcb_value):
        """查找ba_ptzcb_register记录"""
        cursor.execute("""
            SELECT id, shudi_id, platform_id, shopindex_id, order_id, customer_id, 
                   information_id, information_part2_id
            FROM ba_ptzcb_register 
            WHERE project_id = %s OR id = %s
        """, (ptzcb_value, ptzcb_value))
        return cursor.fetchone()
    
    def get_platform_info(self, cursor, platform_id):
        """获取平台信息"""
        cursor.execute("SELECT code FROM ba_platform WHERE id = %s", (platform_id,))
        result = cursor.fetchone()
        if not result:
            raise Exception(f"未找到平台ID {platform_id} 的信息")
        return result
    
    def get_shudi_info(self, cursor, shudi_id):
        """获取属地信息"""
        cursor.execute("SELECT territory_abbreviation FROM ba_shudi WHERE id = %s", (shudi_id,))
        result = cursor.fetchone()
        if not result:
            raise Exception(f"未找到属地ID {shudi_id} 的信息")
        return result
    
    def generate_inventory_num(self, platform_code, territory_abbr):
        """生成库存编号: 平台代码-属地简称-年月-序号"""
        today = datetime.now()
        year_month = today.strftime('%y%m')  # 2602
        
        # 生成分组键
        group_key = f"{platform_code}-{territory_abbr}-{year_month}"
        
        # 获取当前序号
        if group_key not in self.inventory_counters:
            self.inventory_counters[group_key] = 0
        
        self.inventory_counters[group_key] += 1
        sequence = str(self.inventory_counters[group_key]).zfill(4)  # 0001
        
        return f"{platform_code}-{territory_abbr}-{year_month}-{sequence}"
    
    def find_bank_id(self, cursor, legal_name, bank_name, row_num):
        """查找银行ID（可选，找不到不报错）"""
        try:
            # 如果银行名称为空，直接返回None（不报错）
            if not bank_name or bank_name.strip() == '' or bank_name == 'nan':
                return None
            
            # 第一步：根据法人姓名查找ba_rlb_customer
            cursor.execute("SELECT id FROM ba_rlb_customer WHERE legal_name = %s", (legal_name,))
            customer_record = cursor.fetchone()
            if not customer_record:
                return None
            
            customer_id = customer_record['id']
            
            # 第二步：根据银行名称查找ba_zhb_bank_name
            cursor.execute("SELECT id FROM ba_zhb_bank_name WHERE bank_name = %s", (bank_name,))
            bank_name_record = cursor.fetchone()
            if not bank_name_record:
                return None
            
            bank_name_id = bank_name_record['id']
            
            # 第三步：查找ba_zhb_bank记录
            cursor.execute("""
                SELECT id FROM ba_zhb_bank 
                WHERE rlb_customer_id = %s AND bank_name_id = %s
                ORDER BY update_time DESC
                LIMIT 1
            """, (customer_id, bank_name_id))
            
            bank_record = cursor.fetchone()
            return bank_record['id'] if bank_record else None
                
        except Exception as e:
            return None
    
    def get_currency_id(self, cursor, bank_id):
        """从银行记录获取货币ID"""
        if not bank_id:
            return None
        cursor.execute("SELECT currency_id FROM ba_zhb_bank WHERE id = %s", (bank_id,))
        result = cursor.fetchone()
        return result['currency_id'] if result else None
    
    def convert_status(self, status_value):
        """转换状态值"""
        status_map = {
            '正常': 4,      # 已售
            '暂停': 4,      # 已售
            '闭店申请中': 4, # 已售
            '退店申请中': 4, # 已售
            '闭店回款中': 4, # 已售
            '退店回款': 4,   # 已售
            '已闭店': 5,    # 封号
            '已退店': 3     # 可售
        }
        return status_map.get(status_value, 4)  # 默认已售
    
    def insert_account_mag(self, cursor, data):
        """插入ba_account_mag记录，返回插入记录的ID"""
        current_time = int(datetime.now().timestamp())
        
        cursor.execute("""
            INSERT INTO ba_account_mag (
                inventory_num, shudi_id, platform_id, register_id, account_status,
                shopindex_id, bank_id, order_id, currency_id, legal_id,
                information_id, information_part2_id, register_department, source_reg_id,
                remark, create_time, update_time
            ) VALUES (
                %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
            )
        """, (
            data['inventory_num'], data['shudi_id'], data['platform_id'], data['register_id'],
            data['account_status'], data['shopindex_id'], data['bank_id'], data['order_id'],
            data['currency_id'], data['legal_id'], data['information_id'], data['information_part2_id'],
            data['register_department'], data['source_reg_id'], '系统脚本补全', current_time, current_time
        ))
        
        # 返回插入记录的ID
        return cursor.lastrowid


class SyncInventoryWidget(QWidget):
    """同步库存账号功能组件"""
    
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
        
        # 列选择区域
        column_group = QGroupBox("列选择")
        column_layout = QFormLayout()
        
        self.ptzcb_column_combo = QComboBox()
        self.ptzcb_column_combo.setEnabled(False)
        column_layout.addRow("PTZCB查找列:", self.ptzcb_column_combo)
        
        self.legal_column_combo = QComboBox()
        self.legal_column_combo.setEnabled(False)
        column_layout.addRow("法人姓名列:", self.legal_column_combo)
        
        self.bank_column_combo = QComboBox()
        self.bank_column_combo.setEnabled(False)
        column_layout.addRow("银行名称列:", self.bank_column_combo)
        
        self.status_column_combo = QComboBox()
        self.status_column_combo.setEnabled(False)
        column_layout.addRow("状态列:", self.status_column_combo)
        
        column_group.setLayout(column_layout)
        layout.addWidget(column_group)
        
        # 参数设置区域
        param_group = QGroupBox("参数设置")
        param_layout = QFormLayout()
        
        self.register_department_edit = QLineEdit()
        self.register_department_edit.setPlaceholderText("请输入注册部门名称")
        self.register_department_edit.textChanged.connect(self.update_start_button_state)
        param_layout.addRow("注册部门:", self.register_department_edit)
        
        param_group.setLayout(param_layout)
        layout.addWidget(param_group)
        
        # 功能说明区域
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        
        info_label = QLabel("""
<b>同步库存账号功能说明:</b><br>
<b>使用步骤:</b><br>
1. 选择Excel文件后，系统会自动读取表头<br>
2. 在"列选择"区域选择对应的列：<br>
   • <b>PTZCB查找列</b>: 用于查找ba_ptzcb_register表的project_id或id<br>
   • <b>法人姓名列</b>: 用于查找ba_rlb_customer表<br>
   • <b>银行名称列</b>: 用于查找ba_zhb_bank_name表<br>
   • <b>状态列</b>: 用于转换account_status<br>
3. 设置注册部门名称<br><br>
<b>库存编号生成规则:</b><br>
• 格式: 平台代码-属地简称-年月-序号<br>
• 示例: RK-JP-2602-0001<br><br>
<b>状态转换规则:</b><br>
• 正常/暂停/闭店申请中/退店申请中/闭店回款中/退店回款 → 已售(4)<br>
• 已闭店 → 封号(5)<br>
• 已退店 → 可售(3)<br><br>
<b>数据关联:</b><br>
• 从ba_ptzcb_register获取: shudi_id, platform_id, shopindex_id等<br>
• 从ba_rlb_customer和ba_zhb_bank获取: bank_id, currency_id<br>
• source_reg_id固定为1
        """)
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # 操作按钮区域
        button_layout = QHBoxLayout()
        
        self.start_btn = QPushButton("开始同步")
        self.start_btn.clicked.connect(self.start_process)
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
    
    def select_excel_file(self):
        """选择Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Excel文件",
            "",
            "Excel文件 (*.xlsx *.xls)"
        )
        
        if file_path:
            try:
                # 读取Excel表头
                df = pd.read_excel(file_path, engine='openpyxl', nrows=0)
                self.excel_columns = list(df.columns)
                
                # 更新UI
                self.excel_file = file_path
                self.file_label.setText(os.path.basename(file_path))
                self.file_label.setStyleSheet("color: green;")
                
                # 填充列选择下拉框
                for combo in [self.ptzcb_column_combo, self.legal_column_combo, 
                             self.bank_column_combo, self.status_column_combo]:
                    combo.clear()
                    combo.addItems(self.excel_columns)
                    combo.setEnabled(True)
                
                self.log_text.append(f"成功读取Excel表头，共 {len(self.excel_columns)} 列")
                self.log_text.append(f"列名: {', '.join(self.excel_columns)}")
                
                self.update_start_button_state()
                
            except Exception as e:
                QMessageBox.critical(self, "错误", f"读取Excel文件失败: {str(e)}")
                self.excel_file = None
                self.file_label.setText("未选择文件")
                self.file_label.setStyleSheet("color: gray;")
    
    def update_start_button_state(self):
        """更新开始按钮状态"""
        datasource_selected = self.datasource_combo.currentData() is not None
        file_selected = self.excel_file is not None
        department_filled = bool(self.register_department_edit.text().strip())
        self.start_btn.setEnabled(datasource_selected and file_selected and department_filled)
    
    def start_process(self):
        """开始同步"""
        datasource = self.datasource_combo.currentData()
        if not datasource or not self.excel_file:
            QMessageBox.warning(self, "警告", "请选择数据源和Excel文件！")
            return
        
        # 获取选择的列和参数
        ptzcb_column = self.ptzcb_column_combo.currentText()
        legal_column = self.legal_column_combo.currentText()
        bank_column = self.bank_column_combo.currentText()
        status_column = self.status_column_combo.currentText()
        register_department = self.register_department_edit.text().strip()
        
        if not all([ptzcb_column, legal_column, bank_column, status_column, register_department]):
            QMessageBox.warning(self, "警告", "请完整填写所有选项！")
            return
        
        # 确认对话框
        reply = QMessageBox.question(
            self,
            "确认同步",
            f"确定要同步库存账号吗？\n\n"
            f"数据库: {datasource.name}\n"
            f"Excel文件: {os.path.basename(self.excel_file)}\n"
            f"PTZCB查找列: {ptzcb_column}\n"
            f"法人姓名列: {legal_column}\n"
            f"银行名称列: {bank_column}\n"
            f"状态列: {status_column}\n"
            f"注册部门: {register_department}\n\n"
            f"操作不可撤销，请确认数据正确！",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        # 禁用按钮，显示进度条
        self.start_btn.setEnabled(False)
        self.export_completed_btn.setEnabled(False)
        self.export_failed_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在同步中...")
        self.log_text.clear()
        self.completed_data = []
        self.failed_records = []
        
        # 启动工作线程
        self.worker = SyncInventoryWorker(
            datasource, self.excel_file, ptzcb_column, legal_column, 
            bank_column, status_column, register_department
        )
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log_message.connect(self.append_log)
        self.worker.finished.connect(self.on_process_finished)
        self.worker.error.connect(self.on_process_error)
        self.worker.start()
    
    def append_log(self, message):
        """添加日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
    
    def on_process_finished(self, results):
        """处理完成"""
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        # 保存结果数据
        self.completed_data = results['completed_data']
        self.failed_records = results['failed_records']
        
        success_count = results['success_count']
        failed_count = results['failed_count']
        total_count = results['total_rows']
        
        result_text = f"同步完成！成功: {success_count}, 失败: {failed_count}, 总计: {total_count}"
        self.result_label.setText(result_text)
        self.result_label.setStyleSheet("color: green;" if failed_count == 0 else "color: orange;")
        
        # 启用导出按钮
        self.export_completed_btn.setEnabled(len(self.completed_data) > 0)
        self.export_failed_btn.setEnabled(len(self.failed_records) > 0)
        
        self.log_text.append(f"\n=== 同步完成 ===")
        self.log_text.append(result_text)
        
        if failed_count > 0:
            self.log_text.append(f"\n失败记录:")
            for record in results['failed_records'][:10]:  # 只显示前10个失败记录
                self.log_text.append(f"第{record['row']}行: {record['error']}")
    
    def on_process_error(self, error_message):
        """处理错误"""
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.result_label.setText("同步失败！")
        self.result_label.setStyleSheet("color: red;")
        
        QMessageBox.critical(self, "错误", f"同步过程出错: {error_message}")
        self.log_text.append(f"错误: {error_message}")
    
    def clear_log(self):
        """清空日志"""
        self.log_text.clear()
    
    def export_completed_data(self):
        """导出完整结果"""
        if not self.completed_data:
            QMessageBox.warning(self, "警告", "没有可导出的完整结果数据！")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "导出完整结果",
            f"同步库存账号_完整结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel文件 (*.xlsx)"
        )
        
        if file_path:
            try:
                df = pd.DataFrame(self.completed_data)
                df.to_excel(file_path, index=False, engine='openpyxl')
                QMessageBox.information(self, "成功", f"完整结果已导出到: {file_path}")
                self.log_text.append(f"完整结果已导出到: {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出失败: {str(e)}")
    
    def export_failed_records(self):
        """导出失败记录"""
        if not self.failed_records:
            QMessageBox.warning(self, "警告", "没有可导出的失败记录！")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "导出失败记录",
            f"同步库存账号_失败记录_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel文件 (*.xlsx)"
        )
        
        if file_path:
            try:
                # 准备失败记录数据
                failed_data = []
                for record in self.failed_records:
                    failed_row = record['data'].copy()
                    failed_row['失败行号'] = record['row']
                    failed_row['失败原因'] = record['error']
                    failed_row['处理时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    failed_data.append(failed_row)
                
                df = pd.DataFrame(failed_data)
                df.to_excel(file_path, index=False, engine='openpyxl')
                QMessageBox.information(self, "成功", f"失败记录已导出到: {file_path}")
                self.log_text.append(f"失败记录已导出到: {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出失败: {str(e)}")
