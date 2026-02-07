"""
现有客户银行信息补全功能界面
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


class BankInfoSyncWorker(QThread):
    """银行信息同步工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)  # 返回统计结果
    error = pyqtSignal(str)
    
    def __init__(self, target_datasource, source_datasource, db_manager):
        super().__init__()
        self.target_datasource = target_datasource
        self.source_datasource = source_datasource
        self.db_manager = db_manager
        self.results = {
            'customer_total': 0,
            'customer_processed': 0,
            'customer_skipped': 0,
            'bank_cards_found': 0,
            'bank_cards_inserted': 0,
            'bank_cards_updated': 0,
            'bank_cards_skipped': 0,
            'failed_records': []
        }
    
    def run(self):
        try:
            self.log_message.emit("开始连接目标数据库...")
            self.progress.emit(5)
            
            # 连接目标数据库
            target_connection = pymysql.connect(
                host=self.target_datasource.host,
                port=self.target_datasource.port,
                user=self.target_datasource.username,
                password=self.target_datasource.password,
                database=self.target_datasource.database,
                charset=self.target_datasource.charset
            )
            target_cursor = target_connection.cursor()
            self.log_message.emit("目标数据库连接成功")
            
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
            self.progress.emit(10)
            
            # 获取目标数据库中的所有客户
            target_cursor.execute("""
                SELECT id, legal_name 
                FROM ba_rlb_customer 
                WHERE (delete_time IS NULL OR delete_time = 0)
            """)
            customers = target_cursor.fetchall()
            
            self.results['customer_total'] = len(customers)
            self.log_message.emit(f"找到 {len(customers)} 个客户记录")
            self.progress.emit(15)
            
            if len(customers) == 0:
                self.log_message.emit("没有找到客户记录")
                self.finished.emit(self.results)
                return
            
            # 处理每个客户
            for index, customer in enumerate(customers):
                try:
                    self.process_customer(target_cursor, source_cursor, customer)
                    
                    # 更新进度
                    progress = 15 + int((index + 1) / len(customers) * 80)
                    self.progress.emit(progress)
                    
                    # 每50个客户提交一次
                    if (index + 1) % 50 == 0:
                        target_connection.commit()
                        self.log_message.emit(f"已处理 {index + 1} 个客户，提交事务")
                        
                except Exception as e:
                    customer_id, legal_name = customer
                    error_msg = f"处理客户 {legal_name}(ID:{customer_id}) 失败: {str(e)}"
                    self.log_message.emit(error_msg)
                    self.results['failed_records'].append({
                        'customer_id': customer_id,
                        'legal_name': legal_name,
                        'error': str(e)
                    })
            
            # 最终提交
            target_connection.commit()
            
            target_cursor.close()
            target_connection.close()
            source_cursor.close()
            source_connection.close()
            
            self.progress.emit(100)
            self.finished.emit(self.results)
            
        except Exception as e:
            self.error.emit(str(e))
    
    def process_customer(self, target_cursor, source_cursor, customer):
        """处理单个客户的银行信息"""
        customer_id, legal_name = customer
        
        try:
            # 第一步：根据legal_name在源数据库中找到对应的法人记录
            source_cursor.execute("""
                SELECT id FROM ea_dy_legal 
                WHERE name = %s AND status = 1 AND (delete_time IS NULL OR delete_time = 0)
            """, (legal_name,))
            
            legal_result = source_cursor.fetchone()
            if not legal_result:
                self.results['customer_skipped'] += 1
                self.log_message.emit(f"客户 {legal_name}: 在源数据库中未找到对应法人记录")
                return
            
            source_legal_id = legal_result[0]
            
            # 第二步：根据legal_id获取所有银行卡信息
            source_cursor.execute("""
                SELECT lc.id, lc.currency_id, lc.name, lc.account, lc.balance, 
                       lc.bankcard_id, lc.bank_card_type, lc.remark, lc.status,
                       bc.name as bankcard_name, cur.code as currency_code
                FROM ea_dy_legal_cards lc
                LEFT JOIN ea_dy_bankcard bc ON lc.bankcard_id = bc.id
                LEFT JOIN ea_dy_currency cur ON lc.currency_id = cur.id
                WHERE lc.legal_id = %s AND (lc.delete_time IS NULL OR lc.delete_time = 0)
            """, (source_legal_id,))
            
            bank_cards = source_cursor.fetchall()
            if not bank_cards:
                self.results['customer_skipped'] += 1
                self.log_message.emit(f"客户 {legal_name}: 未找到银行卡信息")
                return
            
            self.results['customer_processed'] += 1
            self.results['bank_cards_found'] += len(bank_cards)
            self.log_message.emit(f"客户 {legal_name}: 找到 {len(bank_cards)} 张银行卡")
            
            # 第三步：处理每张银行卡
            for card in bank_cards:
                self.process_bank_card(target_cursor, customer_id, legal_name, card)
                
        except Exception as e:
            raise Exception(f"处理客户银行信息失败: {str(e)}")
    
    def process_bank_card(self, target_cursor, customer_id, legal_name, card):
        """处理单张银行卡"""
        # 先记录源数据库中的银行卡信息
        (card_id, currency_id, name, account, balance, bankcard_id, 
         bank_card_type, remark, source_status, bankcard_name, currency_code) = card
        
        # 创建基础记录信息（用于失败时记录）
        base_record = {
            'customer_id': customer_id,
            'legal_name': legal_name,
            'source_card_id': card_id,
            'bank_account': account,
            'bank_name': bankcard_name,
            'currency_code': currency_code,
            'bank_card_type': bank_card_type,
            'balance': balance,
            'source_remark': remark
        }
        
        try:
            
            # 获取目标数据库中的currency_id
            target_currency_id = None
            if currency_code:
                target_cursor.execute("""
                    SELECT id FROM ba_currency 
                    WHERE currency_en = %s AND status = 1 AND (delete_time IS NULL OR delete_time = 0)
                """, (currency_code,))
                currency_result = target_cursor.fetchone()
                if currency_result:
                    target_currency_id = currency_result[0]
                else:
                    self.log_message.emit(f"客户 {legal_name}: 货币 {currency_code} 在目标数据库中未找到，将使用空值")
            else:
                self.log_message.emit(f"客户 {legal_name}: 银行卡 {account} 缺少货币信息")
            
            # 获取目标数据库中的bank_name_id
            target_bank_name_id = None
            if bankcard_name:
                target_cursor.execute("""
                    SELECT id FROM ba_zhb_bank_name 
                    WHERE bank_name = %s AND status = 1 AND (delete_time IS NULL OR delete_time = 0)
                """, (bankcard_name,))
                bank_name_result = target_cursor.fetchone()
                if bank_name_result:
                    target_bank_name_id = bank_name_result[0]
                else:
                    error_msg = f"银行 {bankcard_name} 在目标数据库中未找到"
                    self.log_message.emit(f"客户 {legal_name}: {error_msg}，跳过银行卡 {account}")
                    self.results['bank_cards_skipped'] += 1
                    failed_record = base_record.copy()
                    failed_record['error'] = error_msg
                    self.results['failed_records'].append(failed_record)
                    return
            else:
                error_msg = "银行卡缺少银行名称信息"
                self.log_message.emit(f"客户 {legal_name}: {error_msg}，跳过银行卡 {account}")
                self.results['bank_cards_skipped'] += 1
                failed_record = base_record.copy()
                failed_record['error'] = error_msg
                self.results['failed_records'].append(failed_record)
                return
            
            # 转换bank_card_type (0,1,2 -> opt0,opt1,opt2)
            target_bank_card_type = 'opt2'  # 默认值
            if bank_card_type == 0:
                target_bank_card_type = 'opt2'
            elif bank_card_type == 1:
                target_bank_card_type = 'opt0'
            elif bank_card_type == 2:
                target_bank_card_type = 'opt1'
            
            # 检查是否已存在相同的银行卡
            check_conditions = []
            check_params = []
            
            # 必须匹配的条件
            check_conditions.append("rlb_customer_id = %s")
            check_params.append(customer_id)
            
            check_conditions.append("bank_id = %s")
            check_params.append(account)
            
            check_conditions.append("bank_card_type = %s")
            check_params.append(target_bank_card_type)
            
            # 可选条件（如果有值的话）
            if target_currency_id:
                check_conditions.append("currency_id = %s")
                check_params.append(target_currency_id)
            
            if target_bank_name_id:
                check_conditions.append("bank_name_id = %s")
                check_params.append(target_bank_name_id)
            
            check_sql = f"""
                SELECT id, bank_balance FROM ba_zhb_bank 
                WHERE {' AND '.join(check_conditions)} AND (delete_time IS NULL OR delete_time = 0)
            """
            
            target_cursor.execute(check_sql, check_params)
            existing_result = target_cursor.fetchone()
            
            current_time = int(datetime.now().timestamp())
            
            if existing_result:
                # 已存在，更新余额
                existing_id, current_balance = existing_result
                if balance and balance != current_balance:
                    target_cursor.execute("""
                        UPDATE ba_zhb_bank 
                        SET bank_balance = %s, update_time = %s, status = %s 
                        WHERE id = %s
                    """, (balance, current_time, source_status, existing_id))
                    
                    self.results['bank_cards_updated'] += 1
                    self.log_message.emit(f"客户 {legal_name}: 更新银行卡余额 {account} -> {balance}")
                else:
                    # 即使余额相同，也要确保状态同步
                    target_cursor.execute("""
                        UPDATE ba_zhb_bank 
                        SET status = %s, update_time = %s 
                        WHERE id = %s
                    """, (source_status, current_time, existing_id))
                    
                    self.results['bank_cards_skipped'] += 1
                    self.log_message.emit(f"客户 {legal_name}: 银行卡 {account} 已存在，同步状态")
            else:
                # 不存在，插入新记录
                # 检查必填字段
                if not target_bank_name_id:
                    error_msg = "银行卡缺少bank_name_id"
                    self.log_message.emit(f"客户 {legal_name}: {error_msg}，跳过银行卡 {account}")
                    self.results['bank_cards_skipped'] += 1
                    failed_record = base_record.copy()
                    failed_record['error'] = error_msg
                    self.results['failed_records'].append(failed_record)
                    return
                
                if not account:
                    error_msg = "银行卡缺少account信息"
                    self.log_message.emit(f"客户 {legal_name}: {error_msg}，跳过")
                    self.results['bank_cards_skipped'] += 1
                    failed_record = base_record.copy()
                    failed_record['error'] = error_msg
                    self.results['failed_records'].append(failed_record)
                    return
                
                insert_data = {
                    'bank_name_id': target_bank_name_id,
                    'bank_id': account,
                    'rlb_customer_id': customer_id,
                    'remark': '脚本自动补充',
                    'admin_id': 1,
                    'admin_dept_id': 1,
                    'bank_balance': balance if balance is not None else 0.00,
                    'currency_id': target_currency_id,
                    'bank_card_type': target_bank_card_type,
                    'create_time': current_time,
                    'update_time': current_time,
                    'status': source_status
                }
                
                # 构建插入SQL - 只插入非空值
                fields = []
                values = []
                placeholders = []
                
                for field, value in insert_data.items():
                    # 必填字段必须有值
                    if field in ['bank_name_id', 'bank_id', 'rlb_customer_id'] and value is None:
                        error_msg = f"银行卡缺少必填字段 {field}"
                        self.log_message.emit(f"客户 {legal_name}: {error_msg}，跳过银行卡 {account}")
                        self.results['bank_cards_skipped'] += 1
                        failed_record = base_record.copy()
                        failed_record['error'] = error_msg
                        self.results['failed_records'].append(failed_record)
                        return
                    
                    # 只插入非空值
                    if value is not None:
                        fields.append(field)
                        values.append(value)
                        placeholders.append('%s')
                
                insert_sql = f"""
                    INSERT INTO ba_zhb_bank ({', '.join(fields)}) 
                    VALUES ({', '.join(placeholders)})
                """
                
                target_cursor.execute(insert_sql, values)
                
                self.results['bank_cards_inserted'] += 1
                self.log_message.emit(f"客户 {legal_name}: 插入新银行卡 {account} (银行: {bankcard_name}), 余额: {balance}")
                
        except Exception as e:
            error_msg = f"处理银行卡时发生异常: {str(e)}"
            self.log_message.emit(f"客户 {legal_name}: {error_msg}")
            failed_record = base_record.copy()
            failed_record['error'] = error_msg
            self.results['failed_records'].append(failed_record)
            raise e


class BankInfoSyncWidget(QWidget):
    """现有客户银行信息补全功能组件"""
    
    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.failed_records = []
        self.init_ui()
        self.load_datasources()
    
    def init_ui(self):
        """初始化UI"""
        layout = QVBoxLayout()
        
        # 数据源选择区域
        datasource_group = QGroupBox("选择数据源")
        datasource_layout = QFormLayout()
        
        # 目标数据源
        self.target_datasource_combo = QComboBox()
        self.target_datasource_combo.currentTextChanged.connect(self.on_datasource_changed)
        datasource_layout.addRow("目标数据库:", self.target_datasource_combo)
        
        self.test_target_btn = QPushButton("测试目标连接")
        self.test_target_btn.clicked.connect(self.test_target_connection)
        self.test_target_btn.setEnabled(False)
        datasource_layout.addRow("", self.test_target_btn)
        
        # 源数据源
        self.source_datasource_combo = QComboBox()
        self.source_datasource_combo.currentTextChanged.connect(self.on_datasource_changed)
        datasource_layout.addRow("源数据库:", self.source_datasource_combo)
        
        self.test_source_btn = QPushButton("测试源连接")
        self.test_source_btn.clicked.connect(self.test_source_connection)
        self.test_source_btn.setEnabled(False)
        datasource_layout.addRow("", self.test_source_btn)
        
        datasource_group.setLayout(datasource_layout)
        layout.addWidget(datasource_group)
        
        # 功能说明区域
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        
        info_label = QLabel("""
<b>现有客户银行信息补全说明:</b><br>
<b>数据流向:</b> 源数据库 → 目标数据库<br><br>
<b>处理逻辑:</b><br>
• 遍历目标库 ba_rlb_customer 表中的所有客户<br>
• 根据 legal_name 在源库 ea_dy_legal 表中查找对应法人<br>
• 获取源库 ea_dy_legal_cards 中的银行卡信息<br>
• 关联查询 ea_dy_bankcard 和 ea_dy_currency 获取完整信息<br>
• 转换数据格式并插入目标库 ba_zhb_bank 表<br><br>
<b>重复检查条件:</b><br>
• rlb_customer_id + bank_id + bank_card_type + currency_id + bank_name_id<br>
• 如果已存在，则更新余额；如果不存在，则插入新记录<br><br>
<b>字段映射:</b><br>
• bank_card_type: 0→opt0(个人卡), 1→opt1(企业卡), 2→opt2(第三方卡)<br>
• currency_id: 通过 ea_dy_currency.code 匹配 ba_currency.currency_en<br>
• bank_name_id: 通过 ea_dy_bankcard.name 匹配 ba_zhb_bank_name.bank_name<br>
• remark: 固定为"脚本自动补充"<br>
• admin_id, admin_dept_id: 固定为1
        """)
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # 操作按钮区域
        button_layout = QHBoxLayout()
        
        self.start_btn = QPushButton("开始同步")
        self.start_btn.clicked.connect(self.start_sync)
        self.start_btn.setEnabled(False)
        button_layout.addWidget(self.start_btn)
        
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
        self.target_datasource_combo.clear()
        self.source_datasource_combo.clear()
        
        self.target_datasource_combo.addItem("请选择目标数据源", None)
        self.source_datasource_combo.addItem("请选择源数据源", None)
        
        datasources = self.db_manager.get_all_datasources()
        for ds in datasources:
            display_name = f"{ds.name} ({ds.host}:{ds.port})"
            self.target_datasource_combo.addItem(display_name, ds)
            self.source_datasource_combo.addItem(display_name, ds)
    
    def on_datasource_changed(self):
        """数据源选择变化"""
        target_data = self.target_datasource_combo.currentData()
        source_data = self.source_datasource_combo.currentData()
        
        self.test_target_btn.setEnabled(target_data is not None)
        self.test_source_btn.setEnabled(source_data is not None)
        self.start_btn.setEnabled(target_data is not None and source_data is not None)
    
    def test_target_connection(self):
        """测试目标数据库连接"""
        datasource = self.target_datasource_combo.currentData()
        if not datasource:
            return
        
        success, message = self.db_manager.test_connection(datasource)
        if success:
            QMessageBox.information(self, "目标连接成功", message)
        else:
            QMessageBox.critical(self, "目标连接失败", message)
    
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
    
    def start_sync(self):
        """开始同步"""
        target_datasource = self.target_datasource_combo.currentData()
        source_datasource = self.source_datasource_combo.currentData()
        
        if not target_datasource or not source_datasource:
            QMessageBox.warning(self, "警告", "请选择目标数据源和源数据源！")
            return
        
        # 确认对话框
        reply = QMessageBox.question(
            self,
            "确认同步",
            f"确定要同步银行信息吗？\n\n"
            f"目标数据库: {target_datasource.name}\n"
            f"源数据库: {source_datasource.name}\n\n"
            f"此操作将从源数据库同步银行卡信息到目标数据库。\n"
            f"已存在的记录会更新余额，不存在的会插入新记录。\n"
            f"操作可能需要较长时间，请耐心等待。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        # 禁用按钮，显示进度条
        self.start_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在同步中...")
        self.log_text.clear()
        self.failed_records = []
        
        # 启动工作线程
        self.worker = BankInfoSyncWorker(target_datasource, source_datasource, self.db_manager)
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log_message.connect(self.append_log)
        self.worker.finished.connect(self.on_sync_finished)
        self.worker.error.connect(self.on_sync_error)
        self.worker.start()
    
    def append_log(self, message):
        """添加日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
    
    def on_sync_finished(self, results):
        """同步完成"""
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.failed_records = results['failed_records']
        
        result_text = (
            f"同步完成！\n"
            f"客户总数：{results['customer_total']} 个\n"
            f"已处理：{results['customer_processed']} 个\n"
            f"跳过：{results['customer_skipped']} 个\n"
            f"找到银行卡：{results['bank_cards_found']} 张\n"
            f"新增银行卡：{results['bank_cards_inserted']} 张\n"
            f"更新银行卡：{results['bank_cards_updated']} 张\n"
            f"跳过银行卡：{results['bank_cards_skipped']} 张\n"
            f"失败记录：{len(results['failed_records'])} 个"
        )
        
        self.result_label.setText(result_text)
        if len(results['failed_records']) > 0:
            self.result_label.setStyleSheet("color: orange;")
            self.export_failed_btn.setEnabled(True)
        else:
            self.result_label.setStyleSheet("color: green;")
        
        QMessageBox.information(self, "同步完成", result_text)
    
    def on_sync_error(self, error_msg):
        """同步错误"""
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.result_label.setText("同步失败！")
        self.result_label.setStyleSheet("color: red;")
        
        QMessageBox.critical(self, "同步错误", f"同步失败：{error_msg}")
    
    def export_failed_records(self):
        """导出失败记录"""
        if not self.failed_records:
            QMessageBox.information(self, "提示", "没有失败记录可导出")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存失败记录",
            f"bank_sync_failed_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel文件 (*.xlsx)"
        )
        
        if file_path:
            try:
                # 准备导出数据
                export_data = []
                for record in self.failed_records:
                    # 转换银行卡类型显示
                    card_type_display = ''
                    if record.get('bank_card_type') == 0:
                        card_type_display = '0-个人卡'
                    elif record.get('bank_card_type') == 1:
                        card_type_display = '1-企业卡'
                    elif record.get('bank_card_type') == 2:
                        card_type_display = '2-第三方卡'
                    else:
                        card_type_display = str(record.get('bank_card_type', ''))
                    
                    export_data.append({
                        '客户ID': record.get('customer_id', ''),
                        '法人姓名': record.get('legal_name', ''),
                        '源银行卡ID': record.get('source_card_id', ''),
                        '银行账号': record.get('bank_account', ''),
                        '银行名称': record.get('bank_name', ''),
                        '货币代码': record.get('currency_code', ''),
                        '银行卡类型': card_type_display,
                        '余额': record.get('balance', ''),
                        '源备注': record.get('source_remark', ''),
                        '失败原因': record.get('error', '')
                    })
                
                # 创建DataFrame并导出
                df = pd.DataFrame(export_data)
                df.to_excel(file_path, index=False, engine='openpyxl')
                
                QMessageBox.information(self, "导出成功", f"失败记录已导出到：\n{file_path}")
                
            except Exception as e:
                QMessageBox.critical(self, "导出失败", f"导出失败记录时出错：{str(e)}")
    
    def clear_log(self):
        """清空日志"""
        self.log_text.clear()
