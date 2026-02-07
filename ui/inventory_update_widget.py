"""
根据库存ID补充库存内容功能界面
"""
import os
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog, 
                             QListWidget, QListWidgetItem, QCheckBox, QSplitter,
                             QScrollArea, QTableWidget, QTableWidgetItem,
                             QHeaderView)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
import pandas as pd
import pymysql


class InventoryUpdateWorker(QThread):
    """库存更新工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    
    def __init__(self, excel_file, id_column, field_mappings, db_config):
        super().__init__()
        self.excel_file = excel_file
        self.id_column = id_column  # Excel中的ID列名
        self.field_mappings = field_mappings  # {excel_column: db_field} 映射关系
        self.db_config = db_config
        self.results = {
            'total_rows': 0,
            'updated_count': 0,
            'failed_count': 0,
            'not_found_count': 0,
            'failed_records': []
        }
    
    def run(self):
        try:
            self.log_message.emit("开始读取Excel文件...")
            self.progress.emit(10)
            
            # 读取Excel文件
            df = pd.read_excel(self.excel_file, engine='openpyxl')
            self.log_message.emit(f"成功读取Excel，共 {len(df)} 行数据")
            
            # 检查ID列是否存在
            if self.id_column not in df.columns:
                raise Exception(f"Excel中未找到ID列: {self.id_column}")
            
            # 检查映射字段是否存在
            missing_columns = [col for col in self.field_mappings.keys() if col not in df.columns]
            if missing_columns:
                raise Exception(f"Excel中未找到以下列: {', '.join(missing_columns)}")
            
            self.log_message.emit("连接数据库...")
            self.progress.emit(20)
            
            # 连接数据库
            connection = pymysql.connect(**self.db_config)
            cursor = connection.cursor()
            
            self.log_message.emit("开始更新库存记录...")
            self.progress.emit(30)
            
            total_rows = len(df)
            updated_count = 0
            failed_count = 0
            not_found_count = 0
            failed_records = []
            
            for index, row in df.iterrows():
                try:
                    # 获取ID值
                    inventory_id = row[self.id_column]
                    if pd.isna(inventory_id):
                        self.log_message.emit(f"第{index+1}行: ID为空，跳过")
                        failed_count += 1
                        failed_records.append({
                            'row': index + 1,
                            'id': 'NULL',
                            'error': 'ID为空'
                        })
                        continue
                    
                    # 检查记录是否存在
                    cursor.execute("SELECT id FROM ba_account_mag WHERE id = %s", (inventory_id,))
                    if not cursor.fetchone():
                        self.log_message.emit(f"第{index+1}行: 库存ID {inventory_id} 不存在")
                        not_found_count += 1
                        failed_records.append({
                            'row': index + 1,
                            'id': inventory_id,
                            'error': '记录不存在'
                        })
                        continue
                    
                    # 构建更新SQL
                    update_fields = []
                    update_values = []
                    
                    for excel_col, db_field in self.field_mappings.items():
                        value = row[excel_col]
                        # 处理空值
                        if pd.isna(value):
                            value = None
                        elif isinstance(value, str):
                            value = value.strip()
                        
                        update_fields.append(f"{db_field} = %s")
                        update_values.append(value)
                    
                    if update_fields:
                        # 添加更新时间
                        update_fields.append("update_time = %s")
                        update_values.append(int(datetime.now().timestamp()))
                        
                        # 执行更新
                        sql = f"UPDATE ba_account_mag SET {', '.join(update_fields)} WHERE id = %s"
                        update_values.append(inventory_id)
                        
                        cursor.execute(sql, update_values)
                        updated_count += 1
                        
                        if updated_count % 10 == 0:
                            self.log_message.emit(f"已更新 {updated_count} 条记录")
                    
                except Exception as e:
                    failed_count += 1
                    error_msg = str(e)
                    self.log_message.emit(f"第{index+1}行更新失败: {error_msg}")
                    failed_records.append({
                        'row': index + 1,
                        'id': inventory_id if 'inventory_id' in locals() else 'Unknown',
                        'error': error_msg
                    })
                
                # 更新进度
                progress = 30 + int((index + 1) / total_rows * 60)
                self.progress.emit(progress)
            
            # 提交事务
            connection.commit()
            self.log_message.emit("数据库事务已提交")
            
            # 关闭数据库连接
            cursor.close()
            connection.close()
            
            self.results['total_rows'] = total_rows
            self.results['updated_count'] = updated_count
            self.results['failed_count'] = failed_count
            self.results['not_found_count'] = not_found_count
            self.results['failed_records'] = failed_records
            
            self.log_message.emit("更新完成！")
            self.progress.emit(100)
            
            self.finished.emit(self.results)
            
        except Exception as e:
            self.error.emit(str(e))


class InventoryUpdateWidget(QWidget):
    """库存更新界面"""
    
    def __init__(self, db_manager=None):
        super().__init__()
        self.db_manager = db_manager
        self.excel_file = ""
        self.excel_columns = []
        self.field_mappings = {}  # {excel_column: db_field}
        self.worker = None
        self.selected_datasource = None
        
        # 数据库字段定义
        self.db_fields = {
            'inventory_num': '库存编号',
            'shudi_id': '属地关联id',
            'platform_id': '平台关联id',
            'register_id': '注册表关联id',
            'account_status': '账号状态',
            'init_status': '初始状态',
            'banned_reason': '封号原因',
            'mobile_num': '手机编号',
            'bank_verification_id': '银行验证关联id',
            'currency': '币种',
            'environment': '环境',
            'product_type': '产品品类',
            'email_address_password': '邮箱地址密码',
            'website_url': '普货网站',
            'shop_password': '店铺密码',
            'shop_account': '店铺账号',
            'shop_url': '店铺链接',
            'phone_num': '电话号码',
            'id_account_password': '账号密码',
            'origin_customer': '原客户',
            'now_customer': '现客户',
            'is_new_customer': '是否新客户',
            'origin_is_new_customer': '原是否新客户',
            'origin_customer_balance': '原客户余额',
            'assignment_date': '给店日期',
            'origin_assignment_date': '原给店日期',
            'remark': '账号备注',
            'weekly_income_flag': '周度入账标志',
            'withdrawable_limit': '可提现额度',
            'sale_id': '给店列表关联id',
            'origin_sale_id': '原给店编号',
            'api_key': 'APIkey',
            'api_secret': 'APISecret',
            'signature_name': '签名名称',
            'proxy_ip': '代理IP',
            'register_department': '注册部门',
            'bottom_table': '注册底表',
            'shopindex_id': '店铺主体库关联id',
            'bank_id': '银行卡关联id',
            'order_id': '订单编号关联id',
            'source_reg_id': '库存来源',
            'currency_id': '币种ID',
            'account_suspend_times': '过审次数',
            'situation_remarks': '情况备注',
            'legal_id': '法人id',
            'information_id': '日区企业法人id',
            'information_part2_id': '欧美企业法人id',
            'reg_get_store_time': '拿店日期',
            'reg_email': '注册邮箱'
        }
        
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # 功能说明
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        
        info_label = QLabel("""
<b>根据库存ID补充库存内容功能:</b><br>
<b>使用步骤:</b> 1.导入Excel文件 → 2.指定ID列 → 3.设置字段映射 → 4.开始更新<br>
<b>功能说明:</b> 根据Excel中的库存ID，将Excel数据更新到数据库ba_account_mag表中<br>
<b>注意事项:</b> 确保Excel中的ID列包含有效的库存记录ID
        """)
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: #666; font-size: 11px;")
        info_layout.addWidget(info_label)
        
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # 数据库选择区域
        db_group = QGroupBox("数据库选择")
        db_layout = QFormLayout()
        
        self.datasource_combo = QComboBox()
        self.datasource_combo.currentTextChanged.connect(self.on_datasource_changed)
        db_layout.addRow("数据源:", self.datasource_combo)
        
        # 连接状态显示
        self.connection_status_label = QLabel("未连接")
        self.connection_status_label.setStyleSheet("color: gray;")
        db_layout.addRow("连接状态:", self.connection_status_label)
        
        db_group.setLayout(db_layout)
        layout.addWidget(db_group)
        
        # 加载数据源列表
        self.load_datasources()
        
        # 文件选择区域
        file_group = QGroupBox("文件选择")
        file_layout = QFormLayout()
        
        # Excel文件选择
        excel_layout = QHBoxLayout()
        self.excel_label = QLabel("未选择文件")
        self.excel_label.setStyleSheet("color: gray;")
        excel_layout.addWidget(self.excel_label)
        
        self.select_excel_btn = QPushButton("选择Excel文件")
        self.select_excel_btn.clicked.connect(self.select_excel_file)
        excel_layout.addWidget(self.select_excel_btn)
        
        file_layout.addRow("Excel文件:", excel_layout)
        
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # ID列选择区域
        id_group = QGroupBox("ID列设置")
        id_layout = QFormLayout()
        
        self.id_column_combo = QComboBox()
        self.id_column_combo.setEnabled(False)
        id_layout.addRow("库存ID列:", self.id_column_combo)
        
        id_group.setLayout(id_layout)
        layout.addWidget(id_group)
        
        # 字段映射区域
        mapping_group = QGroupBox("字段映射设置")
        mapping_layout = QVBoxLayout()
        
        mapping_info_label = QLabel("设置Excel列与数据库字段的对应关系:")
        mapping_info_label.setStyleSheet("font-weight: bold; color: blue;")
        mapping_layout.addWidget(mapping_info_label)
        
        # 创建映射表格
        self.mapping_table = QTableWidget()
        self.mapping_table.setColumnCount(3)
        self.mapping_table.setHorizontalHeaderLabels(["Excel列", "数据库字段", "字段说明"])
        
        # 设置表格属性
        header = self.mapping_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        
        self.mapping_table.setMaximumHeight(300)
        mapping_layout.addWidget(self.mapping_table)
        
        # 映射操作按钮
        mapping_btn_layout = QHBoxLayout()
        
        self.add_mapping_btn = QPushButton("添加映射")
        self.add_mapping_btn.clicked.connect(self.add_mapping_row)
        self.add_mapping_btn.setEnabled(False)
        mapping_btn_layout.addWidget(self.add_mapping_btn)
        
        self.remove_mapping_btn = QPushButton("删除选中")
        self.remove_mapping_btn.clicked.connect(self.remove_mapping_row)
        self.remove_mapping_btn.setEnabled(False)
        mapping_btn_layout.addWidget(self.remove_mapping_btn)
        
        mapping_btn_layout.addStretch()
        mapping_layout.addLayout(mapping_btn_layout)
        
        mapping_group.setLayout(mapping_layout)
        layout.addWidget(mapping_group)
        
        # 操作按钮区域
        button_layout = QHBoxLayout()
        
        self.start_btn = QPushButton("开始更新")
        self.start_btn.clicked.connect(self.start_update)
        self.start_btn.setEnabled(False)
        self.start_btn.setStyleSheet("""
            QPushButton {
                background-color: #28a745;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 5px;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #218838;
            }
            QPushButton:disabled {
                background-color: #ccc;
                color: #666;
            }
        """)
        button_layout.addWidget(self.start_btn)
        
        button_layout.addStretch()
        layout.addLayout(button_layout)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        # 结果显示区域
        result_group = QGroupBox("处理结果")
        result_layout = QVBoxLayout()
        
        self.result_label = QLabel("等待开始...")
        self.result_label.setStyleSheet("font-weight: bold;")
        result_layout.addWidget(self.result_label)
        
        # 导出失败记录按钮
        export_layout = QHBoxLayout()
        
        self.export_failed_btn = QPushButton("导出失败记录")
        self.export_failed_btn.clicked.connect(self.export_failed_records)
        self.export_failed_btn.setEnabled(False)
        self.export_failed_btn.setStyleSheet("""
            QPushButton {
                background-color: #dc3545;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #c82333;
            }
            QPushButton:disabled {
                background-color: #ccc;
                color: #666;
            }
        """)
        export_layout.addWidget(self.export_failed_btn)
        
        export_layout.addStretch()
        result_layout.addLayout(export_layout)
        
        result_group.setLayout(result_layout)
        layout.addWidget(result_group)
        
        # 日志区域
        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout()
        
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(150)
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)
        
        self.setLayout(layout)
    
    def load_datasources(self):
        """加载数据源列表"""
        if not self.db_manager:
            return
        
        try:
            datasources = self.db_manager.get_all_datasources()
            self.datasource_combo.clear()
            self.datasource_combo.addItem("请选择数据源", None)
            
            for datasource in datasources:
                display_text = f"{datasource.name} ({datasource.host}:{datasource.port})"
                self.datasource_combo.addItem(display_text, datasource)
                
        except Exception as e:
            QMessageBox.warning(self, "警告", f"加载数据源失败: {str(e)}")
    
    def on_datasource_changed(self):
        """数据源选择改变"""
        current_data = self.datasource_combo.currentData()
        if current_data:
            self.selected_datasource = current_data
            self.connection_status_label.setText(f"已选择: {current_data.name}")
            self.connection_status_label.setStyleSheet("color: green;")
        else:
            self.selected_datasource = None
            self.connection_status_label.setText("未连接")
            self.connection_status_label.setStyleSheet("color: gray;")
        
        self.update_start_button_state()
    
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
                # 读取Excel文件获取列名
                df = pd.read_excel(file_path, engine='openpyxl', nrows=0)
                self.excel_columns = list(df.columns)
                
                self.excel_file = file_path
                self.excel_label.setText(os.path.basename(file_path))
                self.excel_label.setStyleSheet("color: green;")
                
                # 填充ID列下拉框
                self.id_column_combo.clear()
                self.id_column_combo.addItems(self.excel_columns)
                self.id_column_combo.setEnabled(True)
                
                # 启用映射按钮
                self.add_mapping_btn.setEnabled(True)
                self.remove_mapping_btn.setEnabled(True)
                
                # 清空现有映射
                self.mapping_table.setRowCount(0)
                
                self.update_start_button_state()
                
                self.log_text.append(f"已选择Excel文件: {os.path.basename(file_path)}")
                self.log_text.append(f"检测到列: {', '.join(self.excel_columns)}")
                
            except Exception as e:
                QMessageBox.critical(self, "错误", f"读取Excel文件失败: {str(e)}")
                self.excel_file = ""
                self.excel_label.setText("未选择文件")
                self.excel_label.setStyleSheet("color: gray;")
    
    def add_mapping_row(self):
        """添加映射行"""
        row = self.mapping_table.rowCount()
        self.mapping_table.insertRow(row)
        
        # Excel列下拉框
        excel_combo = QComboBox()
        excel_combo.addItems(self.excel_columns)
        excel_combo.currentTextChanged.connect(self.update_start_button_state)
        self.mapping_table.setCellWidget(row, 0, excel_combo)
        
        # 数据库字段下拉框
        db_combo = QComboBox()
        db_combo.addItems(list(self.db_fields.keys()))
        db_combo.currentTextChanged.connect(self.on_db_field_changed)
        db_combo.currentTextChanged.connect(self.update_start_button_state)
        self.mapping_table.setCellWidget(row, 1, db_combo)
        
        # 字段说明
        desc_item = QTableWidgetItem("")
        desc_item.setFlags(desc_item.flags() & ~Qt.ItemIsEditable)  # 只读
        self.mapping_table.setItem(row, 2, desc_item)
        
        self.update_start_button_state()
    
    def remove_mapping_row(self):
        """删除选中的映射行"""
        current_row = self.mapping_table.currentRow()
        if current_row >= 0:
            self.mapping_table.removeRow(current_row)
            self.update_start_button_state()
    
    def on_db_field_changed(self):
        """数据库字段改变时更新说明"""
        sender = self.sender()
        if sender:
            # 找到发送信号的控件所在行
            for row in range(self.mapping_table.rowCount()):
                if self.mapping_table.cellWidget(row, 1) == sender:
                    field_name = sender.currentText()
                    desc = self.db_fields.get(field_name, "")
                    desc_item = self.mapping_table.item(row, 2)
                    if desc_item:
                        desc_item.setText(desc)
                    break
    
    def update_start_button_state(self):
        """更新开始按钮状态"""
        # 检查必要的控件是否已创建
        if not hasattr(self, 'start_btn') or not hasattr(self, 'id_column_combo') or not hasattr(self, 'mapping_table'):
            return
            
        can_start = bool(
            self.selected_datasource and
            self.excel_file and
            self.id_column_combo.currentText() and
            self.mapping_table.rowCount() > 0
        )
        
        # 检查是否有有效的映射
        if can_start:
            valid_mappings = 0
            for row in range(self.mapping_table.rowCount()):
                excel_combo = self.mapping_table.cellWidget(row, 0)
                db_combo = self.mapping_table.cellWidget(row, 1)
                if excel_combo and db_combo and excel_combo.currentText() and db_combo.currentText():
                    valid_mappings += 1
            
            can_start = valid_mappings > 0
        
        self.start_btn.setEnabled(can_start)
    
    def start_update(self):
        """开始更新"""
        # 获取ID列
        id_column = self.id_column_combo.currentText()
        
        # 获取字段映射
        field_mappings = {}
        for row in range(self.mapping_table.rowCount()):
            excel_combo = self.mapping_table.cellWidget(row, 0)
            db_combo = self.mapping_table.cellWidget(row, 1)
            
            if excel_combo and db_combo:
                excel_col = excel_combo.currentText()
                db_field = db_combo.currentText()
                
                if excel_col and db_field:
                    field_mappings[excel_col] = db_field
        
        if not field_mappings:
            QMessageBox.warning(self, "警告", "请至少设置一个字段映射")
            return
        
        # 确认对话框
        mapping_text = "\n".join([f"  {excel_col} → {db_field}" for excel_col, db_field in field_mappings.items()])
        reply = QMessageBox.question(
            self,
            "确认更新",
            f"确定要开始更新库存记录吗？\n\n"
            f"Excel文件: {os.path.basename(self.excel_file)}\n"
            f"ID列: {id_column}\n"
            f"字段映射:\n{mapping_text}\n\n"
            f"操作不可撤销，请确认！",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        # 获取选择的数据库配置
        if not self.selected_datasource:
            QMessageBox.warning(self, "警告", "请先选择数据源")
            return
        
        db_config = {
            'host': self.selected_datasource.host,
            'port': self.selected_datasource.port,
            'user': self.selected_datasource.username,
            'password': self.selected_datasource.password,
            'database': self.selected_datasource.database,
            'charset': 'utf8mb4'
        }
        
        # 禁用按钮，显示进度条
        self.start_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在更新中...")
        self.log_text.clear()
        
        # 启动工作线程
        self.worker = InventoryUpdateWorker(
            self.excel_file, id_column, field_mappings, db_config
        )
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
        """更新完成"""
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.update_results = results  # 保存结果用于导出
        
        total_rows = results['total_rows']
        updated_count = results['updated_count']
        failed_count = results['failed_count']
        not_found_count = results['not_found_count']
        
        result_text = f"更新完成！总行数: {total_rows}, 成功更新: {updated_count}, 失败: {failed_count}, 记录不存在: {not_found_count}"
        self.result_label.setText(result_text)
        self.result_label.setStyleSheet("color: green; font-weight: bold;")
        
        # 启用导出失败记录按钮
        if failed_count > 0 or not_found_count > 0:
            self.export_failed_btn.setEnabled(True)
            self.export_failed_btn.setText(f"导出失败记录 ({failed_count + not_found_count}条)")
        
        self.log_text.append(f"\n=== 更新完成 ===")
        self.log_text.append(result_text)
        self.log_text.append(f"成功率: {updated_count/total_rows*100:.1f}%")
        
        QMessageBox.information(self, "完成", result_text)
    
    def export_failed_records(self):
        """导出失败记录"""
        if not hasattr(self, 'update_results') or not self.update_results['failed_records']:
            QMessageBox.warning(self, "警告", "没有可导出的失败记录")
            return
        
        # 选择保存文件
        default_name = f"库存更新失败记录_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存失败记录",
            default_name,
            "Excel文件 (*.xlsx)"
        )
        
        if file_path:
            try:
                # 导出失败记录
                failed_df = pd.DataFrame(self.update_results['failed_records'])
                failed_df.to_excel(file_path, index=False, engine='openpyxl')
                
                count = len(self.update_results['failed_records'])
                QMessageBox.information(
                    self, 
                    "导出成功", 
                    f"已成功导出 {count} 条失败记录到:\n{file_path}"
                )
                
                self.log_text.append(f"失败记录已导出到: {file_path}")
                
            except Exception as e:
                QMessageBox.critical(self, "导出失败", f"导出失败记录时出错: {str(e)}")
    
    def on_update_error(self, error_message):
        """更新错误"""
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.result_label.setText("更新失败！")
        self.result_label.setStyleSheet("color: red; font-weight: bold;")
        
        QMessageBox.critical(self, "错误", f"更新过程出错: {error_message}")
        self.log_text.append(f"错误: {error_message}")
