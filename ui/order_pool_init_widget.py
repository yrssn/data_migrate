"""
导入平台注册部Excel表初始化订单池功能界面
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


class OrderPoolInitWorker(QThread):
    """订单池初始化工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)  # 返回统计结果
    error = pyqtSignal(str)
    
    def __init__(self, datasource, db_manager, excel_file):
        super().__init__()
        self.datasource = datasource
        self.db_manager = db_manager
        self.excel_file = excel_file
        self.results = {
            'total_rows': 0,
            'success_count': 0,
            'failed_count': 0,
            'failed_records': [],
            'completed_data': []  # 存储包含订单ID的完整数据
        }
        # 订单号生成相关
        self.current_date = datetime.now().strftime('%Y%m%d')
        self.daily_counter = 1
    
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
            
            # 连接数据库
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
            
            # 获取当前日期已有的最大订单号，确定起始计数器
            self.init_daily_counter(cursor)
            
            # 处理每一行数据
            for index, row in df.iterrows():
                try:
                    order_id, primary_key_id = self.process_row(cursor, index + 1, row)
                    self.results['success_count'] += 1
                    
                    # 记录包含订单ID和主键ID的完整数据
                    completed_row = row.to_dict()
                    completed_row['生成的订单ID'] = order_id
                    completed_row['数据库主键ID'] = primary_key_id
                    completed_row['处理状态'] = '成功'
                    completed_row['处理时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    self.results['completed_data'].append(completed_row)
                    
                    # 更新进度
                    progress = 15 + int((index + 1) / len(df) * 80)
                    self.progress.emit(progress)
                    
                    # 每100行提交一次
                    if (index + 1) % 100 == 0:
                        connection.commit()
                        self.log_message.emit(f"已处理 {index + 1} 行，提交事务")
                        
                except Exception as e:
                    self.results['failed_count'] += 1
                    error_msg = f"第{index + 1}行处理失败: {str(e)}"
                    self.log_message.emit(error_msg)
                    
                    # 记录失败的数据
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
            
            # 最终提交
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
            # 查找当日已有的最大订单号
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
                # 提取最后3位数字
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
            # 如果当日超过999，切换到下一天
            current_datetime = datetime.now()
            next_day = current_datetime.replace(day=current_datetime.day + 1)
            self.current_date = next_day.strftime('%Y%m%d')
            self.daily_counter = 1
            self.log_message.emit(f"当日订单号已满，切换到下一天: {self.current_date}")
        
        order_id = f"OD{self.current_date}{self.daily_counter:03d}"
        self.daily_counter += 1
        return order_id
    
    def process_row(self, cursor, row_num, row):
        """处理单行数据"""
        try:
            # 获取Excel列数据
            platform_name = str(row.iloc[3]) if len(row) > 3 else ""  # D列：平台名称
            company_country = str(row.iloc[5]) if len(row) > 5 else ""  # F列：企业国别
            attribute_value = str(row.iloc[6]) if len(row) > 6 else ""  # G列：属性
            territory_abbr = str(row.iloc[7]) if len(row) > 7 else ""  # H列：属地简称
            
            if not platform_name or platform_name == 'nan':
                raise Exception("D列(平台名称)为空")
            if not attribute_value or attribute_value == 'nan':
                raise Exception("G列(属性)为空")
            if not territory_abbr or territory_abbr == 'nan':
                raise Exception("H列(属地简称)为空")
            
            # 记录企业国别信息（如果有的话）
            country_info = f", 企业国别={company_country}" if company_country and company_country != 'nan' else ""
            self.log_message.emit(f"第{row_num}行: 处理平台={platform_name}, 属性={attribute_value}, 属地={territory_abbr}{country_info}")
            
            # 第一步：根据平台名称获取platform信息
            platform_info = self.get_platform_info(cursor, platform_name)
            if not platform_info:
                raise Exception(f"未找到平台: {platform_name}")
            
            platform_id, platform_type = platform_info
            
            # 第二步：根据属地简称获取shudi_id
            shudi_id = self.get_shudi_id(cursor, territory_abbr)
            if not shudi_id:
                raise Exception(f"未找到属地: {territory_abbr}")
            
            # 第三步：验证和转换属性值
            attribute_enum = self.convert_attribute(attribute_value)
            if not attribute_enum:
                raise Exception(f"无效的属性值: {attribute_value}")
            
            # 第四步：生成订单号
            order_id = self.generate_order_id()
            
            # 第五步：插入订单记录
            primary_key_id = self.insert_order(cursor, order_id, platform_id, platform_type, shudi_id, 
                                              attribute_enum, attribute_value, company_country, row_num)
            
            self.log_message.emit(f"第{row_num}行: 成功创建订单 {order_id}，主键ID: {primary_key_id}")
            
            # 返回订单号和主键ID
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
                    attribute_enum, attribute_value, company_country, row_num):
        """插入订单记录"""
        current_time = int(datetime.now().timestamp())
        
        # 构建插入数据
        insert_data = {
            'order_id': order_id,
            'shop_type': platform_type,  # 从ba_platform.platform_type获取
            'regist_department': 'opt0',  # 默认平台注册部
            'platform_id': platform_id,
            'shudi_id': shudi_id,
            'attribute': attribute_enum,
            'order_number': 1,
            'rejected_number': 0,
            'done_number': 1,
            'status': 4,  # 已完成
            'demand': '系统初始化数据',
            'feedback': '',
            'admin_id': 1,
            'admin_dept_id': 1,
            'create_time': current_time,
            'update_time': current_time,
            'regesit_status': 4,  # 已完成
            'need_data_type': 1,
            'rlb_status': 1
        }
        
        # 添加企业国别字段（F列数据）
        if company_country and company_country != 'nan':
            insert_data['company_country'] = company_country
            self.log_message.emit(f"第{row_num}行: 设置企业国别为 {company_country}")
        else:
            # 如果F列为空或无效，记录日志
            if attribute_value == '企业法人':
                self.log_message.emit(f"第{row_num}行: 企业法人但F列企业国别为空")
        
        # 构建SQL
        fields = list(insert_data.keys())
        placeholders = ['%s'] * len(fields)
        values = list(insert_data.values())
        
        insert_sql = f"""
            INSERT INTO ba_order ({', '.join(fields)}) 
            VALUES ({', '.join(placeholders)})
        """
        
        cursor.execute(insert_sql, values)
        
        # 获取插入后的主键ID
        primary_key_id = cursor.lastrowid
        
        self.log_message.emit(f"第{row_num}行: 插入订单数据 - 平台ID:{platform_id}, 属地ID:{shudi_id}, 属性:{attribute_enum}")
        
        # 返回主键ID
        return primary_key_id


class OrderPoolInitWidget(QWidget):
    """导入平台注册部Excel表初始化订单池功能组件"""
    
    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.excel_file = None
        self.failed_records = []
        self.completed_data = []
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
        
        # 功能说明区域
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        
        info_label = QLabel("""
<b>导入平台注册部Excel表初始化订单池说明:</b><br>
<b>Excel列要求:</b><br>
• D列: 平台名称 (用于查找ba_platform表)<br>
• F列: 企业国别 (企业法人时使用)<br>
• G列: 属性 (个人/个人事业主/企业法人)<br>
• H列: 属地简称 (用于查找ba_shudi表)<br><br>
<b>订单号生成规则:</b><br>
• 格式: OD + 日期(YYYYMMDD) + 3位序号<br>
• 示例: OD20251001001, OD20251001002...<br>
• 每日最多999个订单，超过自动切换到下一天<br><br>
<b>数据映射规则:</b><br>
• shop_type: 从ba_platform.platform_type获取<br>
• regist_department: 固定为opt0(平台注册部)<br>
• platform_id: 从ba_platform.id获取<br>
• shudi_id: 从ba_shudi.id获取<br>
• attribute: 个人→opt0, 个人事业主→opt1, 企业法人→opt2<br>
• company_country: 企业法人时从F列获取<br><br>
<b>固定字段值:</b><br>
• order_number: 1, done_number: 1<br>
• status: 4(已完成), regesit_status: 4(已完成)<br>
• demand: "系统初始化数据"<br>
• admin_id: 1, admin_dept_id: 1<br><br>
<b>错误处理:</b><br>
• 平台名称未找到 → 跳过并记录<br>
• 属地简称未找到 → 跳过并记录<br>
• 属性值无效 → 跳过并记录<br>
• 支持导出失败记录Excel进行分析<br><br>
<b>结果导出:</b><br>
• <b>导出完整结果</b>: 包含原始数据+生成的订单ID+数据库主键ID+处理状态<br>
• <b>导出失败记录</b>: 仅导出处理失败的记录进行分析<br>
• 完整结果Excel可用于后续业务流程追踪<br>
• 数据库主键ID可用于直接关联ba_order表记录
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
            self.excel_file = file_path
            self.file_label.setText(os.path.basename(file_path))
            self.file_label.setStyleSheet("color: green;")
            self.update_start_button_state()
    
    def update_start_button_state(self):
        """更新开始按钮状态"""
        datasource_selected = self.datasource_combo.currentData() is not None
        file_selected = self.excel_file is not None
        self.start_btn.setEnabled(datasource_selected and file_selected)
    
    def start_import(self):
        """开始导入"""
        datasource = self.datasource_combo.currentData()
        if not datasource or not self.excel_file:
            QMessageBox.warning(self, "警告", "请选择数据源和Excel文件！")
            return
        
        # 确认对话框
        reply = QMessageBox.question(
            self,
            "确认导入",
            f"确定要导入订单数据吗？\n\n"
            f"数据库: {datasource.name}\n"
            f"Excel文件: {os.path.basename(self.excel_file)}\n\n"
            f"此操作将根据Excel数据批量创建订单记录。\n"
            f"订单号格式: OD{datetime.now().strftime('%Y%m%d')}XXX\n"
            f"操作不可撤销，请确认数据正确！",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        # 禁用按钮，显示进度条
        self.start_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在导入中...")
        self.log_text.clear()
        self.failed_records = []
        self.completed_data = []
        
        # 启动工作线程
        self.worker = OrderPoolInitWorker(datasource, self.db_manager, self.excel_file)
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
        
        # 启用导出完整结果按钮
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
        """导出完整结果（包含订单ID）"""
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
                # 创建DataFrame并导出
                df = pd.DataFrame(self.completed_data)
                df.to_excel(file_path, index=False, engine='openpyxl')
                
                QMessageBox.information(self, "导出成功", f"完整结果已导出到：\n{file_path}\n\n包含字段：\n• 原始Excel所有列\n• 生成的订单ID\n• 数据库主键ID\n• 处理状态\n• 处理时间\n• 失败原因（如有）")
                
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
