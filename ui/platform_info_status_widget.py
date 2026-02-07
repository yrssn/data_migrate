import sys
import pandas as pd
import pymysql
from datetime import datetime
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *


class PlatformInfoStatusWorker(QThread):
    """平台详情状态更新工作线程"""
    
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    
    def __init__(self, datasource, excel_file, target_status):
        super().__init__()
        self.datasource = datasource
        self.excel_file = excel_file
        self.target_status = target_status
        self.results = {
            'success_count': 0,
            'failed_count': 0,
            'completed_data': [],
            'failed_records': []
        }
    
    def run(self):
        try:
            self.log_message.emit("开始读取Excel文件...")
            df = pd.read_excel(self.excel_file)
            self.log_message.emit(f"读取到 {len(df)} 行数据")
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
            cursor = connection.cursor()
            self.log_message.emit("数据库连接成功")
            self.progress.emit(15)
            
            # 处理每一行数据
            for index, row in df.iterrows():
                try:
                    result = self.process_row(cursor, index + 1, row)
                    if result:
                        self.results['success_count'] += 1
                        # 记录成功数据
                        completed_row = row.to_dict()
                        completed_row.update(result)
                        completed_row['处理状态'] = '成功'
                        completed_row['处理时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        self.results['completed_data'].append(completed_row)
                    
                    # 更新进度
                    progress = 15 + int((index + 1) / len(df) * 80)
                    self.progress.emit(progress)
                    
                except Exception as e:
                    self.results['failed_count'] += 1
                    error_msg = f"第{index + 1}行处理失败: {str(e)}"
                    self.log_message.emit(error_msg)
                    
                    # 记录失败数据
                    failed_row = row.to_dict()
                    failed_row['处理状态'] = '失败'
                    failed_row['失败原因'] = str(e)
                    failed_row['处理时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    self.results['completed_data'].append(failed_row)
                    
                    self.results['failed_records'].append({
                        'row': index + 1,
                        'error': str(e),
                        'data': row.to_dict()
                    })
            
            # 提交事务
            connection.commit()
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
            
            # 获取Excel列数据
            col_a = clean_cell_value(row.iloc[0]) if len(row) > 0 else ""  # A列：平台标识
            col_b = clean_cell_value(row.iloc[1]) if len(row) > 1 else ""  # B列：法人姓名
            col_h = clean_cell_value(row.iloc[7]) if len(row) > 7 else ""  # H列：属性判断
            col_i = clean_cell_value(row.iloc[8]) if len(row) > 8 else ""  # I列：公司名称
            col_territory = clean_cell_value(row.iloc[9]) if len(row) > 9 else ""  # J列：属地简称
            
            self.log_message.emit(f"第{row_num}行: 处理数据 - 平台={col_a}, 法人={col_b}, 属性={col_h}, 属地={col_territory}")
            
            # 第一步：查找shopindex_id
            shopindex_id = self.find_shopindex_id(cursor, col_b, col_i, col_h, row_num)
            
            # 第二步：查找shudi_id
            shudi_id = self.find_shudi_id(cursor, col_territory, row_num)
            
            # 第三步：查找platform_id
            platform_id = self.find_platform_id(cursor, col_a, row_num)
            
            # 第四步：查找并更新platform_info
            platform_info_id = self.update_platform_info(cursor, shopindex_id, shudi_id, platform_id, row_num)
            
            return {
                'shopindex_id': shopindex_id,
                'shudi_id': shudi_id,
                'platform_id': platform_id,
                'platform_info_id': platform_info_id,
                '更新状态': self.target_status
            }
            
        except Exception as e:
            raise e
    
    def find_shopindex_id(self, cursor, legal_name, company_name, attribute_or_country, row_num):
        """根据属性判断查找shopindex_id"""
        try:
            # 先通过legal_name找到customer_id
            cursor.execute("""
                SELECT id FROM ba_rlb_customer 
                WHERE legal_name = %s AND (delete_time IS NULL OR delete_time = 0)
            """, (legal_name,))
            customer_result = cursor.fetchone()
            
            if not customer_result:
                raise Exception(f"未找到客户: {legal_name}")
            
            customer_id = customer_result[0]
            self.log_message.emit(f"第{row_num}行: 找到客户ID={customer_id}")
            
            # 判断是否企业法人
            if attribute_or_country != '企业法人':
                # 非企业法人：查找belong_information为null的shopindex
                cursor.execute("""
                    SELECT id FROM ba_shopindex 
                    WHERE legal_id = %s AND belong_information IS NULL 
                    AND (delete_time IS NULL OR delete_time = 0)
                """, (customer_id,))
                shopindex_result = cursor.fetchone()
                
                if not shopindex_result:
                    raise Exception(f"未找到个人shopindex记录 (customer_id: {customer_id})")
                
                return shopindex_result[0]
            
            else:
                # 企业法人：优先查找ba_rlb_legal_information，找不到再查ba_rlb_legal_information_part2
                self.log_message.emit(f"第{row_num}行: 企业法人，优先查找ba_rlb_legal_information表")
                
                # 先查找ba_rlb_legal_information
                cursor.execute("""
                    SELECT id FROM ba_rlb_legal_information 
                    WHERE legal_id = %s AND company_name = %s 
                    AND (delete_time IS NULL OR delete_time = 0)
                """, (customer_id, company_name))
                legal_info_result = cursor.fetchone()
                
                if legal_info_result:
                    # 找到了，使用ba_rlb_legal_information的记录
                    information_id = legal_info_result[0]
                    self.log_message.emit(f"第{row_num}行: 在ba_rlb_legal_information找到记录ID={information_id}")
                    
                    # 查找对应的shopindex (belong_information = 1)
                    cursor.execute("""
                        SELECT id FROM ba_shopindex 
                        WHERE legal_id = %s AND belong_information = 1 
                        AND (delete_time IS NULL OR delete_time = 0)
                    """, (information_id,))
                    shopindex_result = cursor.fetchone()
                    
                    if not shopindex_result:
                        raise Exception(f"未找到日区企业法人shopindex记录 (information_id: {information_id})")
                    
                    return shopindex_result[0]
                
                else:
                    # ba_rlb_legal_information中没找到，再查ba_rlb_legal_information_part2
                    self.log_message.emit(f"第{row_num}行: ba_rlb_legal_information未找到，查找ba_rlb_legal_information_part2表")
                    
                    cursor.execute("""
                        SELECT id FROM ba_rlb_legal_information_part2 
                        WHERE legal_id = %s AND company_name = %s 
                        AND (delete_time IS NULL OR delete_time = 0)
                    """, (customer_id, company_name))
                    legal_info_part2_result = cursor.fetchone()
                    
                    if not legal_info_part2_result:
                        raise Exception(f"在ba_rlb_legal_information和ba_rlb_legal_information_part2都未找到企业法人信息 (customer_id: {customer_id}, company: {company_name})")
                    
                    information_part2_id = legal_info_part2_result[0]
                    self.log_message.emit(f"第{row_num}行: 在ba_rlb_legal_information_part2找到记录ID={information_part2_id}")
                    
                    # 查找对应的shopindex (belong_information = 2)
                    cursor.execute("""
                        SELECT id FROM ba_shopindex 
                        WHERE legal_id = %s AND belong_information = 2 
                        AND (delete_time IS NULL OR delete_time = 0)
                    """, (information_part2_id,))
                    shopindex_result = cursor.fetchone()
                    
                    if not shopindex_result:
                        raise Exception(f"未找到欧美企业法人shopindex记录 (information_part2_id: {information_part2_id})")
                    
                    return shopindex_result[0]
                    
        except Exception as e:
            raise Exception(f"查找shopindex失败: {str(e)}")
    
    def find_shudi_id(self, cursor, territory_abbreviation, row_num):
        """根据属地简称查找shudi_id"""
        try:
            cursor.execute("""
                SELECT id FROM ba_shudi 
                WHERE territory_abbreviation = %s AND status = 1
            """, (territory_abbreviation,))
            shudi_result = cursor.fetchone()
            
            if not shudi_result:
                raise Exception(f"未找到属地: {territory_abbreviation}")
            
            shudi_id = shudi_result[0]
            self.log_message.emit(f"第{row_num}行: 找到属地ID={shudi_id}")
            return shudi_id
            
        except Exception as e:
            raise Exception(f"查找属地失败: {str(e)}")
    
    def find_platform_id(self, cursor, platform_name, row_num):
        """根据平台标识查找platform_id"""
        try:
            cursor.execute("""
                SELECT id FROM ba_platform 
                WHERE platform = %s AND status = 1 
                AND (delete_time IS NULL OR delete_time = 0)
            """, (platform_name,))
            platform_result = cursor.fetchone()
            
            if not platform_result:
                raise Exception(f"未找到平台: {platform_name}")
            
            platform_id = platform_result[0]
            self.log_message.emit(f"第{row_num}行: 找到平台ID={platform_id}")
            return platform_id
            
        except Exception as e:
            raise Exception(f"查找平台失败: {str(e)}")
    
    def update_platform_info(self, cursor, shopindex_id, shudi_id, platform_id, row_num):
        """查找并更新platform_info状态"""
        try:
            # 查找platform_info记录
            cursor.execute("""
                SELECT id, status FROM ba_platform_info 
                WHERE shopindex_id = %s AND shudi_id = %s AND platform_id = %s 
                AND (delete_time IS NULL OR delete_time = 0)
            """, (shopindex_id, shudi_id, platform_id))
            platform_info_result = cursor.fetchone()
            
            if not platform_info_result:
                raise Exception(f"未找到platform_info记录 (shopindex_id: {shopindex_id}, shudi_id: {shudi_id}, platform_id: {platform_id})")
            
            platform_info_id, current_status = platform_info_result
            self.log_message.emit(f"第{row_num}行: 找到platform_info记录ID={platform_info_id}, 当前状态={current_status}")
            
            # 更新状态
            current_time = int(datetime.now().timestamp())
            cursor.execute("""
                UPDATE ba_platform_info 
                SET status = %s, update_time = %s 
                WHERE id = %s
            """, (self.target_status, current_time, platform_info_id))
            
            self.log_message.emit(f"第{row_num}行: 成功更新状态 {current_status} → {self.target_status}")
            return platform_info_id
            
        except Exception as e:
            raise Exception(f"更新platform_info失败: {str(e)}")


class PlatformInfoStatusWidget(QWidget):
    """平台详情状态更新界面"""
    
    def __init__(self, db_manager):
        super().__init__()
        self.db_manager = db_manager
        self.worker = None
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
        file_group = QGroupBox("选择Excel文件")
        file_layout = QHBoxLayout()
        
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setPlaceholderText("请选择包含平台详情数据的Excel文件")
        file_layout.addWidget(self.file_path_edit)
        
        self.browse_btn = QPushButton("浏览")
        self.browse_btn.clicked.connect(self.browse_file)
        file_layout.addWidget(self.browse_btn)
        
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # 状态设置区域
        status_group = QGroupBox("目标状态设置")
        status_layout = QFormLayout()
        
        self.status_combo = QComboBox()
        self.status_combo.addItem("失败 (0)", 0)
        self.status_combo.addItem("可用 (1)", 1)
        self.status_combo.addItem("已用 (2)", 2)
        self.status_combo.addItem("取消 (3)", 3)
        status_layout.addRow("更新状态:", self.status_combo)
        
        status_group.setLayout(status_layout)
        layout.addWidget(status_group)
        
        # 功能说明区域
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        
        info_label = QLabel("""
<b>验证后的非平台注册部数据平台详情补全说明:</b><br>
<b>Excel列要求:</b><br>
• A列: 平台标识 (用于查找ba_platform表)<br>
• B列: 法人姓名 (用于查找ba_rlb_customer表)<br>
• H列: 属性判断 (企业法人/其他)<br>
• I列: 公司名称 (企业法人时使用)<br>
• J列: 属地简称 (用于查找ba_shudi表)<br><br>
<b>查找规则:</b><br>
• <b>非企业法人</b>: B列 → ba_rlb_customer → ba_shopindex(belong_information=null)<br>
• <b>企业法人</b>: B列+I列 → 优先查找ba_rlb_legal_information → 找不到再查ba_rlb_legal_information_part2<br>
  &nbsp;&nbsp;- 找到ba_rlb_legal_information → ba_shopindex(belong_information=1)<br>
  &nbsp;&nbsp;- 找到ba_rlb_legal_information_part2 → ba_shopindex(belong_information=2)<br>
• <b>属地查找</b>: J列 → ba_shudi.territory_abbreviation<br>
• <b>平台查找</b>: A列 → ba_platform.platform<br><br>
<b>更新操作:</b><br>
• 根据shopindex_id、shudi_id、platform_id查找ba_platform_info记录<br>
• 将找到的记录状态更新为指定值<br>
• 记录更新前后的状态变化
        """)
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # 操作按钮区域
        button_group = QGroupBox("操作")
        button_layout = QHBoxLayout()
        
        self.start_btn = QPushButton("开始更新")
        self.start_btn.clicked.connect(self.start_update)
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
        
        button_group.setLayout(button_layout)
        layout.addWidget(button_group)
        
        # 进度条
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)
        
        # 结果显示
        self.result_label = QLabel("等待开始...")
        layout.addWidget(self.result_label)
        
        # 日志区域
        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout()
        
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(200)
        log_layout.addWidget(self.log_text)
        
        clear_log_btn = QPushButton("清空日志")
        clear_log_btn.clicked.connect(self.clear_log)
        log_layout.addWidget(clear_log_btn)
        
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)
        
        self.setLayout(layout)
    
    def load_datasources(self):
        """加载数据源列表"""
        self.datasource_combo.clear()
        self.datasource_combo.addItem("请选择数据源", None)
        
        datasources = self.db_manager.get_all_datasources()
        for ds in datasources:
            self.datasource_combo.addItem(f"{ds.name} ({ds.host}:{ds.port})", ds)
    
    def on_datasource_changed(self):
        """数据源改变事件"""
        current_data = self.datasource_combo.currentData()
        self.test_connection_btn.setEnabled(current_data is not None)
        self.update_start_button_state()
    
    def test_connection(self):
        """测试数据库连接"""
        datasource = self.datasource_combo.currentData()
        if not datasource:
            return
        
        try:
            connection = pymysql.connect(
                host=datasource.host,
                port=datasource.port,
                user=datasource.username,
                password=datasource.password,
                database=datasource.database,
                charset=datasource.charset,
                connect_timeout=5
            )
            connection.close()
            QMessageBox.information(self, "连接成功", "数据库连接测试成功！")
        except Exception as e:
            QMessageBox.critical(self, "连接失败", f"数据库连接失败：{str(e)}")
    
    def browse_file(self):
        """浏览Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Excel文件",
            "",
            "Excel文件 (*.xlsx *.xls)"
        )
        
        if file_path:
            self.file_path_edit.setText(file_path)
            self.update_start_button_state()
    
    def update_start_button_state(self):
        """更新开始按钮状态"""
        datasource = self.datasource_combo.currentData()
        file_path = self.file_path_edit.text().strip()
        
        self.start_btn.setEnabled(
            datasource is not None and 
            file_path != "" and 
            (self.worker is None or not self.worker.isRunning())
        )
    
    def start_update(self):
        """开始更新"""
        datasource = self.datasource_combo.currentData()
        file_path = self.file_path_edit.text().strip()
        target_status = self.status_combo.currentData()
        
        if not datasource or not file_path:
            QMessageBox.warning(self, "参数错误", "请选择数据源和Excel文件")
            return
        
        # 清空之前的结果
        self.completed_data = []
        self.failed_records = []
        self.log_text.clear()
        self.progress_bar.setValue(0)
        self.result_label.setText("正在处理...")
        self.result_label.setStyleSheet("color: blue;")
        
        # 禁用按钮
        self.start_btn.setEnabled(False)
        self.export_completed_btn.setEnabled(False)
        self.export_failed_btn.setEnabled(False)
        
        # 启动工作线程
        self.worker = PlatformInfoStatusWorker(datasource, file_path, target_status)
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log_message.connect(self.append_log)
        self.worker.finished.connect(self.on_update_finished)
        self.worker.error.connect(self.on_update_error)
        self.worker.start()
    
    def append_log(self, message):
        """添加日志"""
        timestamp = datetime.now().strftime('[%H:%M:%S]')
        self.log_text.append(f"{timestamp} {message}")
        
        # 自动滚动到底部
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
    
    def on_update_finished(self, results):
        """更新完成"""
        self.completed_data = results['completed_data']
        self.failed_records = results['failed_records']
        
        success_count = results['success_count']
        failed_count = results['failed_count']
        total_count = success_count + failed_count
        
        result_text = f"更新完成！\n成功: {success_count} 条\n失败: {failed_count} 条\n总计: {total_count} 条"
        self.result_label.setText(result_text)
        
        if failed_count > 0:
            self.result_label.setStyleSheet("color: orange;")
            self.export_failed_btn.setEnabled(True)
        else:
            self.result_label.setStyleSheet("color: green;")
        
        # 启用导出完整结果按钮
        self.export_completed_btn.setEnabled(True)
        
        # 重新启用开始按钮
        self.update_start_button_state()
        
        QMessageBox.information(self, "更新完成", result_text)
    
    def on_update_error(self, error_msg):
        """更新错误"""
        self.result_label.setText(f"更新失败: {error_msg}")
        self.result_label.setStyleSheet("color: red;")
        self.update_start_button_state()
        QMessageBox.critical(self, "更新错误", f"更新失败：{error_msg}")
    
    def export_completed_data(self):
        """导出完整结果"""
        if not self.completed_data:
            QMessageBox.information(self, "提示", "没有完整结果可导出")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存完整结果",
            f"platform_info_status_completed_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel文件 (*.xlsx)"
        )
        
        if file_path:
            try:
                # 创建DataFrame并导出
                df = pd.DataFrame(self.completed_data)
                df.to_excel(file_path, index=False, engine='openpyxl')
                
                QMessageBox.information(self, "导出成功", f"完整结果已导出到：\n{file_path}\n\n包含字段：\n• 原始Excel所有列\n• shopindex_id、shudi_id、platform_id\n• platform_info_id\n• 更新状态\n• 处理状态\n• 处理时间\n• 失败原因（如有）")
                
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
            f"platform_info_status_failed_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel文件 (*.xlsx)"
        )
        
        if file_path:
            try:
                # 准备导出数据
                export_data = []
                for record in self.failed_records:
                    row_data = record['data'].copy()
                    row_data['行号'] = record['row']
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
