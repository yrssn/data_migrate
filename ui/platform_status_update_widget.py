"""
修改平台详情状态-平台注册部版功能界面
"""
import os
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog,
                             QSpinBox, QCheckBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from database import DatabaseManager
import pandas as pd
import pymysql


class PlatformStatusUpdateWorker(QThread):
    """平台状态更新工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)  # 返回统计结果
    error = pyqtSignal(str)
    
    def __init__(self, datasource, db_manager, excel_file, target_status):
        super().__init__()
        self.datasource = datasource
        self.db_manager = db_manager
        self.excel_file = excel_file
        self.target_status = target_status
        self.results = {
            'total_rows': 0,
            'success_count': 0,
            'failed_count': 0,
            'skipped_count': 0,
            'failed_records': []  # 存储失败的记录详情
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
            
            # 处理每一行数据
            for index, row in df.iterrows():
                try:
                    self.process_row(cursor, index + 1, row)
                    
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
            legal_name = str(row.iloc[1]) if len(row) > 1 else ""  # B列
            platform_name = str(row.iloc[3]) if len(row) > 3 else ""  # D列
            company_name = str(row.iloc[4]) if len(row) > 4 else ""  # E列
            is_corporate = str(row.iloc[6]) if len(row) > 6 else ""  # G列
            territory = str(row.iloc[7]) if len(row) > 7 else ""  # H列
            
            if not legal_name or legal_name == 'nan':
                raise Exception("B列(法人姓名)为空")
            if not platform_name or platform_name == 'nan':
                raise Exception("D列(平台名称)为空")
            if not territory or territory == 'nan':
                raise Exception("H列(属地)为空")
            
            self.log_message.emit(f"第{row_num}行: 处理法人={legal_name}, 平台={platform_name}, 属地={territory}")
            
            # 第一步：根据G列判断类型，获取shopindex_id
            shopindex_id = self.get_shopindex_id(cursor, legal_name, company_name, is_corporate, territory)
            if not shopindex_id:
                raise Exception("未找到对应的shopindex记录")
            
            # 第二步：根据D列获取platform_id
            platform_id = self.get_platform_id(cursor, platform_name)
            if not platform_id:
                raise Exception(f"未找到平台: {platform_name}")
            
            # 第三步：根据H列获取shudi_id
            shudi_id = self.get_shudi_id(cursor, territory)
            if not shudi_id:
                raise Exception(f"未找到属地: {territory}")
            
            # 第四步：查找并更新platform_info状态
            self.update_platform_info_status(cursor, shopindex_id, platform_id, shudi_id, row_num)
            
            self.results['success_count'] += 1
            self.log_message.emit(f"第{row_num}行: 成功更新状态为{self.target_status}")
            
        except Exception as e:
            self.results['failed_count'] += 1
            raise e
    
    def get_shopindex_id(self, cursor, legal_name, company_name, is_corporate, territory):
        """根据复杂逻辑获取shopindex_id"""
        try:
            # 首先根据legal_name找到customer_id
            cursor.execute("""
                SELECT id FROM ba_rlb_customer 
                WHERE legal_name = %s AND (delete_time IS NULL OR delete_time = 0)
            """, (legal_name,))
            customer_result = cursor.fetchone()
            
            if not customer_result:
                raise Exception(f"未找到法人: {legal_name}")
            
            customer_id = customer_result[0]
            
            # 判断是否为企业法人
            if is_corporate == "企业法人":
                # 企业法人逻辑
                if territory == "JP":
                    # JP: 在ba_rlb_legal_information中查找
                    cursor.execute("""
                        SELECT id FROM ba_rlb_legal_information 
                        WHERE company_name = %s AND legal_id = %s AND (delete_time IS NULL OR delete_time = 0)
                    """, (company_name, customer_id))
                    legal_result = cursor.fetchone()
                    
                    if not legal_result:
                        raise Exception(f"未找到JP法人信息: {company_name}")
                    
                    legal_info_id = legal_result[0]
                    
                    # 在shopindex中查找belong_information=1的记录
                    cursor.execute("""
                        SELECT id FROM ba_shopindex 
                        WHERE legal_id = %s AND belong_information = 1 AND status = 1 AND (delete_time IS NULL OR delete_time = 0)
                    """, (legal_info_id,))
                    
                else:
                    # 非JP: 在ba_rlb_legal_information_part2中查找
                    cursor.execute("""
                        SELECT id FROM ba_rlb_legal_information_part2 
                        WHERE company_name = %s AND legal_id = %s AND (delete_time IS NULL OR delete_time = 0)
                    """, (company_name, customer_id))
                    legal_result = cursor.fetchone()
                    
                    if not legal_result:
                        raise Exception(f"未找到非JP法人信息: {company_name}")
                    
                    legal_info_id = legal_result[0]
                    
                    # 在shopindex中查找belong_information=2的记录
                    cursor.execute("""
                        SELECT id FROM ba_shopindex 
                        WHERE legal_id = %s AND belong_information = 2 AND status = 1 AND (delete_time IS NULL OR delete_time = 0)
                    """, (legal_info_id,))
                    
            else:
                # 个人客户逻辑：belong_information为null
                cursor.execute("""
                    SELECT id FROM ba_shopindex 
                    WHERE legal_id = %s AND (belong_information IS NULL OR belong_information = '') 
                    AND status = 1 AND (delete_time IS NULL OR delete_time = 0)
                """, (customer_id,))
            
            shopindex_result = cursor.fetchone()
            if not shopindex_result:
                raise Exception("未找到对应的shopindex记录")
            
            return shopindex_result[0]
            
        except Exception as e:
            raise Exception(f"获取shopindex_id失败: {str(e)}")
    
    def get_platform_id(self, cursor, platform_name):
        """根据平台名称获取platform_id"""
        cursor.execute("""
            SELECT id FROM ba_platform 
            WHERE platform = %s AND status = 1 AND (delete_time IS NULL OR delete_time = 0)
        """, (platform_name,))
        result = cursor.fetchone()
        return result[0] if result else None
    
    def get_shudi_id(self, cursor, territory):
        """根据属地简称获取shudi_id"""
        cursor.execute("""
            SELECT id FROM ba_shudi 
            WHERE territory_abbreviation = %s AND status = 1
        """, (territory,))
        result = cursor.fetchone()
        return result[0] if result else None
    
    def update_platform_info_status(self, cursor, shopindex_id, platform_id, shudi_id, row_num):
        """更新platform_info状态"""
        # 查找对应的platform_info记录
        cursor.execute("""
            SELECT id, status FROM ba_platform_info 
            WHERE shopindex_id = %s AND platform_id = %s AND shudi_id = %s 
            AND (delete_time IS NULL OR delete_time = 0)
        """, (shopindex_id, platform_id, shudi_id))
        
        result = cursor.fetchone()
        if not result:
            raise Exception("未找到对应的platform_info记录")
        
        platform_info_id, current_status = result
        
        # 更新状态
        current_time = int(datetime.now().timestamp())
        cursor.execute("""
            UPDATE ba_platform_info 
            SET status = %s, update_time = %s 
            WHERE id = %s
        """, (self.target_status, current_time, platform_info_id))
        
        self.log_message.emit(f"第{row_num}行: platform_info ID={platform_info_id}, 状态从{current_status}更新为{self.target_status}")


class PlatformStatusUpdateWidget(QWidget):
    """修改平台详情状态功能组件"""
    
    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.excel_file = None
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
        file_group = QGroupBox("Excel文件和状态设置")
        file_layout = QFormLayout()
        
        file_select_layout = QHBoxLayout()
        self.file_label = QLabel("未选择文件")
        self.file_label.setStyleSheet("color: gray;")
        file_select_layout.addWidget(self.file_label)
        
        self.select_file_btn = QPushButton("选择Excel文件")
        self.select_file_btn.clicked.connect(self.select_excel_file)
        file_select_layout.addWidget(self.select_file_btn)
        
        file_layout.addRow("Excel文件:", file_select_layout)
        
        # 目标状态设置
        self.status_spinbox = QSpinBox()
        self.status_spinbox.setRange(0, 10)
        self.status_spinbox.setValue(2)
        file_layout.addRow("目标状态:", self.status_spinbox)
        
        # 状态说明
        status_info = QLabel("状态说明: 0=失败, 1=可用, 2=已用, 3=取消")
        status_info.setStyleSheet("color: blue; font-size: 12px;")
        file_layout.addRow("", status_info)
        
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # 功能说明区域
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        
        info_label = QLabel("""
<b>修改平台详情状态-平台注册部版说明:</b><br>
<b>Excel列对应关系:</b><br>
• B列: 法人姓名 (用于查找ba_rlb_customer)<br>
• D列: 平台名称 (用于查找ba_platform)<br>
• E列: 公司名称 (企业法人时使用)<br>
• G列: 是否企业法人 (判断查找逻辑)<br>
• H列: 属地简称 (用于查找ba_shudi)<br><br>
<b>查找逻辑:</b><br>
• <b>个人客户</b>: B列→ba_rlb_customer→ba_shopindex(belong_information=null)<br>
• <b>企业法人(JP)</b>: B列→ba_rlb_customer, E列→ba_rlb_legal_information→ba_shopindex(belong_information=1)<br>
• <b>企业法人(非JP)</b>: B列→ba_rlb_customer, E列→ba_rlb_legal_information_part2→ba_shopindex(belong_information=2)<br><br>
<b>更新流程:</b><br>
• 根据shopindex_id + platform_id + shudi_id定位ba_platform_info记录<br>
• 将status字段更新为指定值<br>
• 详细日志记录处理过程和结果<br>
• 支持导出失败记录到Excel
        """)
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # 操作按钮区域
        button_layout = QHBoxLayout()
        
        self.start_btn = QPushButton("开始更新")
        self.start_btn.clicked.connect(self.start_update)
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
    
    def start_update(self):
        """开始更新"""
        datasource = self.datasource_combo.currentData()
        if not datasource or not self.excel_file:
            QMessageBox.warning(self, "警告", "请选择数据源和Excel文件！")
            return
        
        target_status = self.status_spinbox.value()
        
        # 确认对话框
        reply = QMessageBox.question(
            self,
            "确认更新",
            f"确定要更新平台详情状态吗？\n\n"
            f"数据库: {datasource.name}\n"
            f"Excel文件: {os.path.basename(self.excel_file)}\n"
            f"目标状态: {target_status}\n\n"
            f"此操作将根据Excel数据批量更新ba_platform_info表的状态字段。\n"
            f"请确保Excel格式正确，操作不可撤销！",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        # 禁用按钮，显示进度条
        self.start_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在更新中...")
        self.log_text.clear()
        self.failed_records = []
        
        # 启动工作线程
        self.worker = PlatformStatusUpdateWorker(datasource, self.db_manager, self.excel_file, target_status)
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
        self.failed_records = results['failed_records']
        
        result_text = (
            f"更新完成！\n"
            f"总行数：{results['total_rows']} 行\n"
            f"成功：{results['success_count']} 行\n"
            f"失败：{results['failed_count']} 行\n"
            f"跳过：{results['skipped_count']} 行"
        )
        
        self.result_label.setText(result_text)
        if results['failed_count'] > 0:
            self.result_label.setStyleSheet("color: orange;")
            self.export_failed_btn.setEnabled(True)
        else:
            self.result_label.setStyleSheet("color: green;")
        
        QMessageBox.information(self, "更新完成", result_text)
    
    def on_update_error(self, error_msg):
        """更新错误"""
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.result_label.setText("更新失败！")
        self.result_label.setStyleSheet("color: red;")
        
        QMessageBox.critical(self, "更新错误", f"更新失败：{error_msg}")
    
    def export_failed_records(self):
        """导出失败记录"""
        if not self.failed_records:
            QMessageBox.information(self, "提示", "没有失败记录可导出")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存失败记录",
            f"failed_records_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
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
