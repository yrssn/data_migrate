"""
平台注册部平台详情预先处理功能界面
找到ba_ptzcb_register表中备注是"系统脚本自动录入"的数据，
根据shopindex_id, shudi_id, platform_id去ba_platform_info表中查找或插入，并将status改为2
"""
import os
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog,
                             QTableWidget, QTableWidgetItem, QHeaderView, QCheckBox,
                             QLineEdit)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from database import DatabaseManager
import pandas as pd
import pymysql


class PlatformInfoPreprocessWorker(QThread):
    """平台详情预处理工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    
    def __init__(self, datasource, db_manager, selected_records, remark_keyword):
        super().__init__()
        self.datasource = datasource
        self.db_manager = db_manager
        self.selected_records = selected_records
        self.remark_keyword = remark_keyword
        self.results = {
            'total_selected': len(selected_records),
            'updated_count': 0,
            'inserted_count': 0,
            'skipped_count': 0,
            'failed_count': 0,
            'failed_records': []
        }
    
    def run(self):
        try:
            self.log_message.emit("开始连接数据库...")
            self.progress.emit(5)
            
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
            self.progress.emit(10)
            
            total = len(self.selected_records)
            for i, record in enumerate(self.selected_records):
                try:
                    self.process_record(cursor, record)
                    
                    progress = 10 + int((i + 1) / total * 85)
                    self.progress.emit(progress)
                    
                except Exception as e:
                    self.results['failed_count'] += 1
                    error_msg = f"记录ID {record['id']} 处理失败: {str(e)}"
                    self.log_message.emit(error_msg)
                    self.results['failed_records'].append({
                        'id': record['id'],
                        'project_id': record.get('project_id', ''),
                        'error': str(e)
                    })
            
            connection.commit()
            cursor.close()
            connection.close()
            
            self.progress.emit(100)
            self.finished.emit(self.results)
            
        except Exception as e:
            self.error.emit(str(e))
    
    def process_record(self, cursor, record):
        """处理单条记录"""
        register_id = record['id']
        shopindex_id = record['shopindex_id']
        shudi_id = record['shudi_id']
        platform_id = record['platform_id']
        register_status = record.get('status', 1)
        
        self.log_message.emit(f"处理记录 ID={register_id}, shopindex_id={shopindex_id}, shudi_id={shudi_id}, platform_id={platform_id}, status={register_status}")
        
        # 检查必要字段
        if not shopindex_id:
            self.results['skipped_count'] += 1
            self.log_message.emit(f"  记录ID {register_id}: shopindex_id为空，跳过")
            return
        
        if not shudi_id:
            self.results['skipped_count'] += 1
            self.log_message.emit(f"  记录ID {register_id}: shudi_id为空，跳过")
            return
        
        if not platform_id:
            self.results['skipped_count'] += 1
            self.log_message.emit(f"  记录ID {register_id}: platform_id为空，跳过")
            return
        
        current_time = int(datetime.now().timestamp())
        
        # 根据ba_ptzcb_register.status决定写入ba_platform_info的status值
        # ba_ptzcb_register.status: 0=待指派,1=进行中,2=审核中,3=银行验证,4=审核失败,5=银行卡验证失败,
        #                           6=完成,7=失败,8=取消,9=暂停,10=店铺验证失败,11=店铺验证中
        # ba_platform_info.status: 0=失效,1=可用,2=已用,3=取消
        if register_status == 8:  # 取消 -> 3
            target_status = 3
        elif register_status == 7:  # 失败 -> 0
            target_status = 0
        else:
            target_status = 2  # 其他情况默认已用
        
        # 在ba_platform_info中查找是否存在对应记录
        cursor.execute("""
            SELECT id, status 
            FROM ba_platform_info 
            WHERE shopindex_id = %s AND shudi_id = %s AND platform_id = %s
              AND (delete_time IS NULL OR delete_time = 0)
        """, (shopindex_id, shudi_id, platform_id))
        
        platform_info = cursor.fetchone()
        
        if platform_info:
            # 找到了，更新status
            if platform_info['status'] == target_status:
                self.log_message.emit(f"  记录ID {register_id}: ba_platform_info中已存在且status已经是{target_status}，跳过")
                self.results['skipped_count'] += 1
            else:
                cursor.execute("""
                    UPDATE ba_platform_info 
                    SET status = %s, update_time = %s 
                    WHERE id = %s
                """, (target_status, current_time, platform_info['id']))
                self.results['updated_count'] += 1
                self.log_message.emit(f"  记录ID {register_id}: 更新ba_platform_info.id={platform_info['id']} status从{platform_info['status']}改为{target_status}")
        else:
            # 没找到，插入新记录
            cursor.execute("""
                INSERT INTO ba_platform_info 
                (shopindex_id, shudi_id, platform_id, status, create_time, update_time, admin_id, admin_dept_id)
                VALUES (%s, %s, %s, %s, %s, %s, 1, 1)
            """, (shopindex_id, shudi_id, platform_id, target_status, current_time, current_time))
            
            new_id = cursor.lastrowid
            self.results['inserted_count'] += 1
            self.log_message.emit(f"  记录ID {register_id}: 插入新记录到ba_platform_info，新ID={new_id}，status={target_status}")


class PlatformInfoPreprocessWidget(QWidget):
    """平台注册部平台详情预先处理功能组件"""
    
    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.loaded_records = []
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
        
        # 查询条件配置
        condition_group = QGroupBox("查询条件")
        condition_layout = QFormLayout()
        
        self.remark_edit = QLineEdit()
        self.remark_edit.setText("系统脚本自动录入")
        self.remark_edit.setPlaceholderText("输入备注关键词")
        condition_layout.addRow("备注包含:", self.remark_edit)
        
        condition_group.setLayout(condition_layout)
        layout.addWidget(condition_group)
        
        # 功能说明区域
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        
        info_label = QLabel("""
<b>平台注册部平台详情预先处理说明:</b><br>
<b>处理逻辑:</b><br>
• 1. 从ba_ptzcb_register表中查找备注(remark)包含指定关键词的记录<br>
• 2. 获取每条记录的 shopindex_id, shudi_id, platform_id, status<br>
• 3. 根据ba_ptzcb_register.status映射ba_platform_info.status:<br>
&nbsp;&nbsp;- status=8(取消) → 3(取消)<br>
&nbsp;&nbsp;- status=7(失败) → 0(失效)<br>
&nbsp;&nbsp;- 其他 → 2(已用)<br>
• 4. 在ba_platform_info表中查找是否存在相同组合的记录:<br>
&nbsp;&nbsp;- 如果存在: 更新status<br>
&nbsp;&nbsp;- 如果不存在: 插入新记录<br><br>
<b>ba_platform_info.status状态说明:</b><br>
• 0=失效, 1=可用, 2=已用, 3=取消
        """)
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # 数据预览区域
        preview_group = QGroupBox("数据预览")
        preview_layout = QVBoxLayout()
        
        preview_button_layout = QHBoxLayout()
        self.load_data_btn = QPushButton("查询数据")
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
        
        self.record_count_label = QLabel("")
        preview_button_layout.addWidget(self.record_count_label)
        
        preview_button_layout.addStretch()
        preview_layout.addLayout(preview_button_layout)
        
        self.records_table = QTableWidget()
        self.records_table.setMaximumHeight(250)
        preview_layout.addWidget(self.records_table)
        
        preview_group.setLayout(preview_layout)
        layout.addWidget(preview_group)
        
        # 操作按钮区域
        button_layout = QHBoxLayout()
        
        self.start_btn = QPushButton("开始处理")
        self.start_btn.clicked.connect(self.start_process)
        self.start_btn.setEnabled(False)
        button_layout.addWidget(self.start_btn)
        
        button_layout.addStretch()
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        button_layout.addWidget(self.progress_bar)
        
        layout.addLayout(button_layout)
        
        # 日志区域
        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout()
        
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(180)
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        
        log_button_layout = QHBoxLayout()
        self.clear_log_btn = QPushButton("清空日志")
        self.clear_log_btn.clicked.connect(self.clear_log)
        log_button_layout.addWidget(self.clear_log_btn)
        
        self.export_log_btn = QPushButton("导出日志")
        self.export_log_btn.clicked.connect(self.export_log)
        log_button_layout.addWidget(self.export_log_btn)
        
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
        self.load_data_btn.setEnabled(datasource is not None)
    
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
    
    def load_data(self):
        """加载数据预览"""
        datasource = self.datasource_combo.currentData()
        if not datasource:
            QMessageBox.warning(self, "警告", "请选择数据源！")
            return
        
        remark_keyword = self.remark_edit.text().strip()
        if not remark_keyword:
            QMessageBox.warning(self, "警告", "请输入备注关键词！")
            return
        
        try:
            connection = pymysql.connect(
                host=datasource.host,
                port=datasource.port,
                user=datasource.username,
                password=datasource.password,
                database=datasource.database,
                charset=datasource.charset
            )
            cursor = connection.cursor(pymysql.cursors.DictCursor)
            
            # 查询备注包含关键词的记录
            query = """
                SELECT id, project_id, shopindex_id, shudi_id, platform_id, 
                       status, remark, company_name, email
                FROM ba_ptzcb_register 
                WHERE remark LIKE %s
                  AND (delete_time IS NULL OR delete_time = 0)
                ORDER BY id DESC
            """
            
            cursor.execute(query, (f"%{remark_keyword}%",))
            records = cursor.fetchall()
            
            cursor.close()
            connection.close()
            
            self.loaded_records = records
            self.display_records(records)
            
            self.record_count_label.setText(f"共 {len(records)} 条记录")
            
            if records:
                self.select_all_btn.setEnabled(True)
                self.deselect_all_btn.setEnabled(True)
                self.start_btn.setEnabled(True)
            else:
                self.select_all_btn.setEnabled(False)
                self.deselect_all_btn.setEnabled(False)
                self.start_btn.setEnabled(False)
            
        except Exception as e:
            QMessageBox.critical(self, "查询失败", f"查询数据失败：{str(e)}")
    
    def display_records(self, records):
        """显示记录到表格"""
        headers = ["选择", "ID", "项目编号", "主体编号ID", "属地ID", "平台ID", "状态", "备注", "公司名称"]
        self.records_table.setRowCount(len(records))
        self.records_table.setColumnCount(len(headers))
        self.records_table.setHorizontalHeaderLabels(headers)
        
        status_map = {
            0: '待指派', 1: '进行中', 2: '审核中', 3: '银行验证',
            4: '审核失败', 5: '银行卡验证失败', 6: '完成', 7: '失败',
            8: '取消', 9: '暂停', 10: '店铺验证失败', 11: '店铺验证中'
        }
        
        for row_idx, record in enumerate(records):
            # 检查是否有完整的ID组合
            has_complete_ids = (
                record['shopindex_id'] is not None and
                record['shudi_id'] is not None and
                record['platform_id'] is not None
            )
            
            checkbox = QCheckBox()
            checkbox.setChecked(has_complete_ids)
            checkbox.setEnabled(has_complete_ids)
            self.records_table.setCellWidget(row_idx, 0, checkbox)
            
            self.records_table.setItem(row_idx, 1, QTableWidgetItem(str(record['id'])))
            self.records_table.setItem(row_idx, 2, QTableWidgetItem(str(record['project_id'] or '')))
            self.records_table.setItem(row_idx, 3, QTableWidgetItem(str(record['shopindex_id'] or '空')))
            self.records_table.setItem(row_idx, 4, QTableWidgetItem(str(record['shudi_id'] or '空')))
            self.records_table.setItem(row_idx, 5, QTableWidgetItem(str(record['platform_id'] or '空')))
            
            status_code = record['status']
            status_text = f"{status_code}-{status_map.get(status_code, '未知')}"
            self.records_table.setItem(row_idx, 6, QTableWidgetItem(status_text))
            
            remark = str(record['remark'] or '')
            if len(remark) > 20:
                remark = remark[:20] + "..."
            self.records_table.setItem(row_idx, 7, QTableWidgetItem(remark))
            
            self.records_table.setItem(row_idx, 8, QTableWidgetItem(str(record['company_name'] or '')))
        
        # 调整列宽
        header = self.records_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(6, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(7, QHeaderView.Stretch)
        header.setSectionResizeMode(8, QHeaderView.Stretch)
    
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
        """获取选中的记录"""
        selected = []
        for row in range(self.records_table.rowCount()):
            checkbox = self.records_table.cellWidget(row, 0)
            if checkbox and checkbox.isChecked():
                selected.append(self.loaded_records[row])
        return selected
    
    def start_process(self):
        """开始处理"""
        datasource = self.datasource_combo.currentData()
        if not datasource:
            QMessageBox.warning(self, "警告", "请选择数据源！")
            return
        
        selected_records = self.get_selected_records()
        if not selected_records:
            QMessageBox.warning(self, "警告", "请选择要处理的记录！")
            return
        
        remark_keyword = self.remark_edit.text().strip()
        
        reply = QMessageBox.question(
            self,
            "确认处理",
            f"确定要处理平台详情吗？\n\n"
            f"数据库: {datasource.name}\n"
            f"备注关键词: {remark_keyword}\n"
            f"选中记录: {len(selected_records)} 条\n\n"
            f"处理逻辑:\n"
            f"• 在ba_platform_info中查找相同组合\n"
            f"• 找到则更新status为2(已用)\n"
            f"• 未找到则插入新记录status=2",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        self.start_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在处理中...")
        self.log_text.clear()
        
        self.worker = PlatformInfoPreprocessWorker(
            datasource, self.db_manager, selected_records, remark_keyword
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
        
        result_text = (
            f"处理完成！\n"
            f"选中记录：{results['total_selected']} 条\n"
            f"更新记录：{results['updated_count']} 条\n"
            f"插入记录：{results['inserted_count']} 条\n"
            f"跳过记录：{results['skipped_count']} 条\n"
            f"处理失败：{results['failed_count']} 条"
        )
        
        self.result_label.setText(result_text)
        if results['failed_count'] > 0:
            self.result_label.setStyleSheet("color: orange;")
        else:
            self.result_label.setStyleSheet("color: green;")
        
        QMessageBox.information(self, "处理完成", result_text)
    
    def on_process_error(self, error_msg):
        """处理错误"""
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.result_label.setText("处理失败！")
        self.result_label.setStyleSheet("color: red;")
        
        QMessageBox.critical(self, "处理错误", f"处理失败：{error_msg}")
    
    def clear_log(self):
        """清空日志"""
        self.log_text.clear()
    
    def export_log(self):
        """导出日志"""
        if not self.log_text.toPlainText():
            QMessageBox.information(self, "提示", "没有日志可以导出")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存日志文件",
            f"平台详情预处理日志_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            "文本文件 (*.txt);;所有文件 (*)"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.log_text.toPlainText())
                QMessageBox.information(self, "成功", "日志导出成功！")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出失败：{str(e)}")
