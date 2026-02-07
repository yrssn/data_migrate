"""
按库存主键更新状态功能界面
"""
import os
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog,
                             QTableWidget, QTableWidgetItem, QHeaderView)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from database import DatabaseManager
import pandas as pd
import pymysql


class InventoryStatusUpdateWorker(QThread):
    """库存状态更新工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    
    def __init__(self, datasource, db_manager, excel_file, inventory_key_column, inventory_status_column, selected_records):
        super().__init__()
        self.datasource = datasource
        self.db_manager = db_manager
        self.excel_file = excel_file
        self.inventory_key_column = inventory_key_column
        self.inventory_status_column = inventory_status_column
        self.selected_records = selected_records
        
        # 状态映射规则
        self.status_mapping = {
            '正常': 4,      # 已售
            '暂停': 1,      # 过审暂停
            '已闭店': 5,    # 封号
            '已退店': 3,    # 可售
            '闭店申请中': 4,  # 已售
            '闭店回款中': 4,  # 已售
            '退店申请中': 4,  # 已售
            '退店回款中': 4,  # 已售
        }
        
        # 默认状态（其他所有情况）
        self.default_status = 4  # 已售
        
        self.results = {
            'total_processed': 0,
            'success_count': 0,
            'failed_count': 0,
            'not_found_count': 0
        }
    
    def run(self):
        """执行状态更新"""
        try:
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
            cursor = connection.cursor(pymysql.cursors.DictCursor)
            self.log_message.emit("数据库连接成功")
            
            # 读取Excel文件
            self.log_message.emit("开始读取Excel文件...")
            df = pd.read_excel(self.excel_file)
            self.log_message.emit(f"成功读取Excel文件，共 {len(df)} 行数据")
            
            # 验证列名
            if self.inventory_key_column not in df.columns:
                raise Exception(f"未找到库存主键列: {self.inventory_key_column}")
            if self.inventory_status_column not in df.columns:
                raise Exception(f"未找到库存状态列: {self.inventory_status_column}")
            
            self.progress.emit(10)
            
            # 处理选中的记录
            total_records = len(self.selected_records)
            self.results['total_processed'] = total_records
            
            for i, record_index in enumerate(self.selected_records):
                try:
                    if record_index >= len(df):
                        continue
                        
                    row = df.iloc[record_index]
                    self.process_record(cursor, record_index + 1, row)
                    
                    # 更新进度
                    progress = 10 + int((i + 1) / total_records * 80)
                    self.progress.emit(progress)
                    
                except Exception as e:
                    self.log_message.emit(f"第{record_index + 1}行处理失败: {str(e)}")
                    self.results['failed_count'] += 1
            
            # 提交事务
            connection.commit()
            self.log_message.emit("所有更改已提交到数据库")
            
            # 关闭数据库连接
            cursor.close()
            connection.close()
            self.log_message.emit("数据库连接已关闭")
            
            self.progress.emit(100)
            self.finished.emit(self.results)
            
        except Exception as e:
            self.error.emit(f"处理过程出错: {str(e)}")
    
    def process_record(self, cursor, row_num, row):
        """处理单条记录"""
        try:
            # 获取库存主键和状态
            inventory_key = str(row[self.inventory_key_column]).strip()
            inventory_status = str(row[self.inventory_status_column]).strip()
            
            if not inventory_key or pd.isna(row[self.inventory_key_column]):
                self.log_message.emit(f"第{row_num}行: 库存主键为空，跳过")
                return
            
            # 映射状态
            account_status = self.status_mapping.get(inventory_status, self.default_status)
            
            # 查询ba_account_mag表中是否存在该记录
            cursor.execute("""
                SELECT id, account_status 
                FROM ba_account_mag 
                WHERE id = %s
            """, (inventory_key,))
            
            existing_record = cursor.fetchone()
            
            if existing_record:
                old_status = existing_record.get('account_status')
                
                # 检查是否是已闭店状态，需要同时更新init_status
                if inventory_status == '已闭店':
                    # 已闭店：更新account_status为5(封号)，同时更新init_status为4
                    cursor.execute("""
                        UPDATE ba_account_mag 
                        SET account_status = %s, init_status = 4, update_time = %s
                        WHERE id = %s
                    """, (account_status, int(datetime.now().timestamp()), inventory_key))
                    
                    self.log_message.emit(f"第{row_num}行: 库存主键 {inventory_key} 状态更新成功 ({old_status} -> {account_status}，init_status -> 4)")
                else:
                    # 其他状态：只更新account_status
                    cursor.execute("""
                        UPDATE ba_account_mag 
                        SET account_status = %s, update_time = %s
                        WHERE id = %s
                    """, (account_status, int(datetime.now().timestamp()), inventory_key))
                    
                    self.log_message.emit(f"第{row_num}行: 库存主键 {inventory_key} 状态更新成功 ({old_status} -> {account_status})")
                
                self.results['success_count'] += 1
            else:
                self.log_message.emit(f"第{row_num}行: 库存主键 {inventory_key} 在数据库中未找到")
                self.results['not_found_count'] += 1
                
        except Exception as e:
            raise Exception(f"处理记录失败: {str(e)}")


class InventoryStatusUpdateWidget(QWidget):
    """库存状态更新界面"""
    
    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.worker = None
        self.excel_file = ""
        self.df = None
        self.selected_records = []
        
        self.init_ui()
    
    def init_ui(self):
        """初始化界面"""
        layout = QVBoxLayout()
        
        # 标题
        title = QLabel("按库存主键更新状态")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        # 数据源选择
        datasource_group = QGroupBox("数据源配置")
        datasource_layout = QFormLayout()
        
        self.datasource_combo = QComboBox()
        self.load_datasources()
        datasource_layout.addRow("数据源:", self.datasource_combo)
        
        datasource_group.setLayout(datasource_layout)
        layout.addWidget(datasource_group)
        
        # 文件选择
        file_group = QGroupBox("文件选择")
        file_layout = QVBoxLayout()
        
        excel_layout = QHBoxLayout()
        self.excel_path_label = QLabel("未选择文件")
        self.select_excel_btn = QPushButton("选择Excel文件")
        self.select_excel_btn.clicked.connect(self.select_excel_file)
        excel_layout.addWidget(QLabel("Excel文件:"))
        excel_layout.addWidget(self.excel_path_label, 1)
        excel_layout.addWidget(self.select_excel_btn)
        file_layout.addLayout(excel_layout)
        
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # 列映射配置
        self.mapping_group = QGroupBox("列映射配置")
        mapping_layout = QFormLayout()
        
        self.inventory_key_combo = QComboBox()
        self.inventory_status_combo = QComboBox()
        
        mapping_layout.addRow("库存主键列:", self.inventory_key_combo)
        mapping_layout.addRow("库存状态列:", self.inventory_status_combo)
        
        self.mapping_group.setLayout(mapping_layout)
        self.mapping_group.setEnabled(False)
        layout.addWidget(self.mapping_group)
        
        # 数据预览
        preview_group = QGroupBox("数据预览")
        preview_layout = QVBoxLayout()
        
        self.preview_table = QTableWidget()
        self.preview_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.preview_table.setAlternatingRowColors(True)
        preview_layout.addWidget(self.preview_table)
        
        # 全选/反选按钮和记录统计
        select_layout = QHBoxLayout()
        self.select_all_btn = QPushButton("全选可见记录")
        self.select_none_btn = QPushButton("全不选")
        self.select_all_visible_btn = QPushButton("全选所有记录(包括不可见)")
        self.record_count_label = QLabel("已选择: 0 / 0 条记录")
        
        self.select_all_btn.clicked.connect(self.select_all_visible_records)
        self.select_none_btn.clicked.connect(self.select_no_records)
        self.select_all_visible_btn.clicked.connect(self.select_all_records)
        
        select_layout.addWidget(self.select_all_btn)
        select_layout.addWidget(self.select_none_btn)
        select_layout.addWidget(self.select_all_visible_btn)
        select_layout.addWidget(self.record_count_label)
        select_layout.addStretch()
        preview_layout.addLayout(select_layout)
        
        preview_group.setLayout(preview_layout)
        layout.addWidget(preview_group)
        
        # 操作按钮
        button_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始更新")
        self.stop_btn = QPushButton("停止")
        self.clear_log_btn = QPushButton("清空日志")
        
        self.start_btn.clicked.connect(self.start_update)
        self.stop_btn.clicked.connect(self.stop_update)
        self.clear_log_btn.clicked.connect(self.clear_log)
        
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(False)
        
        button_layout.addWidget(self.start_btn)
        button_layout.addWidget(self.stop_btn)
        button_layout.addStretch()
        button_layout.addWidget(self.clear_log_btn)
        layout.addLayout(button_layout)
        
        # 进度条
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)
        
        # 日志显示
        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout()
        
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(200)
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)
        
        self.setLayout(layout)
        
        # 连接信号
        self.inventory_key_combo.currentTextChanged.connect(self.check_can_start)
        self.inventory_status_combo.currentTextChanged.connect(self.check_can_start)
    
    def select_excel_file(self):
        """选择Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Excel文件", "", "Excel files (*.xlsx *.xls)"
        )
        
        if file_path:
            self.excel_file = file_path
            self.excel_path_label.setText(os.path.basename(file_path))
            self.load_excel_data()
    
    def load_excel_data(self):
        """加载Excel数据"""
        try:
            self.df = pd.read_excel(self.excel_file)
            
            # 更新列选择下拉框
            columns = list(self.df.columns)
            self.inventory_key_combo.clear()
            self.inventory_status_combo.clear()
            self.inventory_key_combo.addItems(columns)
            self.inventory_status_combo.addItems(columns)
            
            # 启用映射配置
            self.mapping_group.setEnabled(True)
            
            # 显示数据预览
            self.show_data_preview()
            
            self.append_log(f"成功加载Excel文件，共 {len(self.df)} 行数据")
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"加载Excel文件失败：{str(e)}")
    
    def show_data_preview(self):
        """显示数据预览"""
        if self.df is None:
            return
        
        # 设置表格 - 只显示前100行用于预览，但可以选择处理所有记录
        total_rows = len(self.df)
        preview_rows = min(total_rows, 100)
        self.preview_table.setRowCount(preview_rows)
        self.preview_table.setColumnCount(len(self.df.columns) + 1)  # +1 for checkbox
        
        # 设置表头
        headers = ["选择"] + list(self.df.columns)
        self.preview_table.setHorizontalHeaderLabels(headers)
        
        # 填充数据
        for i in range(preview_rows):
            # 添加复选框
            checkbox_item = QTableWidgetItem()
            checkbox_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            checkbox_item.setCheckState(Qt.Checked)
            self.preview_table.setItem(i, 0, checkbox_item)
            
            # 添加数据
            for j, col in enumerate(self.df.columns):
                value = str(self.df.iloc[i, j]) if pd.notna(self.df.iloc[i, j]) else ""
                item = QTableWidgetItem(value)
                item.setFlags(Qt.ItemIsEnabled)
                self.preview_table.setItem(i, j + 1, item)
        
        # 调整列宽
        self.preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        
        # 连接复选框变化信号
        self.preview_table.itemChanged.connect(self.on_item_changed)
        
        # 更新记录统计
        self.update_record_count()
        
        # 默认全选可见记录
        self.select_all_visible_records()
    
    def select_all_visible_records(self):
        """全选记录"""
        for i in range(self.preview_table.rowCount()):
            item = self.preview_table.item(i, 0)
            if item:
                item.setCheckState(Qt.Checked)
        self.update_selected_records()
    
    def select_no_records(self):
        """全不选记录"""
        for i in range(self.preview_table.rowCount()):
            item = self.preview_table.item(i, 0)
            if item:
                item.setCheckState(Qt.Unchecked)
        self.update_selected_records()
    
    def select_all_records(self):
        """全选所有记录（包括不可见的）"""
        if self.df is None:
            self.append_log("错误：Excel数据未加载")
            return
        
        total_rows = len(self.df)
        visible_rows = self.preview_table.rowCount()
        
        self.append_log(f"开始全选所有记录：总行数={total_rows}, 可见行数={visible_rows}")
        
        # 先全选可见的记录
        for i in range(self.preview_table.rowCount()):
            item = self.preview_table.item(i, 0)
            if item:
                item.setCheckState(Qt.Checked)
        
        # 清空当前选择
        self.selected_records = []
        
        # 添加所有记录索引（0到total_rows-1）
        self.selected_records = list(range(total_rows))
        
        self.update_record_count()
        self.check_can_start()
        
        self.append_log(f"✓ 已选择所有 {total_rows} 条记录进行处理（包括 {total_rows - visible_rows} 条不可见记录）")
    
    def update_record_count(self):
        """更新记录统计显示"""
        if self.df is None:
            self.record_count_label.setText("已选择: 0 / 0 条记录")
            return
        
        total_count = len(self.df)
        selected_count = len(self.selected_records)
        self.record_count_label.setText(f"已选择: {selected_count} / {total_count} 条记录")
    
    def on_item_changed(self, item):
        """表格项变化时的处理"""
        if item.column() == 0:  # 只处理复选框列
            self.update_selected_records()
    
    def update_selected_records(self):
        """更新选中的记录"""
        self.selected_records = []
        for i in range(self.preview_table.rowCount()):
            item = self.preview_table.item(i, 0)
            if item and item.checkState() == Qt.Checked:
                self.selected_records.append(i)
        
        self.update_record_count()
        self.check_can_start()
    
    def check_can_start(self):
        """检查是否可以开始处理"""
        try:
            can_start = (
                bool(self.excel_file) and 
                bool(self.inventory_key_combo.currentText()) and
                bool(self.inventory_status_combo.currentText()) and
                len(self.selected_records) > 0
            )
            self.start_btn.setEnabled(can_start)
        except Exception as e:
            # 初始化时可能会出现异常，忽略即可
            pass
    
    def start_update(self):
        """开始更新状态"""
        # 检查是否有选中的记录（不重新计算，保持用户的选择）
        if not self.selected_records:
            QMessageBox.warning(self, "警告", "请至少选择一条记录进行处理")
            return
        
        # 添加调试信息
        self.append_log(f"准备处理 {len(self.selected_records)} 条记录")
        
        # 获取选中的数据源对象
        datasource_data = self.datasource_combo.currentData()
        if not datasource_data:
            QMessageBox.warning(self, "警告", "请选择数据源")
            return
        
        # 根据数据源名称获取完整的数据源对象
        datasources = self.db_manager.get_all_datasources()
        datasource = None
        for ds in datasources:
            if ds.name == datasource_data:
                datasource = ds
                break
        
        if not datasource:
            QMessageBox.warning(self, "警告", "未找到选中的数据源")
            return
        
        # 创建工作线程
        self.worker = InventoryStatusUpdateWorker(
            datasource,
            self.db_manager,
            self.excel_file,
            self.inventory_key_combo.currentText(),
            self.inventory_status_combo.currentText(),
            self.selected_records
        )
        
        # 连接信号
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log_message.connect(self.append_log)
        self.worker.finished.connect(self.on_update_finished)
        self.worker.error.connect(self.on_update_error)
        
        # 更新UI状态
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.select_excel_btn.setEnabled(False)
        self.datasource_combo.setEnabled(False)
        self.inventory_key_combo.setEnabled(False)
        self.inventory_status_combo.setEnabled(False)
        
        # 启动线程
        self.worker.start()
    
    def stop_update(self):
        """停止更新"""
        if self.worker and self.worker.isRunning():
            self.worker.terminate()
            self.worker.wait()
            self.append_log("用户停止了更新操作")
        
        self.reset_ui()
    
    def on_update_finished(self, results):
        """更新完成"""
        self.append_log("=" * 50)
        self.append_log("更新完成！")
        self.append_log(f"总处理记录数: {results['total_processed']}")
        self.append_log(f"成功更新: {results['success_count']} 条")
        self.append_log(f"未找到记录: {results['not_found_count']} 条")
        self.append_log(f"处理失败: {results['failed_count']} 条")
        self.append_log("=" * 50)
        
        QMessageBox.information(
            self,
            "完成",
            f"状态更新完成！\n\n"
            f"总处理记录数: {results['total_processed']}\n"
            f"成功更新: {results['success_count']} 条\n"
            f"未找到记录: {results['not_found_count']} 条\n"
            f"处理失败: {results['failed_count']} 条"
        )
        
        self.reset_ui()
    
    def on_update_error(self, error_msg):
        """更新出错"""
        self.append_log(f"错误: {error_msg}")
        QMessageBox.critical(self, "错误", error_msg)
        self.reset_ui()
    
    def reset_ui(self):
        """重置UI状态"""
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.select_excel_btn.setEnabled(True)
        self.datasource_combo.setEnabled(True)
        self.inventory_key_combo.setEnabled(True)
        self.inventory_status_combo.setEnabled(True)
        self.check_can_start()
    
    def append_log(self, message):
        """添加日志"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
        
        # 自动滚动到底部
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
    
    def clear_log(self):
        """清空日志"""
        self.log_text.clear()
    
    def load_datasources(self):
        """加载数据源列表"""
        self.datasource_combo.clear()
        self.datasource_combo.addItem("请选择数据源", None)
        
        datasources = self.db_manager.get_all_datasources()
        for ds in datasources:
            display_name = f"{ds.name} ({ds.host}:{ds.port})"
            self.datasource_combo.addItem(display_name, ds.name)
