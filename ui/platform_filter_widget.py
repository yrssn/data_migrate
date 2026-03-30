"""
过滤掉平台为银行的数据功能界面
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


class PlatformFilterWorker(QThread):
    """平台过滤工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    
    def __init__(self, excel_file, platform_column, output_file, db_config):
        super().__init__()
        self.excel_file = excel_file
        self.platform_column = platform_column  # Excel中的平台列名
        self.output_file = output_file
        self.db_config = db_config
        self.results = {
            'total_rows': 0,
            'filtered_rows': 0,
            'remaining_rows': 0,
            'bank_platforms': [],
            'filtered_data': []
        }
    
    def run(self):
        try:
            self.log_message.emit("开始读取Excel文件...")
            self.progress.emit(10)
            
            # 读取Excel文件
            df = pd.read_excel(self.excel_file, engine='openpyxl')
            self.log_message.emit(f"成功读取Excel，共 {len(df)} 行数据")
            
            # 检查平台列是否存在
            if self.platform_column not in df.columns:
                raise Exception(f"Excel中未找到平台列: {self.platform_column}")
            
            self.log_message.emit("连接数据库获取平台信息...")
            self.progress.emit(20)
            
            # 连接数据库获取平台信息
            connection = pymysql.connect(**self.db_config)
            cursor = connection.cursor()
            
            # 获取所有平台信息
            cursor.execute("SELECT platform, platform_type FROM ba_platform WHERE status = 1")
            platform_data = cursor.fetchall()
            
            # 构建平台类型映射字典
            platform_type_map = {}
            bank_platforms = []
            
            for platform, platform_type in platform_data:
                platform_type_map[platform] = platform_type
                if platform_type == 'opt1':  # 银行类型
                    bank_platforms.append(platform)
            
            self.log_message.emit(f"获取到 {len(platform_type_map)} 个平台信息")
            self.log_message.emit(f"其中银行平台 {len(bank_platforms)} 个: {', '.join(bank_platforms)}")
            
            cursor.close()
            connection.close()
            
            self.log_message.emit("开始过滤数据...")
            self.progress.emit(40)
            
            total_rows = len(df)
            filtered_count = 0
            filtered_data = []
            
            for index, row in df.iterrows():
                try:
                    # 获取平台值
                    platform_value = row[self.platform_column]
                    
                    # 处理空值
                    if pd.isna(platform_value):
                        platform_value = ""
                    else:
                        platform_value = str(platform_value).strip()
                    
                    # 检查是否为银行平台
                    platform_type = platform_type_map.get(platform_value, None)
                    
                    if platform_type == 'opt1':
                        # 是银行平台，过滤掉
                        filtered_count += 1
                        self.log_message.emit(f"第{index+1}行: 平台 '{platform_value}' 为银行类型，已过滤")
                    else:
                        # 不是银行平台，保留
                        filtered_data.append(row.to_dict())
                        
                        if platform_type is None and platform_value:
                            self.log_message.emit(f"第{index+1}行: 平台 '{platform_value}' 在数据库中未找到，已保留")
                
                except Exception as e:
                    # 处理异常，保留该行数据
                    filtered_data.append(row.to_dict())
                    self.log_message.emit(f"第{index+1}行处理异常: {str(e)}，已保留")
                
                # 更新进度
                progress = 40 + int((index + 1) / total_rows * 40)
                self.progress.emit(progress)
            
            remaining_rows = len(filtered_data)
            
            self.log_message.emit("开始导出过滤后的数据...")
            self.progress.emit(85)
            
            # 导出过滤后的数据
            if filtered_data:
                result_df = pd.DataFrame(filtered_data)
                result_df.to_excel(self.output_file, index=False, engine='openpyxl')
                self.log_message.emit(f"过滤后的数据已导出到: {self.output_file}")
            else:
                # 创建空文件
                empty_df = pd.DataFrame(columns=df.columns)
                empty_df.to_excel(self.output_file, index=False, engine='openpyxl')
                self.log_message.emit(f"所有数据都被过滤，已创建空文件: {self.output_file}")
            
            self.results['total_rows'] = total_rows
            self.results['filtered_rows'] = filtered_count
            self.results['remaining_rows'] = remaining_rows
            self.results['bank_platforms'] = bank_platforms
            self.results['filtered_data'] = filtered_data
            
            self.log_message.emit("过滤完成！")
            self.progress.emit(100)
            
            self.finished.emit(self.results)
            
        except Exception as e:
            self.error.emit(str(e))


class PlatformFilterWidget(QWidget):
    """平台过滤界面"""
    
    def __init__(self, db_manager=None):
        super().__init__()
        self.db_manager = db_manager
        self.excel_file = ""
        self.output_file = ""
        self.excel_columns = []
        self.worker = None
        self.selected_datasource = None
        
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # 功能说明
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        
        info_label = QLabel("""
<b>过滤掉平台为银行的数据功能:</b><br>
<b>使用步骤:</b> 1.导入Excel文件 → 2.选择平台列 → 3.选择输出文件 → 4.开始过滤<br>
<b>功能说明:</b> 根据Excel中指定列的平台值，与数据库ba_platform表对比，过滤掉platform_type为opt1（银行）的数据<br>
<b>处理逻辑:</b> 保留非银行平台的数据，过滤掉银行平台的数据，未找到的平台默认保留
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
        
        # 输出文件选择
        output_layout = QHBoxLayout()
        self.output_label = QLabel("未选择文件")
        self.output_label.setStyleSheet("color: gray;")
        output_layout.addWidget(self.output_label)
        
        self.select_output_btn = QPushButton("选择输出文件")
        self.select_output_btn.clicked.connect(self.select_output_file)
        output_layout.addWidget(self.select_output_btn)
        
        file_layout.addRow("输出文件:", output_layout)
        
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # 平台列选择区域
        platform_group = QGroupBox("平台列设置")
        platform_layout = QFormLayout()
        
        self.platform_column_combo = QComboBox()
        self.platform_column_combo.setEnabled(False)
        self.platform_column_combo.currentTextChanged.connect(self.update_start_button_state)
        platform_layout.addRow("平台列:", self.platform_column_combo)
        
        platform_group.setLayout(platform_layout)
        layout.addWidget(platform_group)
        
        # 操作按钮区域
        button_layout = QHBoxLayout()
        
        self.start_btn = QPushButton("开始过滤")
        self.start_btn.clicked.connect(self.start_filter)
        self.start_btn.setEnabled(False)
        self.start_btn.setStyleSheet("""
            QPushButton {
                background-color: #007ACC;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 5px;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #005a9e;
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
        
        result_group.setLayout(result_layout)
        layout.addWidget(result_group)
        
        # 日志区域
        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout()
        
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(200)
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
                
                # 填充平台列下拉框
                self.platform_column_combo.clear()
                self.platform_column_combo.addItem("请选择列")
                self.platform_column_combo.addItems(self.excel_columns)
                self.platform_column_combo.setEnabled(True)
                
                self.update_start_button_state()
                
                self.log_text.append(f"已选择Excel文件: {os.path.basename(file_path)}")
                self.log_text.append(f"检测到列: {', '.join(self.excel_columns)}")
                
            except Exception as e:
                QMessageBox.critical(self, "错误", f"读取Excel文件失败: {str(e)}")
                self.excel_file = ""
                self.excel_label.setText("未选择文件")
                self.excel_label.setStyleSheet("color: gray;")
    
    def select_output_file(self):
        """选择输出文件"""
        default_name = f"过滤后数据_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "选择输出文件",
            default_name,
            "Excel文件 (*.xlsx)"
        )
        
        if file_path:
            self.output_file = file_path
            self.output_label.setText(os.path.basename(file_path))
            self.output_label.setStyleSheet("color: green;")
            
            self.update_start_button_state()
            
            self.log_text.append(f"已选择输出文件: {os.path.basename(file_path)}")
    
    def update_start_button_state(self):
        """更新开始按钮状态"""
        # 检查必要的控件是否已创建
        if not hasattr(self, 'start_btn') or not hasattr(self, 'platform_column_combo'):
            return
            
        can_start = bool(
            self.selected_datasource and
            self.excel_file and
            self.output_file and
            self.platform_column_combo.currentText()
        )
        
        self.start_btn.setEnabled(can_start)
    
    def start_filter(self):
        """开始过滤"""
        # 获取平台列
        platform_column = self.platform_column_combo.currentText()
        
        # 确认对话框
        reply = QMessageBox.question(
            self,
            "确认过滤",
            f"确定要开始过滤数据吗？\n\n"
            f"Excel文件: {os.path.basename(self.excel_file)}\n"
            f"平台列: {platform_column}\n"
            f"输出文件: {os.path.basename(self.output_file)}\n\n"
            f"将过滤掉所有银行平台（platform_type=opt1）的数据",
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
        self.result_label.setText("正在过滤中...")
        self.log_text.clear()
        
        # 启动工作线程
        self.worker = PlatformFilterWorker(
            self.excel_file, platform_column, self.output_file, db_config
        )
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log_message.connect(self.append_log)
        self.worker.finished.connect(self.on_filter_finished)
        self.worker.error.connect(self.on_filter_error)
        self.worker.start()
    
    def append_log(self, message):
        """添加日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
    
    def on_filter_finished(self, results):
        """过滤完成"""
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        total_rows = results['total_rows']
        filtered_rows = results['filtered_rows']
        remaining_rows = results['remaining_rows']
        bank_platforms = results['bank_platforms']
        
        result_text = f"过滤完成！总行数: {total_rows}, 过滤掉: {filtered_rows}, 保留: {remaining_rows}"
        self.result_label.setText(result_text)
        self.result_label.setStyleSheet("color: green; font-weight: bold;")
        
        self.log_text.append(f"\n=== 过滤完成 ===")
        self.log_text.append(result_text)
        self.log_text.append(f"过滤率: {filtered_rows/total_rows*100:.1f}%")
        self.log_text.append(f"银行平台: {', '.join(bank_platforms) if bank_platforms else '无'}")
        
        QMessageBox.information(
            self, 
            "完成", 
            f"{result_text}\n\n"
            f"过滤掉的银行平台: {', '.join(bank_platforms) if bank_platforms else '无'}\n\n"
            f"结果已保存到:\n{self.output_file}"
        )
    
    def on_filter_error(self, error_message):
        """过滤错误"""
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.result_label.setText("过滤失败！")
        self.result_label.setStyleSheet("color: red; font-weight: bold;")
        
        QMessageBox.critical(self, "错误", f"过滤过程出错: {error_message}")
        self.log_text.append(f"错误: {error_message}")
