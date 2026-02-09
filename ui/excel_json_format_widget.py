"""
Excel数据JSON格式转换功能界面
将Excel中的账号密码列转换为JSON格式
"""
import os
import json
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog, 
                             QTableWidget, QTableWidgetItem, QHeaderView,
                             QCheckBox, QSplitter)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
import pandas as pd


class ExcelJsonFormatWidget(QWidget):
    """Excel数据JSON格式转换功能组件"""
    
    def __init__(self, db_manager):
        super().__init__()
        self.db_manager = db_manager
        self.excel_file = None
        self.df = None
        self.processed_df = None
        self.init_ui()
    
    def init_ui(self):
        """初始化UI"""
        layout = QVBoxLayout()
        
        # 文件选择区域
        file_group = QGroupBox("选择Excel文件")
        file_layout = QFormLayout()
        
        file_select_layout = QHBoxLayout()
        self.file_label = QLabel("未选择文件")
        self.file_btn = QPushButton("选择文件")
        self.file_btn.clicked.connect(self.select_file)
        file_select_layout.addWidget(self.file_label)
        file_select_layout.addWidget(self.file_btn)
        
        file_layout.addRow("Excel文件:", file_select_layout)
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # 字段映射配置区域
        mapping_group = QGroupBox("字段映射配置")
        mapping_layout = QFormLayout()
        
        # id_account_password字段配置
        self.id_account_combo = QComboBox()
        self.id_password_combo = QComboBox()
        mapping_layout.addRow("ID账号列:", self.id_account_combo)
        mapping_layout.addRow("ID密码列:", self.id_password_combo)
        
        # email_address_password字段配置
        self.email_address_combo = QComboBox()
        self.email_password_combo = QComboBox()
        mapping_layout.addRow("邮箱地址列:", self.email_address_combo)
        mapping_layout.addRow("邮箱密码列:", self.email_password_combo)
        
        mapping_group.setLayout(mapping_layout)
        mapping_group.setEnabled(False)
        self.mapping_group = mapping_group
        layout.addWidget(mapping_group)
        
        # 预览和处理区域
        preview_group = QGroupBox("数据预览与处理")
        preview_layout = QVBoxLayout()
        
        # 操作按钮
        button_layout = QHBoxLayout()
        self.process_btn = QPushButton("处理数据")
        self.export_btn = QPushButton("导出结果")
        self.process_btn.clicked.connect(self.process_data)
        self.export_btn.clicked.connect(self.export_results)
        self.process_btn.setEnabled(False)
        self.export_btn.setEnabled(False)
        
        button_layout.addWidget(self.process_btn)
        button_layout.addWidget(self.export_btn)
        button_layout.addStretch()
        preview_layout.addLayout(button_layout)
        
        # 数据预览表格
        self.preview_table = QTableWidget()
        self.preview_table.setAlternatingRowColors(True)
        preview_layout.addWidget(self.preview_table)
        
        preview_group.setLayout(preview_layout)
        layout.addWidget(preview_group)
        
        # 日志区域
        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout()
        
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(150)
        self.log_text.setFont(QFont("Consolas", 9))
        log_layout.addWidget(self.log_text)
        
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)
        
        self.setLayout(layout)
    
    def select_file(self):
        """选择Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "选择Excel文件", 
            "", 
            "Excel files (*.xlsx *.xls)"
        )
        
        if file_path:
            self.load_excel_file(file_path)
    
    def load_excel_file(self, file_path):
        """加载Excel文件"""
        try:
            self.excel_file = file_path
            self.df = pd.read_excel(file_path)
            
            # 更新文件标签
            filename = os.path.basename(file_path)
            self.file_label.setText(f"{filename} ({len(self.df)} 行)")
            
            # 更新列选择下拉框
            columns = list(self.df.columns)
            
            # 清空并重新填充下拉框
            for combo in [self.id_account_combo, self.id_password_combo, 
                         self.email_address_combo, self.email_password_combo]:
                combo.clear()
                combo.addItem("-- 请选择 --", "")
                for col in columns:
                    combo.addItem(col, col)
            
            # 启用映射配置
            self.mapping_group.setEnabled(True)
            
            # 显示原始数据预览
            self.show_original_preview()
            
            # 连接信号以检查是否可以处理
            for combo in [self.id_account_combo, self.id_password_combo, 
                         self.email_address_combo, self.email_password_combo]:
                combo.currentTextChanged.connect(self.check_can_process)
            
            self.append_log(f"成功加载Excel文件: {filename}")
            self.append_log(f"数据行数: {len(self.df)}, 列数: {len(self.df.columns)}")
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"加载Excel文件失败：{str(e)}")
            self.append_log(f"加载文件失败: {str(e)}")
    
    def show_original_preview(self):
        """显示原始数据预览"""
        if self.df is None:
            return
        
        # 只显示前50行
        preview_rows = min(len(self.df), 50)
        self.preview_table.setRowCount(preview_rows)
        self.preview_table.setColumnCount(len(self.df.columns))
        
        # 设置表头
        self.preview_table.setHorizontalHeaderLabels(list(self.df.columns))
        
        # 填充数据
        for i in range(preview_rows):
            for j, col in enumerate(self.df.columns):
                value = str(self.df.iloc[i, j]) if pd.notna(self.df.iloc[i, j]) else ""
                item = QTableWidgetItem(value)
                item.setFlags(Qt.ItemIsEnabled)
                self.preview_table.setItem(i, j, item)
        
        # 调整列宽
        self.preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
    
    def check_can_process(self):
        """检查是否可以开始处理"""
        try:
            # 检查是否至少配置了一组字段映射
            id_account_data = self.id_account_combo.currentData()
            id_password_data = self.id_password_combo.currentData()
            email_address_data = self.email_address_combo.currentData()
            email_password_data = self.email_password_combo.currentData()
            
            # 确保数据不是空字符串
            id_account_ok = bool(id_account_data and id_account_data.strip())
            id_password_ok = bool(id_password_data and id_password_data.strip())
            email_address_ok = bool(email_address_data and email_address_data.strip())
            email_password_ok = bool(email_password_data and email_password_data.strip())
            
            # 检查组合
            id_combo_ok = id_account_ok and id_password_ok
            email_combo_ok = email_address_ok and email_password_ok
            
            # 最终判断
            can_process = bool(self.excel_file) and (id_combo_ok or email_combo_ok)
            
            # 添加调试信息
            self.append_log(f"检查处理条件: Excel文件={bool(self.excel_file)}, ID组合={id_combo_ok}, 邮箱组合={email_combo_ok}, 可处理={can_process}")
            
            # 确保传入布尔值
            self.process_btn.setEnabled(bool(can_process))
            
        except Exception as e:
            self.append_log(f"检查处理条件时出错: {str(e)}")
            # 出错时禁用按钮
            self.process_btn.setEnabled(False)
    
    def process_data(self):
        """处理数据"""
        try:
            if self.df is None:
                QMessageBox.warning(self, "警告", "请先选择Excel文件")
                return
            
            self.append_log("开始处理数据...")
            
            # 复制原始数据
            self.processed_df = self.df.copy()
            
            # 处理id_account_password字段
            if (self.id_account_combo.currentData() and 
                self.id_password_combo.currentData()):
                
                id_account_col = self.id_account_combo.currentData()
                id_password_col = self.id_password_combo.currentData()
                
                self.append_log(f"处理ID账号密码字段: {id_account_col} + {id_password_col}")
                
                id_account_password_data = []
                for index, row in self.processed_df.iterrows():
                    id_account = str(row[id_account_col]) if pd.notna(row[id_account_col]) else ""
                    id_password = str(row[id_password_col]) if pd.notna(row[id_password_col]) else ""
                    
                    # 构建JSON格式
                    json_data = [{
                        "id_account": id_account,
                        "id_account_password": id_password
                    }]
                    
                    id_account_password_data.append(json.dumps(json_data, ensure_ascii=False))
                
                # 添加到DataFrame
                self.processed_df['id_account_password'] = id_account_password_data
                self.append_log(f"✓ 已生成 id_account_password 字段")
            
            # 处理email_address_password字段
            if (self.email_address_combo.currentData() and 
                self.email_password_combo.currentData()):
                
                email_address_col = self.email_address_combo.currentData()
                email_password_col = self.email_password_combo.currentData()
                
                self.append_log(f"处理邮箱地址密码字段: {email_address_col} + {email_password_col}")
                
                email_address_password_data = []
                for index, row in self.processed_df.iterrows():
                    email_address = str(row[email_address_col]) if pd.notna(row[email_address_col]) else ""
                    email_password = str(row[email_password_col]) if pd.notna(row[email_password_col]) else ""
                    
                    # 构建JSON格式
                    json_data = [{
                        "email_address": email_address,
                        "email_address_password": email_password
                    }]
                    
                    email_address_password_data.append(json.dumps(json_data, ensure_ascii=False))
                
                # 添加到DataFrame
                self.processed_df['email_address_password'] = email_address_password_data
                self.append_log(f"✓ 已生成 email_address_password 字段")
            
            # 显示处理后的预览
            self.show_processed_preview()
            
            # 启用导出按钮
            self.export_btn.setEnabled(True)
            
            self.append_log(f"✓ 数据处理完成！共处理 {len(self.processed_df)} 行数据")
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"处理数据失败：{str(e)}")
            self.append_log(f"处理失败: {str(e)}")
    
    def show_processed_preview(self):
        """显示处理后的数据预览"""
        if self.processed_df is None:
            return
        
        # 只显示前50行
        preview_rows = min(len(self.processed_df), 50)
        self.preview_table.setRowCount(preview_rows)
        self.preview_table.setColumnCount(len(self.processed_df.columns))
        
        # 设置表头
        self.preview_table.setHorizontalHeaderLabels(list(self.processed_df.columns))
        
        # 填充数据
        for i in range(preview_rows):
            for j, col in enumerate(self.processed_df.columns):
                value = str(self.processed_df.iloc[i, j]) if pd.notna(self.processed_df.iloc[i, j]) else ""
                # 如果是JSON字段，截断显示
                if col in ['id_account_password', 'email_address_password'] and len(value) > 100:
                    value = value[:100] + "..."
                
                item = QTableWidgetItem(value)
                item.setFlags(Qt.ItemIsEnabled)
                
                # JSON字段用不同颜色标识
                if col in ['id_account_password', 'email_address_password']:
                    item.setBackground(Qt.lightGray)
                
                self.preview_table.setItem(i, j, item)
        
        # 调整列宽
        self.preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
    
    def export_results(self):
        """导出处理结果"""
        if self.processed_df is None:
            QMessageBox.warning(self, "警告", "请先处理数据")
            return
        
        try:
            # 选择保存路径
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            default_filename = f"json_formatted_data_{timestamp}.xlsx"
            
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "保存处理结果",
                default_filename,
                "Excel files (*.xlsx)"
            )
            
            if file_path:
                # 导出到Excel
                self.processed_df.to_excel(file_path, index=False, engine='openpyxl')
                
                self.append_log(f"✓ 结果已导出到: {file_path}")
                QMessageBox.information(self, "成功", f"数据已成功导出到:\n{file_path}")
                
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导出失败：{str(e)}")
            self.append_log(f"导出失败: {str(e)}")
    
    def append_log(self, message):
        """添加日志"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
        self.log_text.ensureCursorVisible()
