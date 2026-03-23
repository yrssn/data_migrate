"""
改变状态栏功能界面
修改Excel中某一列的状态值，从一个状态改为另一个状态
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
import pandas as pd


class StatusChangeWidget(QWidget):
    """改变状态栏功能组件 - 修改Excel状态列"""
    
    def __init__(self, db_manager=None):
        super().__init__()
        self.excel_file = None
        self.excel_data = None
        self.init_ui()
    
    def init_ui(self):
        """初始化UI"""
        layout = QVBoxLayout()
        
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
        
        # 状态列选择
        self.column_combo = QComboBox()
        self.column_combo.setEnabled(False)
        self.column_combo.currentIndexChanged.connect(self.on_column_changed)
        file_layout.addRow("状态所在列:", self.column_combo)
        
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # 状态转换配置区域
        status_group = QGroupBox("状态转换配置")
        status_layout = QFormLayout()
        
        self.from_status_edit = QLineEdit()
        self.from_status_edit.setPlaceholderText("例如: 待定")
        status_layout.addRow("原状态值:", self.from_status_edit)
        
        self.to_status_edit = QLineEdit()
        self.to_status_edit.setPlaceholderText("例如: 进行中")
        status_layout.addRow("目标状态值:", self.to_status_edit)
        
        status_group.setLayout(status_layout)
        layout.addWidget(status_group)
        
        # 功能说明区域
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        
        info_label = QLabel("""
<b>改变状态栏说明:</b><br>
• 1. 选择要修改的Excel文件<br>
• 2. 选择状态所在的列<br>
• 3. 输入原状态值（如"待定"）和目标状态值（如"进行中"）<br>
• 4. 点击"预览变更"查看将要修改的数据<br>
• 5. 确认后点击"执行变更并保存"<br><br>
<b>注意:</b> 会直接修改原Excel文件，建议先备份！
        """)
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # 数据预览区域
        preview_group = QGroupBox("数据预览")
        preview_layout = QVBoxLayout()
        
        preview_button_layout = QHBoxLayout()
        self.preview_btn = QPushButton("预览变更")
        self.preview_btn.clicked.connect(self.preview_changes)
        self.preview_btn.setEnabled(False)
        preview_button_layout.addWidget(self.preview_btn)
        
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
        
        self.start_btn = QPushButton("执行变更并保存")
        self.start_btn.clicked.connect(self.execute_change)
        self.start_btn.setEnabled(False)
        button_layout.addWidget(self.start_btn)
        
        self.save_as_btn = QPushButton("另存为新文件")
        self.save_as_btn.clicked.connect(self.save_as_new_file)
        self.save_as_btn.setEnabled(False)
        button_layout.addWidget(self.save_as_btn)
        
        button_layout.addStretch()
        
        layout.addLayout(button_layout)
        
        # 结果显示区域
        self.result_label = QLabel("")
        self.result_label.setAlignment(Qt.AlignCenter)
        font = QFont()
        font.setBold(True)
        self.result_label.setFont(font)
        layout.addWidget(self.result_label)
        
        layout.addStretch()
        
        self.setLayout(layout)
    
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
            
            try:
                if file_path.endswith('.xlsx'):
                    self.excel_data = pd.read_excel(file_path, engine='openpyxl')
                else:
                    self.excel_data = pd.read_excel(file_path, engine='xlrd')
                
                # 填充列选择下拉框
                self.column_combo.clear()
                for i, col in enumerate(self.excel_data.columns):
                    self.column_combo.addItem(f"{chr(65+i) if i < 26 else ''}列: {col}", i)
                
                self.column_combo.setEnabled(True)
                self.update_button_states()
                self.result_label.setText("")
                
                # 清空预览表格
                self.records_table.setRowCount(0)
                self.record_count_label.setText("")
                
            except Exception as e:
                QMessageBox.critical(self, "文件读取失败", f"读取Excel文件失败：{str(e)}")
                self.excel_file = None
                self.excel_data = None
                self.file_label.setText("未选择文件")
                self.file_label.setStyleSheet("color: gray;")
    
    def on_column_changed(self):
        """列选择变化"""
        self.update_button_states()
        # 清空预览
        self.records_table.setRowCount(0)
        self.record_count_label.setText("")
        self.start_btn.setEnabled(False)
        self.save_as_btn.setEnabled(False)
    
    def update_button_states(self):
        """更新按钮状态"""
        file_selected = self.excel_file is not None and self.excel_data is not None
        column_selected = self.column_combo.currentIndex() >= 0
        
        self.preview_btn.setEnabled(file_selected and column_selected)
    
    def preview_changes(self):
        """预览变更"""
        if self.excel_data is None:
            return
        
        from_status = self.from_status_edit.text().strip()
        to_status = self.to_status_edit.text().strip()
        
        if not from_status:
            QMessageBox.warning(self, "警告", "请输入原状态值！")
            return
        
        if not to_status:
            QMessageBox.warning(self, "警告", "请输入目标状态值！")
            return
        
        column_index = self.column_combo.currentData()
        if column_index is None:
            return
        
        column_name = self.excel_data.columns[column_index]
        
        try:
            # 查找匹配的行
            matching_rows = []
            for index, row in self.excel_data.iterrows():
                cell_value = row.iloc[column_index]
                cell_str = str(cell_value).strip() if pd.notna(cell_value) else ""
                
                if cell_str == from_status:
                    matching_rows.append((index, cell_str))
            
            # 显示预览
            headers = ["Excel行号", f"当前值 ({column_name})", "变更后"]
            self.records_table.setRowCount(len(matching_rows))
            self.records_table.setColumnCount(len(headers))
            self.records_table.setHorizontalHeaderLabels(headers)
            
            for row_idx, (excel_row, current_value) in enumerate(matching_rows):
                self.records_table.setItem(row_idx, 0, QTableWidgetItem(str(excel_row + 2)))  # Excel行号从2开始（1是表头）
                self.records_table.setItem(row_idx, 1, QTableWidgetItem(current_value))
                
                new_item = QTableWidgetItem(to_status)
                new_item.setForeground(Qt.blue)
                self.records_table.setItem(row_idx, 2, new_item)
            
            # 调整列宽
            header = self.records_table.horizontalHeader()
            header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(1, QHeaderView.Stretch)
            header.setSectionResizeMode(2, QHeaderView.Stretch)
            
            self.record_count_label.setText(f"找到 {len(matching_rows)} 条待变更记录")
            
            if matching_rows:
                self.start_btn.setEnabled(True)
                self.save_as_btn.setEnabled(True)
                self.result_label.setText("")
            else:
                self.start_btn.setEnabled(False)
                self.save_as_btn.setEnabled(False)
                self.result_label.setText("未找到匹配的记录")
                self.result_label.setStyleSheet("color: orange;")
            
        except Exception as e:
            QMessageBox.critical(self, "预览失败", f"预览数据失败：{str(e)}")
    
    def execute_change(self):
        """执行变更并保存到原文件"""
        reply = QMessageBox.question(
            self,
            "确认变更",
            f"确定要修改原文件吗？\n\n"
            f"文件: {os.path.basename(self.excel_file)}\n\n"
            f"建议先备份原文件！",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        self.do_change_and_save(self.excel_file)
    
    def save_as_new_file(self):
        """另存为新文件"""
        base_name = os.path.splitext(os.path.basename(self.excel_file))[0]
        default_name = f"{base_name}_状态已修改_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "另存为",
            default_name,
            "Excel文件 (*.xlsx)"
        )
        
        if file_path:
            self.do_change_and_save(file_path)
    
    def do_change_and_save(self, save_path):
        """执行变更并保存"""
        from_status = self.from_status_edit.text().strip()
        to_status = self.to_status_edit.text().strip()
        column_index = self.column_combo.currentData()
        
        try:
            # 执行变更
            change_count = 0
            column_name = self.excel_data.columns[column_index]
            
            for index, row in self.excel_data.iterrows():
                cell_value = row.iloc[column_index]
                cell_str = str(cell_value).strip() if pd.notna(cell_value) else ""
                
                if cell_str == from_status:
                    self.excel_data.at[index, column_name] = to_status
                    change_count += 1
            
            # 保存文件
            self.excel_data.to_excel(save_path, index=False, engine='openpyxl')
            
            result_text = f"变更完成！\n共修改 {change_count} 条记录\n已保存到: {os.path.basename(save_path)}"
            self.result_label.setText(result_text)
            self.result_label.setStyleSheet("color: green;")
            
            QMessageBox.information(self, "成功", result_text)
            
            # 清空预览
            self.records_table.setRowCount(0)
            self.record_count_label.setText("")
            self.start_btn.setEnabled(False)
            self.save_as_btn.setEnabled(False)
            
        except Exception as e:
            QMessageBox.critical(self, "保存失败", f"保存文件失败：{str(e)}")
            self.result_label.setText("保存失败！")
            self.result_label.setStyleSheet("color: red;")
