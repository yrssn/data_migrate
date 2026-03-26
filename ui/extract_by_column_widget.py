"""
去除平台重复数据功能界面
根据某一列的值筛选并提取数据
"""
import os
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog,
                             QTableWidget, QTableWidgetItem, QHeaderView, QCheckBox,
                             QLineEdit, QListWidget, QListWidgetItem, QAbstractItemView,
                             QSplitter)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
import pandas as pd


class ExtractByColumnWidget(QWidget):
    """去除平台重复数据功能组件 - 根据列值筛选提取数据"""
    
    def __init__(self, db_manager=None):
        super().__init__()
        self.excel_file = None
        self.excel_data = None
        self.extracted_data = None
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
        
        # 筛选列选择
        self.column_combo = QComboBox()
        self.column_combo.setEnabled(False)
        self.column_combo.currentIndexChanged.connect(self.on_column_changed)
        file_layout.addRow("筛选列:", self.column_combo)
        
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # 筛选条件配置区域
        filter_group = QGroupBox("筛选条件")
        filter_layout = QVBoxLayout()
        
        # 输入筛选值
        input_layout = QHBoxLayout()
        input_layout.addWidget(QLabel("筛选值（多个用逗号分隔）:"))
        self.filter_values_edit = QLineEdit()
        self.filter_values_edit.setPlaceholderText("例如: TikTok, Amazon, SHEIN 或直接输入单个值")
        input_layout.addWidget(self.filter_values_edit)
        filter_layout.addLayout(input_layout)
        
        # 可选值列表（选择列后自动加载）
        values_layout = QHBoxLayout()
        
        # 左侧：列中的唯一值
        left_layout = QVBoxLayout()
        left_layout.addWidget(QLabel("列中的唯一值（点击添加到筛选）:"))
        self.unique_values_list = QListWidget()
        self.unique_values_list.setMaximumHeight(120)
        self.unique_values_list.setSelectionMode(QAbstractItemView.MultiSelection)
        self.unique_values_list.itemClicked.connect(self.on_unique_value_clicked)
        left_layout.addWidget(self.unique_values_list)
        values_layout.addLayout(left_layout)
        
        # 右侧：已选择的值
        right_layout = QVBoxLayout()
        right_layout.addWidget(QLabel("已选择的筛选值:"))
        self.selected_values_list = QListWidget()
        self.selected_values_list.setMaximumHeight(120)
        self.selected_values_list.itemDoubleClicked.connect(self.remove_selected_value)
        right_layout.addWidget(self.selected_values_list)
        
        btn_layout = QHBoxLayout()
        self.clear_selected_btn = QPushButton("清空已选")
        self.clear_selected_btn.clicked.connect(self.clear_selected_values)
        btn_layout.addWidget(self.clear_selected_btn)
        btn_layout.addStretch()
        right_layout.addLayout(btn_layout)
        
        values_layout.addLayout(right_layout)
        filter_layout.addLayout(values_layout)
        
        filter_group.setLayout(filter_layout)
        layout.addWidget(filter_group)
        
        # 功能说明区域
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        
        info_label = QLabel("""
<b>去除平台重复数据说明:</b><br>
• 1. 选择Excel文件<br>
• 2. 选择要筛选的列（如"平台"列）<br>
• 3. 输入或选择要提取的值（如"TikTok"）<br>
• 4. 点击"提取数据"将匹配的行提取出来<br>
• 5. 可以选择"导出提取结果"或"导出剩余数据"<br><br>
<b>提示:</b> 双击已选择的值可以移除
        """)
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # 操作按钮区域
        button_layout = QHBoxLayout()
        
        self.extract_btn = QPushButton("提取数据")
        self.extract_btn.clicked.connect(self.extract_data)
        self.extract_btn.setEnabled(False)
        button_layout.addWidget(self.extract_btn)
        
        self.export_extracted_btn = QPushButton("导出提取结果")
        self.export_extracted_btn.clicked.connect(self.export_extracted_data)
        self.export_extracted_btn.setEnabled(False)
        button_layout.addWidget(self.export_extracted_btn)
        
        self.export_remaining_btn = QPushButton("导出剩余数据")
        self.export_remaining_btn.clicked.connect(self.export_remaining_data)
        self.export_remaining_btn.setEnabled(False)
        button_layout.addWidget(self.export_remaining_btn)
        
        button_layout.addStretch()
        
        layout.addLayout(button_layout)
        
        # 数据预览区域
        preview_group = QGroupBox("数据预览")
        preview_layout = QVBoxLayout()
        
        self.preview_label = QLabel("提取结果将显示在这里")
        preview_layout.addWidget(self.preview_label)
        
        self.preview_table = QTableWidget()
        self.preview_table.setMaximumHeight(200)
        preview_layout.addWidget(self.preview_table)
        
        preview_group.setLayout(preview_layout)
        layout.addWidget(preview_group)
        
        # 结果显示区域
        self.result_label = QLabel("")
        self.result_label.setAlignment(Qt.AlignCenter)
        font = QFont()
        font.setBold(True)
        self.result_label.setFont(font)
        layout.addWidget(self.result_label)
        
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
                self.column_combo.addItem("请选择列", -1)
                for i, col in enumerate(self.excel_data.columns):
                    col_letter = chr(65 + i) if i < 26 else f"Col{i+1}"
                    self.column_combo.addItem(f"{col_letter}列: {col}", i)
                
                self.column_combo.setEnabled(True)
                self.result_label.setText(f"已加载 {len(self.excel_data)} 行数据")
                self.result_label.setStyleSheet("color: blue;")
                
                # 清空之前的选择
                self.unique_values_list.clear()
                self.selected_values_list.clear()
                self.filter_values_edit.clear()
                self.preview_table.setRowCount(0)
                self.extracted_data = None
                self.export_extracted_btn.setEnabled(False)
                self.export_remaining_btn.setEnabled(False)
                
            except Exception as e:
                QMessageBox.critical(self, "文件读取失败", f"读取Excel文件失败：{str(e)}")
                self.excel_file = None
                self.excel_data = None
                self.file_label.setText("未选择文件")
                self.file_label.setStyleSheet("color: gray;")
    
    def on_column_changed(self):
        """列选择变化"""
        column_index = self.column_combo.currentData()
        
        if column_index is None or column_index < 0 or self.excel_data is None:
            self.unique_values_list.clear()
            self.extract_btn.setEnabled(False)
            return
        
        # 获取该列的唯一值
        try:
            column_name = self.excel_data.columns[column_index]
            unique_values = self.excel_data[column_name].dropna().unique()
            
            # 清空并填充唯一值列表
            self.unique_values_list.clear()
            for value in sorted([str(v) for v in unique_values]):
                if value.strip():
                    item = QListWidgetItem(value)
                    self.unique_values_list.addItem(item)
            
            self.extract_btn.setEnabled(True)
            
        except Exception as e:
            QMessageBox.warning(self, "警告", f"获取列值失败：{str(e)}")
    
    def on_unique_value_clicked(self, item):
        """点击唯一值，添加到已选择列表"""
        value = item.text()
        
        # 检查是否已存在
        for i in range(self.selected_values_list.count()):
            if self.selected_values_list.item(i).text() == value:
                return
        
        self.selected_values_list.addItem(value)
        self.update_filter_edit()
    
    def remove_selected_value(self, item):
        """双击移除已选择的值"""
        row = self.selected_values_list.row(item)
        self.selected_values_list.takeItem(row)
        self.update_filter_edit()
    
    def clear_selected_values(self):
        """清空已选择的值"""
        self.selected_values_list.clear()
        self.filter_values_edit.clear()
    
    def update_filter_edit(self):
        """更新筛选输入框"""
        values = []
        for i in range(self.selected_values_list.count()):
            values.append(self.selected_values_list.item(i).text())
        self.filter_values_edit.setText(", ".join(values))
    
    def get_filter_values(self):
        """获取筛选值列表"""
        text = self.filter_values_edit.text().strip()
        if not text:
            return []
        
        # 按逗号分隔并去除空白
        values = [v.strip() for v in text.split(",") if v.strip()]
        return values
    
    def extract_data(self):
        """提取数据"""
        if self.excel_data is None:
            QMessageBox.warning(self, "警告", "请先选择Excel文件！")
            return
        
        column_index = self.column_combo.currentData()
        if column_index is None or column_index < 0:
            QMessageBox.warning(self, "警告", "请选择筛选列！")
            return
        
        filter_values = self.get_filter_values()
        if not filter_values:
            QMessageBox.warning(self, "警告", "请输入或选择筛选值！")
            return
        
        try:
            column_name = self.excel_data.columns[column_index]
            
            # 筛选匹配的行
            mask = self.excel_data[column_name].astype(str).isin(filter_values)
            self.extracted_data = self.excel_data[mask].copy()
            self.remaining_data = self.excel_data[~mask].copy()
            
            # 显示预览
            self.show_preview(self.extracted_data)
            
            # 更新结果
            result_text = (
                f"提取完成！\n"
                f"原始数据: {len(self.excel_data)} 行\n"
                f"提取数据: {len(self.extracted_data)} 行\n"
                f"剩余数据: {len(self.remaining_data)} 行"
            )
            self.result_label.setText(result_text)
            self.result_label.setStyleSheet("color: green;")
            
            self.preview_label.setText(f"提取结果预览（共 {len(self.extracted_data)} 行，显示前20行）:")
            
            # 启用导出按钮
            self.export_extracted_btn.setEnabled(len(self.extracted_data) > 0)
            self.export_remaining_btn.setEnabled(len(self.remaining_data) > 0)
            
        except Exception as e:
            QMessageBox.critical(self, "提取失败", f"提取数据失败：{str(e)}")
    
    def show_preview(self, df):
        """显示数据预览"""
        preview_df = df.head(20)
        
        self.preview_table.setRowCount(len(preview_df))
        self.preview_table.setColumnCount(len(preview_df.columns))
        self.preview_table.setHorizontalHeaderLabels([str(col) for col in preview_df.columns])
        
        for i in range(len(preview_df)):
            for j, col in enumerate(preview_df.columns):
                value = preview_df.iloc[i, j]
                item = QTableWidgetItem(str(value) if pd.notna(value) else "")
                self.preview_table.setItem(i, j, item)
        
        self.preview_table.resizeColumnsToContents()
    
    def export_extracted_data(self):
        """导出提取的数据"""
        if self.extracted_data is None or len(self.extracted_data) == 0:
            QMessageBox.information(self, "提示", "没有提取的数据可导出")
            return
        
        base_name = os.path.splitext(os.path.basename(self.excel_file))[0]
        default_name = f"{base_name}_提取结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存提取结果",
            default_name,
            "Excel文件 (*.xlsx)"
        )
        
        if file_path:
            try:
                self.extracted_data.to_excel(file_path, index=False, engine='openpyxl')
                QMessageBox.information(self, "成功", f"提取结果已导出！\n共 {len(self.extracted_data)} 行\n保存到: {os.path.basename(file_path)}")
            except Exception as e:
                QMessageBox.critical(self, "导出失败", f"导出失败：{str(e)}")
    
    def export_remaining_data(self):
        """导出剩余的数据"""
        if self.remaining_data is None or len(self.remaining_data) == 0:
            QMessageBox.information(self, "提示", "没有剩余的数据可导出")
            return
        
        base_name = os.path.splitext(os.path.basename(self.excel_file))[0]
        default_name = f"{base_name}_剩余数据_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存剩余数据",
            default_name,
            "Excel文件 (*.xlsx)"
        )
        
        if file_path:
            try:
                self.remaining_data.to_excel(file_path, index=False, engine='openpyxl')
                QMessageBox.information(self, "成功", f"剩余数据已导出！\n共 {len(self.remaining_data)} 行\n保存到: {os.path.basename(file_path)}")
            except Exception as e:
                QMessageBox.critical(self, "导出失败", f"导出失败：{str(e)}")
