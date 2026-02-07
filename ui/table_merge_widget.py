"""
表格对比合并功能界面
"""
import os
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog, 
                             QListWidget, QListWidgetItem, QCheckBox, QSplitter,
                             QScrollArea)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
import pandas as pd


class TableMergeWorker(QThread):
    """表格合并工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    
    def __init__(self, table1_file, table2_file, match_columns1, match_columns2, selected_fields, output_file):
        super().__init__()
        self.table1_file = table1_file
        self.table2_file = table2_file
        self.match_columns1 = match_columns1  # 表1的匹配字段列表
        self.match_columns2 = match_columns2  # 表2的匹配字段列表
        self.selected_fields = selected_fields  # 从表1选择的要补充的字段
        self.output_file = output_file
        self.results = {
            'total_table2_rows': 0,
            'matched_count': 0,
            'unmatched_count': 0,
            'merged_data': [],
            'table1_unmatched': []  # 表1中未匹配的记录
        }
    
    def run(self):
        try:
            self.log_message.emit("开始读取表1...")
            self.progress.emit(10)
            
            # 读取表1
            df1 = pd.read_excel(self.table1_file, engine='openpyxl')
            self.log_message.emit(f"成功读取表1，共 {len(df1)} 行数据")
            
            self.log_message.emit("开始读取表2...")
            self.progress.emit(20)
            
            # 读取表2
            df2 = pd.read_excel(self.table2_file, engine='openpyxl')
            self.log_message.emit(f"成功读取表2，共 {len(df2)} 行数据")
            self.results['total_table2_rows'] = len(df2)
            
            self.progress.emit(30)
            
            # 检查匹配字段是否存在
            missing_match_fields1 = [field for field in self.match_columns1 if field not in df1.columns]
            if missing_match_fields1:
                raise Exception(f"表1中未找到匹配字段: {', '.join(missing_match_fields1)}")
            
            missing_match_fields2 = [field for field in self.match_columns2 if field not in df2.columns]
            if missing_match_fields2:
                raise Exception(f"表2中未找到匹配字段: {', '.join(missing_match_fields2)}")
            
            # 检查匹配字段数量是否一致
            if len(self.match_columns1) != len(self.match_columns2):
                raise Exception(f"表1和表2的匹配字段数量不一致: {len(self.match_columns1)} vs {len(self.match_columns2)}")
            
            # 检查选择的字段是否存在于表1中
            missing_fields = [field for field in self.selected_fields if field not in df1.columns]
            if missing_fields:
                raise Exception(f"表1中未找到以下字段: {', '.join(missing_fields)}")
            
            self.log_message.emit("开始合并数据...")
            self.log_message.emit(f"匹配字段: 表1[{', '.join(self.match_columns1)}] ↔ 表2[{', '.join(self.match_columns2)}]")
            self.progress.emit(40)
            
            # 创建表1的查找字典，以组合匹配字段为键
            table1_dict = {}
            table1_used_keys = set()  # 跟踪被使用的表1记录
            
            for index, row in df1.iterrows():
                # 创建组合键
                key_parts = []
                for col in self.match_columns1:
                    value = str(row[col]).strip() if pd.notna(row[col]) else ""
                    key_parts.append(value)
                
                key = "||".join(key_parts)  # 使用特殊分隔符连接
                
                if key and key != "||".join([""] * len(self.match_columns1)):  # 不是全空
                    # 保存完整的行数据和选择的字段
                    selected_data = {field: row[field] for field in self.selected_fields}
                    table1_dict[key] = {
                        'selected_data': selected_data,
                        'full_row': row.to_dict(),
                        'row_index': index
                    }
            
            self.log_message.emit(f"表1索引创建完成，共 {len(table1_dict)} 个有效匹配键")
            
            # 处理表2的每一行
            merged_data = []
            matched_count = 0
            
            for index, row in df2.iterrows():
                # 获取表2的组合匹配值
                key_parts = []
                for col in self.match_columns2:
                    value = str(row[col]).strip() if pd.notna(row[col]) else ""
                    key_parts.append(value)
                
                match_key = "||".join(key_parts)
                
                # 创建合并后的行数据，先复制表2的所有数据
                merged_row = row.to_dict()
                
                # 查找表1中的匹配数据
                if match_key and match_key != "||".join([""] * len(self.match_columns2)) and match_key in table1_dict:
                    # 找到匹配，补充表1的字段
                    table1_info = table1_dict[match_key]
                    table1_data = table1_info['selected_data']
                    
                    for field, value in table1_data.items():
                        # 添加字段时加上前缀以区分来源
                        merged_row[f"表1_{field}"] = value
                    
                    merged_row['匹配状态'] = '已匹配'
                    matched_count += 1
                    
                    # 记录这个表1记录已被使用
                    table1_used_keys.add(match_key)
                    
                    if matched_count % 100 == 0:
                        self.log_message.emit(f"已处理 {matched_count} 个匹配记录")
                else:
                    # 未找到匹配，补充空值
                    for field in self.selected_fields:
                        merged_row[f"表1_{field}"] = ''
                    merged_row['匹配状态'] = '未匹配'
                
                merged_data.append(merged_row)
                
                # 更新进度
                progress = 40 + int((index + 1) / len(df2) * 40)
                self.progress.emit(progress)
            
            self.results['matched_count'] = matched_count
            self.results['unmatched_count'] = len(df2) - matched_count
            self.results['merged_data'] = merged_data
            
            # 收集表1中未匹配的记录
            self.log_message.emit("收集表1未匹配记录...")
            table1_unmatched = []
            
            for key, info in table1_dict.items():
                if key not in table1_used_keys:
                    # 这个表1记录没有被任何表2记录匹配
                    unmatched_row = info['full_row'].copy()
                    unmatched_row['未匹配原因'] = '在表2中未找到对应记录'
                    unmatched_row['匹配键'] = key
                    table1_unmatched.append(unmatched_row)
            
            self.results['table1_unmatched'] = table1_unmatched
            self.log_message.emit(f"表1未匹配记录: {len(table1_unmatched)} 条")
            
            self.log_message.emit("开始导出结果...")
            self.progress.emit(85)
            
            # 导出合并结果
            result_df = pd.DataFrame(merged_data)
            result_df.to_excel(self.output_file, index=False, engine='openpyxl')
            
            self.log_message.emit(f"结果已导出到: {self.output_file}")
            self.progress.emit(100)
            
            self.finished.emit(self.results)
            
        except Exception as e:
            self.error.emit(str(e))


class TableMergeWidget(QWidget):
    """表格对比合并功能组件"""
    
    def __init__(self):
        super().__init__()
        self.table1_file = ""
        self.table2_file = ""
        self.table1_columns = []
        self.table2_columns = []
        self.output_file = ""
        self.init_ui()
    
    def init_ui(self):
        """初始化UI"""
        layout = QVBoxLayout()
        
        # 文件选择区域
        file_group = QGroupBox("文件选择")
        file_layout = QFormLayout()
        
        # 表1文件选择
        table1_layout = QHBoxLayout()
        self.table1_label = QLabel("未选择文件")
        self.table1_label.setStyleSheet("color: gray;")
        table1_layout.addWidget(self.table1_label)
        
        self.select_table1_btn = QPushButton("选择表1")
        self.select_table1_btn.clicked.connect(self.select_table1_file)
        table1_layout.addWidget(self.select_table1_btn)
        
        file_layout.addRow("表1 (数据源):", table1_layout)
        
        # 表2文件选择
        table2_layout = QHBoxLayout()
        self.table2_label = QLabel("未选择文件")
        self.table2_label.setStyleSheet("color: gray;")
        table2_layout.addWidget(self.table2_label)
        
        self.select_table2_btn = QPushButton("选择表2")
        self.select_table2_btn.clicked.connect(self.select_table2_file)
        table2_layout.addWidget(self.select_table2_btn)
        
        file_layout.addRow("表2 (目标表):", table2_layout)
        
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
        
        # 配置区域
        config_group = QGroupBox("匹配配置")
        config_main_layout = QVBoxLayout()
        
        # 匹配字段说明
        match_info_label = QLabel("设置匹配字段对（支持多组字段组合匹配）:")
        match_info_label.setStyleSheet("font-weight: bold; color: blue;")
        config_main_layout.addWidget(match_info_label)
        
        # 创建滚动区域用于匹配字段对
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setMinimumHeight(120)  # 减少最小高度为按钮留空间
        scroll_area.setMaximumHeight(280)  # 减少最大高度为按钮留空间
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_area.setStyleSheet("""
            QScrollArea {
                border: 1px solid #ccc;
                border-radius: 5px;
                background-color: #f9f9f9;
            }
        """)
        
        # 匹配字段对容器
        scroll_widget = QWidget()
        scroll_widget.setStyleSheet("background-color: #f9f9f9;")
        self.match_pairs_layout = QVBoxLayout()
        self.match_pairs_layout.setSpacing(8)  # 适当间距
        self.match_pairs_layout.setContentsMargins(5, 5, 5, 5)  # 添加容器边距
        self.match_pairs = []  # 存储匹配字段对
        
        scroll_widget.setLayout(self.match_pairs_layout)
        scroll_area.setWidget(scroll_widget)
        config_main_layout.addWidget(scroll_area)
        
        # 添加第一组匹配字段（在滚动区域创建后立即添加）
        self.add_match_pair()
        
        # 添加/删除按钮区域（放在滚动区域下方）
        buttons_container = QWidget()
        buttons_container.setStyleSheet("""
            QWidget {
                background-color: #f0f0f0;
                border: 1px solid #ddd;
                border-radius: 5px;
                margin-top: 5px;
            }
        """)
        buttons_layout = QHBoxLayout()
        buttons_layout.setContentsMargins(10, 8, 10, 8)
        
        self.add_pair_btn = QPushButton("+ 添加字段对")
        self.add_pair_btn.clicked.connect(self.add_match_pair)
        self.add_pair_btn.setEnabled(False)
        self.add_pair_btn.setFixedSize(100, 30)
        self.add_pair_btn.setStyleSheet("""
            QPushButton {
                background-color: #007ACC;
                color: white;
                border: none;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #005a9e;
            }
            QPushButton:disabled {
                background-color: #ccc;
                color: #666;
            }
        """)
        buttons_layout.addWidget(self.add_pair_btn)
        
        self.remove_pair_btn = QPushButton("- 删除最后一组")
        self.remove_pair_btn.clicked.connect(self.remove_last_match_pair)
        self.remove_pair_btn.setEnabled(False)
        self.remove_pair_btn.setFixedSize(110, 30)
        self.remove_pair_btn.setStyleSheet("""
            QPushButton {
                background-color: #dc3545;
                color: white;
                border: none;
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
        buttons_layout.addWidget(self.remove_pair_btn)
        
        buttons_layout.addStretch()
        
        # 匹配字段对计数显示
        self.pairs_count_label = QLabel("当前: 1 组")
        self.pairs_count_label.setStyleSheet("color: green; font-weight: bold; font-size: 12px;")
        buttons_layout.addWidget(self.pairs_count_label)
        
        buttons_container.setLayout(buttons_layout)
        config_main_layout.addWidget(buttons_container)
        
        # 匹配说明
        match_help_label = QLabel("""
<b>使用说明:</b><br>
• 每组匹配字段对表示一个匹配条件，多组时所有组都必须匹配才算成功<br>
• 例如: [姓名↔员工姓名] AND [部门↔所属部门]<br>
• 支持滚动查看多个匹配字段对
        """)
        match_help_label.setWordWrap(True)
        match_help_label.setStyleSheet("color: gray; font-size: 10px;")
        config_main_layout.addWidget(match_help_label)
        
        config_group.setLayout(config_main_layout)
        layout.addWidget(config_group)
        
        # 字段选择区域
        field_group = QGroupBox("选择要补充的表1字段")
        field_layout = QVBoxLayout()
        
        # 创建分割器
        splitter = QSplitter(Qt.Horizontal)
        
        # 左侧：可选字段列表
        left_widget = QWidget()
        left_layout = QVBoxLayout()
        left_layout.addWidget(QLabel("表1可用字段:"))
        
        self.available_fields_list = QListWidget()
        self.available_fields_list.setEnabled(False)
        left_layout.addWidget(self.available_fields_list)
        
        # 添加全选/取消全选按钮
        select_buttons_layout = QHBoxLayout()
        self.select_all_btn = QPushButton("全选")
        self.select_all_btn.clicked.connect(self.select_all_fields)
        self.select_all_btn.setEnabled(False)
        select_buttons_layout.addWidget(self.select_all_btn)
        
        self.clear_all_btn = QPushButton("清空")
        self.clear_all_btn.clicked.connect(self.clear_all_fields)
        self.clear_all_btn.setEnabled(False)
        select_buttons_layout.addWidget(self.clear_all_btn)
        
        left_layout.addLayout(select_buttons_layout)
        left_widget.setLayout(left_layout)
        
        # 右侧：已选字段列表
        right_widget = QWidget()
        right_layout = QVBoxLayout()
        right_layout.addWidget(QLabel("已选择字段:"))
        
        self.selected_fields_list = QListWidget()
        right_layout.addWidget(self.selected_fields_list)
        
        right_widget.setLayout(right_layout)
        
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setSizes([300, 300])
        
        field_layout.addWidget(splitter)
        field_group.setLayout(field_layout)
        layout.addWidget(field_group)
        
        # 功能说明（简化版）
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        
        info_label = QLabel("""
<b>使用步骤:</b> 1.选择表1和表2文件 → 2.设置匹配字段对 → 3.选择要补充的字段 → 4.开始合并<br>
<b>匹配规则:</b> 支持单字段或多字段组合匹配，所有字段对都必须匹配才算成功<br>
<b>输出结果:</b> 以表2为基础，补充表1的选定字段（带"表1_"前缀）和"匹配状态"字段
        """)
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: #666; font-size: 11px;")
        info_layout.addWidget(info_label)
        
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # 操作按钮区域
        button_layout = QHBoxLayout()
        
        self.start_btn = QPushButton("开始合并")
        self.start_btn.clicked.connect(self.start_merge)
        self.start_btn.setEnabled(False)
        button_layout.addWidget(self.start_btn)
        
        button_layout.addStretch()
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        button_layout.addWidget(self.progress_bar)
        
        layout.addLayout(button_layout)
        
        # 结果显示区域
        result_group = QGroupBox("处理结果")
        result_layout = QVBoxLayout()
        
        self.result_label = QLabel("等待开始...")
        self.result_label.setStyleSheet("font-weight: bold;")
        result_layout.addWidget(self.result_label)
        
        # 导出按钮区域
        export_layout = QHBoxLayout()
        
        self.export_table1_unmatched_btn = QPushButton("导出表1未匹配记录")
        self.export_table1_unmatched_btn.clicked.connect(self.export_table1_unmatched)
        self.export_table1_unmatched_btn.setEnabled(False)
        self.export_table1_unmatched_btn.setStyleSheet("""
            QPushButton {
                background-color: #ff8800;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #e67700;
            }
            QPushButton:disabled {
                background-color: #ccc;
                color: #666;
            }
        """)
        export_layout.addWidget(self.export_table1_unmatched_btn)
        
        export_layout.addStretch()
        result_layout.addLayout(export_layout)
        
        result_group.setLayout(result_layout)
        layout.addWidget(result_group)
        
        # 日志区域
        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout()
        
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(100)  # 减少日志区域高度
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)
        
        self.setLayout(layout)
    
    def add_match_pair(self):
        """添加一组匹配字段对"""
        pair_widget = QWidget()
        pair_widget.setMinimumHeight(50)  # 增加最小高度
        pair_widget.setMaximumHeight(60)  # 增加最大高度
        pair_widget.setStyleSheet("""
            QWidget {
                background-color: white;
                border: 1px solid #ddd;
                border-radius: 5px;
                margin: 2px;
            }
            QWidget:hover {
                border: 1px solid #007ACC;
            }
        """)
        pair_layout = QHBoxLayout()
        pair_layout.setContentsMargins(10, 5, 10, 5)  # 增加边距
        
        # 序号标签
        pair_num = len(self.match_pairs) + 1
        num_label = QLabel(f"{pair_num}.")
        num_label.setFixedWidth(30)
        num_label.setStyleSheet("font-weight: bold; color: #666; font-size: 14px;")
        pair_layout.addWidget(num_label)
        
        # 表1标签和下拉框
        table1_label = QLabel("表1:")
        table1_label.setFixedWidth(35)
        table1_label.setStyleSheet("font-weight: bold; color: #333;")
        pair_layout.addWidget(table1_label)
        
        table1_combo = QComboBox()
        table1_combo.setEnabled(False)
        table1_combo.setMinimumWidth(200)  # 增加宽度
        table1_combo.setMinimumHeight(30)  # 增加高度
        table1_combo.currentTextChanged.connect(self.update_start_button_state)
        table1_combo.setStyleSheet("""
            QComboBox {
                padding: 5px;
                border: 1px solid #ccc;
                border-radius: 3px;
                background-color: white;
            }
            QComboBox:focus {
                border: 2px solid #007ACC;
            }
        """)
        pair_layout.addWidget(table1_combo)
        
        # 箭头标签
        arrow_label = QLabel("↔")
        arrow_label.setAlignment(Qt.AlignCenter)
        arrow_label.setFixedWidth(30)
        arrow_label.setStyleSheet("font-weight: bold; color: #007ACC; font-size: 16px;")
        pair_layout.addWidget(arrow_label)
        
        # 表2标签和下拉框
        table2_label = QLabel("表2:")
        table2_label.setFixedWidth(35)
        table2_label.setStyleSheet("font-weight: bold; color: #333;")
        pair_layout.addWidget(table2_label)
        
        table2_combo = QComboBox()
        table2_combo.setEnabled(False)
        table2_combo.setMinimumWidth(200)  # 增加宽度
        table2_combo.setMinimumHeight(30)  # 增加高度
        table2_combo.currentTextChanged.connect(self.update_start_button_state)
        table2_combo.setStyleSheet("""
            QComboBox {
                padding: 5px;
                border: 1px solid #ccc;
                border-radius: 3px;
                background-color: white;
            }
            QComboBox:focus {
                border: 2px solid #007ACC;
            }
        """)
        pair_layout.addWidget(table2_combo)
        
        # 删除按钮（针对单个字段对）
        delete_btn = QPushButton("×")
        delete_btn.setFixedSize(25, 25)
        delete_btn.setStyleSheet("""
            QPushButton {
                background-color: #ff4444;
                color: white;
                border: none;
                border-radius: 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #cc3333;
            }
        """)
        delete_btn.setToolTip("删除此匹配字段对")
        delete_btn.clicked.connect(lambda: self.remove_match_pair(pair_info))
        pair_layout.addWidget(delete_btn)
        
        pair_widget.setLayout(pair_layout)
        
        # 存储匹配字段对信息
        pair_info = {
            'widget': pair_widget,
            'table1_combo': table1_combo,
            'table2_combo': table2_combo,
            'num_label': num_label,
            'delete_btn': delete_btn
        }
        
        self.match_pairs.append(pair_info)
        self.match_pairs_layout.addWidget(pair_widget)
        
        # 如果已有表格数据，填充下拉框
        if hasattr(self, 'table1_columns') and self.table1_columns:
            table1_combo.addItems(self.table1_columns)
            table1_combo.setEnabled(True)
        
        if hasattr(self, 'table2_columns') and self.table2_columns:
            table2_combo.addItems(self.table2_columns)
            table2_combo.setEnabled(True)
        
        self.update_pair_buttons()
        self.update_pairs_count()
        self.update_start_button_state()
    
    def remove_last_match_pair(self):
        """删除最后一组匹配字段对"""
        if len(self.match_pairs) > 1:  # 至少保留一组
            last_pair = self.match_pairs.pop()
            last_pair['widget'].setParent(None)
            last_pair['widget'].deleteLater()
            
            # 更新序号
            for i, pair in enumerate(self.match_pairs):
                pair['num_label'].setText(f"{i+1}.")
            
            self.update_pair_buttons()
            self.update_pairs_count()
            self.update_start_button_state()
    
    def remove_match_pair(self, pair_info):
        """删除指定的匹配字段对"""
        if len(self.match_pairs) <= 1:  # 至少保留一组
            return
        
        # 从列表中移除
        if pair_info in self.match_pairs:
            self.match_pairs.remove(pair_info)
            pair_info['widget'].setParent(None)
            pair_info['widget'].deleteLater()
            
            # 更新序号
            for i, pair in enumerate(self.match_pairs):
                pair['num_label'].setText(f"{i+1}.")
            
            self.update_pair_buttons()
            self.update_pairs_count()
            self.update_start_button_state()
    
    def update_pairs_count(self):
        """更新匹配字段对计数显示"""
        if hasattr(self, 'pairs_count_label'):
            count = len(self.match_pairs)
            self.pairs_count_label.setText(f"当前: {count} 组")
            
            # 根据数量调整颜色
            if count == 1:
                color = "#666"
            elif count <= 3:
                color = "green"
            else:
                color = "#ff8800"
            
            self.pairs_count_label.setStyleSheet(f"color: {color}; font-weight: bold;")
    
    def update_pair_buttons(self):
        """更新添加/删除按钮状态"""
        # 检查按钮是否已创建
        if not hasattr(self, 'add_pair_btn') or not hasattr(self, 'remove_pair_btn'):
            return
        
        has_tables = bool(hasattr(self, 'table1_columns') and self.table1_columns and 
                         hasattr(self, 'table2_columns') and self.table2_columns)
        
        self.add_pair_btn.setEnabled(has_tables)
        self.remove_pair_btn.setEnabled(len(self.match_pairs) > 1)
        
        # 管理单个删除按钮的可见性
        for pair in self.match_pairs:
            if 'delete_btn' in pair:
                pair['delete_btn'].setVisible(len(self.match_pairs) > 1)
    
    def populate_match_combos(self):
        """填充所有匹配字段下拉框"""
        for pair in self.match_pairs:
            # 填充表1下拉框
            if hasattr(self, 'table1_columns') and self.table1_columns:
                pair['table1_combo'].clear()
                pair['table1_combo'].addItems(self.table1_columns)
                pair['table1_combo'].setEnabled(True)
            
            # 填充表2下拉框
            if hasattr(self, 'table2_columns') and self.table2_columns:
                pair['table2_combo'].clear()
                pair['table2_combo'].addItems(self.table2_columns)
                pair['table2_combo'].setEnabled(True)
    
    def select_table1_file(self):
        """选择表1文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择表1 Excel文件",
            "",
            "Excel文件 (*.xlsx *.xls)"
        )
        
        if file_path:
            try:
                # 读取Excel表头
                df = pd.read_excel(file_path, engine='openpyxl', nrows=0)
                self.table1_columns = list(df.columns)
                
                # 更新UI
                self.table1_file = file_path
                self.table1_label.setText(os.path.basename(file_path))
                self.table1_label.setStyleSheet("color: green;")
                
                # 填充匹配字段下拉框
                self.populate_match_combos()
                self.update_pair_buttons()
                
                # 填充可选字段列表
                self.available_fields_list.clear()
                for column in self.table1_columns:
                    item = QListWidgetItem(column)
                    item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                    item.setCheckState(Qt.Unchecked)
                    self.available_fields_list.addItem(item)
                
                self.available_fields_list.setEnabled(True)
                self.select_all_btn.setEnabled(True)
                self.clear_all_btn.setEnabled(True)
                
                # 连接字段选择变化事件
                self.available_fields_list.itemChanged.connect(self.update_selected_fields)
                
                self.log_text.append(f"成功读取表1表头，共 {len(self.table1_columns)} 列")
                self.log_text.append(f"列名: {', '.join(self.table1_columns)}")
                
                self.update_start_button_state()
                
            except Exception as e:
                QMessageBox.critical(self, "错误", f"读取表1文件失败: {str(e)}")
                self.table1_file = ""
                self.table1_label.setText("未选择文件")
                self.table1_label.setStyleSheet("color: gray;")
    
    def select_table2_file(self):
        """选择表2文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择表2 Excel文件",
            "",
            "Excel文件 (*.xlsx *.xls)"
        )
        
        if file_path:
            try:
                # 读取Excel表头
                df = pd.read_excel(file_path, engine='openpyxl', nrows=0)
                self.table2_columns = list(df.columns)
                
                # 更新UI
                self.table2_file = file_path
                self.table2_label.setText(os.path.basename(file_path))
                self.table2_label.setStyleSheet("color: green;")
                
                # 填充匹配字段下拉框
                self.populate_match_combos()
                self.update_pair_buttons()
                
                self.log_text.append(f"成功读取表2表头，共 {len(self.table2_columns)} 列")
                self.log_text.append(f"列名: {', '.join(self.table2_columns)}")
                
                self.update_start_button_state()
                
            except Exception as e:
                QMessageBox.critical(self, "错误", f"读取表2文件失败: {str(e)}")
                self.table2_file = ""
                self.table2_label.setText("未选择文件")
                self.table2_label.setStyleSheet("color: gray;")
    
    def select_output_file(self):
        """选择输出文件"""
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "选择输出文件",
            f"表格合并结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel文件 (*.xlsx)"
        )
        
        if file_path:
            self.output_file = file_path
            self.output_label.setText(os.path.basename(file_path))
            self.output_label.setStyleSheet("color: blue;")
            
            self.update_start_button_state()
    
    def select_all_fields(self):
        """全选字段"""
        for i in range(self.available_fields_list.count()):
            item = self.available_fields_list.item(i)
            item.setCheckState(Qt.Checked)
    
    def clear_all_fields(self):
        """清空选择"""
        for i in range(self.available_fields_list.count()):
            item = self.available_fields_list.item(i)
            item.setCheckState(Qt.Unchecked)
    
    def update_selected_fields(self):
        """更新已选择字段列表"""
        self.selected_fields_list.clear()
        
        selected_fields = []
        for i in range(self.available_fields_list.count()):
            item = self.available_fields_list.item(i)
            if item.checkState() == Qt.Checked:
                selected_fields.append(item.text())
                self.selected_fields_list.addItem(item.text())
        
        self.update_start_button_state()
    
    def update_start_button_state(self):
        """更新开始按钮状态"""
        # 检查必要的控件是否已创建
        if not hasattr(self, 'available_fields_list') or not hasattr(self, 'start_btn'):
            return
        
        # 获取已选择的补充字段
        selected_fields = []
        for i in range(self.available_fields_list.count()):
            item = self.available_fields_list.item(i)
            if item and item.checkState() == Qt.Checked:
                selected_fields.append(item.text())
        
        # 检查所有匹配字段对是否都已选择
        all_pairs_selected = True
        for pair in self.match_pairs:
            if not pair['table1_combo'].currentText() or not pair['table2_combo'].currentText():
                all_pairs_selected = False
                break
        
        can_start = bool(
            self.table1_file and 
            self.table2_file and 
            self.output_file and
            len(selected_fields) > 0 and
            len(self.match_pairs) > 0 and
            all_pairs_selected
        )
        
        self.start_btn.setEnabled(can_start)
    
    def start_merge(self):
        """开始合并"""
        # 获取已选择的补充字段
        selected_fields = []
        for i in range(self.available_fields_list.count()):
            item = self.available_fields_list.item(i)
            if item and item.checkState() == Qt.Checked:
                selected_fields.append(item.text())
        
        # 获取所有匹配字段对
        match_columns1 = []
        match_columns2 = []
        
        for pair in self.match_pairs:
            table1_field = pair['table1_combo'].currentText()
            table2_field = pair['table2_combo'].currentText()
            if table1_field and table2_field:
                match_columns1.append(table1_field)
                match_columns2.append(table2_field)
        
        # 确认对话框
        reply = QMessageBox.question(
            self,
            "确认合并",
            f"确定要开始表格合并吗？\n\n"
            f"表1: {os.path.basename(self.table1_file)}\n"
            f"表2: {os.path.basename(self.table2_file)}\n"
            f"输出: {os.path.basename(self.output_file)}\n"
            f"匹配字段: {' AND '.join([f'{c1}↔{c2}' for c1, c2 in zip(match_columns1, match_columns2)])}\n"
            f"补充字段: {len(selected_fields)} 个\n\n"
            f"操作不可撤销，请确认！",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        # 禁用按钮，显示进度条
        self.start_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在合并中...")
        self.log_text.clear()
        
        # 启动工作线程
        self.worker = TableMergeWorker(
            self.table1_file, self.table2_file, match_columns1, match_columns2, 
            selected_fields, self.output_file
        )
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log_message.connect(self.append_log)
        self.worker.finished.connect(self.on_merge_finished)
        self.worker.error.connect(self.on_merge_error)
        self.worker.start()
    
    def append_log(self, message):
        """添加日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
    
    def on_merge_finished(self, results):
        """合并完成"""
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.merge_results = results  # 保存结果用于导出
        
        total_rows = results['total_table2_rows']
        matched_count = results['matched_count']
        unmatched_count = results['unmatched_count']
        table1_unmatched_count = len(results['table1_unmatched'])
        
        result_text = f"合并完成！表2总行数: {total_rows}, 已匹配: {matched_count}, 未匹配: {unmatched_count}"
        self.result_label.setText(result_text)
        self.result_label.setStyleSheet("color: green; font-weight: bold;")
        
        # 启用导出按钮
        if table1_unmatched_count > 0:
            self.export_table1_unmatched_btn.setEnabled(True)
            self.export_table1_unmatched_btn.setText(f"导出表1未匹配记录 ({table1_unmatched_count}条)")
        
        self.log_text.append(f"\n=== 合并完成 ===")
        self.log_text.append(result_text)
        self.log_text.append(f"表2匹配率: {matched_count/total_rows*100:.1f}%")
        self.log_text.append(f"表1未匹配记录: {table1_unmatched_count} 条")
        
        QMessageBox.information(self, "完成", f"{result_text}\n表1未匹配记录: {table1_unmatched_count} 条\n\n结果已保存到:\n{self.output_file}")
    
    def export_table1_unmatched(self):
        """导出表1未匹配记录"""
        if not hasattr(self, 'merge_results') or not self.merge_results['table1_unmatched']:
            QMessageBox.warning(self, "警告", "没有可导出的表1未匹配记录")
            return
        
        # 选择保存文件
        default_name = f"表1未匹配记录_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存表1未匹配记录",
            default_name,
            "Excel文件 (*.xlsx)"
        )
        
        if file_path:
            try:
                # 导出未匹配记录
                unmatched_df = pd.DataFrame(self.merge_results['table1_unmatched'])
                unmatched_df.to_excel(file_path, index=False, engine='openpyxl')
                
                count = len(self.merge_results['table1_unmatched'])
                QMessageBox.information(
                    self, 
                    "导出成功", 
                    f"已成功导出 {count} 条表1未匹配记录到:\n{file_path}"
                )
                
                self.log_text.append(f"表1未匹配记录已导出到: {file_path}")
                
            except Exception as e:
                QMessageBox.critical(self, "导出失败", f"导出表1未匹配记录时出错: {str(e)}")
    
    def on_merge_error(self, error_message):
        """合并错误"""
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.result_label.setText("合并失败！")
        self.result_label.setStyleSheet("color: red; font-weight: bold;")
        
        QMessageBox.critical(self, "错误", f"合并过程出错: {error_message}")
        self.log_text.append(f"错误: {error_message}")
