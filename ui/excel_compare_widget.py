"""
Excel表对比功能
"""
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QFileDialog, QTextEdit, QProgressBar,
                             QGroupBox, QFormLayout, QMessageBox, QListWidget,
                             QComboBox, QListWidgetItem)
from PyQt5.QtCore import QThread, pyqtSignal
import pandas as pd
from datetime import datetime


class ExcelCompareWorker(QThread):
    """Excel对比工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    
    def __init__(self, excel1_file, excel2_file, compare_columns, output_file):
        super().__init__()
        self.excel1_file = excel1_file
        self.excel2_file = excel2_file
        self.compare_columns = compare_columns  # [(excel1_col, excel2_col), ...]
        self.output_file = output_file
    
    def run(self):
        try:
            self.log_message.emit("开始读取Excel文件...")
            self.progress.emit(5)
            
            # 读取Excel文件1
            try:
                df1 = pd.read_excel(self.excel1_file, engine='openpyxl')
                self.log_message.emit(f"成功读取表1，共 {len(df1)} 行数据")
            except Exception as e:
                self.error.emit(f"读取表1失败: {str(e)}")
                return
            
            self.progress.emit(15)
            
            # 读取Excel文件2
            try:
                df2 = pd.read_excel(self.excel2_file, engine='openpyxl')
                self.log_message.emit(f"成功读取表2，共 {len(df2)} 行数据")
            except Exception as e:
                self.error.emit(f"读取表2失败: {str(e)}")
                return
            
            self.progress.emit(25)
            
            # 验证对比列是否存在
            for col1, col2 in self.compare_columns:
                if col1 not in df1.columns:
                    self.error.emit(f"表1中不存在列: {col1}")
                    return
                if col2 not in df2.columns:
                    self.error.emit(f"表2中不存在列: {col2}")
                    return
            
            self.log_message.emit(f"开始对比，对比条件: {len(self.compare_columns)} 个列")
            for col1, col2 in self.compare_columns:
                self.log_message.emit(f"  表1[{col1}] <-> 表2[{col2}]")
            
            self.progress.emit(35)
            
            # 构建表2的组合键集合（用于快速查找）
            table2_keys = set()
            for index, row in df2.iterrows():
                # 构建组合键：将所有对比列的值组合成元组
                key_values = []
                for _, col2 in self.compare_columns:
                    value = row[col2]
                    # 处理NaN值
                    if pd.isna(value):
                        key_values.append("")
                    else:
                        key_values.append(str(value).strip())
                table2_keys.add(tuple(key_values))
            
            self.log_message.emit(f"表2中共有 {len(table2_keys)} 个唯一组合键")
            self.progress.emit(50)
            
            # 查找表1中不在表2中的行
            not_found_rows = []
            found_count = 0
            
            for index, row in df1.iterrows():
                # 构建表1的组合键
                key_values = []
                for col1, _ in self.compare_columns:
                    value = row[col1]
                    # 处理NaN值
                    if pd.isna(value):
                        key_values.append("")
                    else:
                        key_values.append(str(value).strip())
                
                key = tuple(key_values)
                
                # 检查是否在表2中
                if key not in table2_keys:
                    not_found_rows.append(row)
                else:
                    found_count += 1
                
                # 更新进度
                if (index + 1) % 100 == 0:
                    progress = 50 + int((index + 1) / len(df1) * 35)
                    self.progress.emit(progress)
            
            self.log_message.emit(f"对比完成：")
            self.log_message.emit(f"  表1总行数: {len(df1)}")
            self.log_message.emit(f"  在表2中找到: {found_count} 行")
            self.log_message.emit(f"  在表2中未找到: {len(not_found_rows)} 行")
            
            self.progress.emit(85)
            
            # 导出结果
            if not_found_rows:
                result_df = pd.DataFrame(not_found_rows)
                result_df.to_excel(self.output_file, index=False, engine='openpyxl')
                self.log_message.emit(f"结果已导出到: {self.output_file}")
            else:
                # 即使没有数据也创建一个空表
                result_df = pd.DataFrame(columns=df1.columns)
                result_df.to_excel(self.output_file, index=False, engine='openpyxl')
                self.log_message.emit(f"没有找到不匹配的数据，已导出空表到: {self.output_file}")
            
            self.progress.emit(100)
            
            results = {
                'total_rows': len(df1),
                'found_count': found_count,
                'not_found_count': len(not_found_rows)
            }
            
            self.finished.emit(results)
            
        except Exception as e:
            self.error.emit(f"处理过程出错: {str(e)}")


class ExcelCompareWidget(QWidget):
    """Excel表对比功能界面"""
    
    def __init__(self):
        super().__init__()
        self.worker = None
        self.excel1_file = ""
        self.excel2_file = ""
        self.output_file = ""
        self.excel1_columns = []
        self.excel2_columns = []
        self.init_ui()
    
    def init_ui(self):
        """初始化UI"""
        layout = QVBoxLayout()
        
        # 说明文字
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        info_label = QLabel(
            "此功能用于对比两个Excel表：\n"
            "1. 导入表1和表2\n"
            "2. 选择对比列（可以选择多个列组合）\n"
            "3. 查找表1中在表2中不存在的行\n"
            "4. 导出不匹配的行到新的Excel文件"
        )
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # 文件选择区域
        file_group = QGroupBox("文件选择")
        file_layout = QVBoxLayout()
        
        # 表1选择
        excel1_layout = QHBoxLayout()
        self.excel1_label = QLabel("表1: 未选择")
        self.select_excel1_btn = QPushButton("选择表1")
        self.select_excel1_btn.clicked.connect(self.select_excel1)
        excel1_layout.addWidget(QLabel("表1:"))
        excel1_layout.addWidget(self.excel1_label, 1)
        excel1_layout.addWidget(self.select_excel1_btn)
        file_layout.addLayout(excel1_layout)
        
        # 表2选择
        excel2_layout = QHBoxLayout()
        self.excel2_label = QLabel("表2: 未选择")
        self.select_excel2_btn = QPushButton("选择表2")
        self.select_excel2_btn.clicked.connect(self.select_excel2)
        excel2_layout.addWidget(QLabel("表2:"))
        excel2_layout.addWidget(self.excel2_label, 1)
        excel2_layout.addWidget(self.select_excel2_btn)
        file_layout.addLayout(excel2_layout)
        
        # 输出文件选择
        output_layout = QHBoxLayout()
        self.output_label = QLabel("输出: 未选择")
        self.select_output_btn = QPushButton("选择输出文件")
        self.select_output_btn.clicked.connect(self.select_output)
        self.select_output_btn.setEnabled(False)
        output_layout.addWidget(QLabel("输出:"))
        output_layout.addWidget(self.output_label, 1)
        output_layout.addWidget(self.select_output_btn)
        file_layout.addLayout(output_layout)
        
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # 对比列选择区域
        compare_group = QGroupBox("对比列设置")
        compare_layout = QVBoxLayout()
        
        # 当前对比列列表
        compare_layout.addWidget(QLabel("已选择的对比列组合:"))
        self.compare_list = QListWidget()
        self.compare_list.setMaximumHeight(100)
        compare_layout.addWidget(self.compare_list)
        
        # 添加对比列
        add_compare_layout = QHBoxLayout()
        add_compare_layout.addWidget(QLabel("表1列:"))
        self.excel1_col_combo = QComboBox()
        self.excel1_col_combo.setMinimumWidth(150)
        add_compare_layout.addWidget(self.excel1_col_combo)
        
        add_compare_layout.addWidget(QLabel("<->"))
        
        add_compare_layout.addWidget(QLabel("表2列:"))
        self.excel2_col_combo = QComboBox()
        self.excel2_col_combo.setMinimumWidth(150)
        add_compare_layout.addWidget(self.excel2_col_combo)
        
        self.add_compare_btn = QPushButton("添加对比列")
        self.add_compare_btn.clicked.connect(self.add_compare_column)
        self.add_compare_btn.setEnabled(False)
        add_compare_layout.addWidget(self.add_compare_btn)
        
        self.remove_compare_btn = QPushButton("移除选中")
        self.remove_compare_btn.clicked.connect(self.remove_compare_column)
        add_compare_layout.addWidget(self.remove_compare_btn)
        
        compare_layout.addLayout(add_compare_layout)
        compare_group.setLayout(compare_layout)
        layout.addWidget(compare_group)
        
        # 操作按钮
        button_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始对比")
        self.start_btn.clicked.connect(self.start_compare)
        self.start_btn.setEnabled(False)
        self.stop_btn = QPushButton("停止")
        self.stop_btn.clicked.connect(self.stop_compare)
        self.stop_btn.setEnabled(False)
        button_layout.addWidget(self.start_btn)
        button_layout.addWidget(self.stop_btn)
        layout.addLayout(button_layout)
        
        # 进度条
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)
        
        # 日志区域
        log_group = QGroupBox("日志")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(200)
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)
        
        self.setLayout(layout)
    
    def select_excel1(self):
        """选择表1"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择表1", "", "Excel文件 (*.xlsx *.xls)"
        )
        if file_path:
            self.excel1_file = file_path
            self.excel1_label.setText(file_path.split('/')[-1])
            self.load_excel1_columns()
            self.check_can_start()
    
    def select_excel2(self):
        """选择表2"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择表2", "", "Excel文件 (*.xlsx *.xls)"
        )
        if file_path:
            self.excel2_file = file_path
            self.excel2_label.setText(file_path.split('/')[-1])
            self.load_excel2_columns()
            self.check_can_start()
    
    def select_output(self):
        """选择输出文件"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "选择输出文件", "", "Excel文件 (*.xlsx)"
        )
        if file_path:
            if not file_path.endswith('.xlsx'):
                file_path += '.xlsx'
            self.output_file = file_path
            self.output_label.setText(file_path.split('/')[-1])
            self.check_can_start()
    
    def load_excel1_columns(self):
        """加载表1的列名"""
        try:
            df = pd.read_excel(self.excel1_file, engine='openpyxl', nrows=0)
            self.excel1_columns = list(df.columns)
            self.excel1_col_combo.clear()
            self.excel1_col_combo.addItems(self.excel1_columns)
            self.log_text.append(f"表1列名: {', '.join(self.excel1_columns)}")
            self.add_compare_btn.setEnabled(True)
        except Exception as e:
            QMessageBox.warning(self, "错误", f"读取表1列名失败: {str(e)}")
    
    def load_excel2_columns(self):
        """加载表2的列名"""
        try:
            df = pd.read_excel(self.excel2_file, engine='openpyxl', nrows=0)
            self.excel2_columns = list(df.columns)
            self.excel2_col_combo.clear()
            self.excel2_col_combo.addItems(self.excel2_columns)
            self.log_text.append(f"表2列名: {', '.join(self.excel2_columns)}")
            self.add_compare_btn.setEnabled(True)
        except Exception as e:
            QMessageBox.warning(self, "错误", f"读取表2列名失败: {str(e)}")
    
    def add_compare_column(self):
        """添加对比列"""
        if self.excel1_col_combo.currentText() and self.excel2_col_combo.currentText():
            col1 = self.excel1_col_combo.currentText()
            col2 = self.excel2_col_combo.currentText()
            item_text = f"表1[{col1}] <-> 表2[{col2}]"
            
            # 检查是否已存在
            for i in range(self.compare_list.count()):
                if self.compare_list.item(i).text() == item_text:
                    QMessageBox.warning(self, "提示", "该对比列组合已存在")
                    return
            
            item = QListWidgetItem(item_text)
            item.setData(256, (col1, col2))  # 存储列名对
            self.compare_list.addItem(item)
            self.log_text.append(f"添加对比列: {item_text}")
            self.check_can_start()
    
    def remove_compare_column(self):
        """移除选中的对比列"""
        current_item = self.compare_list.currentItem()
        if current_item:
            self.compare_list.takeItem(self.compare_list.row(current_item))
            self.log_text.append(f"移除对比列: {current_item.text()}")
            self.check_can_start()
    
    def check_can_start(self):
        """检查是否可以开始对比"""
        can_start = bool(
            self.excel1_file and 
            self.excel2_file and 
            self.output_file and
            self.compare_list.count() > 0
        )
        self.start_btn.setEnabled(can_start)
        self.select_output_btn.setEnabled(bool(self.excel1_file and self.excel2_file))
    
    def start_compare(self):
        """开始对比"""
        # 收集对比列
        compare_columns = []
        for i in range(self.compare_list.count()):
            item = self.compare_list.item(i)
            col1, col2 = item.data(256)
            compare_columns.append((col1, col2))
        
        # 禁用按钮
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.select_excel1_btn.setEnabled(False)
        self.select_excel2_btn.setEnabled(False)
        self.select_output_btn.setEnabled(False)
        self.add_compare_btn.setEnabled(False)
        
        # 清空日志和进度
        self.log_text.clear()
        self.progress_bar.setValue(0)
        
        # 创建工作线程
        self.worker = ExcelCompareWorker(
            self.excel1_file, self.excel2_file, compare_columns, self.output_file
        )
        self.worker.progress.connect(self.on_progress)
        self.worker.log_message.connect(self.on_log_message)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()
    
    def stop_compare(self):
        """停止对比"""
        if self.worker and self.worker.isRunning():
            self.worker.terminate()
            self.worker.wait()
            self.log_text.append("处理已停止")
            self.reset_ui()
    
    def on_progress(self, value):
        """更新进度条"""
        self.progress_bar.setValue(value)
    
    def on_log_message(self, message):
        """添加日志消息"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.log_text.append(f"[{timestamp}] {message}")
        # 自动滚动到底部
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
    
    def on_finished(self, results):
        """处理完成"""
        self.log_text.append("=" * 50)
        self.log_text.append("对比完成！")
        self.log_text.append(f"表1总行数: {results['total_rows']}")
        self.log_text.append(f"在表2中找到: {results['found_count']} 行")
        self.log_text.append(f"在表2中未找到: {results['not_found_count']} 行")
        self.log_text.append("=" * 50)
        
        QMessageBox.information(
            self,
            "完成",
            f"对比完成！\n\n"
            f"表1总行数: {results['total_rows']}\n"
            f"在表2中找到: {results['found_count']} 行\n"
            f"在表2中未找到: {results['not_found_count']} 行"
        )
        
        self.reset_ui()
    
    def on_error(self, error_msg):
        """处理错误"""
        self.log_text.append(f"错误: {error_msg}")
        QMessageBox.critical(self, "错误", error_msg)
        self.reset_ui()
    
    def reset_ui(self):
        """重置UI状态"""
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.select_excel1_btn.setEnabled(True)
        self.select_excel2_btn.setEnabled(True)
        self.select_output_btn.setEnabled(True)
        self.add_compare_btn.setEnabled(True)
        self.check_can_start()
