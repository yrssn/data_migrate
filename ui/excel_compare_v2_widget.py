"""
对比表格数据导出2.0
导入两个Excel表格，根据指定字段（支持多条件）对比，找出表1中不存在于表2的行并导出
"""
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QFileDialog, QTextEdit, QProgressBar,
                             QGroupBox, QMessageBox, QComboBox, QListWidget,
                             QListWidgetItem)
from PyQt5.QtCore import QThread, pyqtSignal
import pandas as pd
from datetime import datetime


class CompareV2Worker(QThread):
    """对比工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, excel1_file, excel2_file, compare_columns, output_file):
        super().__init__()
        self.excel1_file = excel1_file
        self.excel2_file = excel2_file
        self.compare_columns = compare_columns  # [(col1, col2), ...]
        self.output_file = output_file

    def run(self):
        try:
            self.log_message.emit("开始读取Excel文件...")
            self.progress.emit(10)

            try:
                df1 = pd.read_excel(self.excel1_file, engine='openpyxl')
                self.log_message.emit(f"表1: {len(df1)} 行")
            except Exception as e:
                self.error.emit(f"读取表1失败: {str(e)}")
                return

            self.progress.emit(30)

            try:
                df2 = pd.read_excel(self.excel2_file, engine='openpyxl')
                self.log_message.emit(f"表2: {len(df2)} 行")
            except Exception as e:
                self.error.emit(f"读取表2失败: {str(e)}")
                return

            self.progress.emit(50)

            # 验证列存在
            for col1, col2 in self.compare_columns:
                if col1 not in df1.columns:
                    self.error.emit(f"表1中不存在列: {col1}")
                    return
                if col2 not in df2.columns:
                    self.error.emit(f"表2中不存在列: {col2}")
                    return

            self.log_message.emit(f"对比条件数: {len(self.compare_columns)}")
            for c1, c2 in self.compare_columns:
                self.log_message.emit(f"  表1[{c1}] <-> 表2[{c2}]")

            def normalize(val):
                if pd.isna(val):
                    return ""
                s = str(val).strip()
                if s.endswith('.0'):
                    try:
                        s = str(int(float(s)))
                    except (ValueError, OverflowError):
                        pass
                return s

            # 构建表2的组合键集合
            cols2 = [c2 for _, c2 in self.compare_columns]
            keys2 = set()
            for _, row in df2.iterrows():
                key = tuple(normalize(row[c]) for c in cols2)
                keys2.add(key)
            self.log_message.emit(f"表2唯一组合键数: {len(keys2)}")
            self.progress.emit(60)

            # 找表1中组合键不在表2的行
            cols1 = [c1 for c1, _ in self.compare_columns]
            not_found = []
            found = 0
            for idx, row in df1.iterrows():
                key = tuple(normalize(row[c]) for c in cols1)
                if key in keys2:
                    found += 1
                else:
                    not_found.append(row)
                if (idx + 1) % 500 == 0:
                    self.progress.emit(60 + int((idx + 1) / len(df1) * 25))

            self.log_message.emit(f"匹配: {found} 行, 未匹配: {len(not_found)} 行")
            self.progress.emit(90)

            result_df = pd.DataFrame(not_found) if not_found else pd.DataFrame(columns=df1.columns)
            result_df.to_excel(self.output_file, index=False, engine='openpyxl')
            self.log_message.emit(f"已导出到: {self.output_file}")
            self.progress.emit(100)

            self.finished.emit({
                'total': len(df1),
                'matched': found,
                'unmatched': len(not_found),
            })

        except Exception as e:
            self.error.emit(f"处理出错: {str(e)}")


class ExcelCompareV2Widget(QWidget):
    """对比表格数据导出2.0"""

    def __init__(self):
        super().__init__()
        self.worker = None
        self.excel1_file = ""
        self.excel2_file = ""
        self.output_file = ""
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # 说明
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        info_label = QLabel(
            "导入表1（数据多）和表2（数据少），添加对比条件（支持多个），\n"
            "找出表1中组合条件不存在于表2的行，导出到新文件。"
        )
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)

        # 文件选择
        file_group = QGroupBox("文件选择")
        file_layout = QVBoxLayout()

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("表1（数据多）:"))
        self.excel1_label = QLabel("未选择")
        row1.addWidget(self.excel1_label, 1)
        self.select_excel1_btn = QPushButton("选择表1")
        self.select_excel1_btn.clicked.connect(self.select_excel1)
        row1.addWidget(self.select_excel1_btn)
        file_layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("表2（数据少）:"))
        self.excel2_label = QLabel("未选择")
        row2.addWidget(self.excel2_label, 1)
        self.select_excel2_btn = QPushButton("选择表2")
        self.select_excel2_btn.clicked.connect(self.select_excel2)
        row2.addWidget(self.select_excel2_btn)
        file_layout.addLayout(row2)

        row3 = QHBoxLayout()
        row3.addWidget(QLabel("输出文件:"))
        self.output_label = QLabel("未选择")
        row3.addWidget(self.output_label, 1)
        self.select_output_btn = QPushButton("选择输出文件")
        self.select_output_btn.clicked.connect(self.select_output)
        row3.addWidget(self.select_output_btn)
        file_layout.addLayout(row3)

        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # 对比条件
        col_group = QGroupBox("对比条件（可添加多个）")
        col_layout = QVBoxLayout()

        # 已添加的条件列表
        self.compare_list = QListWidget()
        self.compare_list.setMaximumHeight(80)
        col_layout.addWidget(self.compare_list)

        # 添加条件行
        add_row = QHBoxLayout()
        add_row.addWidget(QLabel("表1列:"))
        self.col1_combo = QComboBox()
        self.col1_combo.setMinimumWidth(150)
        add_row.addWidget(self.col1_combo)
        add_row.addWidget(QLabel("  <->  "))
        add_row.addWidget(QLabel("表2列:"))
        self.col2_combo = QComboBox()
        self.col2_combo.setMinimumWidth(150)
        add_row.addWidget(self.col2_combo)

        self.add_btn = QPushButton("添加")
        self.add_btn.clicked.connect(self.add_compare_pair)
        self.add_btn.setEnabled(False)
        add_row.addWidget(self.add_btn)

        self.remove_btn = QPushButton("移除选中")
        self.remove_btn.clicked.connect(self.remove_compare_pair)
        add_row.addWidget(self.remove_btn)

        col_layout.addLayout(add_row)
        col_group.setLayout(col_layout)
        layout.addWidget(col_group)

        # 按钮 + 进度条
        btn_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始对比")
        self.start_btn.clicked.connect(self.start_compare)
        self.start_btn.setEnabled(False)
        btn_layout.addWidget(self.start_btn)

        self.stop_btn = QPushButton("停止")
        self.stop_btn.clicked.connect(self.stop_compare)
        self.stop_btn.setEnabled(False)
        btn_layout.addWidget(self.stop_btn)

        btn_layout.addStretch()

        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimumWidth(200)
        btn_layout.addWidget(self.progress_bar)
        layout.addLayout(btn_layout)

        # 日志
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
        file_path, _ = QFileDialog.getOpenFileName(self, "选择表1", "", "Excel文件 (*.xlsx *.xls)")
        if file_path:
            self.excel1_file = file_path
            self.excel1_label.setText(file_path.split('/')[-1].split('\\')[-1])
            self._load_columns(file_path, self.col1_combo, "表1")
            self._update_add_btn()
            self._check_ready()

    def select_excel2(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择表2", "", "Excel文件 (*.xlsx *.xls)")
        if file_path:
            self.excel2_file = file_path
            self.excel2_label.setText(file_path.split('/')[-1].split('\\')[-1])
            self._load_columns(file_path, self.col2_combo, "表2")
            self._update_add_btn()
            self._check_ready()

    def select_output(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "选择输出文件", "", "Excel文件 (*.xlsx)")
        if file_path:
            if not file_path.endswith('.xlsx'):
                file_path += '.xlsx'
            self.output_file = file_path
            self.output_label.setText(file_path.split('/')[-1].split('\\')[-1])
            self._check_ready()

    def _load_columns(self, file_path, combo, name):
        try:
            df = pd.read_excel(file_path, engine='openpyxl', nrows=0)
            combo.clear()
            combo.addItems([str(c) for c in df.columns])
            self._log(f"{name}列名: {', '.join(str(c) for c in df.columns)}")
        except Exception as e:
            QMessageBox.warning(self, "错误", f"读取{name}列名失败: {str(e)}")

    def _update_add_btn(self):
        self.add_btn.setEnabled(self.col1_combo.count() > 0 and self.col2_combo.count() > 0)

    def add_compare_pair(self):
        col1 = self.col1_combo.currentText()
        col2 = self.col2_combo.currentText()
        if not col1 or not col2:
            return
        text = f"表1[{col1}] <-> 表2[{col2}]"
        # 检查重复
        for i in range(self.compare_list.count()):
            if self.compare_list.item(i).text() == text:
                QMessageBox.warning(self, "提示", "该对比条件已存在")
                return
        item = QListWidgetItem(text)
        item.setData(256, (col1, col2))
        self.compare_list.addItem(item)
        self._log(f"添加对比条件: {text}")
        self._check_ready()

    def remove_compare_pair(self):
        current = self.compare_list.currentItem()
        if current:
            self._log(f"移除对比条件: {current.text()}")
            self.compare_list.takeItem(self.compare_list.row(current))
            self._check_ready()

    def _check_ready(self):
        self.start_btn.setEnabled(bool(
            self.excel1_file and self.excel2_file and self.output_file
            and self.compare_list.count() > 0
        ))

    def start_compare(self):
        compare_columns = []
        for i in range(self.compare_list.count()):
            compare_columns.append(self.compare_list.item(i).data(256))

        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.select_excel1_btn.setEnabled(False)
        self.select_excel2_btn.setEnabled(False)
        self.select_output_btn.setEnabled(False)
        self.col1_combo.setEnabled(False)
        self.col2_combo.setEnabled(False)
        self.add_btn.setEnabled(False)
        self.remove_btn.setEnabled(False)
        self.log_text.clear()
        self.progress_bar.setValue(0)

        self.worker = CompareV2Worker(
            self.excel1_file, self.excel2_file, compare_columns, self.output_file
        )
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log_message.connect(self._log)
        self.worker.finished.connect(self._on_finished)
        self.worker.error.connect(self._on_error)
        self.worker.start()

    def stop_compare(self):
        if self.worker and self.worker.isRunning():
            self.worker.terminate()
            self.worker.wait()
            self._log("已停止")
            self._reset_ui()

    def _log(self, msg):
        ts = datetime.now().strftime('%H:%M:%S')
        self.log_text.append(f"[{ts}] {msg}")
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )

    def _on_finished(self, results):
        self._log("=" * 40)
        self._log(f"对比完成！表1共 {results['total']} 行, "
                  f"匹配 {results['matched']} 行, 未匹配 {results['unmatched']} 行")

        QMessageBox.information(
            self, "完成",
            f"表1共 {results['total']} 行\n"
            f"匹配: {results['matched']} 行\n"
            f"未匹配（已导出）: {results['unmatched']} 行"
        )
        self._reset_ui()

    def _on_error(self, msg):
        self._log(f"错误: {msg}")
        QMessageBox.critical(self, "错误", msg)
        self._reset_ui()

    def _reset_ui(self):
        self.stop_btn.setEnabled(False)
        self.select_excel1_btn.setEnabled(True)
        self.select_excel2_btn.setEnabled(True)
        self.select_output_btn.setEnabled(True)
        self.col1_combo.setEnabled(True)
        self.col2_combo.setEnabled(True)
        self.remove_btn.setEnabled(True)
        self._update_add_btn()
        self._check_ready()
