"""
Excel唯一值统计功能界面

导入Excel文件，选择一列（或多列组合），统计总行数、唯一值个数、
空值行数，并列出每个唯一值出现的次数，可导出统计结果Excel。
"""
import os
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox,
                             QGroupBox, QFileDialog, QLineEdit,
                             QListWidget, QListWidgetItem,
                             QTableWidget, QTableWidgetItem, QHeaderView)
from PyQt5.QtCore import Qt
import pandas as pd


class ExcelUniqueCountWidget(QWidget):
    """Excel唯一值统计界面"""

    def __init__(self, db_manager=None):
        super().__init__()
        self.df = None
        self.count_df = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        config_group = QGroupBox("Excel文件选择")
        form_layout = QFormLayout()

        file_layout = QHBoxLayout()
        self.file_edit = QLineEdit()
        self.file_edit.setReadOnly(True)
        self.select_file_btn = QPushButton("选择Excel文件")
        self.select_file_btn.clicked.connect(self.select_file)
        file_layout.addWidget(self.file_edit)
        file_layout.addWidget(self.select_file_btn)
        form_layout.addRow("Excel文件:", file_layout)

        self.columns_list = QListWidget()
        self.columns_list.setMinimumHeight(120)
        form_layout.addRow("统计列(勾选，多列则按组合值):", self.columns_list)

        config_group.setLayout(form_layout)
        layout.addWidget(config_group)

        btn_layout = QHBoxLayout()
        self.count_btn = QPushButton("统计唯一值")
        self.count_btn.clicked.connect(self.do_count)
        self.count_btn.setEnabled(False)
        btn_layout.addWidget(self.count_btn)
        self.export_btn = QPushButton("导出统计结果")
        self.export_btn.clicked.connect(self.export_result)
        self.export_btn.setEnabled(False)
        btn_layout.addWidget(self.export_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        self.summary_label = QLabel("请先选择Excel文件")
        self.summary_label.setStyleSheet("font-weight: bold;")
        layout.addWidget(self.summary_label)

        result_group = QGroupBox("唯一值明细（按出现次数降序）")
        result_layout = QVBoxLayout()
        self.result_table = QTableWidget()
        self.result_table.setColumnCount(2)
        self.result_table.setHorizontalHeaderLabels(["值", "出现次数"])
        header = self.result_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.result_table.setEditTriggers(QTableWidget.NoEditTriggers)
        result_layout.addWidget(self.result_table)
        result_group.setLayout(result_layout)
        layout.addWidget(result_group)

        self.setLayout(layout)

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Excel文件", "", "Excel文件 (*.xlsx *.xls)")
        if not file_path:
            return
        try:
            engine = 'openpyxl' if file_path.lower().endswith('.xlsx') else None
            self.df = pd.read_excel(file_path, engine=engine)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取Excel失败: {str(e)}")
            return

        self.file_edit.setText(file_path)
        self.columns_list.clear()
        for col in self.df.columns:
            item = QListWidgetItem(str(col))
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)
            self.columns_list.addItem(item)
        self.count_btn.setEnabled(True)
        self.export_btn.setEnabled(False)
        self.result_table.setRowCount(0)
        self.summary_label.setText(
            f"已加载: 共 {len(self.df)} 行, {len(self.df.columns)} 列，请勾选要统计的列")

    def get_checked_columns(self):
        cols = []
        for i in range(self.columns_list.count()):
            item = self.columns_list.item(i)
            if item.checkState() == Qt.Checked:
                cols.append(item.text())
        return cols

    def do_count(self):
        if self.df is None:
            QMessageBox.warning(self, "提示", "请先选择Excel文件")
            return
        cols = self.get_checked_columns()
        if not cols:
            QMessageBox.warning(self, "提示", "请勾选要统计的列")
            return

        sub = self.df[cols]
        total_rows = len(sub)
        blank_mask = sub.isna().all(axis=1) | (
            sub.astype(str).apply(lambda s: s.str.strip())
            .replace({'nan': '', 'None': ''}) == '').all(axis=1)
        blank_rows = int(blank_mask.sum())
        valid = sub[~blank_mask]

        if len(cols) == 1:
            keys = valid[cols[0]].astype(str).str.strip()
        else:
            keys = valid.astype(str).apply(
                lambda s: s.str.strip()).agg(' | '.join, axis=1)

        counts = keys.value_counts()
        unique_count = len(counts)
        dup_values = int((counts > 1).sum())

        self.count_df = counts.reset_index()
        self.count_df.columns = ['值', '出现次数']

        self.result_table.setRowCount(0)
        for _, r in self.count_df.iterrows():
            row = self.result_table.rowCount()
            self.result_table.insertRow(row)
            self.result_table.setItem(row, 0, QTableWidgetItem(str(r['值'])))
            self.result_table.setItem(row, 1, QTableWidgetItem(str(r['出现次数'])))

        self.summary_label.setText(
            f"总行数: {total_rows} | 空值行: {blank_rows} | 有效行: {total_rows - blank_rows} | "
            f"唯一值个数: {unique_count} | 出现多次的值: {dup_values} 个")
        self.export_btn.setEnabled(True)

    def export_result(self):
        if self.count_df is None or self.count_df.empty:
            QMessageBox.warning(self, "提示", "没有可导出的统计结果")
            return
        default_name = f"唯一值统计_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        src_dir = os.path.dirname(self.file_edit.text()) or ''
        file_path, _ = QFileDialog.getSaveFileName(
            self, "导出统计结果", os.path.join(src_dir, default_name),
            "Excel文件 (*.xlsx)")
        if not file_path:
            return
        try:
            self.count_df.to_excel(file_path, index=False, engine='openpyxl')
            QMessageBox.information(self, "成功", f"已导出: {file_path}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导出失败: {str(e)}")
