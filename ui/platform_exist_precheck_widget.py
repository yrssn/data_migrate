"""
平台存在预检测：导入 Excel，选择列，与 ba_platform.platform 对照
"""
import os
from collections import Counter
from datetime import datetime

import pandas as pd
import pymysql
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QPushButton, QLabel, QMessageBox, QComboBox,
    QGroupBox, QTableWidget, QTableWidgetItem, QHeaderView,
    QFileDialog, QTextEdit,
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

from database import DatabaseManager


def _cell_to_platform_str(value):
    if pd.isna(value):
        return None
    s = str(value).strip()
    if not s or s.lower() == "nan":
        return None
    return s


class PlatformExistPrecheckWidget(QWidget):
    """平台存在预检测：Excel 列值与 ba_platform 表 platform 字段对比"""

    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.excel_path = None
        self.excel_df = None
        self._last_result_rows = []
        self._last_annotated_df = None
        self.init_ui()
        self.load_datasources()

    def init_ui(self):
        layout = QVBoxLayout()

        ds_group = QGroupBox("选择数据库")
        ds_form = QFormLayout()
        self.datasource_combo = QComboBox()
        self.datasource_combo.currentIndexChanged.connect(self._on_datasource_changed)
        ds_form.addRow("数据库:", self.datasource_combo)
        self.test_connection_btn = QPushButton("测试连接")
        self.test_connection_btn.clicked.connect(self.test_connection)
        self.test_connection_btn.setEnabled(False)
        ds_form.addRow("", self.test_connection_btn)
        ds_group.setLayout(ds_form)
        layout.addWidget(ds_group)

        file_group = QGroupBox("Excel 文件")
        file_form = QFormLayout()
        row_file = QHBoxLayout()
        self.file_label = QLabel("未选择文件")
        self.file_label.setStyleSheet("color: gray;")
        row_file.addWidget(self.file_label)
        self.select_file_btn = QPushButton("选择 Excel")
        self.select_file_btn.clicked.connect(self.select_excel_file)
        row_file.addWidget(self.select_file_btn)
        file_form.addRow("文件:", row_file)
        self.column_combo = QComboBox()
        self.column_combo.setEnabled(False)
        self.column_combo.currentIndexChanged.connect(self._on_column_changed)
        file_form.addRow("平台列:", self.column_combo)
        file_group.setLayout(file_form)
        layout.addWidget(file_group)

        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        info_label = QLabel(
            "<b>平台存在预检测</b><br>"
            "• 从 Excel 选择一列，将其中的值与当前库 <code>ba_platform.platform</code> 对照<br>"
            "• 仅统计 <code>status=1</code> 且未删除（<code>delete_time</code> 为空或 0）的平台<br>"
            "• 匹配规则：去掉首尾空格后，按不区分大小写与库中 platform 比对（与库常用排序规则一致）<br>"
            "• 空单元格单独统计，视为不在库中"
        )
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)

        btn_layout = QHBoxLayout()
        self.check_btn = QPushButton("开始检测")
        self.check_btn.clicked.connect(self.run_check)
        self.check_btn.setEnabled(False)
        btn_layout.addWidget(self.check_btn)
        self.export_missing_btn = QPushButton("导出「不存在」明细")
        self.export_missing_btn.clicked.connect(self.export_missing_detail)
        self.export_missing_btn.setEnabled(False)
        btn_layout.addWidget(self.export_missing_btn)
        self.export_annotated_btn = QPushButton("导出带标记的全表")
        self.export_annotated_btn.clicked.connect(self.export_annotated)
        self.export_annotated_btn.setEnabled(False)
        btn_layout.addWidget(self.export_annotated_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        self.summary_label = QLabel("")
        self.summary_label.setWordWrap(True)
        font = QFont()
        font.setBold(True)
        self.summary_label.setFont(font)
        layout.addWidget(self.summary_label)

        table_group = QGroupBox("唯一平台值检测结果（按 Excel 列去重）")
        table_layout = QVBoxLayout()
        self.result_table = QTableWidget()
        self.result_table.setAlternatingRowColors(True)
        header = self.result_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        table_layout.addWidget(self.result_table)
        table_group.setLayout(table_layout)
        layout.addWidget(table_group)

        log_group = QGroupBox("日志")
        log_layout = QVBoxLayout()
        self.log_te = QTextEdit()
        self.log_te.setReadOnly(True)
        self.log_te.setMaximumHeight(120)
        log_layout.addWidget(self.log_te)
        clear_row = QHBoxLayout()
        self.clear_log_btn = QPushButton("清空日志")
        self.clear_log_btn.clicked.connect(self.log_te.clear)
        clear_row.addWidget(self.clear_log_btn)
        clear_row.addStretch()
        log_layout.addLayout(clear_row)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        self.setLayout(layout)

    def load_datasources(self):
        self.datasource_combo.clear()
        self.datasource_combo.addItem("请选择数据源", None)
        for ds in self.db_manager.get_all_datasources():
            self.datasource_combo.addItem(f"{ds.name} ({ds.host}:{ds.port})", ds)

    def _on_datasource_changed(self):
        ds = self.datasource_combo.currentData()
        self.test_connection_btn.setEnabled(ds is not None)
        self._update_check_enabled()

    def _update_check_enabled(self):
        ok_ds = self.datasource_combo.currentData() is not None
        ok_col = self.column_combo.currentData() is not None and self.column_combo.currentData() >= 0
        ok_file = self.excel_df is not None
        self.check_btn.setEnabled(ok_ds and ok_col and ok_file)

    def test_connection(self):
        ds = self.datasource_combo.currentData()
        if not ds:
            return
        ok, msg = self.db_manager.test_connection(ds)
        if ok:
            QMessageBox.information(self, "连接成功", msg)
        else:
            QMessageBox.critical(self, "连接失败", msg)

    def _log(self, msg: str):
        self.log_te.append(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")

    def select_excel_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择 Excel", "", "Excel (*.xlsx *.xls)"
        )
        if not path:
            return
        try:
            if path.endswith(".xlsx"):
                df = pd.read_excel(path, engine="openpyxl")
            else:
                df = pd.read_excel(path, engine="xlrd")
        except Exception as e:
            QMessageBox.critical(self, "读取失败", str(e))
            return
        self.excel_path = path
        self.excel_df = df
        self.file_label.setText(os.path.basename(path))
        self.file_label.setStyleSheet("color: green;")
        self.column_combo.clear()
        self.column_combo.addItem("请选择列", -1)
        for i, col in enumerate(df.columns):
            letter = chr(65 + i) if i < 26 else f"Col{i+1}"
            self.column_combo.addItem(f"{letter}: {col}", i)
        self.column_combo.setEnabled(True)
        self._last_result_rows = []
        self._last_annotated_df = None
        self.export_missing_btn.setEnabled(False)
        self.export_annotated_btn.setEnabled(False)
        self.result_table.setRowCount(0)
        self.summary_label.clear()
        self._log(f"已加载 Excel，共 {len(df)} 行，{len(df.columns)} 列")
        self._update_check_enabled()

    def _on_column_changed(self):
        self._update_check_enabled()

    def _load_platform_index(self, ds):
        conn = pymysql.connect(
            host=ds.host,
            port=ds.port,
            user=ds.username,
            password=ds.password,
            database=ds.database,
            charset=ds.charset,
        )
        try:
            cur = conn.cursor()
            cur.execute(
                """
                SELECT platform FROM ba_platform
                WHERE status = 1
                  AND (delete_time IS NULL OR delete_time = 0)
                """
            )
            rows = cur.fetchall()
            cur.close()
        finally:
            conn.close()

        lower_to_canonical = {}
        for (p,) in rows:
            if p is None:
                continue
            s = str(p).strip()
            if not s:
                continue
            lower_to_canonical[s.lower()] = s
        return lower_to_canonical

    def run_check(self):
        ds = self.datasource_combo.currentData()
        if not ds:
            QMessageBox.warning(self, "提示", "请选择数据源")
            return
        if self.excel_df is None:
            QMessageBox.warning(self, "提示", "请先选择 Excel")
            return
        col_idx = self.column_combo.currentData()
        if col_idx is None or col_idx < 0:
            QMessageBox.warning(self, "提示", "请选择平台列")
            return

        col_name = self.excel_df.columns[col_idx]
        try:
            self._log("正在从数据库加载 ba_platform…")
            lower_to_canonical = self._load_platform_index(ds)
            self._log(f"库中有效 platform 数（按不区分大小写计）: {len(lower_to_canonical)}")
        except Exception as e:
            QMessageBox.critical(self, "数据库错误", str(e))
            self._log(f"加载 ba_platform 失败: {e}")
            return

        series = self.excel_df[col_name]
        norm_keys = []
        display_for_key = {}
        for v in series:
            s = _cell_to_platform_str(v)
            if s is None:
                key = ""
                disp = "（空）"
            else:
                key = s.lower()
                disp = s
            norm_keys.append(key)
            if key not in display_for_key:
                display_for_key[key] = disp

        counts = Counter(norm_keys)
        result_rows = []
        exist_n = 0
        missing_n = 0
        empty_rows = counts.get("", 0)
        for key, cnt in sorted(
            counts.items(),
            key=lambda x: (0 if x[0] == "" else 1, -x[1], x[0]),
        ):
            disp = display_for_key[key]
            if key == "":
                note = "空单元格"
                missing_n += cnt
                result_rows.append((disp, "否", cnt, note))
            else:
                if key in lower_to_canonical:
                    canon = lower_to_canonical[key]
                    note = f"匹配: {canon}"
                    exist_n += cnt
                    result_rows.append((disp, "是", cnt, note))
                else:
                    note = "库中无此 platform"
                    missing_n += cnt
                    result_rows.append((disp, "否", cnt, note))

        self._last_result_rows = result_rows
        annotated = self.excel_df.copy()
        col_exist = "平台存在预检测_是否存在"
        col_match = "平台存在预检测_库platform"
        annotated[col_exist] = [
            (
                ""
                if k == ""
                else ("是" if k in lower_to_canonical else "否")
            )
            for k in norm_keys
        ]
        annotated[col_match] = [
            "" if k == "" else lower_to_canonical.get(k, "") for k in norm_keys
        ]
        self._last_annotated_df = annotated

        self.result_table.setRowCount(len(result_rows))
        self.result_table.setColumnCount(4)
        self.result_table.setHorizontalHeaderLabels(
            ["平台值(列内代表)", "是否存在", "出现次数", "说明"]
        )
        for i, row in enumerate(result_rows):
            disp, ex, cnt, note = row
            self.result_table.setItem(i, 0, QTableWidgetItem(disp))
            self.result_table.setItem(i, 1, QTableWidgetItem(ex))
            self.result_table.setItem(i, 2, QTableWidgetItem(str(cnt)))
            self.result_table.setItem(i, 3, QTableWidgetItem(note))

        total = len(series)
        uniq = len(counts)
        self.summary_label.setText(
            f"列「{col_name}」：共 {total} 行，唯一值 {uniq} 个；"
            f"行级判断为存在 {exist_n} 行，不存在 {missing_n} 行（含空 {empty_rows} 行）。"
        )
        self.summary_label.setStyleSheet(
            "color: green;" if missing_n == 0 else "color: #b45309;"
        )
        self.export_missing_btn.setEnabled(missing_n > 0)
        self.export_annotated_btn.setEnabled(True)
        self._log("检测完成")
        QMessageBox.information(
            self,
            "检测完成",
            f"行数 {total}，唯一平台值 {uniq}；存在 {exist_n} 行，不存在 {missing_n} 行。",
        )

    def export_missing_detail(self):
        if self._last_annotated_df is None or not self.excel_path:
            QMessageBox.information(self, "提示", "请先完成检测")
            return
        df = self._last_annotated_df
        col_exist = "平台存在预检测_是否存在"
        if col_exist not in df.columns:
            return
        sub = df[df[col_exist] == "否"].copy()
        if len(sub) == 0:
            QMessageBox.information(self, "提示", "没有不存在的数据可导出")
            return
        base = os.path.splitext(os.path.basename(self.excel_path))[0]
        default_name = f"{base}_平台不存在明细_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(
            self, "保存", default_name, "Excel (*.xlsx)"
        )
        if not path:
            return
        try:
            sub.to_excel(path, index=False, engine="openpyxl")
            QMessageBox.information(
                self, "成功", f"已导出 {len(sub)} 行\n{os.path.basename(path)}"
            )
        except Exception as e:
            QMessageBox.critical(self, "导出失败", str(e))

    def export_annotated(self):
        if self._last_annotated_df is None or not self.excel_path:
            QMessageBox.information(self, "提示", "请先完成检测")
            return
        base = os.path.splitext(os.path.basename(self.excel_path))[0]
        default_name = f"{base}_平台预检测标记_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(
            self, "保存", default_name, "Excel (*.xlsx)"
        )
        if not path:
            return
        try:
            self._last_annotated_df.to_excel(path, index=False, engine="openpyxl")
            QMessageBox.information(
                self,
                "成功",
                f"已导出 {len(self._last_annotated_df)} 行\n{os.path.basename(path)}",
            )
        except Exception as e:
            QMessageBox.critical(self, "导出失败", str(e))
