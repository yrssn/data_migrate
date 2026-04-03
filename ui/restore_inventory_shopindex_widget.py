"""
恢复库存的主体编号
按 Excel：条件列筛选后，用库存主键查 ba_account_mag，用库存的 register_id 查 ba_ptzcb_register，
将注册表的 shopindex_id 写回库存；Excel 中注册表主键列须与库存 register_id 一致（用于核对）。
"""
import os
from datetime import datetime

import pandas as pd
import pymysql
from PyQt5.QtCore import QThread, Qt, pyqtSignal
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QComboBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from database import DatabaseManager

FILTER_NONE_LABEL = "不筛选（处理所有行）"


def _parse_int_id(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = str(value).strip()
    if not s or s.lower() == "nan":
        return None
    try:
        return int(float(s))
    except (ValueError, TypeError):
        return None


def _cell_equals_expected(cell_val, expected_str):
    exp = (expected_str if expected_str is not None else "").strip()
    if cell_val is None or (isinstance(cell_val, float) and pd.isna(cell_val)):
        cell_norm = ""
    else:
        cell_norm = str(cell_val).strip()
        if cell_norm.lower() == "nan":
            cell_norm = ""
    if exp == "":
        return cell_norm == ""
    try:
        if float(cell_norm) == float(exp):
            return True
    except (ValueError, TypeError):
        pass
    return cell_norm == exp


class RestoreInventoryShopindexWorker(QThread):
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(
        self,
        datasource,
        excel_path,
        inventory_id_column,
        register_id_column,
        filter_column=None,
        filter_expected="",
    ):
        super().__init__()
        self.datasource = datasource
        self.excel_path = excel_path
        self.inventory_id_column = inventory_id_column
        self.register_id_column = register_id_column
        self.filter_column = filter_column
        self.filter_expected = filter_expected if filter_expected is not None else ""
        self.results = {
            "total_rows": 0,
            "success_count": 0,
            "failed_count": 0,
            "skipped_count": 0,
            "rows": [],
        }

    def run(self):
        try:
            self.log_message.emit("读取 Excel...")
            self.progress.emit(5)
            df = pd.read_excel(self.excel_path, engine="openpyxl")
            self.results["total_rows"] = len(df)

            conn = pymysql.connect(
                host=self.datasource.host,
                port=self.datasource.port,
                user=self.datasource.username,
                password=self.datasource.password,
                database=self.datasource.database,
                charset=self.datasource.charset,
            )
            cursor = conn.cursor(pymysql.cursors.DictCursor)
            self.log_message.emit("数据库已连接")
            self.progress.emit(10)

            total = len(df)
            for i, (_, row) in enumerate(df.iterrows()):
                row_num = i + 1
                inv_id = _parse_int_id(row.get(self.inventory_id_column))
                reg_id = _parse_int_id(row.get(self.register_id_column))
                rec = {
                    "excel行": row_num,
                    "库存id": inv_id,
                    "注册表id": reg_id,
                    "结果": "",
                    "说明": "",
                }
                try:
                    if self.filter_column:
                        if self.filter_column not in row.index:
                            rec["结果"] = "跳过"
                            rec["说明"] = f"表中无列「{self.filter_column}」"
                            self.results["skipped_count"] += 1
                        elif not _cell_equals_expected(
                            row.get(self.filter_column), self.filter_expected
                        ):
                            rec["结果"] = "跳过"
                            rec["说明"] = "条件列不等于设定值"
                            self.results["skipped_count"] += 1
                        elif inv_id is None or reg_id is None:
                            rec["结果"] = "跳过"
                            rec["说明"] = "库存主键或注册表主键为空/无效"
                            self.results["skipped_count"] += 1
                        else:
                            msg = self._process_one(cursor, inv_id, reg_id)
                            rec["结果"] = "成功"
                            rec["说明"] = msg
                            self.results["success_count"] += 1
                    elif inv_id is None or reg_id is None:
                        rec["结果"] = "跳过"
                        rec["说明"] = "库存主键或注册表主键为空/无效"
                        self.results["skipped_count"] += 1
                    else:
                        msg = self._process_one(cursor, inv_id, reg_id)
                        rec["结果"] = "成功"
                        rec["说明"] = msg
                        self.results["success_count"] += 1
                except Exception as e:
                    rec["结果"] = "失败"
                    rec["说明"] = str(e)
                    self.results["failed_count"] += 1
                    self.log_message.emit(
                        f"第{row_num}行 库存id={inv_id} 注册id={reg_id}: {e}"
                    )

                self.results["rows"].append(rec)
                pct = 10 + int((i + 1) / max(total, 1) * 85)
                self.progress.emit(pct)
                if (i + 1) % 20 == 0:
                    conn.commit()

            conn.commit()
            cursor.close()
            conn.close()
            self.progress.emit(100)
            self.finished.emit(self.results)
        except Exception as e:
            self.error.emit(str(e))

    def _process_one(self, cursor, inventory_id, register_id_excel):
        ts = int(datetime.now().timestamp())

        cursor.execute(
            """
            SELECT id, register_id, shopindex_id
            FROM ba_account_mag
            WHERE id = %s
            """,
            (inventory_id,),
        )
        mag = cursor.fetchone()
        if not mag:
            raise Exception(f"未找到库存 ba_account_mag.id={inventory_id}")

        rid_mag = int(mag["register_id"])
        if rid_mag != int(register_id_excel):
            raise Exception(
                f"库存 register_id={rid_mag} 与 Excel 注册表主键列 {register_id_excel} 不一致，已中止本行"
            )

        cursor.execute(
            """
            SELECT id, shopindex_id
            FROM ba_ptzcb_register
            WHERE id = %s AND (delete_time IS NULL OR delete_time = 0)
            """,
            (rid_mag,),
        )
        reg = cursor.fetchone()
        if not reg:
            raise Exception(f"未找到注册表 ba_ptzcb_register.id={rid_mag} 或已删除")

        new_sid = reg["shopindex_id"]
        if new_sid is None:
            raise Exception("注册表 shopindex_id 为空，无法写回库存")

        old_sid = mag["shopindex_id"]
        cursor.execute(
            """
            UPDATE ba_account_mag
            SET shopindex_id = %s, update_time = %s
            WHERE id = %s
            """,
            (new_sid, ts, inventory_id),
        )

        msg = f"shopindex_id {old_sid} → {new_sid}（来自注册表 id={rid_mag}）"
        self.log_message.emit(f"  库存 id={inventory_id}: {msg}")
        return msg


class RestoreInventoryShopindexWidget(QWidget):
    """迁移任务2：恢复库存的主体编号"""

    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.excel_path = None
        self.excel_columns = []
        self._init_ui()
        self._load_datasources()

    def _init_ui(self):
        layout = QVBoxLayout()

        ds_group = QGroupBox("选择数据库")
        ds_form = QFormLayout()
        self.datasource_combo = QComboBox()
        self.datasource_combo.currentTextChanged.connect(self._on_ds_changed)
        ds_form.addRow("数据源:", self.datasource_combo)
        self.test_btn = QPushButton("测试连接")
        self.test_btn.clicked.connect(self._test_connection)
        self.test_btn.setEnabled(False)
        ds_form.addRow("", self.test_btn)
        ds_group.setLayout(ds_form)
        layout.addWidget(ds_group)

        file_group = QGroupBox("Excel 列映射")
        file_form = QFormLayout()
        file_row = QHBoxLayout()
        self.file_label = QLabel("未选择文件")
        self.file_label.setStyleSheet("color: gray;")
        file_row.addWidget(self.file_label)
        pick = QPushButton("选择 Excel")
        pick.clicked.connect(self._pick_excel)
        file_row.addWidget(pick)
        file_form.addRow("文件:", file_row)

        self.filter_col_combo = QComboBox()
        self.filter_col_combo.setEnabled(False)
        self.filter_col_combo.currentIndexChanged.connect(self._update_run_btn)
        file_form.addRow("仅当此列等于下方值时才处理:", self.filter_col_combo)
        self.filter_value_edit = QLineEdit()
        self.filter_value_edit.setPlaceholderText(
            "与条件列比对；留空表示只处理该列为空的行"
        )
        self.filter_value_edit.textChanged.connect(self._update_run_btn)
        file_form.addRow("期望值:", self.filter_value_edit)

        self.inv_id_combo = QComboBox()
        self.inv_id_combo.setEnabled(False)
        self.inv_id_combo.currentIndexChanged.connect(self._update_run_btn)
        file_form.addRow("库存主键 (ba_account_mag.id):", self.inv_id_combo)

        self.reg_id_combo = QComboBox()
        self.reg_id_combo.setEnabled(False)
        self.reg_id_combo.currentIndexChanged.connect(self._update_run_btn)
        file_form.addRow("注册表主键 (须=库存 register_id):", self.reg_id_combo)

        file_group.setLayout(file_form)
        layout.addWidget(file_group)

        info = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        info_label = QLabel(
            """
<b>恢复库存的主体编号</b><br><br>
1. 可选「条件列 + 期望值」，规则同「恢复个人事业主主体编号」。<br>
2. 用 <b>库存主键</b> 查 <b>ba_account_mag</b>。<br>
3. 用库存上的 <b>register_id</b> 查 <b>ba_ptzcb_register</b>（须未删除）。<br>
4. Excel 中 <b>注册表主键</b> 列必须与该行库存的 <b>register_id</b> 一致，否则本行失败（防误操作）。<br>
5. 将注册表的 <b>shopindex_id</b> 写入该库存的 <b>shopindex_id</b>，并更新 <b>update_time</b>。<br><br>
若注册表 <b>shopindex_id</b> 为空则本行失败。
            """
        )
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        info.setLayout(info_layout)
        layout.addWidget(info)

        btn_row = QHBoxLayout()
        self.run_btn = QPushButton("开始执行")
        self.run_btn.setEnabled(False)
        self.run_btn.clicked.connect(self._run)
        btn_row.addWidget(self.run_btn)
        btn_row.addStretch()
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        btn_row.addWidget(self.progress)
        layout.addLayout(btn_row)

        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(200)
        log_layout.addWidget(self.log_text)
        clear_btn = QPushButton("清空日志")
        clear_btn.clicked.connect(self.log_text.clear)
        log_layout.addWidget(clear_btn)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        self.result_label = QLabel("")
        self.result_label.setAlignment(Qt.AlignCenter)
        f = QFont()
        f.setBold(True)
        self.result_label.setFont(f)
        layout.addWidget(self.result_label)

        self.setLayout(layout)

    def _load_datasources(self):
        self.datasource_combo.clear()
        self.datasource_combo.addItem("请选择数据源", None)
        for ds in self.db_manager.get_all_datasources():
            self.datasource_combo.addItem(f"{ds.name} ({ds.host}:{ds.port})", ds)

    def _on_ds_changed(self):
        self.test_btn.setEnabled(self.datasource_combo.currentData() is not None)
        self._update_run_btn()

    def _test_connection(self):
        ds = self.datasource_combo.currentData()
        if not ds:
            return
        ok, msg = self.db_manager.test_connection(ds)
        if ok:
            QMessageBox.information(self, "连接成功", msg)
        else:
            QMessageBox.critical(self, "连接失败", msg)

    def _pick_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择 Excel", "", "Excel (*.xlsx *.xls)"
        )
        if not path:
            return
        try:
            df = pd.read_excel(path, engine="openpyxl", nrows=0)
            self.excel_columns = list(df.columns)
            self.excel_path = path
            self.file_label.setText(os.path.basename(path))
            self.file_label.setStyleSheet("color: green;")

            self.filter_col_combo.clear()
            self.filter_col_combo.addItem(FILTER_NONE_LABEL)
            self.filter_col_combo.addItems(self.excel_columns)
            self.filter_col_combo.setEnabled(True)

            self.inv_id_combo.clear()
            self.inv_id_combo.addItem("请选择列")
            self.inv_id_combo.addItems(self.excel_columns)
            self.inv_id_combo.setEnabled(True)

            self.reg_id_combo.clear()
            self.reg_id_combo.addItem("请选择列")
            self.reg_id_combo.addItems(self.excel_columns)
            self.reg_id_combo.setEnabled(True)
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))
            self.excel_path = None
            self.file_label.setText("未选择文件")
            self.file_label.setStyleSheet("color: gray;")
            self.filter_col_combo.clear()
            self.filter_col_combo.setEnabled(False)
            self.inv_id_combo.clear()
            self.inv_id_combo.setEnabled(False)
            self.reg_id_combo.clear()
            self.reg_id_combo.setEnabled(False)
        self._update_run_btn()

    def _update_run_btn(self):
        ok = bool(
            self.datasource_combo.currentData() is not None
            and self.excel_path
            and self.inv_id_combo.currentText() not in ("", "请选择列")
            and self.reg_id_combo.currentText() not in ("", "请选择列")
        )
        self.run_btn.setEnabled(ok)

    def _run(self):
        ds = self.datasource_combo.currentData()
        inv_c = self.inv_id_combo.currentText()
        reg_c = self.reg_id_combo.currentText()
        if not ds or not self.excel_path or inv_c == "请选择列" or reg_c == "请选择列":
            QMessageBox.warning(self, "提示", "请选择数据源、Excel、库存主键列与注册表主键列")
            return
        ft = self.filter_col_combo.currentText()
        filter_col = None if ft in ("", FILTER_NONE_LABEL) else ft
        filter_expected = self.filter_value_edit.text()
        hint = "将按行更新 ba_account_mag.shopindex_id"
        if filter_col:
            hint += f"\n仅处理列「{filter_col}」=「{filter_expected}」的行"
        hint += "\n是否继续？"
        reply = QMessageBox.question(
            self,
            "确认",
            hint,
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return
        self.run_btn.setEnabled(False)
        self.progress.setVisible(True)
        self.progress.setValue(0)
        self.log_text.clear()
        self.result_label.setText("处理中...")
        self.worker = RestoreInventoryShopindexWorker(
            ds,
            self.excel_path,
            inv_c,
            reg_c,
            filter_col,
            filter_expected,
        )
        self.worker.progress.connect(self.progress.setValue)
        self.worker.log_message.connect(self._log)
        self.worker.finished.connect(self._done)
        self.worker.error.connect(self._err)
        self.worker.start()

    def _log(self, s):
        self.log_text.append(f"[{datetime.now().strftime('%H:%M:%S')}] {s}")

    def _done(self, res):
        self.run_btn.setEnabled(True)
        self.progress.setVisible(False)
        txt = (
            f"完成：共 {res['total_rows']} 行，成功 {res['success_count']}，"
            f"跳过 {res['skipped_count']}，失败 {res['failed_count']}"
        )
        self.result_label.setText(txt)
        self.result_label.setStyleSheet(
            "color: green;" if res["failed_count"] == 0 else "color: orange;"
        )
        QMessageBox.information(self, "完成", txt)
        self._update_run_btn()

    def _err(self, msg):
        self.run_btn.setEnabled(True)
        self.progress.setVisible(False)
        self.result_label.setText("失败")
        self.result_label.setStyleSheet("color: red;")
        QMessageBox.critical(self, "错误", msg)
        self._update_run_btn()
