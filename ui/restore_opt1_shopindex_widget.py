"""
恢复个人事业主注册表与库存主体编号
根据 Excel 中的 ba_ptzcb_register 主键 id，释放原 platform_info、切换为 opt1 主体并占用新 platform_info。
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


def _parse_register_id(value):
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
    """与界面填写的期望值比较；期望值为空字符串时表示只匹配单元格为空。"""
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


class RestoreOpt1ShopindexWorker(QThread):
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(
        self, datasource, excel_path, id_column, filter_column=None, filter_expected=""
    ):
        super().__init__()
        self.datasource = datasource
        self.excel_path = excel_path
        self.id_column = id_column
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
                rid = _parse_register_id(row.get(self.id_column))
                rec = {"excel行": row_num, "register_id": rid, "结果": "", "说明": ""}
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
                            rec["说明"] = (
                                f"条件列「{self.filter_column}」不等于设定值，未处理"
                            )
                            self.results["skipped_count"] += 1
                        elif rid is None:
                            rec["结果"] = "跳过"
                            rec["说明"] = "注册表主键 id 为空或无效"
                            self.results["skipped_count"] += 1
                        else:
                            msg = self._process_one(cursor, rid)
                            rec["结果"] = "成功"
                            rec["说明"] = msg
                            self.results["success_count"] += 1
                    elif rid is None:
                        rec["结果"] = "跳过"
                        rec["说明"] = "注册表主键 id 为空或无效"
                        self.results["skipped_count"] += 1
                    else:
                        msg = self._process_one(cursor, rid)
                        rec["结果"] = "成功"
                        rec["说明"] = msg
                        self.results["success_count"] += 1
                except Exception as e:
                    rec["结果"] = "失败"
                    rec["说明"] = str(e)
                    self.results["failed_count"] += 1
                    self.log_message.emit(f"第{row_num}行 register_id={rid}: {e}")

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

    def _process_one(self, cursor, register_id):
        ts = int(datetime.now().timestamp())

        cursor.execute(
            """
            SELECT id, shopindex_id, shudi_id, platform_id
            FROM ba_ptzcb_register
            WHERE id = %s AND (delete_time IS NULL OR delete_time = 0)
            """,
            (register_id,),
        )
        reg = cursor.fetchone()
        if not reg:
            raise Exception("未找到注册表记录或已删除")

        old_sid = reg["shopindex_id"]
        shudi_id = reg["shudi_id"]
        platform_id = reg["platform_id"]

        if old_sid is None or shudi_id is None or platform_id is None:
            raise Exception("shopindex_id / shudi_id / platform_id 存在空值")

        # 1) 原三元组 platform_info → status=1（可用）
        cursor.execute(
            """
            UPDATE ba_platform_info
            SET status = 1, update_time = %s
            WHERE shopindex_id = %s AND platform_id = %s AND shudi_id = %s
              AND (delete_time IS NULL OR delete_time = 0)
            """,
            (ts, old_sid, platform_id, shudi_id),
        )
        released = cursor.rowcount
        self.log_message.emit(
            f"  id={register_id}: 原 platform_info 释放 status=1，影响行数={released}"
        )

        # 2) 当前主体 → legal_id
        cursor.execute(
            """
            SELECT id, legal_id FROM ba_shopindex
            WHERE id = %s AND (delete_time IS NULL OR delete_time = 0)
            """,
            (old_sid,),
        )
        cur_shop = cursor.fetchone()
        if not cur_shop:
            raise Exception(f"未找到 ba_shopindex.id={old_sid}")
        legal_id = cur_shop["legal_id"]

        # 3) 个人事业主 opt1，belong_information 为空
        cursor.execute(
            """
            SELECT id FROM ba_shopindex
            WHERE legal_id = %s
              AND belong_information IS NULL
              AND `select` = 'opt1'
              AND (delete_time IS NULL OR delete_time = 0)
            ORDER BY id ASC
            LIMIT 1
            """,
            (legal_id,),
        )
        new_row = cursor.fetchone()
        if not new_row:
            raise Exception(
                f"未找到 legal_id={legal_id} 且 belong_information 为空、select=opt1 的 ba_shopindex"
            )
        new_sid = new_row["id"]

        # 4) 更新注册表 shopindex_id
        cursor.execute(
            """
            UPDATE ba_ptzcb_register
            SET shopindex_id = %s, update_time = %s
            WHERE id = %s
            """,
            (new_sid, ts, register_id),
        )

        # 5) 新三元组 platform_info → status=2，否则插入
        cursor.execute(
            """
            SELECT id FROM ba_platform_info
            WHERE shopindex_id = %s AND platform_id = %s AND shudi_id = %s
              AND (delete_time IS NULL OR delete_time = 0)
            """,
            (new_sid, platform_id, shudi_id),
        )
        pinfo = cursor.fetchone()
        if pinfo:
            cursor.execute(
                """
                UPDATE ba_platform_info
                SET status = 2, update_time = %s
                WHERE id = %s
                """,
                (ts, pinfo["id"]),
            )
            tail = f"新 shopindex_id={new_sid}，已更新 platform_info.id={pinfo['id']} status=2"
        else:
            cursor.execute(
                """
                INSERT INTO ba_platform_info
                (platform_id, shudi_id, shopindex_id, status, create_time, update_time,
                 admin_id, admin_dept_id)
                VALUES (%s, %s, %s, 2, %s, %s, 1, 1)
                """,
                (platform_id, shudi_id, new_sid, ts, ts),
            )
            tail = f"新 shopindex_id={new_sid}，已插入 platform_info status=2"

        self.log_message.emit(f"  id={register_id}: {tail}")
        return f"原主体={old_sid} → {tail}"


class RestoreOpt1ShopindexWidget(QWidget):
    """迁移任务2：恢复个人事业主主体编号"""

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

        file_group = QGroupBox("Excel")
        file_form = QFormLayout()
        file_row = QHBoxLayout()
        self.file_label = QLabel("未选择文件")
        self.file_label.setStyleSheet("color: gray;")
        file_row.addWidget(self.file_label)
        pick = QPushButton("选择 Excel")
        pick.clicked.connect(self._pick_excel)
        file_row.addWidget(pick)
        file_form.addRow("文件:", file_row)
        self.id_col_combo = QComboBox()
        self.id_col_combo.setEnabled(False)
        self.id_col_combo.currentIndexChanged.connect(self._update_run_btn)
        file_form.addRow("注册表主键 id 列:", self.id_col_combo)
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
        file_group.setLayout(file_form)
        layout.addWidget(file_group)

        info = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        info_label = QLabel(
            """
<b>恢复个人事业主注册表与库存主体编号</b><br><br>
可选：<b>条件列 + 期望值</b>——仅当该行条件列与期望值一致时才执行（选「不筛选」则处理每一行）。<br>
数字会按数值比较（如 1 与 1.0 视为相同）；期望值为空时只匹配该列为空的行。<br><br>
对满足条件的行，按 <b>ba_ptzcb_register.id</b>：<br>
1. 读取该行的 shopindex_id、shudi_id、platform_id<br>
2. 在 ba_platform_info 中按三者匹配：若存在则将 <b>status 改为 1</b>（释放为可用）<br>
3. 用当前 shopindex_id 查 ba_shopindex 得到 legal_id<br>
4. 查找 <b>legal_id 相同</b>、<b>belong_information 为空</b>、<b>select = opt1</b>（个人事业主）的 ba_shopindex，取 id 最小的一条<br>
5. 将注册表的 <b>shopindex_id</b> 更新为该 id<br>
6. 再按新的 shopindex_id + platform_id + shudi_id 查 ba_platform_info：<br>
&nbsp;&nbsp;有则 <b>status=2</b>；无则 <b>插入</b>一条并 <b>status=2</b><br><br>
请确认 Excel 列选对；操作会写库，建议先备份。
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
            self.id_col_combo.clear()
            self.id_col_combo.addItem("请选择列")
            self.id_col_combo.addItems(self.excel_columns)
            self.id_col_combo.setEnabled(True)
            self.filter_col_combo.clear()
            self.filter_col_combo.addItem(FILTER_NONE_LABEL)
            self.filter_col_combo.addItems(self.excel_columns)
            self.filter_col_combo.setEnabled(True)
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))
            self.excel_path = None
            self.file_label.setText("未选择文件")
            self.file_label.setStyleSheet("color: gray;")
            self.filter_col_combo.clear()
            self.filter_col_combo.setEnabled(False)
        self._update_run_btn()

    def _update_run_btn(self):
        ok = bool(
            self.datasource_combo.currentData() is not None
            and self.excel_path
            and self.id_col_combo.currentText() not in ("", "请选择列")
        )
        self.run_btn.setEnabled(ok)

    def _run(self):
        ds = self.datasource_combo.currentData()
        col = self.id_col_combo.currentText()
        if not ds or not self.excel_path or col == "请选择列":
            QMessageBox.warning(self, "提示", "请选择数据源、Excel 与 id 列")
            return
        ft = self.filter_col_combo.currentText()
        filter_col = None if ft == FILTER_NONE_LABEL else ft
        filter_expected = self.filter_value_edit.text()
        hint = "将按 Excel 逐行更新注册表与 platform_info"
        if filter_col:
            hint += f"\n仅处理列「{filter_col}」=「{filter_expected}」的行"
        hint += "，是否继续？"
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
        self.worker = RestoreOpt1ShopindexWorker(
            ds, self.excel_path, col, filter_col, filter_expected
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
