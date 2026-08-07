"""
货币关联财务系统

以财务系统 ea_dy_currency 为主，按货币代码(code)关联进销存 ba_currency：
- 代码能匹配上(其次按中文名)：回写 ba_currency.finance_currency_id
- 匹配不上：在进销存新增一条并直接绑定

本功能和「账户分类关联财务系统」一样是最先要跑的，
「银行名称关联财务系统」会依赖这里的 finance_currency_id 把财务银行名称上的货币刷到进销存。
"""
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from database import DatabaseManager
import pandas as pd
import pymysql


class CurrencyLinkWorker(QThread):
    """货币关联工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, target_datasource, source_datasource, preview_only=False):
        super().__init__()
        self.target_datasource = target_datasource  # 进销存 ba_
        self.source_datasource = source_datasource  # 财务系统 ea_
        self.preview_only = preview_only
        self.results = {
            'total': 0,
            'bind_count': 0,
            'insert_count': 0,
            'skip_count': 0,
            'fail_count': 0,
            'completed_data': [],
            'failed_records': []
        }

    def _connect(self, ds):
        return pymysql.connect(
            host=ds.host,
            port=ds.port,
            user=ds.username,
            password=ds.password,
            database=ds.database,
            charset=ds.charset
        )

    def run(self):
        target_conn = None
        source_conn = None
        try:
            mode_text = "预检查" if self.preview_only else "同步"
            self.log_message.emit(f"开始{mode_text}...")

            self.log_message.emit("正在连接财务系统数据库...")
            self.progress.emit(5)
            source_conn = self._connect(self.source_datasource)
            source_cursor = source_conn.cursor(pymysql.cursors.DictCursor)

            self.log_message.emit("正在连接进销存数据库...")
            self.progress.emit(10)
            target_conn = self._connect(self.target_datasource)
            target_cursor = target_conn.cursor(pymysql.cursors.DictCursor)

            # 财务货币（以它为主）。ea_dy_currency 只有 id/name/code，没有时间字段
            source_cursor.execute("SELECT id, name, code FROM ea_dy_currency ORDER BY id ASC")
            ea_rows = source_cursor.fetchall()
            total = len(ea_rows)
            self.results['total'] = total
            self.log_message.emit(f"财务系统共 {total} 条货币")

            if total == 0:
                self.progress.emit(100)
                self.finished.emit(self.results)
                return

            # 进销存现有货币（代码/中文名各建一份索引，代码不区分大小写）
            target_cursor.execute("""
                SELECT id, currency_cn, currency_en, finance_currency_id
                FROM ba_currency
                WHERE (delete_time IS NULL OR delete_time = 0)
                ORDER BY id ASC
            """)
            ba_rows = target_cursor.fetchall()
            ba_by_code = {}
            ba_by_name = {}
            for row in ba_rows:
                ba_by_code.setdefault((row['currency_en'] or '').strip().upper(), []).append(row)
                ba_by_name.setdefault((row['currency_cn'] or '').strip(), []).append(row)
            self.log_message.emit(f"进销存现有 {len(ba_rows)} 条货币")
            self.progress.emit(15)

            used_ba_ids = set()
            now = int(datetime.now().timestamp())

            for index, ea in enumerate(ea_rows):
                code = (ea['code'] or '').strip()
                name = (ea['name'] or '').strip()
                result_row = {
                    '财务货币ID': ea['id'],
                    '财务货币代码': code,
                    '财务货币名称': name,
                    '进销存货币ID': '',
                    '原关联财务ID': '',
                }
                try:
                    if not code and not name:
                        self.results['skip_count'] += 1
                        result_row['处理状态'] = '跳过(财务代码和名称都为空)'
                        self.results['completed_data'].append(result_row)
                        continue

                    # 先按代码匹配，代码匹配不上再按中文名匹配
                    candidates = [r for r in ba_by_code.get(code.upper(), []) if r['id'] not in used_ba_ids]
                    match_by = '代码'
                    if not candidates and name:
                        candidates = [r for r in ba_by_name.get(name, []) if r['id'] not in used_ba_ids]
                        match_by = '名称'

                    if candidates:
                        ba_row = candidates[0]
                        used_ba_ids.add(ba_row['id'])
                        result_row['进销存货币ID'] = ba_row['id']
                        result_row['原关联财务ID'] = ba_row['finance_currency_id'] or ''

                        if ba_row['finance_currency_id'] == ea['id']:
                            self.results['skip_count'] += 1
                            result_row['处理状态'] = '跳过(已关联)'
                            self.results['completed_data'].append(result_row)
                            continue

                        if not self.preview_only:
                            target_cursor.execute("""
                                UPDATE ba_currency
                                SET finance_currency_id = %s, update_time = %s
                                WHERE id = %s
                            """, (ea['id'], now, ba_row['id']))
                            target_conn.commit()

                        self.results['bind_count'] += 1
                        if ba_row['finance_currency_id']:
                            result_row['处理状态'] = ('可改绑' if self.preview_only else '已改绑') + \
                                f"(原财务ID={ba_row['finance_currency_id']},按{match_by}匹配)"
                        else:
                            result_row['处理状态'] = ('可关联' if self.preview_only else '已关联') + f"(按{match_by}匹配)"
                        self.results['completed_data'].append(result_row)
                        continue

                    # 进销存没有 → 新增并绑定
                    if not self.preview_only:
                        target_cursor.execute("""
                            INSERT INTO ba_currency
                                (currency_cn, currency_en, status, finance_currency_id,
                                 create_time, update_time)
                            VALUES (%s, %s, %s, %s, %s, %s)
                        """, (name or code, code or name, 1, ea['id'], now, now))
                        target_conn.commit()
                        new_id = target_cursor.lastrowid
                        used_ba_ids.add(new_id)
                        result_row['进销存货币ID'] = new_id

                    self.results['insert_count'] += 1
                    result_row['处理状态'] = '可新增并关联' if self.preview_only else '已新增并关联'
                    self.results['completed_data'].append(result_row)

                except Exception as e:
                    self.results['fail_count'] += 1
                    result_row['处理状态'] = f'异常({str(e)})'
                    self.results['completed_data'].append(result_row)
                    self.results['failed_records'].append(result_row)

                self.progress.emit(15 + int((index + 1) / total * 80))

            # 进销存有、财务没有的（仅提示，不处理）
            ea_codes = {(r['code'] or '').strip().upper() for r in ea_rows}
            ea_names = {(r['name'] or '').strip() for r in ea_rows}
            for row in ba_rows:
                if (row['currency_en'] or '').strip().upper() not in ea_codes \
                        and (row['currency_cn'] or '').strip() not in ea_names:
                    self.log_message.emit(f"提示: 进销存货币「{row['currency_cn']}/{row['currency_en']}」财务系统没有，未处理")

            source_cursor.close()
            source_conn.close()
            target_cursor.close()
            target_conn.close()

            self.progress.emit(100)
            self.log_message.emit(f"\n=== {mode_text}完成 ===")
            self.log_message.emit(f"财务货币总数: {total}")
            self.log_message.emit(f"关联: {self.results['bind_count']}")
            self.log_message.emit(f"新增并关联: {self.results['insert_count']}")
            self.log_message.emit(f"跳过: {self.results['skip_count']}")
            self.log_message.emit(f"失败: {self.results['fail_count']}")

            self.finished.emit(self.results)

        except Exception as e:
            try:
                if target_conn:
                    target_conn.rollback()
            except Exception:
                pass
            self.error.emit(str(e))


class CurrencyLinkWidget(QWidget):
    """货币关联财务系统界面"""

    def __init__(self, db_manager=None):
        super().__init__()
        self.db_manager = db_manager or DatabaseManager()
        self.worker = None
        self.completed_data = []
        self.failed_records = []
        self.init_ui()
        self.load_datasources()

    def init_ui(self):
        layout = QVBoxLayout(self)

        ds_group = QGroupBox("选择数据源")
        ds_layout = QFormLayout()

        self.target_datasource_combo = QComboBox()
        self.target_datasource_combo.currentIndexChanged.connect(self.on_datasource_changed)
        ds_layout.addRow("目标数据库 (进销存 ba_表):", self.target_datasource_combo)

        self.test_target_btn = QPushButton("测试目标连接")
        self.test_target_btn.setEnabled(False)
        self.test_target_btn.clicked.connect(self.test_target_connection)
        ds_layout.addRow("", self.test_target_btn)

        self.source_datasource_combo = QComboBox()
        self.source_datasource_combo.currentIndexChanged.connect(self.on_datasource_changed)
        ds_layout.addRow("源数据库 (财务系统 ea_表):", self.source_datasource_combo)

        self.test_source_btn = QPushButton("测试源连接")
        self.test_source_btn.setEnabled(False)
        self.test_source_btn.clicked.connect(self.test_source_connection)
        ds_layout.addRow("", self.test_source_btn)

        ds_group.setLayout(ds_layout)
        layout.addWidget(ds_group)

        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        info_label = QLabel("""
<b>货币关联财务系统说明:</b><br>
<b>数据流向:</b> 财务系统(源) → 进销存(目标)，<b>以财务系统数据为主，按货币代码匹配</b><br><br>
<b>处理逻辑:</b><br>
• 遍历财务库 ea_dy_currency（该表只有 id/name/code，无时间字段）<br>
• 按 code 匹配进销存 ba_currency.currency_en，匹配不上再按 name 匹配 currency_cn<br>
• 匹配上：回写 ba_currency.finance_currency_id（已关联到其他ID会改绑并记录）<br>
• 匹配不上：在进销存新增一条(currency_cn/currency_en)并直接绑定<br>
• 进销存有、财务没有的货币：仅在日志提示，不处理<br><br>
<b>注意:</b> 本功能和「账户分类关联财务系统」一样要最先跑；
「银行名称关联财务系统」依赖这里的 finance_currency_id 才能把财务银行名称上的货币刷到进销存！<br>
<b>建议先执行「预检查」确认结果后再执行「开始同步」</b>
        """)
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)

        btn_layout = QHBoxLayout()

        self.precheck_btn = QPushButton("预检查")
        self.precheck_btn.setEnabled(False)
        self.precheck_btn.clicked.connect(self.start_precheck)
        btn_layout.addWidget(self.precheck_btn)

        self.start_btn = QPushButton("开始同步")
        self.start_btn.setEnabled(False)
        self.start_btn.clicked.connect(self.start_sync)
        btn_layout.addWidget(self.start_btn)

        self.export_completed_btn = QPushButton("导出完整记录")
        self.export_completed_btn.setEnabled(False)
        self.export_completed_btn.clicked.connect(self.export_completed)
        btn_layout.addWidget(self.export_completed_btn)

        self.export_failed_btn = QPushButton("导出失败记录")
        self.export_failed_btn.setEnabled(False)
        self.export_failed_btn.clicked.connect(self.export_failed)
        btn_layout.addWidget(self.export_failed_btn)

        layout.addLayout(btn_layout)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        self.result_label = QLabel("")
        self.result_label.setAlignment(Qt.AlignCenter)
        font = QFont()
        font.setPointSize(12)
        self.result_label.setFont(font)
        layout.addWidget(self.result_label)

        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(250)
        log_layout.addWidget(self.log_text)

        clear_btn = QPushButton("清空日志")
        clear_btn.clicked.connect(self.log_text.clear)
        log_layout.addWidget(clear_btn)

        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        layout.addStretch()

    def load_datasources(self):
        self.target_datasource_combo.clear()
        self.source_datasource_combo.clear()

        self.target_datasource_combo.addItem("请选择目标数据源", None)
        self.source_datasource_combo.addItem("请选择源数据源", None)

        for ds in self.db_manager.get_all_datasources():
            display = f"{ds.name} ({ds.host}:{ds.port})"
            self.target_datasource_combo.addItem(display, ds)
            self.source_datasource_combo.addItem(display, ds)

    def on_datasource_changed(self):
        target_ok = self.target_datasource_combo.currentData() is not None
        source_ok = self.source_datasource_combo.currentData() is not None
        self.test_target_btn.setEnabled(target_ok)
        self.test_source_btn.setEnabled(source_ok)
        ready = target_ok and source_ok
        self.precheck_btn.setEnabled(ready)
        self.start_btn.setEnabled(ready)

    def _test_connection(self, ds):
        if not ds:
            return
        ok, msg = self.db_manager.test_connection(ds)
        if ok:
            QMessageBox.information(self, "连接成功", msg)
        else:
            QMessageBox.critical(self, "连接失败", msg)

    def test_target_connection(self):
        self._test_connection(self.target_datasource_combo.currentData())

    def test_source_connection(self):
        self._test_connection(self.source_datasource_combo.currentData())

    def start_precheck(self):
        target_ds = self.target_datasource_combo.currentData()
        source_ds = self.source_datasource_combo.currentData()
        if not target_ds or not source_ds:
            QMessageBox.warning(self, "警告", "请选择两个数据源！")
            return
        self._run_worker(target_ds, source_ds, preview_only=True)

    def start_sync(self):
        target_ds = self.target_datasource_combo.currentData()
        source_ds = self.source_datasource_combo.currentData()
        if not target_ds or not source_ds:
            QMessageBox.warning(self, "警告", "请选择两个数据源！")
            return

        reply = QMessageBox.question(
            self, "确认同步",
            f"确定要按财务系统刷进销存货币吗？\n\n"
            f"进销存(目标): {target_ds.name}\n"
            f"财务系统(源): {source_ds.name}\n\n"
            f"将回写 ba_currency.finance_currency_id，并新增财务有、进销存没有的货币\n"
            f"操作不可撤销，请确认！",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        self._run_worker(target_ds, source_ds, preview_only=False)

    def _run_worker(self, target_ds, source_ds, preview_only):
        self.precheck_btn.setEnabled(False)
        self.start_btn.setEnabled(False)
        self.export_completed_btn.setEnabled(False)
        self.export_failed_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在预检查..." if preview_only else "正在处理...")
        self.log_text.clear()
        self.completed_data = []
        self.failed_records = []

        self.worker = CurrencyLinkWorker(target_ds, source_ds, preview_only=preview_only)
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log_message.connect(self.append_log)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def append_log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{ts}] {msg}")

    def on_finished(self, results):
        self.precheck_btn.setEnabled(True)
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)

        self.completed_data = results.get('completed_data', [])
        self.failed_records = results.get('failed_records', [])

        text = (f"财务货币{results['total']}条 | 关联: {results['bind_count']} | "
                f"新增并关联: {results['insert_count']} | 跳过: {results['skip_count']} | "
                f"失败: {results['fail_count']}")
        self.result_label.setStyleSheet("color: orange;" if results['fail_count'] else "color: green;")
        self.result_label.setText(text)

        self.export_completed_btn.setEnabled(len(self.completed_data) > 0)
        self.export_failed_btn.setEnabled(len(self.failed_records) > 0)

    def on_error(self, error_msg):
        self.precheck_btn.setEnabled(True)
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.result_label.setText("处理失败")
        self.result_label.setStyleSheet("color: red;")
        self.append_log(f"错误: {error_msg}")
        QMessageBox.critical(self, "错误", f"处理失败:\n{error_msg}")

    def export_completed(self):
        if not self.completed_data:
            QMessageBox.warning(self, "警告", "没有可导出的数据！")
            return
        default_name = f"货币关联完整记录_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "保存完整记录", default_name, "Excel文件 (*.xlsx)")
        if path:
            pd.DataFrame(self.completed_data).to_excel(path, index=False, engine='openpyxl')
            QMessageBox.information(self, "成功", f"已导出 {len(self.completed_data)} 条记录到:\n{path}")

    def export_failed(self):
        if not self.failed_records:
            QMessageBox.warning(self, "警告", "没有失败记录！")
            return
        default_name = f"货币关联失败记录_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "保存失败记录", default_name, "Excel文件 (*.xlsx)")
        if path:
            pd.DataFrame(self.failed_records).to_excel(path, index=False, engine='openpyxl')
            QMessageBox.information(self, "成功", f"已导出 {len(self.failed_records)} 条失败记录到:\n{path}")
