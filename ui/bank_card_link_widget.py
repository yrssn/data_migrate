"""
银行卡关联财务系统

以财务系统 ea_dy_legal_cards 为主，按「法人姓名 + 卡号」关联进销存 ba_zhb_bank：
- 能匹配上：回写 ba_zhb_bank.finance_card_id，并按财务的账户分类同步 account_classify_id
- 匹配不上：在进销存新增一条银行卡并直接绑定（只处理进销存已存在的法人）

依赖「银行名称关联财务系统」「账户分类关联财务系统」的结果：
银行名称通过 ba_zhb_bank_name.finance_bankcard_id 映射，账户分类通过
ba_account_classify.finance_classify_id 映射，请先跑完那两个功能。
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


# 财务银行卡类型 → 进销存银行卡类型（0=第三方, 1=个人, 2=企业）
FINANCE_CARD_TYPE_MAP = {0: 'opt2', 1: 'opt0', 2: 'opt1'}


class BankCardLinkWorker(QThread):
    """银行卡关联工作线程"""
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
            'classify_update_count': 0,
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
            self.progress.emit(8)
            target_conn = self._connect(self.target_datasource)
            target_cursor = target_conn.cursor(pymysql.cursors.DictCursor)

            # ---------- 预加载映射 ----------
            # 财务币种
            source_cursor.execute("SELECT id, code FROM ea_dy_currency")
            ea_currency_code = {r['id']: (r['code'] or '').strip() for r in source_cursor.fetchall()}

            # 进销存币种：currency_en → id
            target_cursor.execute("""
                SELECT id, currency_en FROM ba_currency
                WHERE (delete_time IS NULL OR delete_time = 0)
            """)
            ba_currency_by_code = {}
            for r in target_cursor.fetchall():
                ba_currency_by_code.setdefault((r['currency_en'] or '').strip(), r['id'])

            # 财务法人
            source_cursor.execute("""
                SELECT id, name FROM ea_dy_legal
                WHERE (delete_time IS NULL OR delete_time = 0)
            """)
            ea_legal_name = {r['id']: (r['name'] or '').strip() for r in source_cursor.fetchall()}

            # 进销存客户：法人姓名 → id
            target_cursor.execute("""
                SELECT id, legal_name FROM ba_rlb_customer
                WHERE (delete_time IS NULL OR delete_time = 0)
            """)
            ba_customer_by_legal = {}
            for r in target_cursor.fetchall():
                ba_customer_by_legal.setdefault((r['legal_name'] or '').strip(), r['id'])
            self.log_message.emit(f"进销存共 {len(ba_customer_by_legal)} 个法人(客户)")

            # 财务银行名称
            source_cursor.execute("""
                SELECT id, name FROM ea_dy_bankcard
                WHERE (delete_time IS NULL OR delete_time = 0)
            """)
            ea_bankcard_name = {r['id']: (r['name'] or '').strip() for r in source_cursor.fetchall()}

            # 进销存银行名称：财务ID → id，名称 → id
            target_cursor.execute("""
                SELECT id, bank_name, finance_bankcard_id FROM ba_zhb_bank_name
                WHERE (delete_time IS NULL OR delete_time = 0)
            """)
            ba_bank_name_by_finance_id = {}
            ba_bank_name_by_name = {}
            for r in target_cursor.fetchall():
                if r['finance_bankcard_id']:
                    ba_bank_name_by_finance_id.setdefault(r['finance_bankcard_id'], r['id'])
                ba_bank_name_by_name.setdefault((r['bank_name'] or '').strip(), r['id'])
            self.log_message.emit(
                f"进销存银行名称已关联财务的有 {len(ba_bank_name_by_finance_id)} 条"
            )

            # 财务账户分类
            source_cursor.execute("""
                SELECT id, name FROM ea_dy_account_classify
                WHERE (delete_time IS NULL OR delete_time = 0)
            """)
            ea_classify_name = {r['id']: (r['name'] or '').strip() for r in source_cursor.fetchall()}

            # 进销存账户分类：财务ID → id，名称 → id
            target_cursor.execute("""
                SELECT id, name, finance_classify_id FROM ba_account_classify
                WHERE (delete_time IS NULL OR delete_time = 0)
            """)
            ba_classify_by_finance_id = {}
            ba_classify_by_name = {}
            for r in target_cursor.fetchall():
                if r['finance_classify_id']:
                    ba_classify_by_finance_id.setdefault(r['finance_classify_id'], r['id'])
                ba_classify_by_name.setdefault((r['name'] or '').strip(), r['id'])

            # ---------- 财务银行卡（以它为主） ----------
            source_cursor.execute("""
                SELECT id, legal_id, currency_id, name, account, status, balance,
                       bankcard_id, account_classify, bank_card_type, remark
                FROM ea_dy_legal_cards
                WHERE (delete_time IS NULL OR delete_time = 0)
                ORDER BY id ASC
            """)
            ea_rows = source_cursor.fetchall()
            total = len(ea_rows)
            self.results['total'] = total
            self.log_message.emit(f"财务系统共 {total} 张银行卡")

            if total == 0:
                self.progress.emit(100)
                self.finished.emit(self.results)
                return

            # 进销存现有银行卡：(客户ID, 卡号) → 记录
            target_cursor.execute("""
                SELECT id, bank_id, rlb_customer_id, bank_name_id, currency_id,
                       bank_card_type, account_classify_id, finance_card_id
                FROM ba_zhb_bank
                WHERE (delete_time IS NULL OR delete_time = 0)
                ORDER BY id ASC
            """)
            ba_rows = target_cursor.fetchall()
            ba_by_key = {}
            for row in ba_rows:
                key = (row['rlb_customer_id'], (row['bank_id'] or '').strip())
                ba_by_key.setdefault(key, []).append(row)
            self.log_message.emit(f"进销存现有 {len(ba_rows)} 张银行卡")
            self.progress.emit(15)

            used_ba_ids = set()
            now = int(datetime.now().timestamp())

            for index, ea in enumerate(ea_rows):
                account = (ea['account'] or '').strip()
                legal_name = ea_legal_name.get(ea['legal_id'], '')
                bank_name = ea_bankcard_name.get(ea['bankcard_id'], '') or (ea['name'] or '').strip()
                classify_name = ea_classify_name.get(ea['account_classify'], '')
                result_row = {
                    '财务银行卡ID': ea['id'],
                    '卡号': account,
                    '法人姓名': legal_name,
                    '银行名称': bank_name,
                    '财务账户分类': f"{ea['account_classify'] or ''}/{classify_name}",
                    '财务银行卡类型': ea['bank_card_type'],
                    '进销存银行卡ID': '',
                    '进销存客户ID': '',
                    '原关联财务ID': '',
                }
                try:
                    if not account:
                        self.results['skip_count'] += 1
                        result_row['处理状态'] = '跳过(财务卡号为空)'
                        self.results['completed_data'].append(result_row)
                        continue

                    if not legal_name:
                        self.results['fail_count'] += 1
                        result_row['处理状态'] = f"失败(财务法人ID={ea['legal_id']}无对应法人)"
                        self.results['completed_data'].append(result_row)
                        self.results['failed_records'].append(result_row)
                        continue

                    customer_id = ba_customer_by_legal.get(legal_name)
                    result_row['进销存客户ID'] = customer_id or ''
                    if not customer_id:
                        self.results['fail_count'] += 1
                        result_row['处理状态'] = f'失败(进销存没有该法人: {legal_name})'
                        self.results['completed_data'].append(result_row)
                        self.results['failed_records'].append(result_row)
                        continue

                    # 账户分类映射
                    ba_classify_id = ba_classify_by_finance_id.get(ea['account_classify'])
                    if not ba_classify_id and classify_name:
                        ba_classify_id = ba_classify_by_name.get(classify_name)
                    classify_warn = ''
                    if not ba_classify_id:
                        classify_warn = f"[账户分类未匹配:{ea['account_classify'] or ''}/{classify_name}]"

                    candidates = [
                        r for r in ba_by_key.get((customer_id, account), [])
                        if r['id'] not in used_ba_ids
                    ]

                    if candidates:
                        ba_row = candidates[0]
                        used_ba_ids.add(ba_row['id'])
                        result_row['进销存银行卡ID'] = ba_row['id']
                        result_row['原关联财务ID'] = ba_row['finance_card_id'] or ''

                        need_bind = ba_row['finance_card_id'] != ea['id']
                        need_classify = bool(ba_classify_id) and \
                            ba_row['account_classify_id'] != ba_classify_id

                        if not need_bind and not need_classify:
                            self.results['skip_count'] += 1
                            result_row['处理状态'] = '跳过(已关联且分类一致)' + classify_warn
                            self.results['completed_data'].append(result_row)
                            continue

                        if not self.preview_only:
                            sets = ['finance_card_id = %s']
                            params = [ea['id']]
                            if need_classify:
                                sets.append('account_classify_id = %s')
                                params.append(ba_classify_id)
                            sets.append('update_time = %s')
                            params.append(now)
                            params.append(ba_row['id'])
                            target_cursor.execute(
                                f"UPDATE ba_zhb_bank SET {', '.join(sets)} WHERE id = %s",
                                params
                            )
                            target_conn.commit()

                        status_parts = []
                        if need_bind:
                            self.results['bind_count'] += 1
                            if ba_row['finance_card_id']:
                                status_parts.append(
                                    ('可改绑' if self.preview_only else '已改绑') +
                                    f"(原财务ID={ba_row['finance_card_id']})"
                                )
                            else:
                                status_parts.append('可关联' if self.preview_only else '已关联')
                        if need_classify:
                            self.results['classify_update_count'] += 1
                            status_parts.append(
                                ('可更新账户分类' if self.preview_only else '已更新账户分类') +
                                f"({ba_row['account_classify_id']}→{ba_classify_id})"
                            )
                        result_row['处理状态'] = '、'.join(status_parts) + classify_warn
                        self.results['completed_data'].append(result_row)
                        continue

                    # 进销存没有 → 新增并绑定
                    bank_name_id = ba_bank_name_by_finance_id.get(ea['bankcard_id'])
                    if not bank_name_id and bank_name:
                        bank_name_id = ba_bank_name_by_name.get(bank_name)
                    if not bank_name_id:
                        self.results['fail_count'] += 1
                        result_row['处理状态'] = f'失败(进销存没有该银行名称: {bank_name}，请先跑银行名称关联)'
                        self.results['completed_data'].append(result_row)
                        self.results['failed_records'].append(result_row)
                        continue

                    currency_code = ea_currency_code.get(ea['currency_id'], '')
                    currency_id = ba_currency_by_code.get(currency_code)
                    card_type = FINANCE_CARD_TYPE_MAP.get(ea['bank_card_type'], 'opt2')
                    status = 1 if ea['status'] is None else int(ea['status'])

                    if not self.preview_only:
                        fields = ['bank_name_id', 'bank_id', 'rlb_customer_id', 'bank_card_type',
                                  'status', 'bank_balance', 'account_classify_id', 'finance_card_id',
                                  'remark', 'admin_id', 'admin_dept_id', 'create_time', 'update_time']
                        values = [bank_name_id, account, customer_id, card_type,
                                  status, ea['balance'] if ea['balance'] is not None else 0.00,
                                  ba_classify_id or 1, ea['id'],
                                  '脚本自动补充', 1, 1, now, now]
                        if currency_id:
                            fields.append('currency_id')
                            values.append(currency_id)
                        target_cursor.execute(
                            f"INSERT INTO ba_zhb_bank ({', '.join(fields)}) "
                            f"VALUES ({', '.join(['%s'] * len(fields))})",
                            values
                        )
                        target_conn.commit()
                        new_id = target_cursor.lastrowid
                        used_ba_ids.add(new_id)
                        result_row['进销存银行卡ID'] = new_id

                    self.results['insert_count'] += 1
                    extra = '' if currency_id else f"[币种未匹配:{currency_code}]"
                    result_row['处理状态'] = ('可新增并关联' if self.preview_only else '已新增并关联') \
                        + classify_warn + extra
                    self.results['completed_data'].append(result_row)

                except Exception as e:
                    self.results['fail_count'] += 1
                    result_row['处理状态'] = f'异常({str(e)})'
                    self.results['completed_data'].append(result_row)
                    self.results['failed_records'].append(result_row)

                self.progress.emit(15 + int((index + 1) / total * 80))

            source_cursor.close()
            source_conn.close()
            target_cursor.close()
            target_conn.close()

            self.progress.emit(100)
            self.log_message.emit(f"\n=== {mode_text}完成 ===")
            self.log_message.emit(f"财务银行卡总数: {total}")
            self.log_message.emit(f"关联: {self.results['bind_count']}")
            self.log_message.emit(f"新增并关联: {self.results['insert_count']}")
            self.log_message.emit(f"更新账户分类: {self.results['classify_update_count']}")
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


class BankCardLinkWidget(QWidget):
    """银行卡关联财务系统界面"""

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
<b>银行卡关联财务系统说明:</b><br>
<b>数据流向:</b> 财务系统(源) → 进销存(目标)，<b>以财务系统数据为主</b><br><br>
<b>处理逻辑:</b><br>
• 遍历财务库 ea_dy_legal_cards 中未删除的银行卡<br>
• legal_id → ea_dy_legal.name → 进销存 ba_rlb_customer.legal_name 得到 rlb_customer_id<br>
• 按「法人(rlb_customer_id) + 卡号(bank_id)」匹配 ba_zhb_bank<br>
• 匹配上：回写 ba_zhb_bank.finance_card_id，并把 account_classify_id 同步成财务的<br>
• 匹配不上：新增 ba_zhb_bank 并直接绑定（银行名称/币种/类型/状态/余额/账户分类一并写入）<br>
• 进销存没有该法人：记为失败（只拉现有法人的卡），请先补客户档案<br><br>
<b>字段映射:</b><br>
• bank_name_id: ea_dy_legal_cards.bankcard_id → ba_zhb_bank_name.finance_bankcard_id，
退化为按银行名称匹配<br>
• account_classify_id: ea_dy_legal_cards.account_classify →
ba_account_classify.finance_classify_id，退化为按分类名称匹配<br>
• currency_id: ea_dy_currency.code 匹配 ba_currency.currency_en<br>
• bank_card_type: 0→opt2(第三方), 1→opt0(个人), 2→opt1(企业)<br>
• remark 固定为「脚本自动补充」，admin_id / admin_dept_id 固定为1<br><br>
<b>注意:</b> 请先执行「账户分类关联财务系统」和「银行名称关联财务系统」！<br>
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
            f"确定要按财务系统刷进销存银行卡吗？\n\n"
            f"进销存(目标): {target_ds.name}\n"
            f"财务系统(源): {source_ds.name}\n\n"
            f"将回写 ba_zhb_bank.finance_card_id、同步 account_classify_id，"
            f"并新增财务有、进销存没有的银行卡\n"
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

        self.worker = BankCardLinkWorker(target_ds, source_ds, preview_only=preview_only)
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

        text = (f"财务银行卡{results['total']}张 | 关联: {results['bind_count']} | "
                f"新增并关联: {results['insert_count']} | 更新分类: {results['classify_update_count']} | "
                f"跳过: {results['skip_count']} | 失败: {results['fail_count']}")
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
        default_name = f"银行卡关联完整记录_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "保存完整记录", default_name, "Excel文件 (*.xlsx)")
        if path:
            pd.DataFrame(self.completed_data).to_excel(path, index=False, engine='openpyxl')
            QMessageBox.information(self, "成功", f"已导出 {len(self.completed_data)} 条记录到:\n{path}")

    def export_failed(self):
        if not self.failed_records:
            QMessageBox.warning(self, "警告", "没有失败记录！")
            return
        default_name = f"银行卡关联失败记录_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "保存失败记录", default_name, "Excel文件 (*.xlsx)")
        if path:
            pd.DataFrame(self.failed_records).to_excel(path, index=False, engine='openpyxl')
            QMessageBox.information(self, "成功", f"已导出 {len(self.failed_records)} 条失败记录到:\n{path}")
