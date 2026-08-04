"""
银行卡关联财务系统

以财务系统 ea_dy_legal_cards 为主，按「法人姓名 + 卡号」关联进销存 ba_zhb_bank：
- 能匹配上：回写 ba_zhb_bank.finance_card_id，并按财务的账户分类同步 account_classify_id
- 匹配不上：在进销存新增一条银行卡并直接绑定（只处理进销存已存在的法人）

依赖「银行名称关联财务系统」「账户分类关联财务系统」的结果：
银行名称通过 ba_zhb_bank_name.finance_bankcard_id 映射，账户分类通过
ba_account_classify.finance_classify_id 映射，请先跑完那两个功能。
"""
from collections import Counter
from datetime import datetime
from decimal import Decimal
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from database import DatabaseManager
import pandas as pd
import pymysql


# 财务银行卡类型(0未知/第三方,1个人,2企业) → 进销存(opt0个人,opt1企业,opt2第三方)
FINANCE_CARD_TYPE_MAP = {0: 'opt2', 1: 'opt0', 2: 'opt1'}
# 进销存 ba_zhb_bank.bank_balance 上限 decimal(10,2)
MAX_BALANCE = Decimal('99999999.99')


def norm_account(value) -> str:
    """卡号归一化：去空白/横线等分隔符"""
    text = '' if value is None else str(value)
    for ch in (' ', '\t', '\u3000', '-', '\uff0d', '\u2010', '\u2212', '_', '.', '\uff0e'):
        text = text.replace(ch, '')
    return text.strip()


def norm_name(value) -> str:
    """名称归一化：去空白"""
    text = '' if value is None else str(value)
    for ch in (' ', '\t', '\u3000'):
        text = text.replace(ch, '')
    return text.strip()


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
            'field_update_count': 0,
            'skip_count': 0,
            'fail_count': 0,
            'completed_data': [],
            'failed_records': [],
            'unlinked_target': []
        }
        self.reason_stats = Counter()

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
            ba_customer_name = {}
            for r in target_cursor.fetchall():
                ba_customer_by_legal.setdefault(norm_name(r['legal_name']), r['id'])
                ba_customer_name[r['id']] = (r['legal_name'] or '').strip()
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
                ba_bank_name_by_name.setdefault(norm_name(r['bank_name']), r['id'])
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
                ba_classify_by_name.setdefault(norm_name(r['name']), r['id'])

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

            # 进销存现有银行卡：按(客户ID,卡号)和卡号两级索引
            target_cursor.execute("""
                SELECT id, bank_id, rlb_customer_id, bank_name_id, currency_id,
                       bank_card_type, account_classify_id, finance_card_id,
                       status, bank_balance
                FROM ba_zhb_bank
                WHERE (delete_time IS NULL OR delete_time = 0)
                ORDER BY id ASC
            """)
            ba_rows = target_cursor.fetchall()
            ba_by_account = {}
            for row in ba_rows:
                ba_by_account.setdefault(norm_account(row['bank_id']), []).append(row)
            self.log_message.emit(f"进销存现有 {len(ba_rows)} 张银行卡")
            self.progress.emit(15)

            used_ba_ids = set()
            bound_ba_by_account = {}
            now = int(datetime.now().timestamp())

            for index, ea in enumerate(ea_rows):
                account = (ea['account'] or '').strip()
                acct_key = norm_account(account)
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

                    customer_id = ba_customer_by_legal.get(norm_name(legal_name))
                    result_row['进销存客户ID'] = customer_id or ''

                    # 账户分类映射
                    ba_classify_id = ba_classify_by_finance_id.get(ea['account_classify'])
                    if not ba_classify_id and classify_name:
                        ba_classify_id = ba_classify_by_name.get(norm_name(classify_name))
                    classify_warn = ''
                    if not ba_classify_id:
                        classify_warn = f"[账户分类未匹配:{ea['account_classify'] or ''}/{classify_name}]"

                    # 银行名称/币种/类型/状态/余额（均以财务为准）
                    bank_name_id = ba_bank_name_by_finance_id.get(ea['bankcard_id'])
                    if not bank_name_id and bank_name:
                        bank_name_id = ba_bank_name_by_name.get(norm_name(bank_name))
                    currency_code = ea_currency_code.get(ea['currency_id'], '')
                    currency_id = ba_currency_by_code.get(currency_code)
                    card_type = FINANCE_CARD_TYPE_MAP.get(ea['bank_card_type'], 'opt2')
                    status = 1 if ea['status'] is None else int(ea['status'])
                    balance = ea['balance'] if ea['balance'] is not None else Decimal('0.00')
                    if abs(Decimal(balance)) > MAX_BALANCE:
                        balance = None  # 超出进销存字段范围，不同步余额

                    # 按卡号匹配，同卡号有多条时按「法人/银行名称/卡类型」打分选最优
                    same_account_rows = ba_by_account.get(acct_key, [])
                    candidates = [r for r in same_account_rows if r['id'] not in used_ba_ids]
                    match_note = ''
                    if len(same_account_rows) > 1:
                        match_note = f"[进销存同卡号有{len(same_account_rows)}条]"

                    if candidates:
                        def score(row):
                            value = 0
                            if customer_id and row['rlb_customer_id'] == customer_id:
                                value += 8
                            if bank_name_id and row['bank_name_id'] == bank_name_id:
                                value += 4
                            if row['bank_card_type'] == card_type:
                                value += 2
                            if row['finance_card_id'] == ea['id']:
                                value += 16
                            elif row['finance_card_id']:
                                value -= 1  # 已绑定到其他财务卡的尽量不动
                            return value

                        candidates.sort(key=score, reverse=True)
                        if customer_id and candidates[0]['rlb_customer_id'] == customer_id:
                            self.reason_stats['法人+卡号匹配'] += 1
                        else:
                            other_legal = ba_customer_name.get(candidates[0]['rlb_customer_id'], '')
                            match_note += f"[仅卡号匹配,进销存法人={other_legal or '空'}]"
                            self.reason_stats['仅卡号匹配'] += 1
                        ba_row = candidates[0]
                        used_ba_ids.add(ba_row['id'])
                        bound_ba_by_account[acct_key] = ba_row['id']
                        result_row['进销存银行卡ID'] = ba_row['id']
                        result_row['原关联财务ID'] = ba_row['finance_card_id'] or ''

                        sets = []
                        params = []
                        status_parts = []

                        need_bind = ba_row['finance_card_id'] != ea['id']
                        if need_bind:
                            sets.append('finance_card_id = %s')
                            params.append(ea['id'])
                            self.results['bind_count'] += 1
                            if ba_row['finance_card_id']:
                                status_parts.append(
                                    ('可改绑' if self.preview_only else '已改绑') +
                                    f"(原财务ID={ba_row['finance_card_id']})"
                                )
                            else:
                                status_parts.append('可关联' if self.preview_only else '已关联')

                        if ba_classify_id and ba_row['account_classify_id'] != ba_classify_id:
                            sets.append('account_classify_id = %s')
                            params.append(ba_classify_id)
                            self.results['classify_update_count'] += 1
                            status_parts.append(
                                ('可改账户分类' if self.preview_only else '已改账户分类') +
                                f"({ba_row['account_classify_id']}→{ba_classify_id})"
                            )

                        # 其余字段以财务为准
                        field_changes = []
                        if ea['bank_card_type'] is not None and ba_row['bank_card_type'] != card_type:
                            sets.append('bank_card_type = %s')
                            params.append(card_type)
                            field_changes.append(f"银行卡类型({ba_row['bank_card_type']}→{card_type})")
                        if bank_name_id and ba_row['bank_name_id'] != bank_name_id:
                            sets.append('bank_name_id = %s')
                            params.append(bank_name_id)
                            field_changes.append(f"银行名称({ba_row['bank_name_id']}→{bank_name_id})")
                        if currency_id and ba_row['currency_id'] != currency_id:
                            sets.append('currency_id = %s')
                            params.append(currency_id)
                            field_changes.append(f"币种({ba_row['currency_id']}→{currency_id})")
                        if ba_row['status'] != status:
                            sets.append('status = %s')
                            params.append(status)
                            field_changes.append(f"状态({ba_row['status']}→{status})")
                        if balance is not None and (
                            ba_row['bank_balance'] is None or
                            Decimal(ba_row['bank_balance']) != Decimal(balance)
                        ):
                            sets.append('bank_balance = %s')
                            params.append(balance)
                            field_changes.append(f"余额({ba_row['bank_balance']}→{balance})")
                        if customer_id and not ba_row['rlb_customer_id']:
                            sets.append('rlb_customer_id = %s')
                            params.append(customer_id)
                            field_changes.append(f"法人(空→{customer_id})")

                        if field_changes:
                            self.results['field_update_count'] += 1
                            status_parts.append(
                                ('可同步' if self.preview_only else '已同步') + ':' + ','.join(field_changes)
                            )

                        if not sets:
                            self.results['skip_count'] += 1
                            self.reason_stats['已关联且信息一致'] += 1
                            result_row['处理状态'] = '跳过(已关联且信息一致)' + classify_warn + match_note
                            self.results['completed_data'].append(result_row)
                            continue

                        if not self.preview_only:
                            sets.append('update_time = %s')
                            params.append(now)
                            params.append(ba_row['id'])
                            target_cursor.execute(
                                f"UPDATE ba_zhb_bank SET {', '.join(sets)} WHERE id = %s",
                                params
                            )
                            target_conn.commit()

                        result_row['处理状态'] = '、'.join(status_parts) + classify_warn + match_note
                        self.results['completed_data'].append(result_row)
                        continue

                    # 进销存没有 → 新增并绑定（需法人和银行名称都能对上）
                    if not customer_id:
                        self.results['skip_count'] += 1
                        reason = '进销存无此卡号且无此法人,未新增' if legal_name \
                            else '财务法人不存在,未新增'
                        self.reason_stats[reason] += 1
                        result_row['处理状态'] = f'跳过未新增({reason}: {legal_name or ea["legal_id"]})'
                        self.results['completed_data'].append(result_row)
                        self.results['failed_records'].append(result_row)
                        continue

                    if not bank_name_id:
                        self.results['skip_count'] += 1
                        self.reason_stats['进销存没有该银行名称,未新增'] += 1
                        result_row['处理状态'] = f'跳过未新增(进销存没有该银行名称: {bank_name}，请先跑银行名称关联)'
                        self.results['completed_data'].append(result_row)
                        self.results['failed_records'].append(result_row)
                        continue

                    if not self.preview_only:
                        fields = ['bank_name_id', 'bank_id', 'rlb_customer_id', 'bank_card_type',
                                  'status', 'bank_balance', 'account_classify_id', 'finance_card_id',
                                  'remark', 'admin_id', 'admin_dept_id', 'create_time', 'update_time']
                        values = [bank_name_id, account, customer_id, card_type,
                                  status, balance if balance is not None else Decimal('0.00'),
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
                    self.reason_stats['进销存无此卡号,新增'] += 1
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

            # ---------- 进销存侧未关联的记录（供导出排查） ----------
            ea_by_account = {}
            for ea in ea_rows:
                ea_by_account.setdefault(norm_account(ea['account']), []).append(ea)

            target_cursor.execute("""
                SELECT b.id, b.bank_id, b.rlb_customer_id, b.bank_card_type,
                       b.account_classify_id, b.status, b.bank_balance, b.finance_card_id,
                       c.legal_name, n.bank_name
                FROM ba_zhb_bank b
                LEFT JOIN ba_rlb_customer c ON c.id = b.rlb_customer_id
                LEFT JOIN ba_zhb_bank_name n ON n.id = b.bank_name_id
                WHERE (b.delete_time IS NULL OR b.delete_time = 0)
                  AND b.finance_card_id IS NULL
                ORDER BY b.id ASC
            """)
            for row in target_cursor.fetchall():
                if self.preview_only and row['id'] in used_ba_ids:
                    continue  # 预检查时这些本次会被关联上
                acct = norm_account(row['bank_id'])
                same_account = ea_by_account.get(acct, [])
                if same_account:
                    ea_hit = same_account[0]
                    hit_desc = (f"财务卡ID={ea_hit['id']},法人="
                                f"{ea_legal_name.get(ea_hit['legal_id'], '') or '?'}")
                    bound_id = bound_ba_by_account.get(acct)
                    dup_count = len(ba_by_account.get(acct, []))
                    if bound_id and bound_id != row['id']:
                        note = (f"进销存同卡号有{dup_count}条重复记录，该财务卡({hit_desc})"
                                f"已绑到进销存ID={bound_id}，一张财务卡只能绑一条")
                    else:
                        note = f"财务有相同卡号({hit_desc})但本次未绑定，需排查"
                else:
                    note = '财务系统没有这个卡号'
                self.results['unlinked_target'].append({
                    '进销存银行卡ID': row['id'],
                    '卡号': row['bank_id'],
                    '法人': row['legal_name'] or '',
                    '银行名称': row['bank_name'] or '',
                    '银行卡类型': row['bank_card_type'],
                    '账户分类ID': row['account_classify_id'],
                    '状态': row['status'],
                    '余额': row['bank_balance'],
                    '未关联原因': note,
                })
            self.log_message.emit(
                f"进销存仍未关联财务的银行卡: {len(self.results['unlinked_target'])} 张（可导出）"
            )

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
            self.log_message.emit(f"同步其余字段(类型/银行/币种/状态/余额): {self.results['field_update_count']}")
            self.log_message.emit(f"跳过: {self.results['skip_count']}")
            self.log_message.emit(f"失败: {self.results['fail_count']}")
            if self.reason_stats:
                self.log_message.emit("分类统计:")
                for reason, count in self.reason_stats.most_common():
                    self.log_message.emit(f"  {reason}: {count}")

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
        self.unlinked_target = []
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
• 按卡号匹配 ba_zhb_bank，同一卡号有多条时按「法人/银行名称/卡类型」打分选最优那条，
法人对不上时仍关联并标注进销存当前法人（一张财务卡只能绑一条进销存记录）<br>
• 卡号匹配时忽略空格/横线等分隔符，法人名忽略空白<br>
• 匹配上：回写 finance_card_id，并把账户分类、银行卡类型、银行名称、币种、状态、余额
全部刷成财务的（以财务为准，只改有差异的字段）<br>
• 匹配不上：新增 ba_zhb_bank 并直接绑定（银行名称/币种/类型/状态/余额/账户分类一并写入）<br>
• 匹配不上且进销存没有该法人/该银行名称：不新增，计入跳过并写入待处理清单（可导出）<br>
• 跑完可用「导出进销存未关联记录」导出 ba_zhb_bank 中 finance_card_id 仍为空的卡，
并标注该卡号在财务是否存在、存在时是哪张卡/哪个法人<br><br>
<b>字段映射:</b><br>
• bank_name_id: ea_dy_legal_cards.bankcard_id → ba_zhb_bank_name.finance_bankcard_id，
退化为按银行名称匹配<br>
• account_classify_id: ea_dy_legal_cards.account_classify →
ba_account_classify.finance_classify_id，退化为按分类名称匹配<br>
• currency_id: ea_dy_currency.code 匹配 ba_currency.currency_en<br>
• bank_card_type: 0未知/第三方→opt2, 1个人→opt0, 2企业→opt1<br>
• bank_balance: 超过进销存 decimal(10,2) 范围的不同步余额<br>
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

        self.export_failed_btn = QPushButton("导出待处理/失败记录")
        self.export_failed_btn.setEnabled(False)
        self.export_failed_btn.clicked.connect(self.export_failed)
        btn_layout.addWidget(self.export_failed_btn)

        self.export_unlinked_btn = QPushButton("导出进销存未关联记录")
        self.export_unlinked_btn.setEnabled(False)
        self.export_unlinked_btn.clicked.connect(self.export_unlinked)
        btn_layout.addWidget(self.export_unlinked_btn)

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
        self.unlinked_target = []
        self.export_unlinked_btn.setEnabled(False)

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
        self.unlinked_target = results.get('unlinked_target', [])

        text = (f"财务银行卡{results['total']}张 | 关联: {results['bind_count']} | "
                f"新增并关联: {results['insert_count']} | 改分类: {results['classify_update_count']} | "
                f"同步其余字段: {results.get('field_update_count', 0)} | "
                f"跳过: {results['skip_count']} | 失败: {results['fail_count']}")
        self.result_label.setStyleSheet("color: orange;" if results['fail_count'] else "color: green;")
        self.result_label.setText(text)

        self.export_completed_btn.setEnabled(len(self.completed_data) > 0)
        self.export_failed_btn.setEnabled(len(self.failed_records) > 0)
        self.export_unlinked_btn.setEnabled(len(self.unlinked_target) > 0)

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

    def export_unlinked(self):
        if not self.unlinked_target:
            QMessageBox.warning(self, "警告", "没有未关联记录！")
            return
        default_name = f"进销存未关联银行卡_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "保存进销存未关联记录", default_name, "Excel文件 (*.xlsx)")
        if path:
            pd.DataFrame(self.unlinked_target).to_excel(path, index=False, engine='openpyxl')
            QMessageBox.information(self, "成功", f"已导出 {len(self.unlinked_target)} 条记录到:\n{path}")

    def export_failed(self):
        if not self.failed_records:
            QMessageBox.warning(self, "警告", "没有失败记录！")
            return
        default_name = f"银行卡关联失败记录_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "保存失败记录", default_name, "Excel文件 (*.xlsx)")
        if path:
            pd.DataFrame(self.failed_records).to_excel(path, index=False, engine='openpyxl')
            QMessageBox.information(self, "成功", f"已导出 {len(self.failed_records)} 条失败记录到:\n{path}")
