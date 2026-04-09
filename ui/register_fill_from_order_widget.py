"""
注册表补充平台ID和属地ID
从 ba_order 表补充 ba_ptzcb_register 中 platform_id 或 shudi_id 为空的记录
通过 order_id 关联
"""
import os
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from database import DatabaseManager
import pandas as pd
import pymysql


class RegisterFillFromOrderWorker(QThread):
    """补充工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, datasource, preview_only=False):
        super().__init__()
        self.datasource = datasource
        self.preview_only = preview_only
        self.results = {
            'total_empty': 0,
            'can_fill': 0,
            'filled_platform': 0,
            'filled_shudi': 0,
            'no_order_id': 0,
            'order_not_found': 0,
            'order_also_empty': 0,
            'completed_data': [],
            'failed_records': []
        }

    def run(self):
        try:
            mode = "预检查" if self.preview_only else "补充"
            self.log_message.emit(f"开始{mode}...")

            self.log_message.emit("正在连接数据库...")
            self.progress.emit(5)
            connection = pymysql.connect(
                host=self.datasource.host,
                port=self.datasource.port,
                user=self.datasource.username,
                password=self.datasource.password,
                database=self.datasource.database,
                charset=self.datasource.charset
            )
            cursor = connection.cursor(pymysql.cursors.DictCursor)
            self.log_message.emit("数据库连接成功")
            self.progress.emit(10)

            # 查找 platform_id 或 shudi_id 为空的注册表记录
            self.log_message.emit("正在查询 platform_id 或 shudi_id 为空的注册记录...")
            cursor.execute("""
                SELECT id, project_id, order_id, platform_id, shudi_id,
                       shopindex_id, status, email, company_name, remark
                FROM ba_ptzcb_register
                WHERE delete_time IS NULL
                  AND (platform_id IS NULL OR shudi_id IS NULL)
            """)
            empty_records = cursor.fetchall()
            total_empty = len(empty_records)
            self.results['total_empty'] = total_empty
            self.log_message.emit(f"找到 {total_empty} 条 platform_id 或 shudi_id 为空的记录")
            self.progress.emit(20)

            if total_empty == 0:
                self.log_message.emit("没有需要补充的记录")
                cursor.close()
                connection.close()
                self.progress.emit(100)
                self.finished.emit(self.results)
                return

            can_fill = 0
            no_order_id = 0
            order_not_found = 0
            order_also_empty = 0
            filled_platform = 0
            filled_shudi = 0

            for index, rec in enumerate(empty_records):
                row = {
                    'register_id': rec['id'],
                    'project_id': rec.get('project_id', ''),
                    'order_id': rec.get('order_id', ''),
                    'old_platform_id': rec.get('platform_id', ''),
                    'old_shudi_id': rec.get('shudi_id', ''),
                    'shopindex_id': rec.get('shopindex_id', ''),
                    'email': rec.get('email', ''),
                    'company_name': rec.get('company_name', ''),
                }

                # 检查是否有 order_id
                if not rec.get('order_id'):
                    no_order_id += 1
                    row['处理状态'] = '跳过(无order_id)'
                    self.results['completed_data'].append(row)
                    self.results['failed_records'].append(row)
                    progress = 20 + int((index + 1) / total_empty * 75)
                    self.progress.emit(progress)
                    continue

                # 从 ba_order 获取 platform_id 和 shudi_id
                cursor.execute("""
                    SELECT platform_id, shudi_id FROM ba_order WHERE id = %s
                """, (rec['order_id'],))
                order = cursor.fetchone()

                if not order:
                    order_not_found += 1
                    row['处理状态'] = f'失败(订单ID={rec["order_id"]}不存在)'
                    self.results['completed_data'].append(row)
                    self.results['failed_records'].append(row)
                    progress = 20 + int((index + 1) / total_empty * 75)
                    self.progress.emit(progress)
                    continue

                row['order_platform_id'] = order.get('platform_id', '')
                row['order_shudi_id'] = order.get('shudi_id', '')

                # 确定需要补充哪些字段
                need_platform = rec.get('platform_id') is None and order.get('platform_id') is not None
                need_shudi = rec.get('shudi_id') is None and order.get('shudi_id') is not None

                if not need_platform and not need_shudi:
                    # 订单表里对应字段也是空的
                    order_also_empty += 1
                    row['处理状态'] = '跳过(订单对应字段也为空)'
                    self.results['completed_data'].append(row)
                    self.results['failed_records'].append(row)
                    progress = 20 + int((index + 1) / total_empty * 75)
                    self.progress.emit(progress)
                    continue

                # 执行更新
                updates = []
                params = []
                if need_platform:
                    updates.append("platform_id = %s")
                    params.append(order['platform_id'])
                    filled_platform += 1
                if need_shudi:
                    updates.append("shudi_id = %s")
                    params.append(order['shudi_id'])
                    filled_shudi += 1

                if not self.preview_only:
                    update_sql = f"UPDATE ba_ptzcb_register SET {', '.join(updates)} WHERE id = %s"
                    params.append(rec['id'])
                    cursor.execute(update_sql, params)
                    connection.commit()

                can_fill += 1
                fill_desc = []
                if need_platform:
                    fill_desc.append(f"platform_id={order['platform_id']}")
                if need_shudi:
                    fill_desc.append(f"shudi_id={order['shudi_id']}")

                if self.preview_only:
                    row['处理状态'] = f'可补充({", ".join(fill_desc)})'
                else:
                    row['处理状态'] = f'成功({", ".join(fill_desc)})'

                self.results['completed_data'].append(row)

                if can_fill % 100 == 0:
                    self.log_message.emit(f"已处理 {can_fill} 条...")

                progress = 20 + int((index + 1) / total_empty * 75)
                self.progress.emit(progress)

            cursor.close()
            connection.close()

            self.results['can_fill'] = can_fill
            self.results['no_order_id'] = no_order_id
            self.results['order_not_found'] = order_not_found
            self.results['order_also_empty'] = order_also_empty
            self.results['filled_platform'] = filled_platform
            self.results['filled_shudi'] = filled_shudi

            self.progress.emit(100)
            self.log_message.emit(f"\n=== {mode}完成 ===")
            self.log_message.emit(f"空值记录总数: {total_empty}")
            self.log_message.emit(f"{'可补充' if self.preview_only else '已补充'}: {can_fill} (platform_id: {filled_platform}, shudi_id: {filled_shudi})")
            self.log_message.emit(f"无order_id: {no_order_id}")
            self.log_message.emit(f"订单不存在: {order_not_found}")
            self.log_message.emit(f"订单字段也为空: {order_also_empty}")

            self.finished.emit(self.results)

        except Exception as e:
            self.error.emit(str(e))


class RegisterFillFromOrderWidget(QWidget):
    """注册表补充平台ID和属地ID界面"""

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

        # 数据源选择
        ds_group = QGroupBox("选择数据库")
        ds_layout = QFormLayout()

        self.datasource_combo = QComboBox()
        self.datasource_combo.currentIndexChanged.connect(self.on_datasource_changed)
        ds_layout.addRow("数据源:", self.datasource_combo)

        self.test_conn_btn = QPushButton("测试连接")
        self.test_conn_btn.setEnabled(False)
        self.test_conn_btn.clicked.connect(self.test_connection)
        ds_layout.addRow("", self.test_conn_btn)

        ds_group.setLayout(ds_layout)
        layout.addWidget(ds_group)

        # 功能说明
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        info_label = QLabel("""
<b>补充注册表的平台ID和属地ID:</b><br><br>
<b>处理逻辑:</b><br>
1. 查找 ba_ptzcb_register 中 platform_id 或 shudi_id 为空的记录<br>
2. 通过 order_id 关联 ba_order 表<br>
3. 将 ba_order 中的 platform_id 和 shudi_id 补充到 ba_ptzcb_register<br><br>
<b>建议先执行「预检查」确认情况后再执行「开始补充」</b>
        """)
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)

        # 操作按钮
        btn_layout = QHBoxLayout()

        self.precheck_btn = QPushButton("预检查")
        self.precheck_btn.setEnabled(False)
        self.precheck_btn.clicked.connect(self.start_precheck)
        btn_layout.addWidget(self.precheck_btn)

        self.start_btn = QPushButton("开始补充")
        self.start_btn.setEnabled(False)
        self.start_btn.clicked.connect(self.start_fill)
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

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # 结果标签
        self.result_label = QLabel("")
        self.result_label.setAlignment(Qt.AlignCenter)
        font = QFont()
        font.setPointSize(12)
        self.result_label.setFont(font)
        layout.addWidget(self.result_label)

        # 日志
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
        self.datasource_combo.clear()
        self.datasource_combo.addItem("请选择数据源", None)
        datasources = self.db_manager.get_all_datasources()
        for ds in datasources:
            self.datasource_combo.addItem(f"{ds.name} ({ds.host}:{ds.port})", ds)

    def on_datasource_changed(self):
        ds_ok = self.datasource_combo.currentData() is not None
        self.test_conn_btn.setEnabled(ds_ok)
        self.precheck_btn.setEnabled(ds_ok)
        self.start_btn.setEnabled(ds_ok)

    def test_connection(self):
        ds = self.datasource_combo.currentData()
        if not ds:
            return
        ok, msg = self.db_manager.test_connection(ds)
        if ok:
            QMessageBox.information(self, "连接成功", msg)
        else:
            QMessageBox.critical(self, "连接失败", msg)

    def start_precheck(self):
        ds = self.datasource_combo.currentData()
        if not ds:
            QMessageBox.warning(self, "警告", "请选择数据源！")
            return
        self._run_worker(ds, preview_only=True)

    def start_fill(self):
        ds = self.datasource_combo.currentData()
        if not ds:
            QMessageBox.warning(self, "警告", "请选择数据源！")
            return

        reply = QMessageBox.question(
            self, "确认补充",
            f"确定要补充 ba_ptzcb_register 的 platform_id 和 shudi_id 吗？\n\n"
            f"数据库: {ds.name}\n"
            f"数据来源: ba_order (通过 order_id 关联)\n\n"
            f"操作不可撤销，请确认！",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        self._run_worker(ds, preview_only=False)

    def _run_worker(self, ds, preview_only):
        self.precheck_btn.setEnabled(False)
        self.start_btn.setEnabled(False)
        self.export_completed_btn.setEnabled(False)
        self.export_failed_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在预检查..." if preview_only else "正在补充...")
        self.log_text.clear()
        self.completed_data = []
        self.failed_records = []

        self.worker = RegisterFillFromOrderWorker(ds, preview_only=preview_only)
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

        total = results['total_empty']
        can_fill = results['can_fill']
        fp = results['filled_platform']
        fs = results['filled_shudi']
        no_oid = results['no_order_id']
        not_found = results['order_not_found']
        also_empty = results['order_also_empty']

        text = (f"空值{total}条 | 可补充: {can_fill} (平台:{fp}, 属地:{fs}) | "
                f"无订单ID: {no_oid} | 订单不存在: {not_found} | 订单也为空: {also_empty}")

        if no_oid + not_found + also_empty > 0:
            self.result_label.setStyleSheet("color: orange;")
        else:
            self.result_label.setStyleSheet("color: green;")
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
        default_name = f"注册表补充记录_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "保存完整记录", default_name, "Excel文件 (*.xlsx)")
        if path:
            df = pd.DataFrame(self.completed_data)
            df.to_excel(path, index=False, engine='openpyxl')
            QMessageBox.information(self, "成功", f"已导出 {len(self.completed_data)} 条记录到:\n{path}")

    def export_failed(self):
        if not self.failed_records:
            QMessageBox.warning(self, "警告", "没有失败记录！")
            return
        default_name = f"注册表补充失败记录_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "保存失败记录", default_name, "Excel文件 (*.xlsx)")
        if path:
            df = pd.DataFrame(self.failed_records)
            df.to_excel(path, index=False, engine='openpyxl')
            QMessageBox.information(self, "成功", f"已导出 {len(self.failed_records)} 条失败记录到:\n{path}")
