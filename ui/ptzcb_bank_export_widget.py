"""
注册表银行数据导出
导出 ba_ptzcb_register 表中 status=6 且 platform_id 对应
ba_platform.platform_type='opt1'（银行）的数据，
关联 ba_platform.platform 作为银行名称、
customer_id 关联 ba_rlb_customer.legal_name 法人姓名，导出Excel
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


class BankExportWorker(QThread):
    """银行数据导出工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, datasource, output_file):
        super().__init__()
        self.datasource = datasource
        self.output_file = output_file

    def run(self):
        try:
            self.log_message.emit("正在连接数据库...")
            self.progress.emit(10)

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
            self.progress.emit(20)

            self.log_message.emit("正在查询数据 (status=6 且平台类型为 opt1=银行)...")

            sql = """
                SELECT
                    r.*,
                    p.platform AS 银行名称,
                    c.legal_name AS 法人姓名
                FROM ba_ptzcb_register r
                INNER JOIN ba_platform p ON r.platform_id = p.id
                LEFT JOIN ba_rlb_customer c ON r.customer_id = c.id
                WHERE r.status = 6
                  AND p.platform_type = 'opt1'
                  AND (r.delete_time IS NULL OR r.delete_time = 0)
                ORDER BY r.id
            """

            cursor.execute(sql)
            rows = cursor.fetchall()
            cursor.close()
            connection.close()
            self.progress.emit(60)

            if not rows:
                self.log_message.emit("没有找到符合条件的数据")
                self.progress.emit(100)
                self.finished.emit({'total_records': 0})
                return

            self.log_message.emit(f"查询到 {len(rows)} 条记录")

            df = pd.DataFrame(rows)
            self.progress.emit(80)

            df.to_excel(self.output_file, index=False, engine='openpyxl')
            self.log_message.emit(f"已导出到: {self.output_file}")
            self.progress.emit(100)

            self.finished.emit({'total_records': len(df)})

        except Exception as e:
            self.error.emit(f"处理出错: {str(e)}")


class PtzcbBankExportWidget(QWidget):
    """注册表银行数据导出组件"""

    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.worker = None
        self.output_file = ""
        self.init_ui()
        self.load_datasources()

    def init_ui(self):
        layout = QVBoxLayout()

        # 说明
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        info_label = QLabel(
            "导出 ba_ptzcb_register 表中 status=6 且 platform_id 对应 "
            "ba_platform.platform_type='opt1'（银行）的数据。\n"
            "同时关联查出银行名称（ba_platform.platform）和 "
            "customer_id 关联的法人姓名（ba_rlb_customer.legal_name），导出到Excel。"
        )
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)

        # 数据源
        ds_group = QGroupBox("选择数据库")
        ds_layout = QFormLayout()
        self.datasource_combo = QComboBox()
        self.datasource_combo.currentTextChanged.connect(self._update_state)
        ds_layout.addRow("数据源:", self.datasource_combo)
        self.test_conn_btn = QPushButton("测试连接")
        self.test_conn_btn.clicked.connect(self.test_connection)
        self.test_conn_btn.setEnabled(False)
        ds_layout.addRow("", self.test_conn_btn)
        ds_group.setLayout(ds_layout)
        layout.addWidget(ds_group)

        # 输出文件
        file_group = QGroupBox("输出文件")
        file_layout = QHBoxLayout()
        self.output_label = QLabel("未选择")
        file_layout.addWidget(self.output_label, 1)
        self.select_output_btn = QPushButton("选择输出文件")
        self.select_output_btn.clicked.connect(self.select_output)
        file_layout.addWidget(self.select_output_btn)
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # 按钮 + 进度
        btn_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始导出")
        self.start_btn.clicked.connect(self.start_export)
        self.start_btn.setEnabled(False)
        btn_layout.addWidget(self.start_btn)

        self.stop_btn = QPushButton("停止")
        self.stop_btn.clicked.connect(self.stop_export)
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

        # 结果
        self.result_label = QLabel("")
        self.result_label.setAlignment(Qt.AlignCenter)
        font = QFont()
        font.setBold(True)
        self.result_label.setFont(font)
        layout.addWidget(self.result_label)

        self.setLayout(layout)

    def load_datasources(self):
        self.datasource_combo.clear()
        self.datasource_combo.addItem("请选择数据源", None)
        for ds in self.db_manager.get_all_datasources():
            self.datasource_combo.addItem(f"{ds.name} ({ds.host}:{ds.port})", ds)

    def _update_state(self):
        ds_ok = self.datasource_combo.currentData() is not None
        self.test_conn_btn.setEnabled(ds_ok)
        self.start_btn.setEnabled(ds_ok and bool(self.output_file))

    def test_connection(self):
        ds = self.datasource_combo.currentData()
        if not ds:
            return
        ok, msg = self.db_manager.test_connection(ds)
        if ok:
            QMessageBox.information(self, "连接成功", msg)
        else:
            QMessageBox.critical(self, "连接失败", msg)

    def select_output(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "选择输出文件",
            f"ptzcb_bank_status6_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel文件 (*.xlsx)"
        )
        if path:
            if not path.endswith('.xlsx'):
                path += '.xlsx'
            self.output_file = path
            self.output_label.setText(path.split('/')[-1].split('\\')[-1])
            self._update_state()

    def start_export(self):
        ds = self.datasource_combo.currentData()
        if not ds or not self.output_file:
            return

        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.select_output_btn.setEnabled(False)
        self.datasource_combo.setEnabled(False)
        self.log_text.clear()
        self.progress_bar.setValue(0)
        self.result_label.setText("正在导出...")

        self.worker = BankExportWorker(ds, self.output_file)
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log_message.connect(self._log)
        self.worker.finished.connect(self._on_finished)
        self.worker.error.connect(self._on_error)
        self.worker.start()

    def stop_export(self):
        if self.worker and self.worker.isRunning():
            self.worker.terminate()
            self.worker.wait()
            self._log("已停止")
            self._reset_ui()

    def _log(self, msg):
        ts = datetime.now().strftime('%H:%M:%S')
        self.log_text.append(f"[{ts}] {msg}")

    def _on_finished(self, results):
        total = results['total_records']
        if total == 0:
            self.result_label.setText("没有找到符合条件的数据")
            self.result_label.setStyleSheet("color: orange;")
        else:
            self.result_label.setText(f"共 {total} 条记录，已导出")
            self.result_label.setStyleSheet("color: green;")
        self._log(f"导出完成！共 {total} 条")
        self._reset_ui()

    def _on_error(self, msg):
        self._log(f"错误: {msg}")
        self.result_label.setText("导出失败")
        self.result_label.setStyleSheet("color: red;")
        QMessageBox.critical(self, "错误", msg)
        self._reset_ui()

    def _reset_ui(self):
        self.stop_btn.setEnabled(False)
        self.select_output_btn.setEnabled(True)
        self.datasource_combo.setEnabled(True)
        self._update_state()
