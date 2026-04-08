"""
账号管理表重复数据检查
根据 shopindex_id, shudi_id, platform_id
查找 ba_account_mag 表中的重复数据，按组导出Excel
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


class AccountMagDuplicateWorker(QThread):
    """重复数据检查工作线程"""
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

            self.log_message.emit("正在查询重复数据 (shopindex_id, shudi_id, platform_id)...")

            sql = """
                SELECT 
                    a.id,
                    a.inventory_num,
                    a.shopindex_id,
                    a.shudi_id,
                    a.platform_id,
                    a.register_id,
                    a.account_status,
                    a.shop_account,
                    a.shop_url,
                    a.now_customer,
                    a.origin_customer,
                    a.register_department,
                    a.remark,
                    a.create_time
                FROM ba_account_mag a
                INNER JOIN (
                    SELECT shopindex_id, shudi_id, platform_id
                    FROM ba_account_mag
                    WHERE shopindex_id IS NOT NULL
                    GROUP BY shopindex_id, shudi_id, platform_id
                    HAVING COUNT(*) > 1
                ) dup ON a.shopindex_id = dup.shopindex_id
                    AND a.shudi_id = dup.shudi_id
                    AND a.platform_id = dup.platform_id
                ORDER BY a.shopindex_id, a.shudi_id, a.platform_id, a.id
            """

            cursor.execute(sql)
            rows = cursor.fetchall()
            cursor.close()
            connection.close()
            self.progress.emit(60)

            if not rows:
                self.log_message.emit("没有找到重复数据")
                self.progress.emit(100)
                self.finished.emit({'total_records': 0, 'total_groups': 0})
                return

            self.log_message.emit(f"查询到 {len(rows)} 条重复记录")

            df = pd.DataFrame(rows)

            # 添加分组编号
            df['分组编号'] = (
                df.groupby(['shopindex_id', 'shudi_id', 'platform_id']).ngroup() + 1
            )

            # 转换时间戳
            if 'create_time' in df.columns:
                df['create_time'] = df['create_time'].apply(
                    lambda x: datetime.fromtimestamp(x).strftime('%Y-%m-%d %H:%M:%S') if x else ''
                )

            # 转换账号状态
            status_map = {
                0: '待更换密码',
                1: '过审暂停',
                2: '出售中',
                3: '可售',
                4: '已售',
                5: '封号'
            }
            if 'account_status' in df.columns:
                df['account_status_text'] = df['account_status'].map(status_map).fillna('')

            # 重新排列列顺序
            cols = ['分组编号', 'id', 'inventory_num', 'shopindex_id', 'shudi_id', 
                    'platform_id', 'register_id', 'account_status', 'account_status_text',
                    'shop_account', 'shop_url', 'now_customer', 'origin_customer',
                    'register_department', 'remark', 'create_time']
            cols = [c for c in cols if c in df.columns]
            df = df[cols]

            total_groups = df['分组编号'].max()
            self.log_message.emit(f"共 {total_groups} 组重复数据")
            self.progress.emit(80)

            # 导出Excel
            self.log_message.emit(f"正在导出到 {self.output_file}...")
            df.to_excel(self.output_file, index=False, engine='openpyxl')
            self.progress.emit(100)

            self.log_message.emit(f"导出完成！共 {len(rows)} 条记录，{total_groups} 组")
            self.finished.emit({
                'total_records': len(rows),
                'total_groups': total_groups
            })

        except Exception as e:
            self.error.emit(str(e))


class AccountMagDuplicateCheckWidget(QWidget):
    """账号管理表重复数据检查界面"""

    def __init__(self):
        super().__init__()
        self.db_manager = DatabaseManager()
        self.output_file = None
        self.worker = None
        self.init_ui()
        self.load_datasources()

    def init_ui(self):
        layout = QVBoxLayout(self)

        # 数据源选择
        ds_group = QGroupBox("选择目标数据库")
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

        # 输出文件选择
        file_group = QGroupBox("输出文件")
        file_layout = QFormLayout()

        file_row = QHBoxLayout()
        self.file_label = QLabel("未选择文件")
        self.file_label.setStyleSheet("color: gray;")
        file_row.addWidget(self.file_label)
        self.select_file_btn = QPushButton("选择保存位置")
        self.select_file_btn.clicked.connect(self.select_output_file)
        file_row.addWidget(self.select_file_btn)
        file_layout.addRow("Excel文件:", file_row)

        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # 操作按钮
        btn_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始检查")
        self.start_btn.setEnabled(False)
        self.start_btn.clicked.connect(self.start_check)
        btn_layout.addWidget(self.start_btn)
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
        self.log_text.setMaximumHeight(200)
        log_layout.addWidget(self.log_text)

        clear_btn = QPushButton("清空日志")
        clear_btn.clicked.connect(self.log_text.clear)
        log_layout.addWidget(clear_btn)

        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        layout.addStretch()

    def load_datasources(self):
        self.datasource_combo.clear()
        datasources = self.db_manager.get_all_datasources()
        for ds in datasources:
            self.datasource_combo.addItem(f"{ds.name} ({ds.host}:{ds.port})", ds)
        self.update_start_state()

    def on_datasource_changed(self):
        self.test_conn_btn.setEnabled(self.datasource_combo.currentData() is not None)
        self.update_start_state()

    def test_connection(self):
        ds = self.datasource_combo.currentData()
        if not ds:
            return
        ok, msg = self.db_manager.test_connection(ds)
        if ok:
            QMessageBox.information(self, "连接成功", msg)
        else:
            QMessageBox.critical(self, "连接失败", msg)

    def select_output_file(self):
        default_name = f"ba_account_mag重复数据_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(
            self, "选择保存位置", default_name, "Excel文件 (*.xlsx)"
        )
        if path:
            self.output_file = path
            self.file_label.setText(os.path.basename(path))
            self.file_label.setStyleSheet("color: green;")
            self.update_start_state()

    def update_start_state(self):
        ds_ok = self.datasource_combo.currentData() is not None
        file_ok = self.output_file is not None
        self.start_btn.setEnabled(ds_ok and file_ok)

    def start_check(self):
        ds = self.datasource_combo.currentData()
        if not ds or not self.output_file:
            QMessageBox.warning(self, "警告", "请选择数据源和输出文件！")
            return

        reply = QMessageBox.question(
            self, "确认",
            f"确定要检查 ba_account_mag 表中的重复数据吗？\n\n"
            f"数据库: {ds.name}\n"
            f"检查字段: shopindex_id, shudi_id, platform_id\n"
            f"输出文件: {os.path.basename(self.output_file)}",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        self.start_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在检查...")
        self.log_text.clear()

        self.worker = AccountMagDuplicateWorker(ds, self.output_file)
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log_message.connect(self.append_log)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def append_log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{ts}] {msg}")

    def on_finished(self, results):
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)

        total = results['total_records']
        groups = results['total_groups']

        if total == 0:
            text = "未发现重复数据"
            self.result_label.setStyleSheet("color: green;")
        else:
            text = f"发现 {groups} 组重复数据，共 {total} 条记录"
            self.result_label.setStyleSheet("color: orange;")

        self.result_label.setText(text)
        self.append_log(f"\n=== 检查完成 ===\n{text}")

        if total > 0:
            QMessageBox.information(
                self, "完成",
                f"检查完成！\n\n{text}\n\n结果已导出到:\n{self.output_file}"
            )

    def on_error(self, error_msg):
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.result_label.setText("检查失败")
        self.result_label.setStyleSheet("color: red;")
        self.append_log(f"错误: {error_msg}")
        QMessageBox.critical(self, "错误", f"检查失败:\n{error_msg}")
