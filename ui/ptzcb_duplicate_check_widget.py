"""
平台注册部重复数据检查
根据 shopindex_id, shudi_id, platform_id, customer_id, information_id, information_part2_id
查找 ba_ptzcb_register 表中的重复数据，按组导出Excel
排除 platform_id = 108
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


class DuplicateCheckWorker(QThread):
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

            self.log_message.emit("正在查询重复数据...")

            sql = """
                SELECT 
                    r.id,
                    r.project_id,
                    r.shopindex_id,
                    r.shudi_id,
                    r.platform_id,
                    r.customer_id,
                    r.information_id,
                    r.information_part2_id,
                    r.company_name,
                    r.country,
                    r.status,
                    r.create_time
                FROM ba_ptzcb_register r
                INNER JOIN (
                    SELECT shopindex_id, shudi_id, platform_id, customer_id, 
                           information_id, information_part2_id
                    FROM ba_ptzcb_register
                    WHERE platform_id != 108
                      AND delete_time IS NULL
                    GROUP BY shopindex_id, shudi_id, platform_id, customer_id, 
                             information_id, information_part2_id
                    HAVING COUNT(*) > 1
                ) dup ON r.shopindex_id <=> dup.shopindex_id
                    AND r.shudi_id <=> dup.shudi_id
                    AND r.platform_id <=> dup.platform_id
                    AND r.customer_id <=> dup.customer_id
                    AND r.information_id <=> dup.information_id
                    AND r.information_part2_id <=> dup.information_part2_id
                WHERE r.platform_id != 108
                  AND r.delete_time IS NULL
                ORDER BY r.shopindex_id, r.shudi_id, r.platform_id, r.customer_id, 
                         r.information_id, r.information_part2_id, r.id
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
            group_cols = ['shopindex_id', 'shudi_id', 'platform_id',
                          'customer_id', 'information_id', 'information_part2_id']

            # 按组编号，同组放一起
            group_id = 0
            prev_key = None
            group_ids = []
            for _, row in df.iterrows():
                key = tuple(row[c] for c in group_cols)
                if key != prev_key:
                    group_id += 1
                    prev_key = key
                group_ids.append(group_id)

            df.insert(0, '重复组号', group_ids)
            self.progress.emit(80)

            total_groups = df['重复组号'].nunique()
            self.log_message.emit(f"共 {total_groups} 组重复，涉及 {len(df)} 条记录")

            df.to_excel(self.output_file, index=False, engine='openpyxl')
            self.log_message.emit(f"已导出到: {self.output_file}")
            self.progress.emit(100)

            self.finished.emit({
                'total_records': len(df),
                'total_groups': total_groups
            })

        except Exception as e:
            self.error.emit(f"处理出错: {str(e)}")


class PtzcbDuplicateCheckWidget(QWidget):
    """平台注册部重复数据检查组件"""

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
            "根据 shopindex_id、shudi_id、platform_id、customer_id、"
            "information_id、information_part2_id 六个字段，\n"
            "查找 ba_ptzcb_register 表中的重复数据（排除 platform_id=108），"
            "按组导出到Excel。\n"
            "导出字段：重复组号、id、project_id、以及上述6个字段、company_name、country、status、create_time"
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
        self.start_btn = QPushButton("开始检查")
        self.start_btn.clicked.connect(self.start_check)
        self.start_btn.setEnabled(False)
        btn_layout.addWidget(self.start_btn)

        self.stop_btn = QPushButton("停止")
        self.stop_btn.clicked.connect(self.stop_check)
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
            f"ptzcb_duplicates_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel文件 (*.xlsx)"
        )
        if path:
            if not path.endswith('.xlsx'):
                path += '.xlsx'
            self.output_file = path
            self.output_label.setText(path.split('/')[-1].split('\\')[-1])
            self._update_state()

    def start_check(self):
        ds = self.datasource_combo.currentData()
        if not ds or not self.output_file:
            return

        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.select_output_btn.setEnabled(False)
        self.datasource_combo.setEnabled(False)
        self.log_text.clear()
        self.progress_bar.setValue(0)
        self.result_label.setText("正在检查...")

        self.worker = DuplicateCheckWorker(ds, self.output_file)
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log_message.connect(self._log)
        self.worker.finished.connect(self._on_finished)
        self.worker.error.connect(self._on_error)
        self.worker.start()

    def stop_check(self):
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
        groups = results['total_groups']
        if total == 0:
            self.result_label.setText("没有找到重复数据")
            self.result_label.setStyleSheet("color: green;")
        else:
            self.result_label.setText(f"共 {groups} 组重复，涉及 {total} 条记录，已导出")
            self.result_label.setStyleSheet("color: orange;")
        self._log(f"检查完成！{groups} 组, {total} 条")
        self._reset_ui()

    def _on_error(self, msg):
        self._log(f"错误: {msg}")
        self.result_label.setText("检查失败")
        self.result_label.setStyleSheet("color: red;")
        QMessageBox.critical(self, "错误", msg)
        self._reset_ui()

    def _reset_ui(self):
        self.stop_btn.setEnabled(False)
        self.select_output_btn.setEnabled(True)
        self.datasource_combo.setEnabled(True)
        self._update_state()
