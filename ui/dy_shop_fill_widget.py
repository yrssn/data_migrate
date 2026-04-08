"""
抖店库存编号填充
根据Excel中的ID列查找ea_dy_shop表记录，将Excel中的值更新到stock_number字段
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


class DyShopPreCheckWorker(QThread):
    """预检查工作线程：检查哪些ID存在，哪些不存在"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, datasource, excel_file, id_column, value_column):
        super().__init__()
        self.datasource = datasource
        self.excel_file = excel_file
        self.id_column = id_column
        self.value_column = value_column

    def run(self):
        try:
            self.log_message.emit("正在读取Excel文件...")
            self.progress.emit(5)

            df = pd.read_excel(self.excel_file, engine='openpyxl')
            total_rows = len(df)
            self.log_message.emit(f"Excel共 {total_rows} 行")

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

            can_fill = 0
            id_not_found = 0
            id_empty = 0
            value_empty = 0
            id_invalid = 0
            not_found_ids = []

            for index, row in df.iterrows():
                record_id = row[self.id_column]
                stock_number = row[self.value_column]

                if pd.isna(record_id) or str(record_id).strip() == '':
                    id_empty += 1
                    continue

                if pd.isna(stock_number) or str(stock_number).strip() == '':
                    value_empty += 1
                    continue

                try:
                    rid = int(float(record_id))
                except:
                    id_invalid += 1
                    continue

                cursor.execute("SELECT id, stock_number FROM ea_dy_shop WHERE id = %s", (rid,))
                existing = cursor.fetchone()

                if existing:
                    can_fill += 1
                else:
                    id_not_found += 1
                    not_found_ids.append(str(rid))

                progress = 10 + int((index + 1) / total_rows * 85)
                self.progress.emit(progress)

            cursor.close()
            connection.close()
            self.progress.emit(100)

            self.log_message.emit(f"\n=== 预检查结果 ===")
            self.log_message.emit(f"总行数: {total_rows}")
            self.log_message.emit(f"可填充: {can_fill} (ID存在且值非空)")
            self.log_message.emit(f"ID为空: {id_empty}")
            self.log_message.emit(f"值为空: {value_empty}")
            self.log_message.emit(f"ID无效: {id_invalid}")
            self.log_message.emit(f"ID不存在: {id_not_found}")
            if not_found_ids:
                show_ids = not_found_ids[:20]
                self.log_message.emit(f"不存在的ID: {', '.join(show_ids)}{'...' if len(not_found_ids) > 20 else ''}")

            self.finished.emit({
                'total': total_rows,
                'can_fill': can_fill,
                'id_empty': id_empty,
                'value_empty': value_empty,
                'id_invalid': id_invalid,
                'id_not_found': id_not_found
            })

        except Exception as e:
            self.error.emit(str(e))


class DyShopFillWorker(QThread):
    """填充工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, datasource, excel_file, id_column, value_column):
        super().__init__()
        self.datasource = datasource
        self.excel_file = excel_file
        self.id_column = id_column
        self.value_column = value_column
        self.results = {
            'success_count': 0,
            'fail_count': 0,
            'skip_count': 0,
            'completed_data': [],
            'failed_records': []
        }

    def run(self):
        try:
            self.log_message.emit("正在读取Excel文件...")
            self.progress.emit(5)

            df = pd.read_excel(self.excel_file, engine='openpyxl')
            total_rows = len(df)
            self.log_message.emit(f"读取到 {total_rows} 条记录")

            if self.id_column not in df.columns:
                self.error.emit(f"Excel中未找到ID列: {self.id_column}")
                return
            if self.value_column not in df.columns:
                self.error.emit(f"Excel中未找到值列: {self.value_column}")
                return

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

            success_count = 0
            fail_count = 0
            skip_count = 0

            for index, row in df.iterrows():
                try:
                    record_id = row[self.id_column]
                    stock_number = row[self.value_column]

                    # 跳过空ID
                    if pd.isna(record_id) or str(record_id).strip() == '':
                        skip_count += 1
                        completed_row = row.to_dict()
                        completed_row['处理状态'] = '跳过(ID为空)'
                        completed_row['处理时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        self.results['completed_data'].append(completed_row)
                        continue

                    # 跳过空值
                    if pd.isna(stock_number) or str(stock_number).strip() == '':
                        skip_count += 1
                        completed_row = row.to_dict()
                        completed_row['处理状态'] = '跳过(值为空)'
                        completed_row['处理时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        self.results['completed_data'].append(completed_row)
                        continue

                    # 转换ID为整数
                    try:
                        record_id = int(float(record_id))
                    except:
                        skip_count += 1
                        completed_row = row.to_dict()
                        completed_row['处理状态'] = f'跳过(ID无效: {record_id})'
                        completed_row['处理时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        self.results['completed_data'].append(completed_row)
                        continue

                    stock_number = str(stock_number).strip()

                    # 检查记录是否存在
                    cursor.execute("SELECT id FROM ea_dy_shop WHERE id = %s", (record_id,))
                    existing = cursor.fetchone()

                    if not existing:
                        fail_count += 1
                        completed_row = row.to_dict()
                        completed_row['处理状态'] = f'失败(ID不存在: {record_id})'
                        completed_row['处理时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        self.results['completed_data'].append(completed_row)
                        self.results['failed_records'].append(completed_row)
                        continue

                    # 更新stock_number
                    cursor.execute(
                        "UPDATE ea_dy_shop SET stock_number = %s WHERE id = %s",
                        (stock_number, record_id)
                    )
                    connection.commit()

                    success_count += 1
                    completed_row = row.to_dict()
                    completed_row['处理状态'] = '成功'
                    completed_row['处理时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    self.results['completed_data'].append(completed_row)

                    if success_count % 50 == 0:
                        self.log_message.emit(f"已处理 {success_count} 条...")

                except Exception as e:
                    fail_count += 1
                    completed_row = row.to_dict()
                    completed_row['处理状态'] = f'失败({str(e)})'
                    completed_row['处理时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    self.results['completed_data'].append(completed_row)
                    self.results['failed_records'].append(completed_row)

                progress = 10 + int((index + 1) / total_rows * 85)
                self.progress.emit(progress)

            cursor.close()
            connection.close()

            self.results['success_count'] = success_count
            self.results['fail_count'] = fail_count
            self.results['skip_count'] = skip_count

            self.progress.emit(100)
            self.log_message.emit(f"\n=== 处理完成 ===")
            self.log_message.emit(f"成功: {success_count}")
            self.log_message.emit(f"失败: {fail_count}")
            self.log_message.emit(f"跳过: {skip_count}")

            self.finished.emit(self.results)

        except Exception as e:
            self.error.emit(str(e))


class DyShopFillWidget(QWidget):
    """抖店库存编号填充界面"""

    def __init__(self, db_manager=None):
        super().__init__()
        self.db_manager = db_manager or DatabaseManager()
        self.excel_file = None
        self.excel_columns = []
        self.worker = None
        self.completed_data = []
        self.failed_records = []
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

        # Excel文件选择
        file_group = QGroupBox("Excel文件")
        file_layout = QFormLayout()

        file_row = QHBoxLayout()
        self.file_label = QLabel("未选择文件")
        self.file_label.setStyleSheet("color: gray;")
        file_row.addWidget(self.file_label)
        self.select_file_btn = QPushButton("选择文件")
        self.select_file_btn.clicked.connect(self.select_excel_file)
        file_row.addWidget(self.select_file_btn)
        file_layout.addRow("Excel文件:", file_row)

        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # 列选择
        col_group = QGroupBox("列选择")
        col_layout = QFormLayout()

        self.id_column_combo = QComboBox()
        self.id_column_combo.setEnabled(False)
        self.id_column_combo.currentTextChanged.connect(self.update_start_state)
        col_layout.addRow("ID列 (ea_dy_shop.id):", self.id_column_combo)

        self.value_column_combo = QComboBox()
        self.value_column_combo.setEnabled(False)
        self.value_column_combo.currentTextChanged.connect(self.update_start_state)
        col_layout.addRow("值列 (stock_number):", self.value_column_combo)

        col_group.setLayout(col_layout)
        layout.addWidget(col_group)

        # 操作按钮
        btn_layout = QHBoxLayout()

        self.precheck_btn = QPushButton("预检查")
        self.precheck_btn.setEnabled(False)
        self.precheck_btn.clicked.connect(self.start_precheck)
        btn_layout.addWidget(self.precheck_btn)

        self.start_btn = QPushButton("开始填充")
        self.start_btn.setEnabled(False)
        self.start_btn.clicked.connect(self.start_process)
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

    def select_excel_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", "", "Excel文件 (*.xlsx *.xls)")
        if not path:
            return
        try:
            df = pd.read_excel(path, engine='openpyxl', nrows=0)
            self.excel_columns = list(df.columns)
            self.excel_file = path
            self.file_label.setText(os.path.basename(path))
            self.file_label.setStyleSheet("color: green;")

            for combo in [self.id_column_combo, self.value_column_combo]:
                combo.clear()
                combo.addItem("请选择列")
                combo.addItems([str(c) for c in self.excel_columns])
                combo.setEnabled(True)

            self.append_log(f"读取Excel表头，共 {len(self.excel_columns)} 列")
            self.update_start_state()

        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取Excel失败:\n{str(e)}")

    def update_start_state(self):
        ds_ok = self.datasource_combo.currentData() is not None
        file_ok = self.excel_file is not None
        id_ok = self.id_column_combo.currentText() != "请选择列" and self.id_column_combo.currentText() != ""
        value_ok = self.value_column_combo.currentText() != "请选择列" and self.value_column_combo.currentText() != ""
        ready = ds_ok and file_ok and id_ok and value_ok
        self.start_btn.setEnabled(ready)
        self.precheck_btn.setEnabled(ready)

    def start_precheck(self):
        ds = self.datasource_combo.currentData()
        if not ds or not self.excel_file:
            QMessageBox.warning(self, "警告", "请选择数据源和Excel文件！")
            return

        id_col = self.id_column_combo.currentText()
        value_col = self.value_column_combo.currentText()
        if id_col == "请选择列" or value_col == "请选择列":
            QMessageBox.warning(self, "警告", "请选择ID列和值列！")
            return

        self.precheck_btn.setEnabled(False)
        self.start_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在预检查...")
        self.log_text.clear()

        self.precheck_worker = DyShopPreCheckWorker(ds, self.excel_file, id_col, value_col)
        self.precheck_worker.progress.connect(self.progress_bar.setValue)
        self.precheck_worker.log_message.connect(self.append_log)
        self.precheck_worker.finished.connect(self.on_precheck_finished)
        self.precheck_worker.error.connect(self.on_precheck_error)
        self.precheck_worker.start()

    def on_precheck_finished(self, results):
        self.precheck_btn.setEnabled(True)
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)

        total = results['total']
        can_fill = results['can_fill']
        id_not_found = results['id_not_found']
        id_empty = results['id_empty']
        value_empty = results['value_empty']
        id_invalid = results['id_invalid']

        text = f"总{total}行 | 可填充: {can_fill} | ID不存在: {id_not_found} | ID空: {id_empty} | 值空: {value_empty} | ID无效: {id_invalid}"
        if id_not_found > 0:
            self.result_label.setStyleSheet("color: orange;")
        else:
            self.result_label.setStyleSheet("color: green;")
        self.result_label.setText(text)

    def on_precheck_error(self, error_msg):
        self.precheck_btn.setEnabled(True)
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.result_label.setText("预检查失败")
        self.result_label.setStyleSheet("color: red;")
        self.append_log(f"预检查错误: {error_msg}")
        QMessageBox.critical(self, "错误", f"预检查失败:\n{error_msg}")

    def start_process(self):
        ds = self.datasource_combo.currentData()
        if not ds or not self.excel_file:
            QMessageBox.warning(self, "警告", "请选择数据源和Excel文件！")
            return

        id_col = self.id_column_combo.currentText()
        value_col = self.value_column_combo.currentText()

        if id_col == "请选择列" or value_col == "请选择列":
            QMessageBox.warning(self, "警告", "请选择ID列和值列！")
            return

        reply = QMessageBox.question(
            self, "确认填充",
            f"确定要填充 ea_dy_shop.stock_number 吗？\n\n"
            f"数据库: {ds.name}\n"
            f"Excel: {os.path.basename(self.excel_file)}\n"
            f"ID列: {id_col}\n"
            f"值列: {value_col}\n\n"
            f"操作将更新数据库记录，请确认！",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        self.start_btn.setEnabled(False)
        self.export_completed_btn.setEnabled(False)
        self.export_failed_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在处理...")
        self.log_text.clear()
        self.completed_data = []
        self.failed_records = []

        self.worker = DyShopFillWorker(ds, self.excel_file, id_col, value_col)
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

        self.completed_data = results.get('completed_data', [])
        self.failed_records = results.get('failed_records', [])

        success = results['success_count']
        fail = results['fail_count']
        skip = results['skip_count']

        text = f"成功: {success}  失败: {fail}  跳过: {skip}"
        if fail > 0:
            self.result_label.setStyleSheet("color: orange;")
        else:
            self.result_label.setStyleSheet("color: green;")
        self.result_label.setText(text)

        self.export_completed_btn.setEnabled(len(self.completed_data) > 0)
        self.export_failed_btn.setEnabled(len(self.failed_records) > 0)

        QMessageBox.information(self, "完成", f"填充完成！\n\n{text}")

    def on_error(self, error_msg):
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
        default_name = f"抖店填充完整记录_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "保存完整记录", default_name, "Excel文件 (*.xlsx)")
        if path:
            df = pd.DataFrame(self.completed_data)
            df.to_excel(path, index=False, engine='openpyxl')
            QMessageBox.information(self, "成功", f"已导出 {len(self.completed_data)} 条记录到:\n{path}")

    def export_failed(self):
        if not self.failed_records:
            QMessageBox.warning(self, "警告", "没有失败记录！")
            return
        default_name = f"抖店填充失败记录_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "保存失败记录", default_name, "Excel文件 (*.xlsx)")
        if path:
            df = pd.DataFrame(self.failed_records)
            df.to_excel(path, index=False, engine='openpyxl')
            QMessageBox.information(self, "成功", f"已导出 {len(self.failed_records)} 条失败记录到:\n{path}")
