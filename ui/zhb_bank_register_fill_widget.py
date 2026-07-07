"""
银行表补充注册表ID
从Excel（可选sheet）读取 ba_zhb_bank 的ID列和 ptzcb_register_id 列，
把 ptzcb_register_id 补充到 ba_zhb_bank.ptzcb_register_id，
非数字的值自动跳过
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


def parse_int(value):
    """解析整数，非数字返回None"""
    if value is None or pd.isna(value):
        return None
    text = str(value).strip()
    if text == '' or text.lower() == 'nan':
        return None
    try:
        num = float(text)
    except ValueError:
        return None
    if num != int(num):
        return None
    return int(num)


class ZhbBankRegisterFillWorker(QThread):
    """银行表补充注册表ID工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, datasource, excel_file, sheet_name, bank_id_column, register_id_column):
        super().__init__()
        self.datasource = datasource
        self.excel_file = excel_file
        self.sheet_name = sheet_name
        self.bank_id_column = bank_id_column
        self.register_id_column = register_id_column
        self.results = {
            'total_rows': 0,
            'success_count': 0,
            'skipped_invalid': 0,
            'skipped_no_bank': 0,
            'skipped_already_has': 0,
            'failed_count': 0,
            'completed_data': []
        }

    def run(self):
        try:
            self.log_message.emit(f"开始读取Excel文件 (sheet: {self.sheet_name})...")
            self.progress.emit(5)

            df = pd.read_excel(self.excel_file, sheet_name=self.sheet_name, engine='openpyxl')
            self.log_message.emit(f"成功读取Excel文件，共 {len(df)} 行数据")
            self.results['total_rows'] = len(df)
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
            self.progress.emit(15)

            total = len(df)
            for index, row in df.iterrows():
                row_num = index + 2  # Excel行号(含表头)
                try:
                    result = self.process_row(cursor, row_num, row)
                    completed_row = row.to_dict()
                    completed_row['处理结果'] = result['status']
                    completed_row['说明'] = result['message']
                    self.results['completed_data'].append(completed_row)

                    progress = 15 + int((index + 1) / total * 80)
                    self.progress.emit(progress)

                    if (index + 1) % 50 == 0:
                        connection.commit()
                        self.log_message.emit(f"已处理 {index + 1}/{total} 行，提交事务")

                except Exception as e:
                    self.results['failed_count'] += 1
                    self.log_message.emit(f"第{row_num}行处理失败: {str(e)}")
                    failed_row = row.to_dict()
                    failed_row['处理结果'] = '失败'
                    failed_row['说明'] = str(e)
                    self.results['completed_data'].append(failed_row)

            connection.commit()
            cursor.close()
            connection.close()

            self.progress.emit(100)
            self.finished.emit(self.results)

        except Exception as e:
            self.error.emit(str(e))

    def process_row(self, cursor, row_num, row):
        bank_id = parse_int(row.get(self.bank_id_column))
        register_id = parse_int(row.get(self.register_id_column))

        if bank_id is None:
            self.results['skipped_invalid'] += 1
            return {'status': '跳过', 'message': f'银行表ID非数字或为空: {row.get(self.bank_id_column)}'}

        if register_id is None:
            self.results['skipped_invalid'] += 1
            return {'status': '跳过', 'message': f'注册表ID非数字或为空: {row.get(self.register_id_column)}'}

        cursor.execute("""
            SELECT id, ptzcb_register_id FROM ba_zhb_bank
            WHERE id = %s AND (delete_time IS NULL OR delete_time = 0)
        """, (bank_id,))
        bank_record = cursor.fetchone()

        if not bank_record:
            self.results['skipped_no_bank'] += 1
            return {'status': '跳过', 'message': f'未找到银行记录 id={bank_id}'}

        if bank_record['ptzcb_register_id']:
            self.results['skipped_already_has'] += 1
            return {
                'status': '跳过',
                'message': f'银行记录id={bank_id}已有ptzcb_register_id={bank_record["ptzcb_register_id"]}'
            }

        current_time = int(datetime.now().timestamp())
        cursor.execute("""
            UPDATE ba_zhb_bank
            SET ptzcb_register_id = %s, update_time = %s
            WHERE id = %s
        """, (register_id, current_time, bank_id))

        self.results['success_count'] += 1
        self.log_message.emit(f"第{row_num}行: 银行id={bank_id} 更新ptzcb_register_id={register_id}")
        return {'status': '成功', 'message': f'更新ptzcb_register_id={register_id}'}


class ZhbBankRegisterFillWidget(QWidget):
    """银行表补充注册表ID组件"""

    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.excel_file = None
        self.completed_data = []
        self.worker = None
        self.init_ui()
        self.load_datasources()

    def init_ui(self):
        layout = QVBoxLayout()

        # 功能说明
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        info_label = QLabel(
            "从Excel读取 ba_zhb_bank 的ID和对应的注册表ID（ba_ptzcb_register.id），\n"
            "把注册表ID补充到 ba_zhb_bank.ptzcb_register_id 字段。\n"
            "支持选择sheet和对应的列；非数字的值（如 #N/A、空值）自动跳过；\n"
            "已有 ptzcb_register_id 的记录不会覆盖。"
        )
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)

        # 数据源选择
        ds_group = QGroupBox("选择目标数据库")
        ds_layout = QFormLayout()
        self.datasource_combo = QComboBox()
        self.datasource_combo.currentTextChanged.connect(self.update_start_btn_state)
        ds_layout.addRow("数据源:", self.datasource_combo)
        self.test_connection_btn = QPushButton("测试连接")
        self.test_connection_btn.clicked.connect(self.test_connection)
        ds_layout.addRow("", self.test_connection_btn)
        ds_group.setLayout(ds_layout)
        layout.addWidget(ds_group)

        # Excel文件与sheet选择
        file_group = QGroupBox("Excel文件选择")
        file_layout = QFormLayout()

        file_row = QHBoxLayout()
        self.file_label = QLabel("未选择文件")
        self.file_label.setStyleSheet("color: gray;")
        file_row.addWidget(self.file_label)
        self.select_file_btn = QPushButton("选择Excel文件")
        self.select_file_btn.clicked.connect(self.select_excel_file)
        file_row.addWidget(self.select_file_btn)
        file_layout.addRow("Excel文件:", file_row)

        self.sheet_combo = QComboBox()
        self.sheet_combo.setEnabled(False)
        self.sheet_combo.currentTextChanged.connect(self.on_sheet_changed)
        file_layout.addRow("选择Sheet:", self.sheet_combo)

        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # 列选择
        column_group = QGroupBox("列选择")
        column_layout = QFormLayout()

        self.bank_id_column_combo = QComboBox()
        self.bank_id_column_combo.setEnabled(False)
        self.bank_id_column_combo.currentTextChanged.connect(self.update_start_btn_state)
        column_layout.addRow("银行表ID列 (ba_zhb_bank.id):", self.bank_id_column_combo)

        self.register_id_column_combo = QComboBox()
        self.register_id_column_combo.setEnabled(False)
        self.register_id_column_combo.currentTextChanged.connect(self.update_start_btn_state)
        column_layout.addRow("注册表ID列 (ptzcb_register_id):", self.register_id_column_combo)

        column_group.setLayout(column_layout)
        layout.addWidget(column_group)

        # 操作按钮
        btn_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始补充")
        self.start_btn.clicked.connect(self.start_process)
        self.start_btn.setEnabled(False)
        btn_layout.addWidget(self.start_btn)

        self.export_btn = QPushButton("导出处理结果")
        self.export_btn.clicked.connect(self.export_completed)
        self.export_btn.setEnabled(False)
        btn_layout.addWidget(self.export_btn)

        btn_layout.addStretch()

        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimumWidth(200)
        btn_layout.addWidget(self.progress_bar)
        layout.addLayout(btn_layout)

        # 日志
        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(180)
        self.log_text.setReadOnly(True)
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

    def test_connection(self):
        ds = self.datasource_combo.currentData()
        if not ds:
            QMessageBox.warning(self, "警告", "请选择数据源！")
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
            excel = pd.ExcelFile(path, engine='openpyxl')
            self.excel_file = path
            self.file_label.setText(os.path.basename(path))
            self.file_label.setStyleSheet("color: green;")

            self.sheet_combo.blockSignals(True)
            self.sheet_combo.clear()
            self.sheet_combo.addItems(excel.sheet_names)
            self.sheet_combo.setEnabled(True)
            self.sheet_combo.blockSignals(False)

            self.log_text.append(f"已加载Excel文件，共 {len(excel.sheet_names)} 个sheet")
            self.on_sheet_changed()

        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取Excel失败: {str(e)}")
            self.excel_file = None
            self.file_label.setText("未选择文件")
            self.file_label.setStyleSheet("color: gray;")

    def on_sheet_changed(self):
        if not self.excel_file or not self.sheet_combo.currentText():
            return
        try:
            df = pd.read_excel(self.excel_file, sheet_name=self.sheet_combo.currentText(),
                               engine='openpyxl', nrows=0)
            columns = [str(c) for c in df.columns]

            for combo in [self.bank_id_column_combo, self.register_id_column_combo]:
                combo.clear()
                combo.addItem("请选择列")
                combo.addItems(columns)
                combo.setEnabled(True)

            self.log_text.append(
                f"sheet「{self.sheet_combo.currentText()}」共 {len(columns)} 列"
            )
            self.update_start_btn_state()

        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取sheet失败: {str(e)}")

    def update_start_btn_state(self):
        ds_ok = self.datasource_combo.currentData() is not None
        file_ok = self.excel_file is not None
        bank_col_ok = self.bank_id_column_combo.currentText() not in ("", "请选择列")
        register_col_ok = self.register_id_column_combo.currentText() not in ("", "请选择列")
        self.start_btn.setEnabled(ds_ok and file_ok and bank_col_ok and register_col_ok)

    def start_process(self):
        ds = self.datasource_combo.currentData()
        sheet_name = self.sheet_combo.currentText()
        bank_col = self.bank_id_column_combo.currentText()
        register_col = self.register_id_column_combo.currentText()

        if not ds or not self.excel_file or bank_col == "请选择列" or register_col == "请选择列":
            QMessageBox.warning(self, "警告", "请选择数据源、Excel文件、sheet和列！")
            return

        reply = QMessageBox.question(
            self,
            "确认补充",
            f"确定要补充 ba_zhb_bank 的 ptzcb_register_id 吗？\n\n"
            f"数据库: {ds.name}\n"
            f"Excel文件: {os.path.basename(self.excel_file)}\n"
            f"Sheet: {sheet_name}\n"
            f"银行表ID列: {bank_col}\n"
            f"注册表ID列: {register_col}\n\n"
            f"非数字的值将自动跳过，已有值的记录不会覆盖。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        self.start_btn.setEnabled(False)
        self.export_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在处理中...")
        self.log_text.clear()
        self.completed_data = []

        self.worker = ZhbBankRegisterFillWorker(
            ds, self.excel_file, sheet_name, bank_col, register_col
        )
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log_message.connect(self.append_log)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def append_log(self, msg):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {msg}")

    def on_finished(self, results):
        self.completed_data = results['completed_data']

        result_text = (
            f"处理完成！\n"
            f"总计: {results['total_rows']} 行\n"
            f"成功: {results['success_count']} 行\n"
            f"跳过(非数字/空值): {results['skipped_invalid']} 行\n"
            f"跳过(银行记录不存在): {results['skipped_no_bank']} 行\n"
            f"跳过(已有ptzcb_register_id): {results['skipped_already_has']} 行\n"
            f"失败: {results['failed_count']} 行"
        )
        self.result_label.setText(result_text)
        self.result_label.setStyleSheet(
            "color: green;" if results['failed_count'] == 0 else "color: orange;"
        )
        self.export_btn.setEnabled(len(self.completed_data) > 0)
        self.update_start_btn_state()
        QMessageBox.information(self, "处理完成", result_text)

    def on_error(self, error_msg):
        self.result_label.setText("处理失败！")
        self.result_label.setStyleSheet("color: red;")
        self.update_start_btn_state()
        QMessageBox.critical(self, "错误", f"处理失败: {error_msg}")

    def export_completed(self):
        if not self.completed_data:
            QMessageBox.information(self, "提示", "没有数据可导出")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "保存处理结果",
            f"zhb_bank_register_fill_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel文件 (*.xlsx)"
        )
        if path:
            try:
                df = pd.DataFrame(self.completed_data)
                df.to_excel(path, index=False, engine='openpyxl')
                QMessageBox.information(self, "成功", f"已导出到: {path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出失败: {str(e)}")
