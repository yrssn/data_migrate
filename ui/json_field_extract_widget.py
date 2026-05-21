"""
JSON数组字段元素提取与删除
从指定表的JSON数组字段中，根据关键字匹配提取元素（导出Excel日志），
然后从JSON数组中删除匹配的元素并更新回数据库。
保持JSON结构完整性。
"""
import os
import json
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog,
                             QLineEdit)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from database import DatabaseManager
import pandas as pd
import pymysql


class JsonFieldExtractWorker(QThread):
    """JSON字段提取与删除工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, datasource, table_name, id_field, json_field,
                 keyword, preview_only=False):
        super().__init__()
        self.datasource = datasource
        self.table_name = table_name
        self.id_field = id_field
        self.json_field = json_field
        self.keyword = keyword
        self.preview_only = preview_only
        self.results = {
            'total_records': 0,
            'matched_records': 0,
            'matched_elements': 0,
            'updated_records': 0,
            'skipped_records': 0,
            'error_records': 0,
            'completed_data': [],
            'failed_records': []
        }

    def run(self):
        try:
            mode = "预检查" if self.preview_only else "执行"
            self.log_message.emit(f"开始{mode}...")
            self.log_message.emit(f"表: {self.table_name}")
            self.log_message.emit(f"JSON字段: {self.json_field}")
            self.log_message.emit(f"匹配关键字: {self.keyword}")

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

            # 查询包含关键字的记录
            self.log_message.emit(f"正在查询包含关键字的记录...")
            query = f"""
                SELECT {self.id_field}, {self.json_field}
                FROM {self.table_name}
                WHERE {self.json_field} LIKE %s
            """
            cursor.execute(query, (f'%{self.keyword}%',))
            records = cursor.fetchall()
            total = len(records)
            self.results['total_records'] = total
            self.log_message.emit(f"找到 {total} 条包含关键字的记录")
            self.progress.emit(20)

            if total == 0:
                self.log_message.emit("没有匹配的记录")
                cursor.close()
                connection.close()
                self.progress.emit(100)
                self.finished.emit(self.results)
                return

            matched_records = 0
            matched_elements = 0
            updated_records = 0
            skipped_records = 0
            error_records = 0

            for index, rec in enumerate(records):
                rec_id = rec[self.id_field]
                raw_json = rec[self.json_field]

                try:
                    # 解析JSON
                    if raw_json is None or raw_json.strip() == '':
                        skipped_records += 1
                        continue

                    arr = json.loads(raw_json)
                    if not isinstance(arr, list):
                        skipped_records += 1
                        row = {
                            'id': rec_id,
                            '原始值': raw_json,
                            '处理状态': '跳过(非数组)'
                        }
                        self.results['completed_data'].append(row)
                        continue

                    # 找出精确匹配的元素
                    matched = [elem for elem in arr if str(elem) == self.keyword]
                    not_matched = [elem for elem in arr if str(elem) != self.keyword]

                    if not matched:
                        skipped_records += 1
                        continue

                    matched_records += 1
                    matched_elements += len(matched)

                    row = {
                        'id': rec_id,
                        '原始JSON': raw_json,
                        '匹配元素数': len(matched),
                        '匹配的元素': json.dumps(matched, ensure_ascii=False),
                        '删除后JSON': json.dumps(not_matched, ensure_ascii=False),
                        '原数组长度': len(arr),
                        '删除后数组长度': len(not_matched),
                    }

                    # 更新数据库
                    if not self.preview_only:
                        new_json = json.dumps(not_matched, ensure_ascii=False)
                        update_sql = f"""
                            UPDATE {self.table_name}
                            SET {self.json_field} = %s
                            WHERE {self.id_field} = %s
                        """
                        cursor.execute(update_sql, (new_json, rec_id))
                        connection.commit()
                        updated_records += 1
                        row['处理状态'] = '已更新'
                    else:
                        row['处理状态'] = '可处理'

                    self.results['completed_data'].append(row)

                except json.JSONDecodeError as e:
                    error_records += 1
                    err_row = {
                        'id': rec_id,
                        '原始值': raw_json[:200] if raw_json else '',
                        '处理状态': f'JSON解析失败({str(e)})'
                    }
                    self.results['completed_data'].append(err_row)
                    self.results['failed_records'].append(err_row)

                except Exception as e:
                    error_records += 1
                    err_row = {
                        'id': rec_id,
                        '处理状态': f'异常({str(e)})'
                    }
                    self.results['completed_data'].append(err_row)
                    self.results['failed_records'].append(err_row)

                progress = 20 + int((index + 1) / total * 75)
                self.progress.emit(progress)

            cursor.close()
            connection.close()

            self.results['matched_records'] = matched_records
            self.results['matched_elements'] = matched_elements
            self.results['updated_records'] = updated_records
            self.results['skipped_records'] = skipped_records
            self.results['error_records'] = error_records

            self.progress.emit(100)
            self.log_message.emit(f"\n=== {mode}完成 ===")
            self.log_message.emit(f"总记录: {total}")
            self.log_message.emit(f"匹配记录: {matched_records}")
            self.log_message.emit(f"匹配元素总数: {matched_elements}")
            if not self.preview_only:
                self.log_message.emit(f"已更新: {updated_records}")
            self.log_message.emit(f"跳过: {skipped_records}")
            self.log_message.emit(f"错误: {error_records}")

            self.finished.emit(self.results)

        except Exception as e:
            self.error.emit(str(e))


class JsonFieldExtractWidget(QWidget):
    """JSON数组字段元素提取与删除界面"""

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

        # 参数设置
        param_group = QGroupBox("参数设置")
        param_layout = QFormLayout()

        self.table_name_edit = QLineEdit("ba_request_data")
        param_layout.addRow("表名:", self.table_name_edit)

        self.id_field_edit = QLineEdit("id")
        param_layout.addRow("ID字段:", self.id_field_edit)

        self.json_field_edit = QLineEdit("request_materials")
        param_layout.addRow("JSON数组字段:", self.json_field_edit)

        self.keyword_edit = QLineEdit()
        self.keyword_edit.setPlaceholderText("输入要匹配的关键字，如: matters_certificates_paper")
        self.keyword_edit.textChanged.connect(self.update_start_state)
        param_layout.addRow("匹配关键字:", self.keyword_edit)

        param_group.setLayout(param_layout)
        layout.addWidget(param_group)

        # 功能说明
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        info_label = QLabel("""
<b>JSON数组字段元素提取与删除:</b><br><br>
1. 根据关键字匹配JSON数组中的元素<br>
2. 导出匹配记录到Excel（完整保留原始JSON结构作为日志）<br>
3. 从JSON数组中删除匹配的元素，更新回数据库<br>
4. <b>不破坏JSON结构</b>：使用json.loads解析 → 过滤 → json.dumps回写<br><br>
<span style="color:red;"><b>生产环境操作，请务必先「预检查」确认后再「执行删除」！</b></span>
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

        self.execute_btn = QPushButton("执行删除")
        self.execute_btn.setEnabled(False)
        self.execute_btn.setStyleSheet("background-color: #ff6b6b; color: white; font-weight: bold;")
        self.execute_btn.clicked.connect(self.start_execute)
        btn_layout.addWidget(self.execute_btn)

        self.export_btn = QPushButton("导出日志Excel")
        self.export_btn.setEnabled(False)
        self.export_btn.clicked.connect(self.export_log)
        btn_layout.addWidget(self.export_btn)

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

    def update_start_state(self):
        ds_ok = self.datasource_combo.currentData() is not None
        keyword_ok = self.keyword_edit.text().strip() != ''
        ready = ds_ok and keyword_ok
        self.precheck_btn.setEnabled(ready)
        self.execute_btn.setEnabled(ready)

    def start_precheck(self):
        ds = self.datasource_combo.currentData()
        keyword = self.keyword_edit.text().strip()
        if not ds or not keyword:
            QMessageBox.warning(self, "警告", "请选择数据源并输入关键字！")
            return
        self._run_worker(ds, keyword, preview_only=True)

    def start_execute(self):
        ds = self.datasource_combo.currentData()
        keyword = self.keyword_edit.text().strip()
        if not ds or not keyword:
            QMessageBox.warning(self, "警告", "请选择数据源并输入关键字！")
            return

        table = self.table_name_edit.text().strip()
        json_field = self.json_field_edit.text().strip()

        reply = QMessageBox.question(
            self, "确认执行",
            f"⚠️ 生产环境操作！\n\n"
            f"确定要从 {table}.{json_field} 中\n"
            f"删除包含「{keyword}」的元素吗？\n\n"
            f"数据库: {ds.name}\n\n"
            f"此操作会直接修改数据库记录！\n"
            f"请确认已执行过「预检查」并导出了日志！",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        # 二次确认
        reply2 = QMessageBox.warning(
            self, "最终确认",
            f"最终确认：将从数据库中删除匹配的JSON元素！\n\n"
            f"关键字: {keyword}\n"
            f"此操作不可撤销！",
            QMessageBox.Ok | QMessageBox.Cancel, QMessageBox.Cancel
        )
        if reply2 != QMessageBox.Ok:
            return

        self._run_worker(ds, keyword, preview_only=False)

    def _run_worker(self, ds, keyword, preview_only):
        table = self.table_name_edit.text().strip()
        id_field = self.id_field_edit.text().strip()
        json_field = self.json_field_edit.text().strip()

        if not table or not id_field or not json_field:
            QMessageBox.warning(self, "警告", "表名、ID字段、JSON字段不能为空！")
            return

        self.precheck_btn.setEnabled(False)
        self.execute_btn.setEnabled(False)
        self.export_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在预检查..." if preview_only else "正在执行...")
        self.log_text.clear()
        self.completed_data = []
        self.failed_records = []

        self.worker = JsonFieldExtractWorker(
            ds, table, id_field, json_field, keyword, preview_only=preview_only
        )
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
        self.execute_btn.setEnabled(True)
        self.progress_bar.setVisible(False)

        self.completed_data = results.get('completed_data', [])
        self.failed_records = results.get('failed_records', [])

        total = results['total_records']
        matched_rec = results['matched_records']
        matched_elem = results['matched_elements']
        updated = results['updated_records']
        skipped = results['skipped_records']
        errors = results['error_records']

        text = (f"总{total}条 | 匹配记录: {matched_rec} | 匹配元素: {matched_elem} | "
                f"已更新: {updated} | 跳过: {skipped} | 错误: {errors}")

        if errors > 0:
            self.result_label.setStyleSheet("color: red;")
        elif matched_rec > 0:
            self.result_label.setStyleSheet("color: orange;")
        else:
            self.result_label.setStyleSheet("color: green;")
        self.result_label.setText(text)

        self.export_btn.setEnabled(len(self.completed_data) > 0)

    def on_error(self, error_msg):
        self.precheck_btn.setEnabled(True)
        self.execute_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.result_label.setText("处理失败")
        self.result_label.setStyleSheet("color: red;")
        self.append_log(f"错误: {error_msg}")
        QMessageBox.critical(self, "错误", f"处理失败:\n{error_msg}")

    def export_log(self):
        if not self.completed_data:
            QMessageBox.warning(self, "警告", "没有可导出的数据！")
            return
        keyword = self.keyword_edit.text().strip() or 'unknown'
        default_name = f"JSON提取日志_{keyword}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "保存日志", default_name, "Excel文件 (*.xlsx)")
        if path:
            df = pd.DataFrame(self.completed_data)
            df.to_excel(path, index=False, engine='openpyxl')
            QMessageBox.information(self, "成功", f"已导出 {len(self.completed_data)} 条记录到:\n{path}")
