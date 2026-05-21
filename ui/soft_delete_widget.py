"""
软删除工具
连接数据库，选择任意表，输入主键ID，执行软删除（设置delete_time）
"""
import time
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QLineEdit,
                             QFileDialog)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from database import DatabaseManager
import pymysql
import openpyxl


class SoftDeleteWorker(QThread):
    """软删除工作线程"""
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, datasource, table_name, primary_key, record_ids, preview_only=False):
        super().__init__()
        self.datasource = datasource
        self.table_name = table_name
        self.primary_key = primary_key
        self.record_ids = record_ids
        self.preview_only = preview_only

    def run(self):
        try:
            mode = "预检查" if self.preview_only else "执行"
            self.log_message.emit(f"开始{mode}软删除...")
            self.log_message.emit(f"表: {self.table_name}")
            self.log_message.emit(f"主键: {self.primary_key}")
            self.log_message.emit(f"ID数量: {len(self.record_ids)}")

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

            delete_time = int(time.time())
            success = 0
            not_found = 0
            already_deleted = 0
            failed = 0

            for rec_id in self.record_ids:
                rec_id = rec_id.strip()
                if not rec_id:
                    continue

                try:
                    # 先查询记录是否存在
                    cursor.execute(
                        f"SELECT {self.primary_key}, delete_time FROM {self.table_name} WHERE {self.primary_key} = %s",
                        (rec_id,)
                    )
                    row = cursor.fetchone()

                    if not row:
                        not_found += 1
                        self.log_message.emit(f"  ID={rec_id}: 记录不存在")
                        continue

                    if row.get('delete_time') and row['delete_time'] != 0:
                        already_deleted += 1
                        self.log_message.emit(f"  ID={rec_id}: 已被软删除(delete_time={row['delete_time']})")
                        continue

                    if self.preview_only:
                        success += 1
                        self.log_message.emit(f"  ID={rec_id}: 可删除 ✓")
                    else:
                        cursor.execute(
                            f"UPDATE {self.table_name} SET delete_time = %s WHERE {self.primary_key} = %s",
                            (delete_time, rec_id)
                        )
                        connection.commit()
                        success += 1
                        self.log_message.emit(f"  ID={rec_id}: 已软删除(delete_time={delete_time})")

                except Exception as e:
                    failed += 1
                    self.log_message.emit(f"  ID={rec_id}: 失败({str(e)})")

            cursor.close()
            connection.close()

            self.log_message.emit(f"\n=== {mode}完成 ===")
            if self.preview_only:
                self.log_message.emit(f"可删除: {success}")
            else:
                self.log_message.emit(f"已删除: {success}")
            self.log_message.emit(f"记录不存在: {not_found}")
            self.log_message.emit(f"已被删除: {already_deleted}")
            self.log_message.emit(f"失败: {failed}")

            self.finished.emit({
                'success': success,
                'not_found': not_found,
                'already_deleted': already_deleted,
                'failed': failed
            })

        except Exception as e:
            self.error.emit(str(e))


class SoftDeleteWidget(QWidget):
    """软删除工具界面"""

    def __init__(self, db_manager=None):
        super().__init__()
        self.db_manager = db_manager or DatabaseManager()
        self.worker = None
        self.table_list = []
        self.column_list = []
        self.excel_wb = None
        self.excel_ws = None
        self.yellow_row_indices = []
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

        # 表和主键选择
        table_group = QGroupBox("表设置")
        table_layout = QFormLayout()

        self.table_combo = QComboBox()
        self.table_combo.setEditable(True)
        self.table_combo.currentTextChanged.connect(self.on_table_changed)
        table_layout.addRow("选择表:", self.table_combo)

        self.load_tables_btn = QPushButton("加载表列表")
        self.load_tables_btn.setEnabled(False)
        self.load_tables_btn.clicked.connect(self.load_tables)
        table_layout.addRow("", self.load_tables_btn)

        self.pk_combo = QComboBox()
        self.pk_combo.setEditable(True)
        self.pk_combo.setCurrentText("id")
        table_layout.addRow("主键字段:", self.pk_combo)

        table_group.setLayout(table_layout)
        layout.addWidget(table_group)

        # ID输入
        id_group = QGroupBox("输入主键ID")
        id_layout = QVBoxLayout()

        # Excel导入
        import_layout = QHBoxLayout()

        self.import_excel_btn = QPushButton("从Excel导入(黄色背景行)")
        self.import_excel_btn.clicked.connect(self.import_excel)
        import_layout.addWidget(self.import_excel_btn)

        self.excel_col_combo = QComboBox()
        self.excel_col_combo.setMinimumWidth(150)
        self.excel_col_combo.addItem("请先导入Excel")
        self.excel_col_combo.setEnabled(False)
        import_layout.addWidget(self.excel_col_combo)

        self.apply_col_btn = QPushButton("确认提取该列ID")
        self.apply_col_btn.setEnabled(False)
        self.apply_col_btn.clicked.connect(self.apply_excel_column)
        import_layout.addWidget(self.apply_col_btn)

        id_layout.addLayout(import_layout)

        self.excel_info_label = QLabel("")
        id_layout.addWidget(self.excel_info_label)

        id_hint = QLabel("也可以手动输入ID，多个ID用英文逗号或换行分隔:")
        id_layout.addWidget(id_hint)

        self.id_input = QTextEdit()
        self.id_input.setPlaceholderText("例如:\n1001\n1002\n1003\n\n或: 1001,1002,1003")
        self.id_input.setMaximumHeight(120)
        self.id_input.textChanged.connect(self.update_btn_state)
        id_layout.addWidget(self.id_input)

        self.id_count_label = QLabel("已输入 0 个ID")
        id_layout.addWidget(self.id_count_label)

        id_group.setLayout(id_layout)
        layout.addWidget(id_group)

        # 操作按钮
        btn_layout = QHBoxLayout()

        self.precheck_btn = QPushButton("预检查")
        self.precheck_btn.setEnabled(False)
        self.precheck_btn.clicked.connect(self.start_precheck)
        btn_layout.addWidget(self.precheck_btn)

        self.execute_btn = QPushButton("执行软删除")
        self.execute_btn.setEnabled(False)
        self.execute_btn.setStyleSheet("background-color: #ff6b6b; color: white; font-weight: bold;")
        self.execute_btn.clicked.connect(self.start_execute)
        btn_layout.addWidget(self.execute_btn)

        layout.addLayout(btn_layout)

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
        self.load_tables_btn.setEnabled(ds_ok)
        self.table_combo.clear()
        self.pk_combo.clear()
        self.pk_combo.addItem("id")
        self.update_btn_state()

    def test_connection(self):
        ds = self.datasource_combo.currentData()
        if not ds:
            return
        ok, msg = self.db_manager.test_connection(ds)
        if ok:
            QMessageBox.information(self, "连接成功", msg)
        else:
            QMessageBox.critical(self, "连接失败", msg)

    def load_tables(self):
        ds = self.datasource_combo.currentData()
        if not ds:
            return
        try:
            connection = pymysql.connect(
                host=ds.host, port=ds.port,
                user=ds.username, password=ds.password,
                database=ds.database, charset=ds.charset
            )
            cursor = connection.cursor()
            cursor.execute("SHOW TABLES")
            tables = [row[0] for row in cursor.fetchall()]
            cursor.close()
            connection.close()

            self.table_combo.clear()
            self.table_combo.addItem("请选择表")
            for t in tables:
                self.table_combo.addItem(t)

            self.append_log(f"加载到 {len(tables)} 张表")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"加载表列表失败:\n{str(e)}")

    def on_table_changed(self, text):
        if not text or text == "请选择表":
            return
        ds = self.datasource_combo.currentData()
        if not ds:
            return
        try:
            connection = pymysql.connect(
                host=ds.host, port=ds.port,
                user=ds.username, password=ds.password,
                database=ds.database, charset=ds.charset
            )
            cursor = connection.cursor()
            cursor.execute(f"SHOW COLUMNS FROM `{text}`")
            columns = [row[0] for row in cursor.fetchall()]
            cursor.close()
            connection.close()

            self.pk_combo.clear()
            for col in columns:
                self.pk_combo.addItem(col)
            # 默认选id
            idx = self.pk_combo.findText("id")
            if idx >= 0:
                self.pk_combo.setCurrentIndex(idx)

        except Exception as e:
            self.append_log(f"加载字段失败: {str(e)}")

        self.update_btn_state()

    def _is_yellow(self, cell):
        """检查单元格背景色是否为黄色"""
        fill = cell.fill
        if fill and fill.fgColor and fill.fgColor.rgb:
            rgb = str(fill.fgColor.rgb)
            if rgb in ('00000000', '0', '00'):
                return False
            # 常见黄色: FFFF00, FFFFFF00, FFFFE599, FFFFF2CC 等
            try:
                if len(rgb) == 8:
                    r = int(rgb[2:4], 16)
                    g = int(rgb[4:6], 16)
                    b = int(rgb[6:8], 16)
                elif len(rgb) == 6:
                    r = int(rgb[0:2], 16)
                    g = int(rgb[2:4], 16)
                    b = int(rgb[4:6], 16)
                else:
                    return False
                # 黄色: R高, G高, B低
                if r >= 200 and g >= 200 and b <= 150:
                    return True
            except (ValueError, IndexError):
                return False
        return False

    def import_excel(self):
        """导入Excel文件，检测黄色背景行"""
        path, _ = QFileDialog.getOpenFileName(
            self, "选择Excel文件", "", "Excel文件 (*.xlsx *.xls)"
        )
        if not path:
            return

        try:
            self.excel_wb = openpyxl.load_workbook(path, data_only=True)
            self.excel_ws = self.excel_wb.active
            ws = self.excel_ws

            # 获取表头
            headers = []
            for col_idx in range(1, ws.max_column + 1):
                header = ws.cell(row=1, column=col_idx).value
                headers.append(str(header) if header is not None else f"列{col_idx}")

            # 检测黄色背景行
            self.yellow_row_indices = []
            for row_idx in range(2, ws.max_row + 1):
                for col_idx in range(1, ws.max_column + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    if self._is_yellow(cell):
                        self.yellow_row_indices.append(row_idx)
                        break

            # 更新列选择
            self.excel_col_combo.clear()
            self.excel_col_combo.setEnabled(True)
            self.apply_col_btn.setEnabled(True)
            for idx, header in enumerate(headers):
                self.excel_col_combo.addItem(f"{header}", idx + 1)

            total_rows = ws.max_row - 1
            yellow_count = len(self.yellow_row_indices)
            self.excel_info_label.setText(
                f"已加载: {path.split('/')[-1].split(chr(92))[-1]}  |  "
                f"总行数: {total_rows}  |  黄色背景行: {yellow_count}"
            )
            self.excel_info_label.setStyleSheet(
                "color: green;" if yellow_count > 0 else "color: orange;"
            )
            self.append_log(f"导入Excel: 总{total_rows}行, 黄色背景{yellow_count}行")

            if yellow_count == 0:
                self.append_log("警告: 未检测到黄色背景行！")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"导入Excel失败:\n{str(e)}")

    def apply_excel_column(self):
        """从选定列提取黄色行的ID"""
        if not self.excel_ws or not self.yellow_row_indices:
            QMessageBox.warning(self, "警告", "没有黄色背景行可提取！")
            return

        col_idx = self.excel_col_combo.currentData()
        if col_idx is None:
            return

        ids = []
        for row_idx in self.yellow_row_indices:
            cell_val = self.excel_ws.cell(row=row_idx, column=col_idx).value
            if cell_val is not None:
                val = str(cell_val).strip()
                if val:
                    ids.append(val)

        if not ids:
            QMessageBox.warning(self, "警告", "选定列在黄色行中没有有效值！")
            return

        self.id_input.setPlainText('\n'.join(ids))
        self.append_log(f"从Excel黄色行提取到 {len(ids)} 个ID (列: {self.excel_col_combo.currentText()})")
        self.update_btn_state()

    def get_ids(self):
        text = self.id_input.toPlainText().strip()
        if not text:
            return []
        # 支持逗号和换行分隔
        text = text.replace('\n', ',').replace('\r', ',')
        ids = [x.strip() for x in text.split(',') if x.strip()]
        return ids

    def update_btn_state(self):
        ds_ok = self.datasource_combo.currentData() is not None
        table_ok = self.table_combo.currentText() not in ('', '请选择表')
        ids = self.get_ids()
        id_ok = len(ids) > 0
        ready = ds_ok and table_ok and id_ok

        self.id_count_label.setText(f"已输入 {len(ids)} 个ID")
        self.precheck_btn.setEnabled(ready)
        self.execute_btn.setEnabled(ready)

    def start_precheck(self):
        self._run_worker(preview_only=True)

    def start_execute(self):
        table = self.table_combo.currentText()
        ids = self.get_ids()

        reply = QMessageBox.question(
            self, "确认软删除",
            f"确定要对 {table} 表中的 {len(ids)} 条记录执行软删除吗？\n\n"
            f"将设置 delete_time 为当前时间戳。\n\n"
            f"请确认已执行过「预检查」！",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        self._run_worker(preview_only=False)

    def _run_worker(self, preview_only):
        ds = self.datasource_combo.currentData()
        table = self.table_combo.currentText()
        pk = self.pk_combo.currentText()
        ids = self.get_ids()

        if not ds or not table or not pk or not ids:
            QMessageBox.warning(self, "警告", "请填写完整参数！")
            return

        self.precheck_btn.setEnabled(False)
        self.execute_btn.setEnabled(False)
        self.result_label.setText("正在预检查..." if preview_only else "正在执行...")
        self.log_text.clear()

        self.worker = SoftDeleteWorker(ds, table, pk, ids, preview_only=preview_only)
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

        s = results['success']
        nf = results['not_found']
        ad = results['already_deleted']
        f = results['failed']

        text = f"成功: {s} | 不存在: {nf} | 已删除: {ad} | 失败: {f}"
        if f > 0:
            self.result_label.setStyleSheet("color: red;")
        elif nf > 0 or ad > 0:
            self.result_label.setStyleSheet("color: orange;")
        else:
            self.result_label.setStyleSheet("color: green;")
        self.result_label.setText(text)

    def on_error(self, error_msg):
        self.precheck_btn.setEnabled(True)
        self.execute_btn.setEnabled(True)
        self.result_label.setText("处理失败")
        self.result_label.setStyleSheet("color: red;")
        self.append_log(f"错误: {error_msg}")
        QMessageBox.critical(self, "错误", f"处理失败:\n{error_msg}")
