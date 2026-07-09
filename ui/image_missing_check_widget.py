"""
图片丢失检查功能界面

动态选择表和字段，检查JSON数组格式的文件字段（如
[{"name":"xxx.jpg","url":"/uploads/...","uid":"...","status":"success"}]）
中的文件在服务器上是否丢失。对丢失的文件按
输出目录/缺失/<法人名>/<字段备注名>/<序号_原文件名>/ 生成文件夹结构，
并导出Excel报告和JSON清单，方便业务补充图片后按位置回填。
"""
import os
import json
import re
import ssl
import urllib.request
import urllib.parse
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog,
                             QLineEdit, QListWidget, QListWidgetItem,
                             QTableWidget, QTableWidgetItem, QHeaderView)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import pandas as pd
import pymysql


def sanitize_name(name: str) -> str:
    """清理文件夹/文件名中的非法字符"""
    name = str(name or '').strip()
    name = re.sub(r'[\\/:*?"<>|\r\n\t]', '_', name)
    return name or '未命名'


class ImageMissingCheckWorker(QThread):
    """图片丢失检查工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, datasource, table_name, check_fields, legal_field,
                 base_url, output_dir, id_field='id'):
        super().__init__()
        self.datasource = datasource
        self.table_name = table_name
        # check_fields: [(字段名, 字段备注), ...]
        self.check_fields = check_fields
        self.legal_field = legal_field
        self.base_url = base_url.rstrip('/')
        self.output_dir = output_dir
        self.id_field = id_field
        self.stop_requested = False
        self.url_cache = {}
        self.results = {
            'total_rows': 0,
            'checked_files': 0,
            'missing_files': 0,
            'invalid_json': 0,
            'missing_records': []
        }

    def request_stop(self):
        self.stop_requested = True

    def check_url_exists(self, url_path: str) -> bool:
        """检查文件是否存在（HTTP地址用HEAD请求，本地路径直接判断文件）"""
        if url_path in self.url_cache:
            return self.url_cache[url_path]

        exists = False
        try:
            if self.base_url.lower().startswith(('http://', 'https://')):
                quoted = urllib.parse.quote(url_path, safe='/')
                full_url = self.base_url + quoted
                ctx = ssl.create_default_context()
                ctx.check_hostname = False
                ctx.verify_mode = ssl.CERT_NONE
                req = urllib.request.Request(full_url, method='HEAD')
                try:
                    with urllib.request.urlopen(req, timeout=10, context=ctx) as resp:
                        exists = 200 <= resp.status < 400
                except urllib.error.HTTPError as e:
                    if e.code == 405:  # 服务器不支持HEAD，改用GET
                        req = urllib.request.Request(full_url, method='GET')
                        with urllib.request.urlopen(req, timeout=10, context=ctx) as resp:
                            exists = 200 <= resp.status < 400
                    else:
                        exists = False
            else:
                local_path = os.path.join(self.base_url, url_path.lstrip('/\\'))
                exists = os.path.isfile(local_path)
        except Exception:
            exists = False

        self.url_cache[url_path] = exists
        return exists

    def parse_file_entries(self, raw_value):
        """解析JSON数组格式的文件字段，返回 (entries, is_valid)"""
        if raw_value is None:
            return [], True
        text = str(raw_value).strip()
        if text == '' or text == '[]' or text.lower() == 'null':
            return [], True
        try:
            data = json.loads(text)
        except Exception:
            return [], False
        if isinstance(data, dict):
            data = [data]
        if not isinstance(data, list):
            return [], False
        entries = []
        for item in data:
            if isinstance(item, dict) and item.get('url'):
                entries.append(item)
        return entries, True

    def run(self):
        try:
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
            self.progress.emit(5)

            field_names = [f[0] for f in self.check_fields]
            select_fields = [self.id_field] + field_names
            if self.legal_field and self.legal_field not in select_fields:
                select_fields.append(self.legal_field)
            columns_sql = ', '.join([f'`{f}`' for f in select_fields])
            cursor.execute(f"SELECT {columns_sql} FROM `{self.table_name}`")
            rows = cursor.fetchall()
            cursor.close()
            connection.close()

            self.results['total_rows'] = len(rows)
            self.log_message.emit(f"查询到 {len(rows)} 行数据，开始检查 {len(self.check_fields)} 个字段...")
            self.progress.emit(10)

            total = len(rows) if rows else 1
            for index, row in enumerate(rows):
                if self.stop_requested:
                    self.log_message.emit("已手动停止检查")
                    break

                row_id = row.get(self.id_field)
                legal_name = sanitize_name(row.get(self.legal_field, '')) if self.legal_field else str(row_id)

                for field_name, field_comment in self.check_fields:
                    entries, is_valid = self.parse_file_entries(row.get(field_name))
                    if not is_valid:
                        self.results['invalid_json'] += 1
                        self.log_message.emit(
                            f"[格式异常] id={row_id} 字段 {field_name} 不是有效的JSON数组，已跳过")
                        continue

                    for idx, entry in enumerate(entries):
                        url_path = str(entry.get('url', ''))
                        file_name = str(entry.get('name', '')) or os.path.basename(url_path)
                        self.results['checked_files'] += 1
                        if not self.check_url_exists(url_path):
                            self.results['missing_files'] += 1
                            record = {
                                'row_id': row_id,
                                'legal_name': legal_name,
                                'field_name': field_name,
                                'field_comment': field_comment or field_name,
                                'index': idx,
                                'file_name': file_name,
                                'url': url_path,
                                'uid': entry.get('uid', ''),
                            }
                            self.results['missing_records'].append(record)
                            self.log_message.emit(
                                f"[丢失] id={row_id} {legal_name} {field_comment or field_name}"
                                f" 第{idx + 1}个文件: {file_name}")

                progress = 10 + int((index + 1) / total * 80)
                self.progress.emit(progress)
                if (index + 1) % 100 == 0:
                    self.log_message.emit(f"已检查 {index + 1}/{len(rows)} 行")

            self.log_message.emit("检查完成，开始生成缺失文件夹结构和报告...")
            self.generate_output()
            self.progress.emit(100)
            self.finished.emit(self.results)

        except Exception as e:
            self.error.emit(str(e))

    def generate_output(self):
        """生成缺失文件夹结构、Excel报告和JSON清单"""
        records = self.results['missing_records']
        if not records:
            self.log_message.emit("没有发现丢失的文件，不生成文件夹")
            return

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        root_dir = os.path.join(self.output_dir, f"缺失_{self.table_name}_{timestamp}")
        os.makedirs(root_dir, exist_ok=True)

        for record in records:
            legal_dir = os.path.join(root_dir, sanitize_name(record['legal_name']))
            field_dir = os.path.join(legal_dir, sanitize_name(record['field_comment']))
            # 每个丢失文件对应一个固定位置的子文件夹，业务把补充的图片放进去即可，
            # 文件名可以不一样，回填时按文件夹对应的 行ID/字段/序号 精确回填
            entry_dir = os.path.join(
                field_dir,
                f"{record['index'] + 1:02d}_{sanitize_name(record['file_name'])}")
            os.makedirs(entry_dir, exist_ok=True)

            readme_path = os.path.join(entry_dir, '说明.txt')
            with open(readme_path, 'w', encoding='utf-8') as f:
                f.write(f"表名: {self.table_name}\n")
                f.write(f"行ID: {record['row_id']}\n")
                f.write(f"法人: {record['legal_name']}\n")
                f.write(f"字段: {record['field_name']} ({record['field_comment']})\n")
                f.write(f"数组位置: 第{record['index'] + 1}个\n")
                f.write(f"原文件名: {record['file_name']}\n")
                f.write(f"原URL: {record['url']}\n")
                f.write("请将补充的图片文件放到本文件夹中（文件名可以不一样，一个文件夹只放一个文件）\n")

            record['folder'] = os.path.relpath(entry_dir, root_dir)

        # JSON清单，供后续回填脚本使用
        manifest = {
            'table': self.table_name,
            'id_field': self.id_field,
            'legal_field': self.legal_field,
            'base_url': self.base_url,
            'generated_at': timestamp,
            'missing': records,
        }
        manifest_path = os.path.join(root_dir, '缺失清单.json')
        with open(manifest_path, 'w', encoding='utf-8') as f:
            json.dump(manifest, f, ensure_ascii=False, indent=2, default=str)

        # Excel报告
        df = pd.DataFrame(records)
        df = df.rename(columns={
            'row_id': '行ID',
            'legal_name': '法人',
            'field_name': '字段名',
            'field_comment': '字段备注',
            'index': '数组位置(从0开始)',
            'file_name': '原文件名',
            'url': '原URL',
            'uid': 'uid',
            'folder': '对应文件夹',
        })
        excel_path = os.path.join(root_dir, '缺失清单.xlsx')
        df.to_excel(excel_path, index=False, engine='openpyxl')

        self.results['output_dir'] = root_dir
        self.log_message.emit(f"已生成缺失文件夹结构和报告: {root_dir}")


class ImageMissingCheckWidget(QWidget):
    """图片丢失检查界面"""

    def __init__(self, db_manager):
        super().__init__()
        self.db_manager = db_manager
        self.worker = None
        self.columns_info = []  # [(字段名, 类型, 备注), ...]
        self.init_ui()
        self.load_datasources()

    def init_ui(self):
        layout = QVBoxLayout()

        config_group = QGroupBox("检查配置")
        form_layout = QFormLayout()

        self.datasource_combo = QComboBox()
        form_layout.addRow("数据源:", self.datasource_combo)

        table_layout = QHBoxLayout()
        self.table_combo = QComboBox()
        self.table_combo.setMinimumWidth(300)
        self.load_tables_btn = QPushButton("加载表")
        self.load_tables_btn.clicked.connect(self.load_tables)
        table_layout.addWidget(self.table_combo)
        table_layout.addWidget(self.load_tables_btn)
        table_layout.addStretch()
        form_layout.addRow("检查表:", table_layout)

        field_btn_layout = QHBoxLayout()
        self.load_fields_btn = QPushButton("加载字段")
        self.load_fields_btn.clicked.connect(self.load_fields)
        field_btn_layout.addWidget(self.load_fields_btn)
        field_btn_layout.addStretch()
        form_layout.addRow("", field_btn_layout)

        self.fields_list = QListWidget()
        self.fields_list.setMinimumHeight(150)
        form_layout.addRow("检查字段(勾选):", self.fields_list)

        self.legal_field_combo = QComboBox()
        form_layout.addRow("法人名字段:", self.legal_field_combo)

        self.id_field_combo = QComboBox()
        form_layout.addRow("主键字段:", self.id_field_combo)

        self.base_url_edit = QLineEdit()
        self.base_url_edit.setPlaceholderText(
            "如 https://img.example.com 或本地目录 D:\\www\\public（url会拼接到后面）")
        form_layout.addRow("服务器拼接地址:", self.base_url_edit)

        output_layout = QHBoxLayout()
        self.output_dir_edit = QLineEdit()
        self.output_dir_edit.setReadOnly(True)
        self.select_output_btn = QPushButton("选择输出目录")
        self.select_output_btn.clicked.connect(self.select_output_dir)
        output_layout.addWidget(self.output_dir_edit)
        output_layout.addWidget(self.select_output_btn)
        form_layout.addRow("输出目录:", output_layout)

        config_group.setLayout(form_layout)
        layout.addWidget(config_group)

        btn_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始检查")
        self.start_btn.clicked.connect(self.start_check)
        btn_layout.addWidget(self.start_btn)
        self.stop_btn = QPushButton("停止")
        self.stop_btn.clicked.connect(self.stop_check)
        self.stop_btn.setEnabled(False)
        btn_layout.addWidget(self.stop_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(150)
        layout.addWidget(self.log_text)

        result_group = QGroupBox("缺失结果")
        result_layout = QVBoxLayout()
        self.result_table = QTableWidget()
        self.result_table.setColumnCount(7)
        self.result_table.setHorizontalHeaderLabels([
            "行ID", "法人", "字段名", "字段备注", "位置", "原文件名", "原URL"
        ])
        self.result_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.result_table.horizontalHeader().setStretchLastSection(True)
        self.result_table.setEditTriggers(QTableWidget.NoEditTriggers)
        result_layout.addWidget(self.result_table)
        result_group.setLayout(result_layout)
        layout.addWidget(result_group)

        self.setLayout(layout)

    def load_datasources(self):
        self.datasource_combo.clear()
        for ds in self.db_manager.get_all_datasources():
            self.datasource_combo.addItem(f"{ds.name} ({ds.host}:{ds.port}/{ds.database})", ds)

    def get_selected_datasource(self):
        return self.datasource_combo.currentData()

    def load_tables(self):
        datasource = self.get_selected_datasource()
        if not datasource:
            QMessageBox.warning(self, "提示", "请先选择数据源")
            return
        try:
            connection = pymysql.connect(
                host=datasource.host, port=datasource.port,
                user=datasource.username, password=datasource.password,
                database=datasource.database, charset=datasource.charset,
                connect_timeout=10)
            cursor = connection.cursor()
            cursor.execute("SHOW TABLES")
            tables = [row[0] for row in cursor.fetchall()]
            cursor.close()
            connection.close()
            self.table_combo.clear()
            self.table_combo.addItems(tables)
            self.log(f"加载了 {len(tables)} 张表")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"加载表失败: {str(e)}")

    def load_fields(self):
        datasource = self.get_selected_datasource()
        table_name = self.table_combo.currentText()
        if not datasource or not table_name:
            QMessageBox.warning(self, "提示", "请先选择数据源和表")
            return
        try:
            connection = pymysql.connect(
                host=datasource.host, port=datasource.port,
                user=datasource.username, password=datasource.password,
                database=datasource.database, charset=datasource.charset,
                connect_timeout=10)
            cursor = connection.cursor()
            cursor.execute(f"SHOW FULL COLUMNS FROM `{table_name}`")
            columns = cursor.fetchall()
            cursor.close()
            connection.close()

            # SHOW FULL COLUMNS: Field, Type, Collation, Null, Key, Default, Extra, Privileges, Comment
            self.columns_info = [(col[0], col[1], col[8] or '') for col in columns]

            self.fields_list.clear()
            self.legal_field_combo.clear()
            self.id_field_combo.clear()
            for name, col_type, comment in self.columns_info:
                display = f"{name}  [{col_type}]" + (f"  ({comment})" if comment else "")
                item = QListWidgetItem(display)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Unchecked)
                item.setData(Qt.UserRole, (name, comment))
                self.fields_list.addItem(item)

                legal_display = f"{name}" + (f" ({comment})" if comment else "")
                self.legal_field_combo.addItem(legal_display, name)
                self.id_field_combo.addItem(legal_display, name)

            for i, (name, _, _) in enumerate(self.columns_info):
                if name == 'id':
                    self.id_field_combo.setCurrentIndex(i)
                if name == 'legal_name':
                    self.legal_field_combo.setCurrentIndex(i)

            self.log(f"加载了表 {table_name} 的 {len(self.columns_info)} 个字段")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"加载字段失败: {str(e)}")

    def select_output_dir(self):
        dir_path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if dir_path:
            self.output_dir_edit.setText(dir_path)

    def get_checked_fields(self):
        fields = []
        for i in range(self.fields_list.count()):
            item = self.fields_list.item(i)
            if item.checkState() == Qt.Checked:
                fields.append(item.data(Qt.UserRole))
        return fields

    def start_check(self):
        datasource = self.get_selected_datasource()
        table_name = self.table_combo.currentText()
        check_fields = self.get_checked_fields()
        legal_field = self.legal_field_combo.currentData()
        id_field = self.id_field_combo.currentData()
        base_url = self.base_url_edit.text().strip()
        output_dir = self.output_dir_edit.text().strip()

        if not datasource:
            QMessageBox.warning(self, "提示", "请选择数据源")
            return
        if not table_name:
            QMessageBox.warning(self, "提示", "请选择表")
            return
        if not check_fields:
            QMessageBox.warning(self, "提示", "请勾选要检查的字段")
            return
        if not legal_field:
            QMessageBox.warning(self, "提示", "请选择法人名字段")
            return
        if not id_field:
            QMessageBox.warning(self, "提示", "请选择主键字段")
            return
        if not base_url:
            QMessageBox.warning(self, "提示", "请填写服务器拼接地址或本地目录")
            return
        if not output_dir:
            QMessageBox.warning(self, "提示", "请选择输出目录")
            return

        self.result_table.setRowCount(0)
        self.log_text.clear()
        self.progress_bar.setValue(0)
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)

        self.worker = ImageMissingCheckWorker(
            datasource, table_name, check_fields, legal_field,
            base_url, output_dir, id_field)
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log_message.connect(self.log)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def stop_check(self):
        if self.worker:
            self.worker.request_stop()
            self.stop_btn.setEnabled(False)

    def log(self, message):
        self.log_text.append(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")

    def on_finished(self, results):
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)

        for record in results['missing_records']:
            row = self.result_table.rowCount()
            self.result_table.insertRow(row)
            self.result_table.setItem(row, 0, QTableWidgetItem(str(record['row_id'])))
            self.result_table.setItem(row, 1, QTableWidgetItem(str(record['legal_name'])))
            self.result_table.setItem(row, 2, QTableWidgetItem(str(record['field_name'])))
            self.result_table.setItem(row, 3, QTableWidgetItem(str(record['field_comment'])))
            self.result_table.setItem(row, 4, QTableWidgetItem(str(record['index'] + 1)))
            self.result_table.setItem(row, 5, QTableWidgetItem(str(record['file_name'])))
            self.result_table.setItem(row, 6, QTableWidgetItem(str(record['url'])))

        summary = (f"检查完成: 共 {results['total_rows']} 行, "
                   f"检查文件 {results['checked_files']} 个, "
                   f"丢失 {results['missing_files']} 个, "
                   f"JSON格式异常 {results['invalid_json']} 个")
        self.log(summary)
        output_dir = results.get('output_dir')
        if output_dir:
            self.log(f"缺失文件夹和报告已生成到: {output_dir}")
            QMessageBox.information(self, "完成", f"{summary}\n\n输出目录:\n{output_dir}")
        else:
            QMessageBox.information(self, "完成", summary)

    def on_error(self, message):
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.log(f"检查失败: {message}")
        QMessageBox.critical(self, "错误", f"检查失败: {message}")
