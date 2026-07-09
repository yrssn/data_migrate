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
                             QLineEdit, QListWidget, QListWidgetItem, QCheckBox,
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
                 base_url, output_dir, id_field='id',
                 dept_config=None, legal_link_config=None, concat_field=''):
        super().__init__()
        self.datasource = datasource
        self.table_name = table_name
        # check_fields: [(字段名, 字段备注), ...]
        self.check_fields = check_fields
        self.legal_field = legal_field
        self.base_url = base_url.rstrip('/')
        self.output_dir = output_dir
        self.id_field = id_field
        # 部门分组: {'admin_field','admin_table','admin_pk','dept_field'}
        self.dept_config = dept_config
        # 法人名关联: {'link_field','link_table','link_pk','name_field'}
        self.legal_link_config = legal_link_config
        # 拼接本表字段（如 company_name）
        self.concat_field = concat_field
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

    def dept_label(self, dept_value):
        """部门目录名：4=日联一部, 5=日联二部, 其它值单独一份，查不到归为其他"""
        if dept_value is None or str(dept_value).strip() == '':
            return '其他'
        dept_str = str(dept_value).strip()
        if dept_str == '4':
            return '日联一部'
        if dept_str == '5':
            return '日联二部'
        return f'部门_{sanitize_name(dept_str)}'

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
            if self.concat_field and self.concat_field not in select_fields:
                select_fields.append(self.concat_field)
            if self.legal_link_config and self.legal_link_config['link_field'] not in select_fields:
                select_fields.append(self.legal_link_config['link_field'])
            if self.dept_config and self.dept_config['admin_field'] not in select_fields:
                select_fields.append(self.dept_config['admin_field'])
            columns_sql = ', '.join([f'`{f}`' for f in select_fields])
            cursor.execute(f"SELECT {columns_sql} FROM `{self.table_name}`")
            rows = cursor.fetchall()

            # 预加载管理员→部门映射
            admin_dept_map = {}
            if self.dept_config:
                cfg = self.dept_config
                cursor.execute(
                    f"SELECT `{cfg['admin_pk']}`, `{cfg['dept_field']}` FROM `{cfg['admin_table']}`")
                for r in cursor.fetchall():
                    admin_dept_map[str(r[cfg['admin_pk']])] = r[cfg['dept_field']]
                self.log_message.emit(
                    f"已加载 {cfg['admin_table']} 部门映射 {len(admin_dept_map)} 条")

            # 预加载关联表→法人名映射
            legal_name_map = {}
            if self.legal_link_config:
                cfg = self.legal_link_config
                cursor.execute(
                    f"SELECT `{cfg['link_pk']}`, `{cfg['name_field']}` FROM `{cfg['link_table']}`")
                for r in cursor.fetchall():
                    legal_name_map[str(r[cfg['link_pk']])] = r[cfg['name_field']]
                self.log_message.emit(
                    f"已加载 {cfg['link_table']} 法人名映射 {len(legal_name_map)} 条")

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

                # 法人名：本表字段或关联表字段，可拼接本表字段（如company_name）
                if self.legal_link_config:
                    link_value = row.get(self.legal_link_config['link_field'])
                    base_name = legal_name_map.get(str(link_value), '') if link_value is not None else ''
                    if not base_name:
                        base_name = f"未关联_{link_value}" if link_value is not None else '未关联'
                else:
                    base_name = row.get(self.legal_field, '') if self.legal_field else str(row_id)
                if self.concat_field:
                    concat_value = row.get(self.concat_field) or ''
                    if concat_value:
                        base_name = f"{base_name}_{concat_value}" if base_name else str(concat_value)
                legal_name = sanitize_name(base_name)

                # 部门
                dept_name = ''
                if self.dept_config:
                    admin_value = row.get(self.dept_config['admin_field'])
                    dept_value = admin_dept_map.get(str(admin_value)) if admin_value is not None else None
                    dept_name = self.dept_label(dept_value)

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
                                'dept': dept_name,
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
            base_dir = root_dir
            if record.get('dept'):
                base_dir = os.path.join(root_dir, sanitize_name(record['dept']))
            legal_dir = os.path.join(base_dir, sanitize_name(record['legal_name']))
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
                if record.get('dept'):
                    f.write(f"部门: {record['dept']}\n")
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
            'dept': '部门',
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
        form_layout.addRow("法人名字段(本表):", self.legal_field_combo)

        self.id_field_combo = QComboBox()
        form_layout.addRow("主键字段:", self.id_field_combo)

        # 法人名关联配置（如 ba_rlb_legal_information 通过 legal_id 关联 ba_rlb_customer 取 legal_name）
        self.legal_link_checkbox = QCheckBox(
            "从关联表取法人名（如 legal_id → ba_rlb_customer.legal_name，勾选后忽略上面本表法人名字段）")
        form_layout.addRow("", self.legal_link_checkbox)

        legal_link_layout = QHBoxLayout()
        legal_link_layout.addWidget(QLabel("本表关联字段:"))
        self.link_field_edit = QLineEdit("legal_id")
        legal_link_layout.addWidget(self.link_field_edit)
        legal_link_layout.addWidget(QLabel("关联表:"))
        self.link_table_edit = QLineEdit("ba_rlb_customer")
        legal_link_layout.addWidget(self.link_table_edit)
        legal_link_layout.addWidget(QLabel("关联表主键:"))
        self.link_pk_edit = QLineEdit("id")
        legal_link_layout.addWidget(self.link_pk_edit)
        legal_link_layout.addWidget(QLabel("法人名字段:"))
        self.link_name_edit = QLineEdit("legal_name")
        legal_link_layout.addWidget(self.link_name_edit)
        form_layout.addRow("", legal_link_layout)

        self.concat_field_edit = QLineEdit()
        self.concat_field_edit.setPlaceholderText("可选，如 company_name，目录名为 法人名_拼接字段值")
        form_layout.addRow("目录名拼接本表字段:", self.concat_field_edit)

        # 部门分组配置（admin_id → ba_admin.dept_id，4=日联一部 5=日联二部，其它单独一份）
        self.dept_checkbox = QCheckBox(
            "按部门分目录（admin_id → 管理员表查部门，4=日联一部 5=日联二部，其它部门单独一份，查不到归为其他）")
        form_layout.addRow("", self.dept_checkbox)

        dept_layout = QHBoxLayout()
        dept_layout.addWidget(QLabel("本表管理员字段:"))
        self.admin_field_edit = QLineEdit("admin_id")
        dept_layout.addWidget(self.admin_field_edit)
        dept_layout.addWidget(QLabel("管理员表:"))
        self.admin_table_edit = QLineEdit("ba_admin")
        dept_layout.addWidget(self.admin_table_edit)
        dept_layout.addWidget(QLabel("管理员表主键:"))
        self.admin_pk_edit = QLineEdit("id")
        dept_layout.addWidget(self.admin_pk_edit)
        dept_layout.addWidget(QLabel("部门字段:"))
        self.dept_field_edit = QLineEdit("dept_id")
        dept_layout.addWidget(self.dept_field_edit)
        form_layout.addRow("", dept_layout)

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
        self.result_table.setColumnCount(8)
        self.result_table.setHorizontalHeaderLabels([
            "行ID", "部门", "法人", "字段名", "字段备注", "位置", "原文件名", "原URL"
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

        legal_link_config = None
        if self.legal_link_checkbox.isChecked():
            legal_link_config = {
                'link_field': self.link_field_edit.text().strip(),
                'link_table': self.link_table_edit.text().strip(),
                'link_pk': self.link_pk_edit.text().strip(),
                'name_field': self.link_name_edit.text().strip(),
            }
            if not all(legal_link_config.values()):
                QMessageBox.warning(self, "提示", "请填写完整的法人名关联配置")
                return
        elif not legal_field:
            QMessageBox.warning(self, "提示", "请选择法人名字段或启用关联表取法人名")
            return

        dept_config = None
        if self.dept_checkbox.isChecked():
            dept_config = {
                'admin_field': self.admin_field_edit.text().strip(),
                'admin_table': self.admin_table_edit.text().strip(),
                'admin_pk': self.admin_pk_edit.text().strip(),
                'dept_field': self.dept_field_edit.text().strip(),
            }
            if not all(dept_config.values()):
                QMessageBox.warning(self, "提示", "请填写完整的部门分组配置")
                return

        concat_field = self.concat_field_edit.text().strip()
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
            base_url, output_dir, id_field,
            dept_config=dept_config,
            legal_link_config=legal_link_config,
            concat_field=concat_field)
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
            self.result_table.setItem(row, 1, QTableWidgetItem(str(record.get('dept', ''))))
            self.result_table.setItem(row, 2, QTableWidgetItem(str(record['legal_name'])))
            self.result_table.setItem(row, 3, QTableWidgetItem(str(record['field_name'])))
            self.result_table.setItem(row, 4, QTableWidgetItem(str(record['field_comment'])))
            self.result_table.setItem(row, 5, QTableWidgetItem(str(record['index'] + 1)))
            self.result_table.setItem(row, 6, QTableWidgetItem(str(record['file_name'])))
            self.result_table.setItem(row, 7, QTableWidgetItem(str(record['url'])))

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
