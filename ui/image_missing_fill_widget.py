"""
图片回填功能界面

配合「图片丢失检查」使用：业务把补充的图片放进检查生成的
缺失_<表名>_<时间戳>/<法人>/<字段备注>/<序号_原文件名>/ 文件夹后，
本功能读取 缺失清单.json，扫描各文件夹中的图片，按数据库现有规则
重命名（<字段备注>_<YYYYMMDD>_<unix时间戳>.<扩展名>），复制生成
uploads/<年>/<月>/<日>/ 目录结构（直接上传到服务器即可），
并可选择同时按 行ID/字段/数组位置 精确更新数据库JSON字段的 name/url。
"""
import os
import json
import re
import time
import shutil
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog,
                             QLineEdit, QCheckBox,
                             QTableWidget, QTableWidgetItem, QHeaderView)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import pandas as pd
import pymysql


IGNORED_FILES = {'说明.txt', 'thumbs.db', 'desktop.ini', '.ds_store'}


def sanitize_name(name: str) -> str:
    """清理文件名中的非法字符"""
    name = str(name or '').strip()
    name = re.sub(r'[\\/:*?"<>|\r\n\t]', '_', name)
    return name or '未命名'


def php_json_dumps(data) -> str:
    """按数据库现有格式序列化JSON（unicode转义 + 斜杠转义，与PHP json_encode一致）"""
    text = json.dumps(data, ensure_ascii=True, separators=(',', ':'))
    return text.replace('/', '\\/')


class ImageMissingFillWorker(QThread):
    """图片回填工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, datasource, manifest_path, uploads_output_dir, update_db):
        super().__init__()
        self.datasource = datasource
        self.manifest_path = manifest_path
        self.uploads_output_dir = uploads_output_dir
        self.update_db = update_db
        self.results = {
            'total_missing': 0,
            'found_files': 0,
            'copied_files': 0,
            'db_updated': 0,
            'not_provided': 0,
            'failed': 0,
            'records': []
        }

    def find_provided_file(self, folder: str):
        """在缺失文件夹中查找业务放入的图片文件"""
        if not os.path.isdir(folder):
            return None
        for name in sorted(os.listdir(folder)):
            path = os.path.join(folder, name)
            if os.path.isfile(path) and name.lower() not in IGNORED_FILES:
                return path
        return None

    def build_new_name(self, field_comment: str, ext: str, used_names: set) -> tuple:
        """按数据库规则生成新文件名和url: <备注>_<YYYYMMDD>_<unix秒>.<ext>"""
        now = datetime.now()
        date_str = now.strftime('%Y%m%d')
        ts = int(time.time())
        base = sanitize_name(field_comment)
        while True:
            new_name = f"{base}_{date_str}_{ts}{ext}"
            if new_name not in used_names:
                used_names.add(new_name)
                break
            ts += 1
        url = f"/uploads/{now.strftime('%Y')}/{now.strftime('%m')}/{now.strftime('%d')}/{new_name}"
        return new_name, url

    def run(self):
        connection = None
        try:
            with open(self.manifest_path, 'r', encoding='utf-8') as f:
                manifest = json.load(f)

            table = manifest['table']
            id_field = manifest.get('id_field', 'id')
            missing = manifest.get('missing', [])
            root_dir = os.path.dirname(os.path.abspath(self.manifest_path))
            self.results['total_missing'] = len(missing)
            self.log_message.emit(f"清单加载成功: 表 {table}, 共 {len(missing)} 个缺失文件")
            self.progress.emit(5)

            cursor = None
            if self.update_db:
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

            used_names = set()
            total = len(missing) if missing else 1
            for i, record in enumerate(missing):
                row_id = record.get('row_id')
                field_name = record.get('field_name')
                field_comment = record.get('field_comment') or field_name
                index = int(record.get('index', 0))
                folder_rel = record.get('folder', '')
                folder = os.path.join(root_dir, folder_rel)

                result_row = {
                    '行ID': row_id,
                    '法人': record.get('legal_name', ''),
                    '字段名': field_name,
                    '字段备注': field_comment,
                    '数组位置': index + 1,
                    '文件夹': folder_rel,
                    '找到的文件': '',
                    '新文件名': '',
                    '新URL': '',
                    '处理结果': '',
                }

                try:
                    provided = self.find_provided_file(folder)
                    if not provided:
                        self.results['not_provided'] += 1
                        result_row['处理结果'] = '未提供图片'
                        self.results['records'].append(result_row)
                        continue

                    self.results['found_files'] += 1
                    result_row['找到的文件'] = os.path.basename(provided)

                    ext = os.path.splitext(provided)[1].lower() or '.jpg'
                    new_name, url = self.build_new_name(field_comment, ext, used_names)
                    result_row['新文件名'] = new_name
                    result_row['新URL'] = url

                    # 复制到 uploads/年/月/日/ 目录
                    dest_path = os.path.join(
                        self.uploads_output_dir, url.lstrip('/').replace('/', os.sep))
                    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
                    shutil.copy2(provided, dest_path)
                    self.results['copied_files'] += 1

                    if self.update_db and cursor is not None:
                        updated = self.update_db_record(
                            cursor, table, id_field, row_id,
                            field_name, index, new_name, url)
                        if updated:
                            self.results['db_updated'] += 1
                            result_row['处理结果'] = '已复制并更新数据库'
                        else:
                            result_row['处理结果'] = '已复制，数据库更新失败'
                            self.results['failed'] += 1
                    else:
                        result_row['处理结果'] = '已复制（未更新数据库）'

                    self.log_message.emit(
                        f"[回填] id={row_id} {field_comment} 第{index + 1}个: "
                        f"{os.path.basename(provided)} -> {url}")
                except Exception as e:
                    self.results['failed'] += 1
                    result_row['处理结果'] = f'失败: {str(e)}'
                    self.log_message.emit(f"[失败] id={row_id} {field_name}: {str(e)}")

                self.results['records'].append(result_row)
                self.progress.emit(5 + int((i + 1) / total * 90))

            if self.update_db and connection is not None:
                connection.commit()
                cursor.close()
                connection.close()
                connection = None

            # 导出回填结果Excel
            if self.results['records']:
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                excel_path = os.path.join(root_dir, f'回填结果_{timestamp}.xlsx')
                pd.DataFrame(self.results['records']).to_excel(
                    excel_path, index=False, engine='openpyxl')
                self.results['report_path'] = excel_path
                self.log_message.emit(f"回填结果已导出: {excel_path}")

            self.progress.emit(100)
            self.finished.emit(self.results)

        except Exception as e:
            if connection is not None:
                try:
                    connection.rollback()
                    connection.close()
                except Exception:
                    pass
            self.error.emit(str(e))

    def update_db_record(self, cursor, table, id_field, row_id,
                         field_name, index, new_name, url) -> bool:
        """按行ID/字段/数组位置精确更新JSON字段"""
        cursor.execute(
            f"SELECT `{field_name}` FROM `{table}` WHERE `{id_field}` = %s",
            (row_id,))
        row = cursor.fetchone()
        if not row:
            self.log_message.emit(f"[警告] id={row_id} 记录不存在")
            return False

        raw = row.get(field_name)
        entries = []
        if raw is not None and str(raw).strip() not in ('', 'null', 'NULL'):
            try:
                data = json.loads(str(raw))
                if isinstance(data, dict):
                    data = [data]
                if isinstance(data, list):
                    entries = data
            except Exception:
                self.log_message.emit(
                    f"[警告] id={row_id} 字段 {field_name} JSON解析失败，将重建该位置")
                entries = []

        while len(entries) <= index:
            entries.append({})

        entry = entries[index] if isinstance(entries[index], dict) else {}
        entry['name'] = new_name
        entry['url'] = url
        entry['uid'] = entry.get('uid') or str(int(time.time() * 1000))
        entry['status'] = 'success'
        entries[index] = entry

        new_value = php_json_dumps(entries)
        cursor.execute(
            f"UPDATE `{table}` SET `{field_name}` = %s WHERE `{id_field}` = %s",
            (new_value, row_id))
        return True


class ImageMissingFillWidget(QWidget):
    """图片回填界面"""

    def __init__(self, db_manager):
        super().__init__()
        self.db_manager = db_manager
        self.worker = None
        self.init_ui()
        self.load_datasources()

    def init_ui(self):
        layout = QVBoxLayout()

        info_label = QLabel(
            "使用说明: 先用「图片丢失检查」生成缺失文件夹，业务把补充的图片放进对应文件夹"
            "（每个文件夹放一张，文件名随意），然后在这里选择 缺失清单.json，"
            "工具会按规则重命名并生成 uploads/年/月/日/ 目录（上传到服务器即可），"
            "并可选择同时更新数据库JSON字段。")
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: #666;")
        layout.addWidget(info_label)

        config_group = QGroupBox("回填配置")
        form_layout = QFormLayout()

        self.datasource_combo = QComboBox()
        form_layout.addRow("数据源:", self.datasource_combo)

        manifest_layout = QHBoxLayout()
        self.manifest_edit = QLineEdit()
        self.manifest_edit.setReadOnly(True)
        self.select_manifest_btn = QPushButton("选择缺失清单.json")
        self.select_manifest_btn.clicked.connect(self.select_manifest)
        manifest_layout.addWidget(self.manifest_edit)
        manifest_layout.addWidget(self.select_manifest_btn)
        form_layout.addRow("缺失清单:", manifest_layout)

        output_layout = QHBoxLayout()
        self.output_dir_edit = QLineEdit()
        self.output_dir_edit.setReadOnly(True)
        self.select_output_btn = QPushButton("选择输出目录")
        self.select_output_btn.clicked.connect(self.select_output_dir)
        output_layout.addWidget(self.output_dir_edit)
        output_layout.addWidget(self.select_output_btn)
        form_layout.addRow("uploads输出目录:", output_layout)

        self.update_db_checkbox = QCheckBox("同时更新数据库JSON字段（不勾选则只生成uploads目录）")
        self.update_db_checkbox.setChecked(True)
        form_layout.addRow("", self.update_db_checkbox)

        config_group.setLayout(form_layout)
        layout.addWidget(config_group)

        btn_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始回填")
        self.start_btn.clicked.connect(self.start_fill)
        btn_layout.addWidget(self.start_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(120)
        layout.addWidget(self.log_text)

        result_group = QGroupBox("回填结果")
        result_layout = QVBoxLayout()
        self.result_table = QTableWidget()
        self.result_table.setColumnCount(8)
        self.result_table.setHorizontalHeaderLabels([
            "行ID", "法人", "字段备注", "位置", "找到的文件", "新文件名", "新URL", "处理结果"
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

    def select_manifest(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择缺失清单.json", "", "JSON文件 (*.json)")
        if file_path:
            self.manifest_edit.setText(file_path)

    def select_output_dir(self):
        dir_path = QFileDialog.getExistingDirectory(self, "选择uploads输出目录")
        if dir_path:
            self.output_dir_edit.setText(dir_path)

    def start_fill(self):
        datasource = self.datasource_combo.currentData()
        manifest_path = self.manifest_edit.text().strip()
        output_dir = self.output_dir_edit.text().strip()
        update_db = self.update_db_checkbox.isChecked()

        if not manifest_path:
            QMessageBox.warning(self, "提示", "请选择缺失清单.json")
            return
        if not output_dir:
            QMessageBox.warning(self, "提示", "请选择uploads输出目录")
            return
        if update_db and not datasource:
            QMessageBox.warning(self, "提示", "更新数据库需要先选择数据源")
            return

        if update_db:
            reply = QMessageBox.question(
                self, "确认",
                "将会直接更新数据库JSON字段（按行ID/字段/数组位置），确认继续吗？",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply != QMessageBox.Yes:
                return

        self.result_table.setRowCount(0)
        self.log_text.clear()
        self.progress_bar.setValue(0)
        self.start_btn.setEnabled(False)

        self.worker = ImageMissingFillWorker(
            datasource, manifest_path, output_dir, update_db)
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log_message.connect(self.log)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def log(self, message):
        self.log_text.append(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")

    def on_finished(self, results):
        self.start_btn.setEnabled(True)

        for record in results['records']:
            row = self.result_table.rowCount()
            self.result_table.insertRow(row)
            self.result_table.setItem(row, 0, QTableWidgetItem(str(record['行ID'])))
            self.result_table.setItem(row, 1, QTableWidgetItem(str(record['法人'])))
            self.result_table.setItem(row, 2, QTableWidgetItem(str(record['字段备注'])))
            self.result_table.setItem(row, 3, QTableWidgetItem(str(record['数组位置'])))
            self.result_table.setItem(row, 4, QTableWidgetItem(str(record['找到的文件'])))
            self.result_table.setItem(row, 5, QTableWidgetItem(str(record['新文件名'])))
            self.result_table.setItem(row, 6, QTableWidgetItem(str(record['新URL'])))
            self.result_table.setItem(row, 7, QTableWidgetItem(str(record['处理结果'])))

        summary = (f"回填完成: 缺失 {results['total_missing']} 个, "
                   f"找到图片 {results['found_files']} 个, "
                   f"复制 {results['copied_files']} 个, "
                   f"数据库更新 {results['db_updated']} 条, "
                   f"未提供 {results['not_provided']} 个, "
                   f"失败 {results['failed']} 个")
        self.log(summary)
        report_path = results.get('report_path')
        msg = summary + (f"\n\n结果报告:\n{report_path}" if report_path else "")
        QMessageBox.information(self, "完成", msg)

    def on_error(self, message):
        self.start_btn.setEnabled(True)
        self.log(f"回填失败: {message}")
        QMessageBox.critical(self, "错误", f"回填失败: {message}")
