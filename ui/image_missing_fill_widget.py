"""
图片回填功能界面

配合「图片丢失检查」使用：业务把补充的图片放进检查生成的
缺失_<表名>_<时间戳>/<法人>/<字段备注>/<序号_原文件名>/ 文件夹后，
本功能读取 缺失清单.json（或 缺失清单.xlsx），扫描各文件夹中的图片，
不管业务给的文件名和类型是什么，一律按数据库里原有的 url/文件名
（含扩展名）重命名，严格按原URL的路径生成目录结构（uploads/年/月/日
或 storage/default/日期 等任意结构都支持），直接上传到服务器即可，
数据库数据完全不动。
"""
import os
import json
import shutil
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog,
                             QLineEdit,
                             QTableWidget, QTableWidgetItem, QHeaderView)
from PyQt5.QtCore import QThread, pyqtSignal
import pandas as pd


IGNORED_FILES = {'说明.txt', 'thumbs.db', 'desktop.ini', '.ds_store'}

# 缺失清单.xlsx 列名 → 清单字段名（与图片丢失检查导出的Excel列对应）
EXCEL_COLUMN_MAP = {
    '行ID': 'row_id',
    '部门': 'dept',
    '法人': 'legal_name',
    '字段名': 'field_name',
    '字段备注': 'field_comment',
    '数组位置(从0开始)': 'index',
    '原文件名': 'file_name',
    '原URL': 'url',
    'uid': 'uid',
    '对应文件夹': 'folder',
}


def load_manifest(manifest_path: str) -> dict:
    """加载缺失清单，支持 json 和 Excel（图片丢失检查导出的 缺失清单.xlsx）"""
    if manifest_path.lower().endswith(('.xlsx', '.xls')):
        df = pd.read_excel(manifest_path, dtype=str)
        df = df.rename(columns=EXCEL_COLUMN_MAP)
        if 'url' not in df.columns:
            raise ValueError("Excel中找不到「原URL」列，请选择图片丢失检查导出的 缺失清单.xlsx")
        missing = []
        for _, row in df.iterrows():
            record = {}
            for key in ('row_id', 'dept', 'legal_name', 'field_name',
                        'field_comment', 'file_name', 'url', 'uid', 'folder'):
                value = row.get(key)
                record[key] = '' if value is None or pd.isna(value) else str(value)
            index_value = row.get('index')
            try:
                record['index'] = int(float(index_value))
            except (TypeError, ValueError):
                record['index'] = 0
            missing.append(record)
        return {'table': '', 'missing': missing}
    with open(manifest_path, 'r', encoding='utf-8') as f:
        return json.load(f)


class ImageMissingFillWorker(QThread):
    """图片回填工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, manifest_path, uploads_output_dir):
        super().__init__()
        self.manifest_path = manifest_path
        self.uploads_output_dir = uploads_output_dir
        self.results = {
            'total_missing': 0,
            'found_files': 0,
            'copied_files': 0,
            'ext_mismatch': 0,
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

    def run(self):
        try:
            manifest = load_manifest(self.manifest_path)

            table = manifest.get('table', '')
            missing = manifest.get('missing', [])
            root_dir = os.path.dirname(os.path.abspath(self.manifest_path))
            self.results['total_missing'] = len(missing)
            self.log_message.emit(f"清单加载成功: 表 {table}, 共 {len(missing)} 个缺失文件")
            self.progress.emit(5)

            filled_records = []
            remaining_records = []

            total = len(missing) if missing else 1
            for i, record in enumerate(missing):
                row_id = record.get('row_id')
                field_comment = record.get('field_comment') or record.get('field_name')
                index = int(record.get('index', 0))
                url = str(record.get('url', ''))
                folder_rel = record.get('folder', '')
                folder = os.path.join(root_dir, folder_rel)

                result_row = {
                    '行ID': row_id,
                    '法人': record.get('legal_name', ''),
                    '字段备注': field_comment,
                    '数组位置': index + 1,
                    '文件夹': folder_rel,
                    '找到的文件': '',
                    '数据库文件名': os.path.basename(url),
                    '生成路径': '',
                    '处理结果': '',
                }

                try:
                    if not url:
                        self.results['failed'] += 1
                        result_row['处理结果'] = '清单中缺少url，无法生成'
                        self.results['records'].append(result_row)
                        remaining_records.append(record)
                        continue

                    if not folder_rel:
                        self.results['failed'] += 1
                        result_row['处理结果'] = '清单中无对应文件夹（生成时建目录失败）'
                        self.results['records'].append(result_row)
                        remaining_records.append(record)
                        continue

                    provided = self.find_provided_file(folder)
                    if not provided:
                        self.results['not_provided'] += 1
                        result_row['处理结果'] = '未提供图片'
                        self.results['records'].append(result_row)
                        remaining_records.append(record)
                        continue

                    self.results['found_files'] += 1
                    result_row['找到的文件'] = os.path.basename(provided)

                    # 严格按照数据库里原有的url路径和文件名生成，数据库不动
                    dest_path = os.path.join(
                        self.uploads_output_dir, url.lstrip('/\\').replace('/', os.sep))
                    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
                    shutil.copy2(provided, dest_path)
                    self.results['copied_files'] += 1
                    result_row['生成路径'] = dest_path

                    provided_ext = os.path.splitext(provided)[1].lower()
                    db_ext = os.path.splitext(url)[1].lower()
                    if provided_ext and db_ext and provided_ext != db_ext:
                        self.results['ext_mismatch'] += 1
                        result_row['处理结果'] = f'已生成（注意: 业务给的是{provided_ext}，已按数据库改成{db_ext}）'
                        self.log_message.emit(
                            f"[类型不一致] id={row_id} {field_comment} 第{index + 1}个: "
                            f"业务给的{provided_ext}，已重命名为数据库的{db_ext}")
                    else:
                        result_row['处理结果'] = '已生成'

                    filled_records.append(record)

                    self.log_message.emit(
                        f"[回填] id={row_id} {field_comment} 第{index + 1}个: "
                        f"{os.path.basename(provided)} -> {url}")
                except Exception as e:
                    self.results['failed'] += 1
                    result_row['处理结果'] = f'失败: {str(e)}'
                    self.log_message.emit(f"[失败] id={row_id} {field_comment}: {str(e)}")
                    remaining_records.append(record)

                self.results['records'].append(result_row)
                self.progress.emit(5 + int((i + 1) / total * 90))

            # 导出回填结果Excel
            if self.results['records']:
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                excel_path = os.path.join(root_dir, f'回填结果_{timestamp}.xlsx')
                pd.DataFrame(self.results['records']).to_excel(
                    excel_path, index=False, engine='openpyxl')
                self.results['report_path'] = excel_path
                self.log_message.emit(f"回填结果已导出: {excel_path}")

                # 归档：已回填清单 + 剩余缺失清单（下次直接选剩余清单继续跑）
                if filled_records:
                    filled_manifest = dict(manifest)
                    filled_manifest['missing'] = filled_records
                    filled_json = os.path.join(root_dir, f'已回填清单_{timestamp}.json')
                    with open(filled_json, 'w', encoding='utf-8') as f:
                        json.dump(filled_manifest, f, ensure_ascii=False, indent=2)
                    filled_xlsx = os.path.join(root_dir, f'已回填清单_{timestamp}.xlsx')
                    pd.DataFrame(filled_records).to_excel(
                        filled_xlsx, index=False, engine='openpyxl')
                    self.log_message.emit(
                        f"本次已回填 {len(filled_records)} 条，已导出: {filled_json} / {filled_xlsx}")

                remaining_manifest = dict(manifest)
                remaining_manifest['missing'] = remaining_records
                remaining_json = os.path.join(root_dir, f'剩余缺失清单_{timestamp}.json')
                with open(remaining_json, 'w', encoding='utf-8') as f:
                    json.dump(remaining_manifest, f, ensure_ascii=False, indent=2)
                remaining_xlsx = os.path.join(root_dir, f'剩余缺失清单_{timestamp}.xlsx')
                pd.DataFrame(remaining_records).to_excel(
                    remaining_xlsx, index=False, engine='openpyxl')
                self.log_message.emit(
                    f"剩余未回填 {len(remaining_records)} 条，下次回填直接选: {remaining_json}")

            self.progress.emit(100)
            self.finished.emit(self.results)

        except Exception as e:
            self.error.emit(str(e))


class ImageMissingFillWidget(QWidget):
    """图片回填界面"""

    def __init__(self, db_manager=None):
        super().__init__()
        self.worker = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        info_label = QLabel(
            "使用说明: 先用「图片丢失检查」生成缺失文件夹，业务把补充的图片放进对应文件夹"
            "（每个文件夹放一张，文件名和类型随意），然后在这里选择 缺失清单.json 或 缺失清单.xlsx"
            "（需和缺失文件夹在同一目录），工具会按数据库里原有的url和文件名重命名，"
            "严格按原URL路径生成目录结构（uploads/年/月/日 或 storage/default/日期 等都支持），"
            "直接上传到服务器即可，数据库数据不会被修改。")
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: #666;")
        layout.addWidget(info_label)

        config_group = QGroupBox("回填配置")
        form_layout = QFormLayout()

        manifest_layout = QHBoxLayout()
        self.manifest_edit = QLineEdit()
        self.manifest_edit.setReadOnly(True)
        self.select_manifest_btn = QPushButton("选择缺失清单(json/xlsx)")
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
        form_layout.addRow("输出目录:", output_layout)

        config_group.setLayout(form_layout)
        layout.addWidget(config_group)

        btn_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始生成回填目录")
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
            "行ID", "法人", "字段备注", "位置", "找到的文件", "数据库文件名", "生成路径", "处理结果"
        ])
        self.result_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.result_table.horizontalHeader().setStretchLastSection(True)
        self.result_table.setEditTriggers(QTableWidget.NoEditTriggers)
        result_layout.addWidget(self.result_table)
        result_group.setLayout(result_layout)
        layout.addWidget(result_group)

        self.setLayout(layout)

    def select_manifest(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择缺失清单", "", "缺失清单 (*.json *.xlsx *.xls)")
        if file_path:
            self.manifest_edit.setText(file_path)

    def select_output_dir(self):
        dir_path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if dir_path:
            self.output_dir_edit.setText(dir_path)

    def start_fill(self):
        manifest_path = self.manifest_edit.text().strip()
        output_dir = self.output_dir_edit.text().strip()

        if not manifest_path:
            QMessageBox.warning(self, "提示", "请选择缺失清单(json或xlsx)")
            return
        if not output_dir:
            QMessageBox.warning(self, "提示", "请选择输出目录")
            return

        self.result_table.setRowCount(0)
        self.log_text.clear()
        self.progress_bar.setValue(0)
        self.start_btn.setEnabled(False)

        self.worker = ImageMissingFillWorker(manifest_path, output_dir)
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
            self.result_table.setItem(row, 5, QTableWidgetItem(str(record['数据库文件名'])))
            self.result_table.setItem(row, 6, QTableWidgetItem(str(record['生成路径'])))
            self.result_table.setItem(row, 7, QTableWidgetItem(str(record['处理结果'])))

        summary = (f"生成完成: 缺失 {results['total_missing']} 个, "
                   f"找到图片 {results['found_files']} 个, "
                   f"生成 {results['copied_files']} 个, "
                   f"类型不一致 {results['ext_mismatch']} 个, "
                   f"未提供 {results['not_provided']} 个, "
                   f"失败 {results['failed']} 个")
        self.log(summary)
        report_path = results.get('report_path')
        msg = summary + (f"\n\n结果报告:\n{report_path}" if report_path else "")
        QMessageBox.information(self, "完成", msg)

    def on_error(self, message):
        self.start_btn.setEnabled(True)
        self.log(f"生成失败: {message}")
        QMessageBox.critical(self, "错误", f"生成失败: {message}")
