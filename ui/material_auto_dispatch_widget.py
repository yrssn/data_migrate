"""
资料自动分发功能界面

配合「图片丢失检查」/「图片回填」使用：业务把补充资料按
法人一个文件夹发回来（文件夹名、文件名不一定和数据库一致），
本功能读取 缺失清单.json，先按法人目录名匹配业务文件夹，
再按文件名与 字段备注/原文件名 的相似度，把每个文件自动
放进缺失目录里对应的 <序号_原文件名>/ 文件夹。
默认预览模式只生成匹配报告不复制，确认后再执行复制。
"""
import os
import json
import re
import shutil
import difflib
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog,
                             QLineEdit, QCheckBox, QDoubleSpinBox, QComboBox,
                             QTableWidget, QTableWidgetItem, QHeaderView)
from PyQt5.QtCore import QThread, pyqtSignal
import pandas as pd


IGNORED_FILES = {'说明.txt', 'thumbs.db', 'desktop.ini', '.ds_store'}


def normalize(text: str) -> str:
    """归一化用于匹配的文本：去空格、标点、大小写"""
    text = str(text or '').lower()
    return re.sub(r'[\s\-_—·．.,，、()（）\[\]【】]+', '', text)


class MaterialAutoDispatchWorker(QThread):
    """资料自动分发工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, manifest_path, material_dir, threshold, preview,
                 dept_filter=''):
        super().__init__()
        self.manifest_path = manifest_path
        self.material_dir = material_dir
        self.threshold = threshold
        self.preview = preview
        self.dept_filter = dept_filter
        self.results = {
            'total_missing': 0,
            'matched': 0,
            'copied': 0,
            'unmatched_records': 0,
            'unmatched_files': 0,
            'records': []
        }

    def strip_parts(self, text: str, parts) -> str:
        """从文本中剥掉法人/公司名等片段，留下资料类型关键词"""
        norm = normalize(text)
        for part in sorted((normalize(p) for p in parts), key=len, reverse=True):
            if part:
                norm = norm.replace(part, '')
        return norm

    def similarity(self, a: str, b: str) -> float:
        a, b = normalize(a), normalize(b)
        if not a or not b:
            return 0.0
        if a in b or b in a:
            return 1.0
        return difflib.SequenceMatcher(None, a, b).ratio()

    def collect_files(self, folder: str):
        """递归收集业务文件夹下所有文件"""
        files = []
        for root, _, names in os.walk(folder):
            for name in names:
                if name.lower() in IGNORED_FILES:
                    continue
                files.append(os.path.join(root, name))
        return files

    def run(self):
        try:
            with open(self.manifest_path, 'r', encoding='utf-8') as f:
                manifest = json.load(f)
            missing = manifest.get('missing', [])
            root_dir = os.path.dirname(os.path.abspath(self.manifest_path))
            self.results['total_missing'] = len(missing)
            self.log_message.emit(f"清单加载成功，共 {len(missing)} 个缺失文件")
            if self.dept_filter:
                missing = [r for r in missing
                           if str(r.get('dept', '')) == self.dept_filter]
                self.log_message.emit(
                    f"只处理部门「{self.dept_filter}」，共 {len(missing)} 个缺失文件")
            self.progress.emit(5)

            # 按 部门+法人 分组缺失记录（业务文件夹和法人是一对一的，
            # 不能跨部门复用）
            legal_records = {}
            for record in missing:
                key = (str(record.get('dept', '')), str(record.get('legal_name', '')))
                legal_records.setdefault(key, []).append(record)

            # 业务资料文件夹（任意层级都参与法人匹配）
            material_folders = []
            for walk_root, walk_dirs, _ in os.walk(self.material_dir):
                for d in sorted(walk_dirs):
                    material_folders.append(os.path.relpath(
                        os.path.join(walk_root, d), self.material_dir))
            self.log_message.emit(f"业务资料目录下共 {len(material_folders)} 个文件夹（含子目录）")

            # 法人 ↔ 业务文件夹匹配：只认包含关系（法人名或公司名整段出现在
            # 文件夹名里，或文件夹名整段出现在法人名里），不做模糊比对，
            # 避免公司名里 TRADING 之类的通用词导致乱匹配；并全局唯一分配，
            # 一个文件夹树只归一个法人
            candidates = []
            for group_key in legal_records:
                legal_name = group_key[1]
                name_parts = [legal_name] + [
                    p for p in re.split(r'[_／/\\]+', legal_name) if p.strip()]
                for folder in material_folders:
                    base = normalize(os.path.basename(folder))
                    score = 0.0
                    for part in name_parts:
                        p = normalize(part)
                        if len(p) >= 2 and len(base) >= 2 and (p in base or base in p):
                            # 重合的部分越长越优先
                            score = max(score, min(len(p), len(base)))
                    if score > 0:
                        # 同分时浅层文件夹优先（文件收集是递归的，浅层能把嵌套的都包进来）
                        candidates.append((-score, folder.count(os.sep), len(folder),
                                           group_key, folder))
            candidates.sort(key=lambda c: (c[0], c[1], c[2]))

            folder_map = {}
            taken_folders = []

            def conflict_owner(folder, group_key):
                fn = folder + os.sep
                for taken, owner in taken_folders:
                    if owner == group_key:
                        continue
                    tn = taken + os.sep
                    if fn.startswith(tn) or tn.startswith(fn):
                        return owner
                return None

            for neg_score, _, _, group_key, folder in candidates:
                if group_key in folder_map:
                    continue
                owner = conflict_owner(folder, group_key)
                if owner is not None:
                    self.log_message.emit(
                        f"[跳过] {group_key[0]}/{group_key[1]} 候选文件夹 {folder} "
                        f"已被 {owner[0]}/{owner[1]} 占用")
                    continue
                folder_map[group_key] = folder
                taken_folders.append((folder, group_key))
                self.log_message.emit(
                    f"[法人匹配] {group_key[0]}/{group_key[1]} <-> {folder} "
                    f"(名称重合长度 {int(-neg_score)})")

            for group_key in legal_records:
                if group_key not in folder_map:
                    self.log_message.emit(
                        f"[法人未匹配] {group_key[0]}/{group_key[1]} 找不到对应业务文件夹")

            total = len(missing) if missing else 1
            done = 0
            for group_key, records in legal_records.items():
                legal_name = group_key[1]
                name_parts = [p for p in re.split(r'[_／/\\]+', legal_name) if p.strip()]
                folder = folder_map.get(group_key)
                # 只有法人文件夹确定匹配上才分发，否则一律不动
                files = self.collect_files(
                    os.path.join(self.material_dir, folder)) if folder else []
                used_files = set()

                # 每条缺失记录找最相似文件（唯一分配，按最高分优先）
                pairs = []
                for r_idx, record in enumerate(records):
                    target_text = f"{record.get('field_comment', '')} {record.get('file_name', '')}"
                    for f_idx, path in enumerate(files):
                        base = os.path.splitext(os.path.basename(path))[0]
                        stripped = self.strip_parts(base, name_parts)
                        score = max(
                            self.similarity(base, record.get('file_name', '')),
                            self.similarity(base, record.get('field_comment', '')),
                            self.similarity(base, target_text),
                            self.similarity(stripped, record.get('field_comment', '')),
                            self.similarity(
                                stripped,
                                self.strip_parts(record.get('file_name', ''), name_parts)))
                        pairs.append((score, r_idx, f_idx))
                pairs.sort(reverse=True)

                assigned = {}
                for score, r_idx, f_idx in pairs:
                    if score < self.threshold:
                        break
                    if r_idx in assigned or f_idx in used_files:
                        continue
                    assigned[r_idx] = (f_idx, score)
                    used_files.add(f_idx)

                for r_idx, record in enumerate(records):
                    done += 1
                    result_row = {
                        '行ID': record.get('row_id'),
                        '部门': record.get('dept', ''),
                        '法人': legal_name,
                        '字段备注': record.get('field_comment', ''),
                        '位置': int(record.get('index', 0)) + 1,
                        '原文件名': record.get('file_name', ''),
                        '业务文件夹': folder or '',
                        '匹配文件': '',
                        '相似度': '',
                        '结果': '',
                    }
                    if folder is None:
                        result_row['结果'] = '法人未匹配到业务文件夹'
                        self.results['unmatched_records'] += 1
                    elif r_idx not in assigned:
                        result_row['结果'] = '未找到相似文件'
                        self.results['unmatched_records'] += 1
                    else:
                        f_idx, score = assigned[r_idx]
                        src = files[f_idx]
                        result_row['匹配文件'] = os.path.relpath(src, self.material_dir)
                        result_row['相似度'] = f"{score:.2f}"
                        self.results['matched'] += 1
                        dest_folder = os.path.join(root_dir, record.get('folder', ''))
                        if self.preview:
                            result_row['结果'] = '预览: 待复制'
                        else:
                            try:
                                os.makedirs(dest_folder, exist_ok=True)
                                # 一个说明.txt只对应一个文件：先删掉旧文件
                                for old in os.listdir(dest_folder):
                                    if old.lower() in IGNORED_FILES:
                                        continue
                                    old_path = os.path.join(dest_folder, old)
                                    if os.path.isfile(old_path):
                                        os.remove(old_path)
                                shutil.copy2(src, os.path.join(
                                    dest_folder, os.path.basename(src)))
                                self.results['copied'] += 1
                                result_row['结果'] = '已复制'
                            except Exception as e:
                                result_row['结果'] = f'复制失败: {str(e)}'
                        self.log_message.emit(
                            f"[匹配] {legal_name} {record.get('field_comment', '')} "
                            f"第{int(record.get('index', 0)) + 1}个 <- "
                            f"{os.path.basename(src)} (相似度 {score:.2f}) "
                            f"{result_row['结果']}")

                    self.results['records'].append(result_row)
                    self.progress.emit(5 + int(done / total * 90))

                # 该法人剩余未用到的文件
                if folder:
                    for f_idx, path in enumerate(files):
                        if f_idx not in used_files:
                            self.results['unmatched_files'] += 1
                            self.results['records'].append({
                                '行ID': '', '部门': '', '法人': legal_name,
                                '字段备注': '', '位置': '', '原文件名': '',
                                '业务文件夹': folder,
                                '匹配文件': os.path.relpath(path, self.material_dir),
                                '相似度': '',
                                '结果': '业务文件未被使用',
                            })

            # 导出报告
            if self.results['records']:
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                mode = '预览' if self.preview else '执行'
                excel_path = os.path.join(root_dir, f'自动分发{mode}报告_{timestamp}.xlsx')
                pd.DataFrame(self.results['records']).to_excel(
                    excel_path, index=False, engine='openpyxl')
                self.results['report_path'] = excel_path
                self.log_message.emit(f"报告已导出: {excel_path}")

            self.progress.emit(100)
            self.finished.emit(self.results)

        except Exception as e:
            self.error.emit(str(e))


class MaterialAutoDispatchWidget(QWidget):
    """资料自动分发界面"""

    def __init__(self, db_manager=None):
        super().__init__()
        self.worker = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        info_label = QLabel(
            "使用说明: 业务按法人一个文件夹发回资料后，选择 缺失清单.json 和业务资料根目录，"
            "工具按法人文件夹名和文件名相似度自动把文件分发到缺失目录对应的位置。"
            "默认预览模式只出报告不复制，确认无误后取消勾选再执行；"
            "匹配不上的在报告里列出，手动放入即可。分发完成后再用「图片回填」生成uploads目录。")
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: #666;")
        layout.addWidget(info_label)

        config_group = QGroupBox("分发配置")
        form_layout = QFormLayout()

        manifest_layout = QHBoxLayout()
        self.manifest_edit = QLineEdit()
        self.manifest_edit.setReadOnly(True)
        self.select_manifest_btn = QPushButton("选择缺失清单.json")
        self.select_manifest_btn.clicked.connect(self.select_manifest)
        manifest_layout.addWidget(self.manifest_edit)
        manifest_layout.addWidget(self.select_manifest_btn)
        form_layout.addRow("缺失清单:", manifest_layout)

        material_layout = QHBoxLayout()
        self.material_edit = QLineEdit()
        self.material_edit.setReadOnly(True)
        self.select_material_btn = QPushButton("选择业务资料根目录")
        self.select_material_btn.clicked.connect(self.select_material_dir)
        material_layout.addWidget(self.material_edit)
        material_layout.addWidget(self.select_material_btn)
        form_layout.addRow("业务资料目录:", material_layout)

        self.dept_combo = QComboBox()
        self.dept_combo.addItem("全部部门")
        form_layout.addRow("只处理部门:", self.dept_combo)

        self.threshold_spin = QDoubleSpinBox()
        self.threshold_spin.setRange(0.1, 1.0)
        self.threshold_spin.setSingleStep(0.05)
        self.threshold_spin.setValue(0.6)
        form_layout.addRow("相似度阈值:", self.threshold_spin)

        self.preview_checkbox = QCheckBox("预览模式（只生成匹配报告，不复制文件）")
        self.preview_checkbox.setChecked(True)
        form_layout.addRow("", self.preview_checkbox)

        config_group.setLayout(form_layout)
        layout.addWidget(config_group)

        btn_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始分发")
        self.start_btn.clicked.connect(self.start_dispatch)
        btn_layout.addWidget(self.start_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(120)
        layout.addWidget(self.log_text)

        result_group = QGroupBox("分发结果")
        result_layout = QVBoxLayout()
        self.result_table = QTableWidget()
        self.result_table.setColumnCount(9)
        self.result_table.setHorizontalHeaderLabels([
            "行ID", "法人", "字段备注", "位置", "原文件名",
            "业务文件夹", "匹配文件", "相似度", "结果"
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
            self, "选择缺失清单.json", "", "JSON文件 (*.json)")
        if not file_path:
            return
        self.manifest_edit.setText(file_path)
        self.dept_combo.clear()
        self.dept_combo.addItem("全部部门")
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                manifest = json.load(f)
            depts = sorted({str(r.get('dept', '')) for r in manifest.get('missing', [])
                            if str(r.get('dept', '')).strip()})
            for dept in depts:
                self.dept_combo.addItem(dept)
        except Exception as e:
            QMessageBox.warning(self, "提示", f"读取清单失败: {str(e)}")

    def select_material_dir(self):
        dir_path = QFileDialog.getExistingDirectory(self, "选择业务资料根目录")
        if dir_path:
            self.material_edit.setText(dir_path)

    def start_dispatch(self):
        manifest_path = self.manifest_edit.text().strip()
        material_dir = self.material_edit.text().strip()
        if not manifest_path:
            QMessageBox.warning(self, "提示", "请选择缺失清单.json")
            return
        if not material_dir:
            QMessageBox.warning(self, "提示", "请选择业务资料根目录")
            return

        preview = self.preview_checkbox.isChecked()
        if not preview:
            reply = QMessageBox.question(
                self, "确认",
                "将把匹配到的业务文件复制进缺失目录对应的文件夹，确认执行吗？\n"
                "（建议先用预览模式检查匹配报告）",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply != QMessageBox.Yes:
                return

        self.result_table.setRowCount(0)
        self.log_text.clear()
        self.progress_bar.setValue(0)
        self.start_btn.setEnabled(False)

        dept_filter = ''
        if self.dept_combo.currentIndex() > 0:
            dept_filter = self.dept_combo.currentText()

        self.worker = MaterialAutoDispatchWorker(
            manifest_path, material_dir, self.threshold_spin.value(), preview,
            dept_filter)
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
            self.result_table.setItem(row, 3, QTableWidgetItem(str(record['位置'])))
            self.result_table.setItem(row, 4, QTableWidgetItem(str(record['原文件名'])))
            self.result_table.setItem(row, 5, QTableWidgetItem(str(record['业务文件夹'])))
            self.result_table.setItem(row, 6, QTableWidgetItem(str(record['匹配文件'])))
            self.result_table.setItem(row, 7, QTableWidgetItem(str(record['相似度'])))
            self.result_table.setItem(row, 8, QTableWidgetItem(str(record['结果'])))

        mode = '预览' if self.preview_checkbox.isChecked() else '执行'
        summary = (f"{mode}完成: 缺失 {results['total_missing']} 个, "
                   f"匹配 {results['matched']} 个, "
                   f"已复制 {results['copied']} 个, "
                   f"未匹配记录 {results['unmatched_records']} 个, "
                   f"未使用业务文件 {results['unmatched_files']} 个")
        self.log(summary)
        report_path = results.get('report_path')
        msg = summary + (f"\n\n报告:\n{report_path}" if report_path else "")
        QMessageBox.information(self, "完成", msg)

    def on_error(self, message):
        self.start_btn.setEnabled(True)
        self.log(f"分发失败: {message}")
        QMessageBox.critical(self, "错误", f"分发失败: {message}")
