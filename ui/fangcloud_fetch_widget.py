"""
云盘资料抓取功能界面

配合「图片丢失检查」/「资料自动分发」使用：读取 缺失清单.json（或
缺失清单.xlsx），用 Playwright 打开亿方云(v2.fangcloud.com)，人工登录后
按 原文件名 逐条走 文件(夹)名称 搜索(qf=file_name)，结果按归一化相似度
匹配（法人名出现在路径里加分），相似度最高且达到阈值的下载到
输出目录/<法人>/<原文件名>，之后直接用「资料自动分发」分发。
登录状态保存在本机浏览器数据目录，下次不用重复登录。
"""
import os
import json
import threading
import urllib.parse
import difflib
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog,
                             QLineEdit, QDoubleSpinBox,
                             QTableWidget, QTableWidgetItem, QHeaderView)
from PyQt5.QtCore import QThread, pyqtSignal
import pandas as pd

from .material_auto_dispatch_widget import normalize
from .image_missing_check_widget import sanitize_name
from .image_missing_fill_widget import load_manifest


SEARCH_URL = "https://v2.fangcloud.com/desktop/search?query_words={query}&qf=file_name"
BROWSER_DATA_DIR = os.path.join(os.path.expanduser('~'), '.fangcloud_browser')


def similarity(a: str, b: str) -> float:
    return difflib.SequenceMatcher(None, normalize(a), normalize(b)).ratio()


class FangcloudFetchWorker(QThread):
    """云盘抓取工作线程（Playwright 必须整个生命周期在同一线程里）"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    login_ready = pyqtSignal()          # 浏览器已打开，等人工登录
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, manifest_path, output_dir, threshold=0.6):
        super().__init__()
        self.manifest_path = manifest_path
        self.output_dir = output_dir
        self.threshold = threshold
        self.login_done_event = threading.Event()
        self.stop_requested = False
        self.results = {
            'total': 0,
            'downloaded': 0,
            'skipped_exists': 0,
            'no_match': 0,
            'failed': 0,
            'records': []
        }

    def confirm_login(self):
        self.login_done_event.set()

    def request_stop(self):
        self.stop_requested = True
        self.login_done_event.set()

    def collect_tasks(self):
        """从缺失清单取 (法人, 原文件名) 去重任务列表"""
        manifest = load_manifest(self.manifest_path)
        tasks = []
        seen = set()
        for record in manifest.get('missing', []):
            legal_name = str(record.get('legal_name', '') or '未知法人')
            file_name = str(record.get('file_name', '') or '')
            if not file_name:
                file_name = os.path.basename(str(record.get('url', '')))
            if not file_name:
                continue
            key = (legal_name, file_name)
            if key in seen:
                continue
            seen.add(key)
            tasks.append({'legal_name': legal_name, 'file_name': file_name,
                          'dept': str(record.get('dept', '') or '')})
        return tasks

    def pick_best_row(self, page, target_name, legal_name):
        """在搜索结果里选相似度最高的一条，返回 (row, 命中文件名, 分数)"""
        rows = page.locator('li.list-item.search-item')
        count = rows.count()
        best = (None, '', 0.0)
        for i in range(min(count, 30)):
            row = rows.nth(i)
            name_el = row.locator('.file-name .name').first
            title = name_el.get_attribute('title') or name_el.inner_text()
            score = similarity(title, target_name)
            # 法人名出现在路径面包屑里加分
            if legal_name:
                crumbs = row.locator('.breadcrumb a')
                crumb_text = ' '.join(
                    crumbs.nth(j).get_attribute('title') or ''
                    for j in range(crumbs.count()))
                if normalize(legal_name) and normalize(legal_name) in normalize(crumb_text):
                    score += 0.15
            if score > best[2]:
                best = (row, title, score)
        return best

    def download_row(self, page, row, dest_path):
        """点结果行的 更多(onDropDown) 菜单里的 下载"""
        row.scroll_into_view_if_needed()
        row.hover()
        more_btn = row.locator('[data-action="onDropDown"]').first
        more_btn.click()
        page.wait_for_timeout(600)
        # 弹出菜单里文字为「下载」的可见项（不依赖菜单容器class）
        download_item = None
        for selector in ('[data-action="download"]',
                         'li:text-is("下载")',
                         'li:has-text("下载")',
                         ':text-is("下载")'):
            candidates = page.locator(selector)
            for i in range(candidates.count()):
                item = candidates.nth(i)
                try:
                    if item.is_visible():
                        download_item = item
                        break
                except Exception:
                    continue
            if download_item is not None:
                break
        if download_item is None:
            raise RuntimeError('弹出菜单里没找到可见的「下载」项')
        with page.expect_download(timeout=60000) as download_info:
            download_item.click(timeout=10000)
        download = download_info.value
        os.makedirs(os.path.dirname(dest_path), exist_ok=True)
        download.save_as(dest_path)

    def run(self):
        try:
            from playwright.sync_api import sync_playwright
        except ImportError:
            self.error.emit("未安装 playwright，请先执行: pip install playwright && playwright install chromium")
            return

        try:
            tasks = self.collect_tasks()
            self.results['total'] = len(tasks)
            self.log_message.emit(f"清单加载成功，共 {len(tasks)} 个待抓取文件（已按法人+文件名去重）")

            with sync_playwright() as p:
                context = p.chromium.launch_persistent_context(
                    BROWSER_DATA_DIR, headless=False,
                    accept_downloads=True,
                    args=['--start-maximized'], no_viewport=True)
                page = context.pages[0] if context.pages else context.new_page()
                page.goto('https://v2.fangcloud.com/desktop', timeout=60000)
                self.log_message.emit("浏览器已打开，请在浏览器里完成登录，然后点「已登录，开始抓取」")
                self.login_ready.emit()
                self.login_done_event.wait()
                if self.stop_requested:
                    context.close()
                    self.finished.emit(self.results)
                    return

                total = len(tasks) if tasks else 1
                for i, task in enumerate(tasks):
                    if self.stop_requested:
                        self.log_message.emit("已手动停止")
                        break
                    legal_name = task['legal_name']
                    file_name = task['file_name']
                    result_row = {
                        '部门': task['dept'],
                        '法人': legal_name,
                        '原文件名': file_name,
                        '命中文件名': '',
                        '相似度': '',
                        '保存路径': '',
                        '结果': '',
                    }
                    dest_path = os.path.join(
                        self.output_dir, sanitize_name(legal_name),
                        sanitize_name(file_name, max_len=150))
                    try:
                        if os.path.isfile(dest_path):
                            self.results['skipped_exists'] += 1
                            result_row['保存路径'] = dest_path
                            result_row['结果'] = '已存在，跳过'
                            self.results['records'].append(result_row)
                            continue

                        url = SEARCH_URL.format(query=urllib.parse.quote(file_name))
                        try:
                            # 亿方云是单页应用，打开搜索URL后会再做一次内部跳转，
                            # goto会报"导航被打断"，忽略之，以结果列表出现为准
                            page.goto(url, timeout=60000, wait_until='commit')
                        except Exception as e:
                            if 'interrupted by another navigation' not in str(e):
                                raise
                        found = False
                        for _ in range(2):
                            try:
                                page.wait_for_selector('li.list-item.search-item',
                                                       timeout=15000)
                                found = True
                                break
                            except Exception:
                                # 可能被内部跳转带走了，重新进一次搜索页
                                try:
                                    page.goto(url, timeout=60000, wait_until='commit')
                                except Exception as e:
                                    if 'interrupted by another navigation' not in str(e):
                                        raise
                        if not found:
                            self.results['no_match'] += 1
                            result_row['结果'] = '搜索无结果'
                            self.results['records'].append(result_row)
                            self.log_message.emit(f"[无结果] {legal_name} / {file_name}")
                            continue

                        row, hit_name, score = self.pick_best_row(page, file_name, legal_name)
                        result_row['命中文件名'] = hit_name
                        result_row['相似度'] = f"{score:.2f}"
                        if row is None or score < self.threshold:
                            self.results['no_match'] += 1
                            result_row['结果'] = f'相似度不足(阈值{self.threshold})'
                            self.results['records'].append(result_row)
                            self.log_message.emit(
                                f"[不匹配] {legal_name} / {file_name} 最高相似度 {score:.2f}")
                            continue

                        self.download_row(page, row, dest_path)
                        self.results['downloaded'] += 1
                        result_row['保存路径'] = dest_path
                        result_row['结果'] = '已下载'
                        self.results['records'].append(result_row)
                        self.log_message.emit(
                            f"[下载] {legal_name} / {hit_name} (相似度{score:.2f})")
                    except Exception as e:
                        self.results['failed'] += 1
                        result_row['结果'] = f'失败: {str(e)}'
                        self.results['records'].append(result_row)
                        self.log_message.emit(f"[失败] {legal_name} / {file_name}: {str(e)}")

                    self.progress.emit(int((i + 1) / total * 100))

                context.close()

            # 导出抓取报告
            if self.results['records']:
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                report_path = os.path.join(self.output_dir, f'抓取报告_{timestamp}.xlsx')
                os.makedirs(self.output_dir, exist_ok=True)
                pd.DataFrame(self.results['records']).to_excel(
                    report_path, index=False, engine='openpyxl')
                self.results['report_path'] = report_path
                self.log_message.emit(f"抓取报告已导出: {report_path}")

            self.progress.emit(100)
            self.finished.emit(self.results)
        except Exception as e:
            self.error.emit(str(e))


class FangcloudFetchWidget(QWidget):
    """云盘资料抓取界面"""

    def __init__(self, db_manager=None):
        super().__init__()
        self.worker = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        info_label = QLabel(
            "使用说明: 选择 缺失清单(json/xlsx) 和输出目录，点「打开浏览器抓取」会弹出浏览器，"
            "在浏览器里登录亿方云后回来点「已登录，开始抓取」。工具按原文件名走 文件(夹)名称 搜索，"
            "相似度最高（法人名出现在路径里加分）且达到阈值的下载到 输出目录/法人/原文件名，"
            "完成后用「资料自动分发」分发即可。登录状态会保存，下次无需重复登录。")
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: #666;")
        layout.addWidget(info_label)

        config_group = QGroupBox("抓取配置")
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

        self.threshold_spin = QDoubleSpinBox()
        self.threshold_spin.setRange(0.1, 1.0)
        self.threshold_spin.setSingleStep(0.05)
        self.threshold_spin.setValue(0.60)
        form_layout.addRow("相似度阈值:", self.threshold_spin)

        config_group.setLayout(form_layout)
        layout.addWidget(config_group)

        btn_layout = QHBoxLayout()
        self.start_btn = QPushButton("打开浏览器抓取")
        self.start_btn.clicked.connect(self.start_fetch)
        btn_layout.addWidget(self.start_btn)
        self.login_done_btn = QPushButton("已登录，开始抓取")
        self.login_done_btn.clicked.connect(self.on_login_done)
        self.login_done_btn.setEnabled(False)
        btn_layout.addWidget(self.login_done_btn)
        self.stop_btn = QPushButton("停止")
        self.stop_btn.clicked.connect(self.stop_fetch)
        self.stop_btn.setEnabled(False)
        btn_layout.addWidget(self.stop_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(120)
        layout.addWidget(self.log_text)

        result_group = QGroupBox("抓取结果")
        result_layout = QVBoxLayout()
        self.result_table = QTableWidget()
        self.result_table.setColumnCount(7)
        self.result_table.setHorizontalHeaderLabels([
            "部门", "法人", "原文件名", "命中文件名", "相似度", "保存路径", "结果"])
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

    def start_fetch(self):
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
        self.stop_btn.setEnabled(True)

        self.worker = FangcloudFetchWorker(
            manifest_path, output_dir, self.threshold_spin.value())
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log_message.connect(self.log)
        self.worker.login_ready.connect(self.on_login_ready)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_login_ready(self):
        self.login_done_btn.setEnabled(True)

    def on_login_done(self):
        self.login_done_btn.setEnabled(False)
        if self.worker:
            self.worker.confirm_login()

    def stop_fetch(self):
        if self.worker:
            self.worker.request_stop()
        self.stop_btn.setEnabled(False)

    def log(self, message):
        self.log_text.append(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")

    def on_finished(self, results):
        self.start_btn.setEnabled(True)
        self.login_done_btn.setEnabled(False)
        self.stop_btn.setEnabled(False)

        for record in results['records']:
            row = self.result_table.rowCount()
            self.result_table.insertRow(row)
            for col, key in enumerate(
                    ["部门", "法人", "原文件名", "命中文件名", "相似度", "保存路径", "结果"]):
                self.result_table.setItem(row, col, QTableWidgetItem(str(record[key])))

        summary = (f"抓取完成: 共 {results['total']} 个, "
                   f"下载 {results['downloaded']} 个, "
                   f"已存在跳过 {results['skipped_exists']} 个, "
                   f"未匹配 {results['no_match']} 个, "
                   f"失败 {results['failed']} 个")
        self.log(summary)
        report_path = results.get('report_path')
        msg = summary + (f"\n\n抓取报告:\n{report_path}" if report_path else "")
        QMessageBox.information(self, "完成", msg)

    def on_error(self, message):
        self.start_btn.setEnabled(True)
        self.login_done_btn.setEnabled(False)
        self.stop_btn.setEnabled(False)
        self.log(f"抓取失败: {message}")
        QMessageBox.critical(self, "错误", f"抓取失败: {message}")
