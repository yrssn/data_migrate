"""
补全主体编号的平台详情功能界面
支持全部主体或指定单个主体，可选择要补全的平台
"""
import time
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QLineEdit,
                             QCheckBox, QScrollArea, QRadioButton, QButtonGroup)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from database import DatabaseManager
import pymysql


class LoadPlatformWorker(QThread):
    """加载平台列表工作线程"""
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, datasource, db_manager):
        super().__init__()
        self.datasource = datasource
        self.db_manager = db_manager

    def run(self):
        try:
            self.log_message.emit("正在连接数据库加载平台列表...")

            connection = pymysql.connect(
                host=self.datasource.host,
                port=self.datasource.port,
                user=self.datasource.username,
                password=self.datasource.password,
                database=self.datasource.database,
                charset=self.datasource.charset
            )
            cursor = connection.cursor()
            self.log_message.emit("数据库连接成功")

            cursor.execute("""
                SELECT id, platform, financial_platform, code 
                FROM ba_platform 
                WHERE status = 1 AND (delete_time IS NULL OR delete_time = 0)
                ORDER BY id
            """)
            platform_list = cursor.fetchall()

            cursor.close()
            connection.close()

            if not platform_list:
                self.error.emit("未找到任何启用的平台记录")
                return

            self.log_message.emit(f"成功加载 {len(platform_list)} 个平台")
            self.finished.emit({'platform_list': platform_list})

        except Exception as e:
            self.error.emit(str(e))


class PreCheckWorker(QThread):
    """Pre-check worker: read-only check of missing combinations."""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, datasource, db_manager, selected_platforms, scope_all=True, shopindex_id_str=''):
        super().__init__()
        self.datasource = datasource
        self.db_manager = db_manager
        self.selected_platforms = selected_platforms
        self.scope_all = scope_all
        self.shopindex_id_str = shopindex_id_str.strip()

    def run(self):
        try:
            self.log_message.emit("[PreCheck] Connecting to database...")
            self.progress.emit(5)

            connection = pymysql.connect(
                host=self.datasource.host,
                port=self.datasource.port,
                user=self.datasource.username,
                password=self.datasource.password,
                database=self.datasource.database,
                charset=self.datasource.charset
            )
            cursor = connection.cursor()
            self.log_message.emit("[PreCheck] Connected")
            self.progress.emit(10)

            if self.scope_all:
                cursor.execute("""
                    SELECT id, shopindex_id, admin_id, admin_dept_id
                    FROM ba_shopindex
                    WHERE status = 1 AND (delete_time IS NULL OR delete_time = 0)
                    ORDER BY id
                """)
            else:
                cursor.execute("""
                    SELECT id, shopindex_id, admin_id, admin_dept_id
                    FROM ba_shopindex
                    WHERE shopindex_id = %s AND (delete_time IS NULL OR delete_time = 0)
                """, (self.shopindex_id_str,))

            shopindex_records = cursor.fetchall()
            if not shopindex_records:
                cursor.close()
                connection.close()
                if self.scope_all:
                    self.error.emit("No active shopindex records found")
                else:
                    self.error.emit(f"Shopindex not found: {self.shopindex_id_str}")
                return

            self.progress.emit(15)
            cursor.execute("""
                SELECT id, territory_name, territory_abbreviation
                FROM ba_shudi
                WHERE status = 1
                ORDER BY id
            """)
            shudi_list = cursor.fetchall()
            if not shudi_list:
                cursor.close()
                connection.close()
                self.error.emit("No active shudi records found")
                return

            self.progress.emit(20)
            platform_list = self.selected_platforms
            if not platform_list:
                cursor.close()
                connection.close()
                self.error.emit("No platforms selected")
                return

            per_shopindex = len(shudi_list) * len(platform_list)
            total_checks = len(shopindex_records) * per_shopindex

            shopindex_ids = [r[0] for r in shopindex_records]
            shudi_ids = [s[0] for s in shudi_list]
            platform_ids = [p[0] for p in platform_list]
            shudi_name_map = {s[0]: s[1] for s in shudi_list}
            platform_name_map = {p[0]: p[1] for p in platform_list}

            def chunked(seq, size):
                for i in range(0, len(seq), size):
                    yield seq[i:i + size]

            existing_set = set()
            batch_size = 500
            id_chunks = list(chunked(shopindex_ids, batch_size))
            total_chunk = max(len(id_chunks), 1)
            self.log_message.emit(f"[PreCheck] Batch query existing rows in {total_chunk} chunks")

            for idx, si_chunk in enumerate(id_chunks, 1):
                si_ph = ",".join(["%s"] * len(si_chunk))
                p_ph = ",".join(["%s"] * len(platform_ids))
                sql = f"""
                    SELECT shopindex_id, shudi_id, platform_id
                    FROM ba_platform_info
                    WHERE shopindex_id IN ({si_ph})
                      AND platform_id IN ({p_ph})
                      AND (delete_time IS NULL OR delete_time = 0)
                """
                cursor.execute(sql, tuple(si_chunk) + tuple(platform_ids))
                for row in cursor.fetchall():
                    existing_set.add((row[0], row[1], row[2]))

                progress = 20 + int(idx / total_chunk * 30)
                self.progress.emit(progress)

            total_existing = 0
            total_missing = 0
            missing_by_shopindex = {}
            processed = 0
            progress_step = max(total_checks // 200, 1)

            for si_pk, si_id_val, si_admin_id, si_admin_dept_id in shopindex_records:
                for shudi_id in shudi_ids:
                    shudi_name = shudi_name_map.get(shudi_id, "")
                    for platform_id in platform_ids:
                        platform_name = platform_name_map.get(platform_id, "")
                        processed += 1
                        key = (si_pk, shudi_id, platform_id)

                        if key in existing_set:
                            total_existing += 1
                        else:
                            total_missing += 1
                            if si_pk not in missing_by_shopindex:
                                missing_by_shopindex[si_pk] = {
                                    'shopindex_id_val': si_id_val,
                                    'admin_id': si_admin_id,
                                    'admin_dept_id': si_admin_dept_id,
                                    'items': []
                                }
                            missing_by_shopindex[si_pk]['items'].append({
                                'shudi_id': shudi_id,
                                'shudi_name': shudi_name,
                                'platform_id': platform_id,
                                'platform_name': platform_name
                            })

                        if processed % progress_step == 0 or processed == total_checks:
                            progress = 50 + int(processed / max(total_checks, 1) * 45)
                            self.progress.emit(progress)

            cursor.close()
            connection.close()

            self.progress.emit(100)
            self.finished.emit({
                'shopindex_count': len(shopindex_records),
                'total_checks': total_checks,
                'total_existing': total_existing,
                'total_missing': total_missing,
                'missing_by_shopindex': missing_by_shopindex,
                'scope_all': self.scope_all
            })

        except Exception as e:
            self.error.emit(str(e))


class PlatformInfoFillWorker(QThread):
    """Fill missing ba_platform_info rows using batched SQL."""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, datasource, db_manager, missing_by_shopindex):
        super().__init__()
        self.datasource = datasource
        self.db_manager = db_manager
        self.missing_by_shopindex = missing_by_shopindex
        self.total = sum(len(info['items']) for info in missing_by_shopindex.values())
        self.results = {
            'total': self.total,
            'inserted_count': 0,
            'failed_count': 0,
            'shopindex_count': len(missing_by_shopindex)
        }

    def run(self):
        try:
            self.log_message.emit("[Fill] Connecting to database...")
            self.progress.emit(5)

            connection = pymysql.connect(
                host=self.datasource.host,
                port=self.datasource.port,
                user=self.datasource.username,
                password=self.datasource.password,
                database=self.datasource.database,
                charset=self.datasource.charset
            )
            cursor = connection.cursor()
            connection.autocommit(False)
            self.log_message.emit("[Fill] Connected")
            self.progress.emit(10)

            current_time = int(time.time())
            processed = 0
            batch_size = 1000

            def chunked(seq, size):
                for i in range(0, len(seq), size):
                    yield seq[i:i + size]

            insert_sql = """
                INSERT INTO ba_platform_info
                (platform_id, shudi_id, shopindex_id, status,
                 create_time, update_time, admin_id, admin_dept_id)
                VALUES (%s, %s, %s, 1, %s, %s, %s, %s)
            """

            for si_pk, info in self.missing_by_shopindex.items():
                si_id_val = info['shopindex_id_val']
                si_admin_id = info.get('admin_id') or 0
                si_admin_dept_id = info.get('admin_dept_id') or 0
                items = info['items']
                if not items:
                    continue

                shudi_ids = sorted({item['shudi_id'] for item in items})
                platform_ids = sorted({item['platform_id'] for item in items})
                existing_pairs = set()

                s_ph = ",".join(["%s"] * len(shudi_ids))
                p_ph = ",".join(["%s"] * len(platform_ids))
                check_sql = f"""
                    SELECT shudi_id, platform_id
                    FROM ba_platform_info
                    WHERE shopindex_id = %s
                      AND shudi_id IN ({s_ph})
                      AND platform_id IN ({p_ph})
                      AND (delete_time IS NULL OR delete_time = 0)
                """
                cursor.execute(check_sql, (si_pk, *shudi_ids, *platform_ids))
                for row in cursor.fetchall():
                    existing_pairs.add((row[0], row[1]))

                pending_rows = []
                for item in items:
                    pair = (item['shudi_id'], item['platform_id'])
                    if pair not in existing_pairs:
                        pending_rows.append((
                            item['platform_id'], item['shudi_id'], si_pk,
                            current_time, current_time, si_admin_id, si_admin_dept_id
                        ))

                self.log_message.emit(
                    f"[Fill] shopindex={si_id_val}(id={si_pk}) pending {len(pending_rows)}/{len(items)}"
                )

                for batch in chunked(pending_rows, batch_size):
                    try:
                        cursor.executemany(insert_sql, batch)
                        connection.commit()
                        self.results['inserted_count'] += len(batch)
                    except Exception as batch_error:
                        connection.rollback()
                        for row in batch:
                            try:
                                cursor.execute(insert_sql, row)
                                connection.commit()
                                self.results['inserted_count'] += 1
                            except Exception:
                                connection.rollback()
                                self.results['failed_count'] += 1
                        self.log_message.emit(f"[Fill] Batch fallback due to: {str(batch_error)}")

                    processed += len(batch)
                    progress = 10 + int(processed / max(self.total, 1) * 85)
                    self.progress.emit(progress)

            cursor.close()
            connection.close()

            self.progress.emit(100)
            self.finished.emit(self.results)

        except Exception as e:
            self.error.emit(str(e))


class PlatformInfoFillWidget(QWidget):
    """补全主体编号的平台详情功能组件 - 支持全部/单个主体"""

    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.platform_checkboxes = []
        self.all_platforms = []
        self.precheck_data = None
        self.init_ui()
        self.load_datasources()

    def init_ui(self):
        """初始化UI"""
        layout = QVBoxLayout()

        # 数据源选择区域
        datasource_group = QGroupBox("选择数据库")
        datasource_layout = QFormLayout()

        self.datasource_combo = QComboBox()
        self.datasource_combo.currentTextChanged.connect(self.on_datasource_changed)
        datasource_layout.addRow("数据库:", self.datasource_combo)

        self.test_btn = QPushButton("测试连接")
        self.test_btn.clicked.connect(self.test_connection)
        self.test_btn.setEnabled(False)
        datasource_layout.addRow("", self.test_btn)

        datasource_group.setLayout(datasource_layout)
        layout.addWidget(datasource_group)

        # 范围选择区域
        scope_group = QGroupBox("主体编号范围")
        scope_layout = QVBoxLayout()

        scope_radio_layout = QHBoxLayout()
        self.scope_all_radio = QRadioButton("全部主体编号")
        self.scope_single_radio = QRadioButton("指定主体编号")
        self.scope_all_radio.setChecked(True)
        self.scope_btn_group = QButtonGroup()
        self.scope_btn_group.addButton(self.scope_all_radio)
        self.scope_btn_group.addButton(self.scope_single_radio)
        self.scope_all_radio.toggled.connect(self.on_scope_changed)
        scope_radio_layout.addWidget(self.scope_all_radio)
        scope_radio_layout.addWidget(self.scope_single_radio)
        scope_radio_layout.addStretch()
        scope_layout.addLayout(scope_radio_layout)

        input_layout = QHBoxLayout()
        input_layout.addWidget(QLabel("主体编号:"))
        self.shopindex_input = QLineEdit()
        self.shopindex_input.setPlaceholderText("请输入主体编号 (ba_shopindex.shopindex_id)")
        self.shopindex_input.setEnabled(False)
        self.shopindex_input.textChanged.connect(self.on_input_changed)
        input_layout.addWidget(self.shopindex_input)
        scope_layout.addLayout(input_layout)

        scope_group.setLayout(scope_layout)
        layout.addWidget(scope_group)

        # 平台选择区域
        platform_group = QGroupBox("平台选择 (请先加载平台列表)")
        platform_main_layout = QVBoxLayout()

        platform_btn_layout = QHBoxLayout()

        self.load_platform_btn = QPushButton("加载平台列表")
        self.load_platform_btn.clicked.connect(self.load_platforms)
        self.load_platform_btn.setEnabled(False)
        platform_btn_layout.addWidget(self.load_platform_btn)

        self.select_all_btn = QPushButton("全选")
        self.select_all_btn.clicked.connect(self.select_all_platforms)
        self.select_all_btn.setEnabled(False)
        platform_btn_layout.addWidget(self.select_all_btn)

        self.deselect_all_btn = QPushButton("取消全选")
        self.deselect_all_btn.clicked.connect(self.deselect_all_platforms)
        self.deselect_all_btn.setEnabled(False)
        platform_btn_layout.addWidget(self.deselect_all_btn)

        self.platform_count_label = QLabel("未加载")
        self.platform_count_label.setStyleSheet("color: gray;")
        platform_btn_layout.addWidget(self.platform_count_label)

        platform_btn_layout.addStretch()
        platform_main_layout.addLayout(platform_btn_layout)

        self.platform_scroll = QScrollArea()
        self.platform_scroll.setWidgetResizable(True)
        self.platform_scroll.setMaximumHeight(150)
        self.platform_scroll_content = QWidget()
        self.platform_scroll_layout = QVBoxLayout()
        self.platform_scroll_content.setLayout(self.platform_scroll_layout)
        self.platform_scroll.setWidget(self.platform_scroll_content)
        platform_main_layout.addWidget(self.platform_scroll)

        platform_group.setLayout(platform_main_layout)
        layout.addWidget(platform_group)

        # 功能说明区域
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()

        info_label = QLabel("""
<b>补全主体编号的平台详情:</b><br>
<b>操作步骤:</b><br>
• 1. 选择数据库，选择范围（全部主体 或 指定主体编号）<br>
• 2. 点击【加载平台列表】获取所有平台，勾选需要补全的平台<br>
• 3. 点击【预检查】查看已有/缺失的记录数<br>
• 4. 确认无误后点击【确认执行补全】插入缺失记录<br><br>
<b>补全逻辑:</b> 主体编号 × 选中的平台 × 所有属地 的组合，插入前逐条检查已存在则跳过
        """)
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)

        info_group.setLayout(info_layout)
        layout.addWidget(info_group)

        # 操作按钮区域
        button_layout = QHBoxLayout()

        self.precheck_btn = QPushButton("预检查")
        self.precheck_btn.clicked.connect(self.start_precheck)
        self.precheck_btn.setEnabled(False)
        button_layout.addWidget(self.precheck_btn)

        self.execute_btn = QPushButton("确认执行补全")
        self.execute_btn.clicked.connect(self.start_fill)
        self.execute_btn.setEnabled(False)
        self.execute_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        button_layout.addWidget(self.execute_btn)

        button_layout.addStretch()

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        button_layout.addWidget(self.progress_bar)

        layout.addLayout(button_layout)

        # 日志区域
        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout()

        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(200)
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)

        log_button_layout = QHBoxLayout()
        self.clear_log_btn = QPushButton("清空日志")
        self.clear_log_btn.clicked.connect(self.clear_log)
        log_button_layout.addWidget(self.clear_log_btn)

        log_button_layout.addStretch()
        log_layout.addLayout(log_button_layout)

        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        # 结果显示区域
        self.result_label = QLabel("")
        self.result_label.setAlignment(Qt.AlignCenter)
        font = QFont()
        font.setBold(True)
        self.result_label.setFont(font)
        layout.addWidget(self.result_label)

        self.setLayout(layout)

    def load_datasources(self):
        """加载数据源列表"""
        self.datasource_combo.clear()
        self.datasource_combo.addItem("请选择数据源", None)

        datasources = self.db_manager.get_all_datasources()
        for ds in datasources:
            display_name = f"{ds.name} ({ds.host}:{ds.port})"
            self.datasource_combo.addItem(display_name, ds)

    def on_datasource_changed(self):
        """数据源选择变化"""
        data = self.datasource_combo.currentData()
        self.test_btn.setEnabled(data is not None)
        self.load_platform_btn.setEnabled(data is not None)
        self.clear_platform_checkboxes()
        self.reset_precheck()
        self.update_button_states()

    def test_connection(self):
        """测试数据库连接"""
        datasource = self.datasource_combo.currentData()
        if not datasource:
            return

        success, message = self.db_manager.test_connection(datasource)
        if success:
            QMessageBox.information(self, "连接成功", message)
        else:
            QMessageBox.critical(self, "连接失败", message)

    def on_scope_changed(self):
        """范围切换"""
        is_single = self.scope_single_radio.isChecked()
        self.shopindex_input.setEnabled(is_single)
        if not is_single:
            self.shopindex_input.clear()
        self.reset_precheck()
        self.update_button_states()

    def on_input_changed(self):
        """输入变化时清除预检查结果"""
        self.reset_precheck()
        self.update_button_states()

    def reset_precheck(self):
        """清除预检查结果"""
        self.execute_btn.setEnabled(False)
        self.precheck_data = None

    def clear_platform_checkboxes(self):
        """清空平台复选框列表"""
        for cb, _ in self.platform_checkboxes:
            cb.deleteLater()
        self.platform_checkboxes = []
        self.all_platforms = []
        self.select_all_btn.setEnabled(False)
        self.deselect_all_btn.setEnabled(False)
        self.platform_count_label.setText("未加载")
        self.platform_count_label.setStyleSheet("color: gray;")

    def load_platforms(self):
        """从数据库加载平台列表"""
        datasource = self.datasource_combo.currentData()
        if not datasource:
            QMessageBox.warning(self, "警告", "请先选择数据源！")
            return

        self.load_platform_btn.setEnabled(False)
        self.platform_count_label.setText("加载中...")
        self.platform_count_label.setStyleSheet("color: blue;")

        self.load_platform_worker = LoadPlatformWorker(datasource, self.db_manager)
        self.load_platform_worker.log_message.connect(self.append_log)
        self.load_platform_worker.finished.connect(self.on_platforms_loaded)
        self.load_platform_worker.error.connect(self.on_platforms_load_error)
        self.load_platform_worker.start()

    def on_platforms_loaded(self, results):
        """平台列表加载完成"""
        self.load_platform_btn.setEnabled(True)
        self.clear_platform_checkboxes()

        platform_list = results['platform_list']
        self.all_platforms = platform_list

        for p in platform_list:
            p_id, p_name, p_financial, p_code = p
            cb = QCheckBox(f"[{p_id}] {p_name} (财务: {p_financial or '-'}, 编码: {p_code or '-'})")
            cb.setChecked(True)
            cb.stateChanged.connect(self.on_platform_selection_changed)
            self.platform_scroll_layout.addWidget(cb)
            self.platform_checkboxes.append((cb, p))

        self.select_all_btn.setEnabled(True)
        self.deselect_all_btn.setEnabled(True)
        self.update_platform_count_label()
        self.update_button_states()

    def on_platforms_load_error(self, error_msg):
        """平台列表加载失败"""
        self.load_platform_btn.setEnabled(True)
        self.platform_count_label.setText("加载失败")
        self.platform_count_label.setStyleSheet("color: red;")
        QMessageBox.critical(self, "加载失败", f"加载平台列表失败：{error_msg}")

    def select_all_platforms(self):
        """全选所有平台"""
        for cb, _ in self.platform_checkboxes:
            cb.setChecked(True)

    def deselect_all_platforms(self):
        """取消全选所有平台"""
        for cb, _ in self.platform_checkboxes:
            cb.setChecked(False)

    def on_platform_selection_changed(self):
        """平台选择变化时更新计数和按钮状态"""
        self.update_platform_count_label()
        self.reset_precheck()
        self.update_button_states()

    def update_platform_count_label(self):
        """更新平台选择计数标签"""
        selected = self.get_selected_platforms()
        total = len(self.platform_checkboxes)
        self.platform_count_label.setText(f"已选 {len(selected)}/{total} 个平台")
        if len(selected) == 0:
            self.platform_count_label.setStyleSheet("color: red;")
        elif len(selected) == total:
            self.platform_count_label.setStyleSheet("color: green;")
        else:
            self.platform_count_label.setStyleSheet("color: orange;")

    def get_selected_platforms(self):
        """获取已选中的平台列表"""
        selected = []
        for cb, platform_data in self.platform_checkboxes:
            if cb.isChecked():
                selected.append(platform_data)
        return selected

    def update_button_states(self):
        """更新按钮状态"""
        datasource_selected = self.datasource_combo.currentData() is not None
        has_platforms = len(self.get_selected_platforms()) > 0

        if self.scope_all_radio.isChecked():
            # 全部模式：只需数据源和平台
            self.precheck_btn.setEnabled(datasource_selected and has_platforms)
        else:
            # 单个模式：还需要输入主体编号
            has_input = bool(self.shopindex_input.text().strip())
            self.precheck_btn.setEnabled(datasource_selected and has_platforms and has_input)

    def start_precheck(self):
        """开始预检查"""
        datasource = self.datasource_combo.currentData()
        selected_platforms = self.get_selected_platforms()

        if not datasource:
            QMessageBox.warning(self, "警告", "请选择数据源！")
            return

        if not selected_platforms:
            QMessageBox.warning(self, "警告", "请至少选择一个平台！")
            return

        scope_all = self.scope_all_radio.isChecked()
        shopindex_id_str = ''

        if not scope_all:
            shopindex_id_str = self.shopindex_input.text().strip()
            if not shopindex_id_str:
                QMessageBox.warning(self, "警告", "请输入主体编号！")
                return

        # 禁用控件
        self.precheck_btn.setEnabled(False)
        self.execute_btn.setEnabled(False)
        self.shopindex_input.setEnabled(False)
        self.scope_all_radio.setEnabled(False)
        self.scope_single_radio.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        scope_desc = "全部主体" if scope_all else f"主体编号={shopindex_id_str}"
        self.result_label.setText(f"正在预检查中... ({scope_desc})")
        self.log_text.clear()
        self.precheck_data = None

        self.precheck_worker = PreCheckWorker(
            datasource, self.db_manager, selected_platforms,
            scope_all=scope_all, shopindex_id_str=shopindex_id_str
        )
        self.precheck_worker.progress.connect(self.progress_bar.setValue)
        self.precheck_worker.log_message.connect(self.append_log)
        self.precheck_worker.finished.connect(self.on_precheck_finished)
        self.precheck_worker.error.connect(self.on_precheck_error)
        self.precheck_worker.start()

    def on_precheck_finished(self, results):
        """预检查完成"""
        self.precheck_btn.setEnabled(True)
        self.shopindex_input.setEnabled(self.scope_single_radio.isChecked())
        self.scope_all_radio.setEnabled(True)
        self.scope_single_radio.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.precheck_data = results

        si_count = results['shopindex_count']
        total_checks = results['total_checks']
        existing = results['total_existing']
        missing = results['total_missing']
        missing_si = len(results['missing_by_shopindex'])

        scope_desc = "全部主体" if results['scope_all'] else "指定主体"
        result_text = (
            f"预检查完成！ ({scope_desc})\n"
            f"主体编号数: {si_count}\n"
            f"总检查组合: {total_checks}\n"
            f"已存在: {existing} 条\n"
            f"需补全: {missing} 条"
        )
        if missing > 0:
            result_text += f"\n涉及 {missing_si} 个主体"

        self.result_label.setText(result_text)

        if missing > 0:
            self.result_label.setStyleSheet("color: orange;")
            self.execute_btn.setEnabled(True)
            QMessageBox.information(
                self, "预检查完成",
                f"{result_text}\n\n如需补全，请点击【确认执行补全】按钮。"
            )
        else:
            self.result_label.setStyleSheet("color: green;")
            self.execute_btn.setEnabled(False)
            QMessageBox.information(
                self, "预检查完成",
                f"{result_text}\n\n所有组合均已存在，无需补全。"
            )

    def on_precheck_error(self, error_msg):
        """预检查错误"""
        self.precheck_btn.setEnabled(True)
        self.shopindex_input.setEnabled(self.scope_single_radio.isChecked())
        self.scope_all_radio.setEnabled(True)
        self.scope_single_radio.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.result_label.setText("预检查失败！")
        self.result_label.setStyleSheet("color: red;")
        QMessageBox.critical(self, "预检查错误", f"预检查失败：{error_msg}")

    def start_fill(self):
        """开始执行补全"""
        if not self.precheck_data or not self.precheck_data['missing_by_shopindex']:
            QMessageBox.warning(self, "警告", "请先执行预检查！")
            return

        datasource = self.datasource_combo.currentData()
        if not datasource:
            QMessageBox.warning(self, "警告", "请选择数据源！")
            return

        missing = self.precheck_data['total_missing']
        missing_si = len(self.precheck_data['missing_by_shopindex'])
        scope_desc = "全部主体" if self.precheck_data['scope_all'] else "指定主体"

        reply = QMessageBox.question(
            self,
            "确认执行补全",
            f"确定要执行补全吗？ ({scope_desc})\n\n"
            f"数据库: {datasource.name}\n"
            f"涉及主体: {missing_si} 个\n"
            f"待补全: {missing} 条\n\n"
            f"此操作将向 ba_platform_info 表插入最多 {missing} 条新记录。\n"
            f"每条插入前会再次检查是否已存在，确保不重复。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply != QMessageBox.Yes:
            return

        # 禁用控件
        self.precheck_btn.setEnabled(False)
        self.execute_btn.setEnabled(False)
        self.shopindex_input.setEnabled(False)
        self.scope_all_radio.setEnabled(False)
        self.scope_single_radio.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在执行补全...")

        self.worker = PlatformInfoFillWorker(
            datasource, self.db_manager,
            self.precheck_data['missing_by_shopindex']
        )
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log_message.connect(self.append_log)
        self.worker.finished.connect(self.on_fill_finished)
        self.worker.error.connect(self.on_fill_error)
        self.worker.start()

    def append_log(self, message):
        """添加日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")

    def on_fill_finished(self, results):
        """补全完成"""
        self.precheck_btn.setEnabled(True)
        self.shopindex_input.setEnabled(self.scope_single_radio.isChecked())
        self.scope_all_radio.setEnabled(True)
        self.scope_single_radio.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.execute_btn.setEnabled(False)
        self.precheck_data = None

        result_text = (
            f"补全完成！\n"
            f"涉及主体: {results['shopindex_count']} 个\n"
            f"待补全：{results['total']} 条\n"
            f"成功插入：{results['inserted_count']} 条\n"
            f"失败：{results['failed_count']} 条"
        )

        self.result_label.setText(result_text)
        if results['failed_count'] > 0:
            self.result_label.setStyleSheet("color: orange;")
        else:
            self.result_label.setStyleSheet("color: green;")

        QMessageBox.information(self, "补全完成", result_text)

    def on_fill_error(self, error_msg):
        """补全错误"""
        self.precheck_btn.setEnabled(True)
        self.shopindex_input.setEnabled(self.scope_single_radio.isChecked())
        self.scope_all_radio.setEnabled(True)
        self.scope_single_radio.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.execute_btn.setEnabled(False)
        self.result_label.setText("补全失败！")
        self.result_label.setStyleSheet("color: red;")

        QMessageBox.critical(self, "补全错误", f"补全失败：{error_msg}")

    def clear_log(self):
        """清空日志"""
        self.log_text.clear()
