"""
补全单个主体编号的平台详情功能界面
"""
import time
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QLineEdit,
                             QCheckBox, QScrollArea, QFrame)
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
    """预检查工作线程 - 只查询不写入"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, datasource, db_manager, shopindex_id_str, selected_platforms):
        super().__init__()
        self.datasource = datasource
        self.db_manager = db_manager
        self.shopindex_id_str = shopindex_id_str.strip()
        self.selected_platforms = selected_platforms  # list of (id, platform, financial_platform, code)

    def run(self):
        try:
            self.log_message.emit("【预检查】开始连接数据库...")
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
            self.log_message.emit("【预检查】数据库连接成功")
            self.progress.emit(10)

            # 第一步：查找主体编号
            self.log_message.emit(f"【预检查】正在查找主体编号: {self.shopindex_id_str}")
            cursor.execute("""
                SELECT id, shopindex_id, `select`, legal_id 
                FROM ba_shopindex 
                WHERE shopindex_id = %s AND (delete_time IS NULL OR delete_time = 0)
            """, (self.shopindex_id_str,))
            shopindex_result = cursor.fetchone()

            if not shopindex_result:
                self.error.emit(f"未找到主体编号: {self.shopindex_id_str}")
                cursor.close()
                connection.close()
                return

            shopindex_pk = shopindex_result[0]
            shopindex_id_val = shopindex_result[1]
            shopindex_select = shopindex_result[2]
            shopindex_legal_id = shopindex_result[3]
            self.log_message.emit(
                f"【预检查】找到主体: id={shopindex_pk}, 编号={shopindex_id_val}, "
                f"属性={shopindex_select}, 法人/客户ID={shopindex_legal_id}"
            )
            self.progress.emit(20)

            # 第二步：获取所有启用的ba_shudi记录
            self.log_message.emit("【预检查】正在获取所有属地信息...")
            cursor.execute("""
                SELECT id, territory_name, territory_abbreviation 
                FROM ba_shudi 
                WHERE status = 1
                ORDER BY id
            """)
            shudi_list = cursor.fetchall()

            if not shudi_list:
                self.error.emit("未找到任何启用的属地记录")
                cursor.close()
                connection.close()
                return

            self.log_message.emit(f"【预检查】找到 {len(shudi_list)} 个属地")
            for s in shudi_list:
                self.log_message.emit(f"  属地ID={s[0]}, 名称={s[1]}, 简称={s[2]}")
            self.progress.emit(30)

            # 第三步：使用用户选择的平台列表
            platform_list = self.selected_platforms
            self.log_message.emit(f"【预检查】用户选择了 {len(platform_list)} 个平台")
            for p in platform_list:
                self.log_message.emit(f"  平台ID={p[0]}, 标识={p[1]}, 财务平台={p[2]}")
            self.progress.emit(40)

            # 第四步：逐一检查每个组合是否已存在
            total_combinations = len(shudi_list) * len(platform_list)
            self.log_message.emit(
                f"【预检查】共需检查 {len(shudi_list)} x {len(platform_list)} = {total_combinations} 个组合"
            )

            existing_count = 0
            missing_count = 0
            missing_details = []
            existing_details = []
            processed = 0

            for shudi in shudi_list:
                shudi_id = shudi[0]
                shudi_name = shudi[1]

                for platform in platform_list:
                    platform_id = platform[0]
                    platform_name = platform[1]

                    processed += 1

                    cursor.execute("""
                        SELECT id FROM ba_platform_info 
                        WHERE shopindex_id = %s AND shudi_id = %s AND platform_id = %s
                        AND (delete_time IS NULL OR delete_time = 0)
                    """, (shopindex_pk, shudi_id, platform_id))
                    existing = cursor.fetchone()

                    if existing:
                        existing_count += 1
                        self.log_message.emit(
                            f"[{processed}/{total_combinations}] 已存在: "
                            f"属地={shudi_name}, 平台={platform_name} (id={existing[0]})"
                        )
                        existing_details.append({
                            'shudi_id': shudi_id,
                            'shudi_name': shudi_name,
                            'platform_id': platform_id,
                            'platform_name': platform_name,
                            'existing_id': existing[0]
                        })
                    else:
                        missing_count += 1
                        self.log_message.emit(
                            f"[{processed}/{total_combinations}] 缺失: "
                            f"属地={shudi_name}, 平台={platform_name}"
                        )
                        missing_details.append({
                            'shudi_id': shudi_id,
                            'shudi_name': shudi_name,
                            'platform_id': platform_id,
                            'platform_name': platform_name
                        })

                    progress = 40 + int(processed / total_combinations * 55)
                    self.progress.emit(progress)

            cursor.close()
            connection.close()

            self.log_message.emit("")
            self.log_message.emit("========== 预检查结果 ==========")
            self.log_message.emit(f"总组合数: {total_combinations}")
            self.log_message.emit(f"已存在: {existing_count} 条")
            self.log_message.emit(f"需补全: {missing_count} 条")
            self.log_message.emit("================================")

            self.progress.emit(100)
            self.finished.emit({
                'shopindex_pk': shopindex_pk,
                'shopindex_id_val': shopindex_id_val,
                'total_combinations': total_combinations,
                'existing_count': existing_count,
                'missing_count': missing_count,
                'missing_details': missing_details,
                'existing_details': existing_details,
                'shudi_list': shudi_list,
                'platform_list': platform_list
            })

        except Exception as e:
            self.error.emit(str(e))


class PlatformInfoFillWorker(QThread):
    """补全单个主体编号平台详情工作线程 - 执行实际插入"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, datasource, db_manager, shopindex_pk, missing_details):
        super().__init__()
        self.datasource = datasource
        self.db_manager = db_manager
        self.shopindex_pk = shopindex_pk
        self.missing_details = missing_details
        self.results = {
            'total': len(missing_details),
            'inserted_count': 0,
            'failed_count': 0,
            'details': []
        }

    def run(self):
        try:
            self.log_message.emit("【执行补全】开始连接数据库...")
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
            self.log_message.emit("【执行补全】数据库连接成功")
            self.progress.emit(10)

            current_time = int(time.time())
            total = len(self.missing_details)

            for i, item in enumerate(self.missing_details):
                shudi_id = item['shudi_id']
                shudi_name = item['shudi_name']
                platform_id = item['platform_id']
                platform_name = item['platform_name']

                try:
                    cursor.execute("""
                        INSERT INTO ba_platform_info 
                        (platform_id, shudi_id, shopindex_id, status, 
                         create_time, update_time, admin_id, admin_dept_id)
                        VALUES (%s, %s, %s, 1, %s, %s, 0, 0)
                    """, (platform_id, shudi_id, self.shopindex_pk,
                          current_time, current_time))
                    connection.commit()

                    self.results['inserted_count'] += 1
                    self.log_message.emit(
                        f"[{i+1}/{total}] 插入成功: "
                        f"属地={shudi_name}, 平台={platform_name}"
                    )
                    self.results['details'].append({
                        'shudi_id': shudi_id,
                        'shudi_name': shudi_name,
                        'platform_id': platform_id,
                        'platform_name': platform_name,
                        'status': '插入成功'
                    })
                except Exception as e:
                    connection.rollback()
                    self.results['failed_count'] += 1
                    self.log_message.emit(
                        f"[{i+1}/{total}] 插入失败: "
                        f"属地={shudi_name}, 平台={platform_name}, 原因: {str(e)}"
                    )
                    self.results['details'].append({
                        'shudi_id': shudi_id,
                        'shudi_name': shudi_name,
                        'platform_id': platform_id,
                        'platform_name': platform_name,
                        'status': f'插入失败: {str(e)}'
                    })

                progress = 10 + int((i + 1) / total * 85)
                self.progress.emit(progress)

            cursor.close()
            connection.close()

            self.progress.emit(100)
            self.finished.emit(self.results)

        except Exception as e:
            self.error.emit(str(e))


class PlatformInfoFillWidget(QWidget):
    """补全单个主体编号的平台详情功能组件"""

    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.platform_checkboxes = []  # 存储 (checkbox, platform_data) 元组
        self.all_platforms = []  # 从数据库加载的全部平台
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

        # 主体编号输入区域
        input_group = QGroupBox("主体编号输入")
        input_layout = QFormLayout()

        self.shopindex_input = QLineEdit()
        self.shopindex_input.setPlaceholderText("请输入主体编号 (ba_shopindex.shopindex_id)")
        self.shopindex_input.textChanged.connect(self.on_input_changed)
        input_layout.addRow("主体编号:", self.shopindex_input)

        input_group.setLayout(input_layout)
        layout.addWidget(input_group)

        # 平台选择区域
        platform_group = QGroupBox("平台选择 (请先加载平台列表)")
        platform_main_layout = QVBoxLayout()

        # 加载平台 + 全选/取消全选 按钮行
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

        # 平台复选框滚动区域
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
<b>补全单个主体编号的平台详情说明:</b><br>
<b>操作步骤:</b><br>
• 1. 选择数据库并输入主体编号<br>
• 2. 点击【加载平台列表】获取所有平台，勾选需要补全的平台<br>
• 3. 点击【预检查】查看已有/缺失的记录数<br>
• 4. 确认无误后点击【确认执行补全】插入缺失记录<br><br>
<b>补全逻辑:</b> 选中的平台 × 所有属地 的组合，已存在则跳过
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

        # 进度条
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
        # 数据源变化时清空平台列表
        self.clear_platform_checkboxes()
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

    def on_input_changed(self):
        """输入变化时清除预检查结果"""
        self.execute_btn.setEnabled(False)
        self.precheck_data = None
        self.update_button_states()

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
            cb.setChecked(True)  # 默认全选
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
        # 平台选择变化时清除预检查结果
        self.execute_btn.setEnabled(False)
        self.precheck_data = None
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
        has_input = bool(self.shopindex_input.text().strip())
        has_platforms = len(self.get_selected_platforms()) > 0
        self.precheck_btn.setEnabled(datasource_selected and has_input and has_platforms)

    def start_precheck(self):
        """开始预检查"""
        datasource = self.datasource_combo.currentData()
        shopindex_id_str = self.shopindex_input.text().strip()
        selected_platforms = self.get_selected_platforms()

        if not datasource or not shopindex_id_str:
            QMessageBox.warning(self, "警告", "请选择数据源并输入主体编号！")
            return

        if not selected_platforms:
            QMessageBox.warning(self, "警告", "请至少选择一个平台！")
            return

        # 禁用按钮，显示进度条
        self.precheck_btn.setEnabled(False)
        self.execute_btn.setEnabled(False)
        self.shopindex_input.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在预检查中...")
        self.log_text.clear()
        self.precheck_data = None

        # 启动预检查线程
        self.precheck_worker = PreCheckWorker(
            datasource, self.db_manager, shopindex_id_str, selected_platforms
        )
        self.precheck_worker.progress.connect(self.progress_bar.setValue)
        self.precheck_worker.log_message.connect(self.append_log)
        self.precheck_worker.finished.connect(self.on_precheck_finished)
        self.precheck_worker.error.connect(self.on_precheck_error)
        self.precheck_worker.start()

    def on_precheck_finished(self, results):
        """预检查完成"""
        self.precheck_btn.setEnabled(True)
        self.shopindex_input.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.precheck_data = results

        total = results['total_combinations']
        existing = results['existing_count']
        missing = results['missing_count']

        result_text = (
            f"预检查完成！\n"
            f"主体编号: {results['shopindex_id_val']} (id={results['shopindex_pk']})\n"
            f"总组合数: {total} (选中平台数 × 属地数)\n"
            f"已存在: {existing} 条\n"
            f"需补全: {missing} 条"
        )
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
                f"{result_text}\n\n选中平台的所有组合均已存在，无需补全。"
            )

    def on_precheck_error(self, error_msg):
        """预检查错误"""
        self.precheck_btn.setEnabled(True)
        self.shopindex_input.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.result_label.setText("预检查失败！")
        self.result_label.setStyleSheet("color: red;")
        QMessageBox.critical(self, "预检查错误", f"预检查失败：{error_msg}")

    def start_fill(self):
        """开始执行补全"""
        if not self.precheck_data or not self.precheck_data['missing_details']:
            QMessageBox.warning(self, "警告", "请先执行预检查！")
            return

        datasource = self.datasource_combo.currentData()
        if not datasource:
            QMessageBox.warning(self, "警告", "请选择数据源！")
            return

        missing = self.precheck_data['missing_count']
        shopindex_id_val = self.precheck_data['shopindex_id_val']

        # 确认对话框
        reply = QMessageBox.question(
            self,
            "确认执行补全",
            f"确定要为主体编号 [{shopindex_id_val}] 补全平台详情吗？\n\n"
            f"数据库: {datasource.name}\n"
            f"待补全: {missing} 条\n\n"
            f"此操作将向 ba_platform_info 表插入 {missing} 条新记录。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply != QMessageBox.Yes:
            return

        # 禁用按钮，显示进度条
        self.precheck_btn.setEnabled(False)
        self.execute_btn.setEnabled(False)
        self.shopindex_input.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在执行补全...")

        # 启动补全线程
        self.worker = PlatformInfoFillWorker(
            datasource, self.db_manager,
            self.precheck_data['shopindex_pk'],
            self.precheck_data['missing_details']
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
        self.shopindex_input.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.execute_btn.setEnabled(False)
        self.precheck_data = None

        result_text = (
            f"补全完成！\n"
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
        self.shopindex_input.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.execute_btn.setEnabled(False)
        self.result_label.setText("补全失败！")
        self.result_label.setStyleSheet("color: red;")

        QMessageBox.critical(self, "补全错误", f"补全失败：{error_msg}")

    def clear_log(self):
        """清空日志"""
        self.log_text.clear()
