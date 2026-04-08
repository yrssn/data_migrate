"""
同步进销存库存编号到财务系统
遍历ba_account_mag，通过legal_id和platform_id关联找到对应名称，
然后在财务系统中匹配ea_dy_legal和ea_dy_site_type，
找到ea_dy_shop记录并将inventory_num填充到stock_number
"""
import os
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from database import DatabaseManager
import pandas as pd
import pymysql


class SyncInventoryToFinanceWorker(QThread):
    """同步工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, source_datasource, target_datasource, preview_only=False):
        super().__init__()
        self.source_datasource = source_datasource  # 进销存 ba_
        self.target_datasource = target_datasource  # 财务系统 ea_
        self.preview_only = preview_only
        self.results = {
            'total': 0,
            'success_count': 0,
            'fail_count': 0,
            'skip_count': 0,
            'multi_match_count': 0,
            'completed_data': [],
            'failed_records': []
        }

    def run(self):
        try:
            mode_text = "预检查" if self.preview_only else "同步"
            self.log_message.emit(f"开始{mode_text}...")

            # 连接进销存数据库
            self.log_message.emit("正在连接进销存数据库...")
            self.progress.emit(5)
            source_conn = pymysql.connect(
                host=self.source_datasource.host,
                port=self.source_datasource.port,
                user=self.source_datasource.username,
                password=self.source_datasource.password,
                database=self.source_datasource.database,
                charset=self.source_datasource.charset
            )
            source_cursor = source_conn.cursor(pymysql.cursors.DictCursor)
            self.log_message.emit("进销存数据库连接成功")

            # 连接财务系统数据库
            self.log_message.emit("正在连接财务系统数据库...")
            self.progress.emit(10)
            target_conn = pymysql.connect(
                host=self.target_datasource.host,
                port=self.target_datasource.port,
                user=self.target_datasource.username,
                password=self.target_datasource.password,
                database=self.target_datasource.database,
                charset=self.target_datasource.charset
            )
            target_cursor = target_conn.cursor(pymysql.cursors.DictCursor)
            self.log_message.emit("财务系统数据库连接成功")
            self.progress.emit(15)

            # 查询所有有inventory_num的ba_account_mag记录
            self.log_message.emit("正在查询ba_account_mag中有库存编号的记录...")
            source_cursor.execute("""
                SELECT a.id, a.inventory_num, a.legal_id, a.platform_id,
                       a.shopindex_id, a.shudi_id, a.account_status,
                       a.register_department, a.remark
                FROM ba_account_mag a
                WHERE a.inventory_num IS NOT NULL 
                  AND a.inventory_num != ''
                  AND a.delete_time IS NULL
            """)
            mag_records = source_cursor.fetchall()
            total = len(mag_records)
            self.results['total'] = total
            self.log_message.emit(f"找到 {total} 条有库存编号的记录")

            if total == 0:
                self.log_message.emit("没有需要同步的记录")
                self.progress.emit(100)
                self.finished.emit(self.results)
                return

            # 预加载财务系统的法人和店铺类型映射
            self.log_message.emit("正在加载财务系统法人列表...")
            target_cursor.execute("""
                SELECT id, name FROM ea_dy_legal 
                WHERE (delete_time IS NULL OR delete_time = 0)
            """)
            ea_legal_map = {}
            for row in target_cursor.fetchall():
                ea_legal_map[row['name']] = row['id']
            self.log_message.emit(f"财务系统共 {len(ea_legal_map)} 个法人")

            self.log_message.emit("正在加载财务系统店铺类型列表...")
            target_cursor.execute("""
                SELECT id, name FROM ea_dy_site_type 
                WHERE (delete_time IS NULL OR delete_time = 0)
            """)
            ea_site_type_map = {}
            for row in target_cursor.fetchall():
                ea_site_type_map[row['name']] = row['id']
            self.log_message.emit(f"财务系统共 {len(ea_site_type_map)} 个店铺类型")
            self.progress.emit(20)

            success_count = 0
            fail_count = 0
            skip_count = 0
            multi_match_count = 0

            for index, mag in enumerate(mag_records):
                try:
                    result_row = {
                        'ba_account_mag_id': mag['id'],
                        'inventory_num': mag['inventory_num'],
                        'legal_id': mag['legal_id'],
                        'platform_id': mag['platform_id'],
                    }

                    # 1. 根据legal_id查ba_rlb_customer获取legal_name
                    legal_name = None
                    if mag['legal_id']:
                        source_cursor.execute("""
                            SELECT legal_name FROM ba_rlb_customer 
                            WHERE id = %s
                        """, (mag['legal_id'],))
                        cust = source_cursor.fetchone()
                        if cust:
                            legal_name = cust['legal_name']

                    result_row['legal_name'] = legal_name or ''

                    if not legal_name:
                        skip_count += 1
                        result_row['处理状态'] = '跳过(法人ID无对应客户)'
                        self.results['completed_data'].append(result_row)
                        continue

                    # 2. 根据platform_id查ba_platform获取platform名称
                    platform_name = None
                    if mag['platform_id']:
                        source_cursor.execute("""
                            SELECT platform FROM ba_platform 
                            WHERE id = %s
                        """, (mag['platform_id'],))
                        plat = source_cursor.fetchone()
                        if plat:
                            platform_name = plat['platform']

                    result_row['platform_name'] = platform_name or ''

                    if not platform_name:
                        skip_count += 1
                        result_row['处理状态'] = '跳过(平台ID无对应平台)'
                        self.results['completed_data'].append(result_row)
                        continue

                    # 3. 在财务系统中查找法人ID
                    ea_legal_id = ea_legal_map.get(legal_name)
                    result_row['ea_legal_id'] = ea_legal_id or ''

                    if not ea_legal_id:
                        fail_count += 1
                        result_row['处理状态'] = f'失败(财务系统未找到法人: {legal_name})'
                        self.results['completed_data'].append(result_row)
                        self.results['failed_records'].append(result_row)
                        continue

                    # 4. 在财务系统中查找店铺类型ID
                    ea_site_type_id = ea_site_type_map.get(platform_name)
                    result_row['ea_site_type_id'] = ea_site_type_id or ''

                    if not ea_site_type_id:
                        fail_count += 1
                        result_row['处理状态'] = f'失败(财务系统未找到店铺类型: {platform_name})'
                        self.results['completed_data'].append(result_row)
                        self.results['failed_records'].append(result_row)
                        continue

                    # 5. 在ea_dy_shop中找到匹配的记录
                    target_cursor.execute("""
                        SELECT id, name, stock_number, status
                        FROM ea_dy_shop 
                        WHERE legal_id = %s AND siteType = %s
                          AND (delete_time IS NULL OR delete_time = 0)
                    """, (ea_legal_id, ea_site_type_id))
                    shops = target_cursor.fetchall()

                    if not shops:
                        fail_count += 1
                        result_row['处理状态'] = f'失败(财务系统未找到店铺: 法人ID={ea_legal_id}, 类型ID={ea_site_type_id})'
                        self.results['completed_data'].append(result_row)
                        self.results['failed_records'].append(result_row)
                        continue

                    if len(shops) > 1:
                        multi_match_count += 1
                        self.log_message.emit(
                            f"ID={mag['id']}: 法人={legal_name}, 类型={platform_name} "
                            f"匹配到 {len(shops)} 条店铺记录，全部填充"
                        )

                    # 6. 对所有匹配的店铺填充stock_number
                    shop_ids = []
                    shop_names = []
                    shop_old_stocks = []
                    for shop in shops:
                        shop_ids.append(str(shop['id']))
                        shop_names.append(shop['name'] or '')
                        shop_old_stocks.append(shop['stock_number'] or '')

                        if not self.preview_only:
                            target_cursor.execute(
                                "UPDATE ea_dy_shop SET stock_number = %s WHERE id = %s",
                                (mag['inventory_num'], shop['id'])
                            )

                    if not self.preview_only:
                        target_conn.commit()

                    result_row['ea_shop_ids'] = ','.join(shop_ids)
                    result_row['ea_shop_names'] = ','.join(shop_names)
                    result_row['ea_shop_old_stocks'] = ','.join(shop_old_stocks)
                    result_row['匹配店铺数'] = len(shops)

                    success_count += 1
                    if self.preview_only:
                        result_row['处理状态'] = f'可填充({len(shops)}条)'
                    else:
                        result_row['处理状态'] = f'成功({len(shops)}条)'

                    self.results['completed_data'].append(result_row)

                    if success_count % 100 == 0:
                        self.log_message.emit(f"已处理 {success_count} 条...")

                except Exception as e:
                    fail_count += 1
                    err_row = {
                        'ba_account_mag_id': mag['id'],
                        'inventory_num': mag['inventory_num'],
                        '处理状态': f'异常({str(e)})'
                    }
                    self.results['completed_data'].append(err_row)
                    self.results['failed_records'].append(err_row)

                progress = 20 + int((index + 1) / total * 75)
                self.progress.emit(progress)

            source_cursor.close()
            source_conn.close()
            target_cursor.close()
            target_conn.close()

            self.results['success_count'] = success_count
            self.results['fail_count'] = fail_count
            self.results['skip_count'] = skip_count
            self.results['multi_match_count'] = multi_match_count

            self.progress.emit(100)
            self.log_message.emit(f"\n=== {mode_text}完成 ===")
            self.log_message.emit(f"总记录: {total}")
            self.log_message.emit(f"{'可填充' if self.preview_only else '成功'}: {success_count}")
            self.log_message.emit(f"失败: {fail_count}")
            self.log_message.emit(f"跳过: {skip_count}")
            self.log_message.emit(f"多条匹配: {multi_match_count}")

            self.finished.emit(self.results)

        except Exception as e:
            self.error.emit(str(e))


class SyncInventoryToFinanceWidget(QWidget):
    """同步进销存库存编号到财务系统界面"""

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
        ds_group = QGroupBox("选择数据源")
        ds_layout = QFormLayout()

        self.source_datasource_combo = QComboBox()
        self.source_datasource_combo.currentIndexChanged.connect(self.on_datasource_changed)
        ds_layout.addRow("进销存数据库 (ba_表):", self.source_datasource_combo)

        self.test_source_btn = QPushButton("测试连接")
        self.test_source_btn.setEnabled(False)
        self.test_source_btn.clicked.connect(self.test_source_connection)
        ds_layout.addRow("", self.test_source_btn)

        self.target_datasource_combo = QComboBox()
        self.target_datasource_combo.currentIndexChanged.connect(self.on_datasource_changed)
        ds_layout.addRow("财务系统数据库 (ea_表):", self.target_datasource_combo)

        self.test_target_btn = QPushButton("测试连接")
        self.test_target_btn.setEnabled(False)
        self.test_target_btn.clicked.connect(self.test_target_connection)
        ds_layout.addRow("", self.test_target_btn)

        ds_group.setLayout(ds_layout)
        layout.addWidget(ds_group)

        # 功能说明
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        info_label = QLabel("""
<b>同步进销存库存编号到财务系统:</b><br><br>
<b>处理逻辑:</b><br>
• 1. 遍历 ba_account_mag 中有库存编号(inventory_num)的记录<br>
• 2. 根据 legal_id → ba_rlb_customer 获取法人姓名(legal_name)<br>
• 3. 根据 platform_id → ba_platform 获取平台名称(platform)<br>
• 4. 用法人姓名在 ea_dy_legal 中查找对应ID<br>
• 5. 用平台名称在 ea_dy_site_type 中查找对应ID<br>
• 6. 用这两个ID在 ea_dy_shop 中找到对应店铺<br>
• 7. 将 inventory_num 填充到 ea_dy_shop.stock_number<br><br>
<b>建议先执行「预检查」确认匹配情况后再执行「开始同步」</b>
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

        self.start_btn = QPushButton("开始同步")
        self.start_btn.setEnabled(False)
        self.start_btn.clicked.connect(self.start_sync)
        btn_layout.addWidget(self.start_btn)

        self.export_completed_btn = QPushButton("导出完整记录")
        self.export_completed_btn.setEnabled(False)
        self.export_completed_btn.clicked.connect(self.export_completed)
        btn_layout.addWidget(self.export_completed_btn)

        self.export_failed_btn = QPushButton("导出失败记录")
        self.export_failed_btn.setEnabled(False)
        self.export_failed_btn.clicked.connect(self.export_failed)
        btn_layout.addWidget(self.export_failed_btn)

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
        self.source_datasource_combo.clear()
        self.target_datasource_combo.clear()

        self.source_datasource_combo.addItem("请选择进销存数据源", None)
        self.target_datasource_combo.addItem("请选择财务系统数据源", None)

        datasources = self.db_manager.get_all_datasources()
        for ds in datasources:
            display = f"{ds.name} ({ds.host}:{ds.port})"
            self.source_datasource_combo.addItem(display, ds)
            self.target_datasource_combo.addItem(display, ds)

    def on_datasource_changed(self):
        source_ok = self.source_datasource_combo.currentData() is not None
        target_ok = self.target_datasource_combo.currentData() is not None
        self.test_source_btn.setEnabled(source_ok)
        self.test_target_btn.setEnabled(target_ok)
        ready = source_ok and target_ok
        self.precheck_btn.setEnabled(ready)
        self.start_btn.setEnabled(ready)

    def test_source_connection(self):
        ds = self.source_datasource_combo.currentData()
        if not ds:
            return
        ok, msg = self.db_manager.test_connection(ds)
        if ok:
            QMessageBox.information(self, "连接成功", msg)
        else:
            QMessageBox.critical(self, "连接失败", msg)

    def test_target_connection(self):
        ds = self.target_datasource_combo.currentData()
        if not ds:
            return
        ok, msg = self.db_manager.test_connection(ds)
        if ok:
            QMessageBox.information(self, "连接成功", msg)
        else:
            QMessageBox.critical(self, "连接失败", msg)

    def start_precheck(self):
        source_ds = self.source_datasource_combo.currentData()
        target_ds = self.target_datasource_combo.currentData()
        if not source_ds or not target_ds:
            QMessageBox.warning(self, "警告", "请选择两个数据源！")
            return

        self._run_worker(source_ds, target_ds, preview_only=True)

    def start_sync(self):
        source_ds = self.source_datasource_combo.currentData()
        target_ds = self.target_datasource_combo.currentData()
        if not source_ds or not target_ds:
            QMessageBox.warning(self, "警告", "请选择两个数据源！")
            return

        reply = QMessageBox.question(
            self, "确认同步",
            f"确定要同步库存编号到财务系统吗？\n\n"
            f"进销存: {source_ds.name}\n"
            f"财务系统: {target_ds.name}\n\n"
            f"将把 ba_account_mag.inventory_num 更新到 ea_dy_shop.stock_number\n"
            f"操作不可撤销，请确认！",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        self._run_worker(source_ds, target_ds, preview_only=False)

    def _run_worker(self, source_ds, target_ds, preview_only):
        self.precheck_btn.setEnabled(False)
        self.start_btn.setEnabled(False)
        self.export_completed_btn.setEnabled(False)
        self.export_failed_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在处理..." if not preview_only else "正在预检查...")
        self.log_text.clear()
        self.completed_data = []
        self.failed_records = []

        self.worker = SyncInventoryToFinanceWorker(source_ds, target_ds, preview_only=preview_only)
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
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)

        self.completed_data = results.get('completed_data', [])
        self.failed_records = results.get('failed_records', [])

        total = results['total']
        success = results['success_count']
        fail = results['fail_count']
        skip = results['skip_count']
        multi = results['multi_match_count']

        text = f"总{total}条 | 成功: {success} | 失败: {fail} | 跳过: {skip} | 多匹配: {multi}"
        if fail > 0:
            self.result_label.setStyleSheet("color: orange;")
        else:
            self.result_label.setStyleSheet("color: green;")
        self.result_label.setText(text)

        self.export_completed_btn.setEnabled(len(self.completed_data) > 0)
        self.export_failed_btn.setEnabled(len(self.failed_records) > 0)

    def on_error(self, error_msg):
        self.precheck_btn.setEnabled(True)
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.result_label.setText("处理失败")
        self.result_label.setStyleSheet("color: red;")
        self.append_log(f"错误: {error_msg}")
        QMessageBox.critical(self, "错误", f"处理失败:\n{error_msg}")

    def export_completed(self):
        if not self.completed_data:
            QMessageBox.warning(self, "警告", "没有可导出的数据！")
            return
        default_name = f"库存编号同步完整记录_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "保存完整记录", default_name, "Excel文件 (*.xlsx)")
        if path:
            df = pd.DataFrame(self.completed_data)
            df.to_excel(path, index=False, engine='openpyxl')
            QMessageBox.information(self, "成功", f"已导出 {len(self.completed_data)} 条记录到:\n{path}")

    def export_failed(self):
        if not self.failed_records:
            QMessageBox.warning(self, "警告", "没有失败记录！")
            return
        default_name = f"库存编号同步失败记录_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "保存失败记录", default_name, "Excel文件 (*.xlsx)")
        if path:
            df = pd.DataFrame(self.failed_records)
            df.to_excel(path, index=False, engine='openpyxl')
            QMessageBox.information(self, "成功", f"已导出 {len(self.failed_records)} 条失败记录到:\n{path}")
