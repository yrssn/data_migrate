"""
注册表属性相同检测
动态选择 ba_ptzcb_register 表的字段进行重复数据检查
platform_id 字段特殊处理：对应 ba_platform.is_repeatable=1 的平台记录跳过
"""
import os
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog,
                             QCheckBox, QScrollArea, QGridLayout)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from database import DatabaseManager
import pandas as pd
import pymysql


class RegisterAttrDuplicateWorker(QThread):
    """注册表属性相同检测工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, datasource, selected_fields, output_file):
        super().__init__()
        self.datasource = datasource
        self.selected_fields = selected_fields  # 用户选择的检查字段列表
        self.output_file = output_file

    def run(self):
        try:
            self.log_message.emit("正在连接数据库...")
            self.progress.emit(5)

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
            self.progress.emit(10)

            # 如果选择的字段包含 platform_id，先查询可重复的平台
            repeatable_platform_ids = set()
            if 'platform_id' in self.selected_fields:
                self.log_message.emit("正在查询可重复平台 (is_repeatable=1)...")
                cursor.execute("""
                    SELECT id, platform FROM ba_platform 
                    WHERE is_repeatable = 1 AND delete_time IS NULL
                """)
                repeatable_rows = cursor.fetchall()
                repeatable_platform_ids = {row['id'] for row in repeatable_rows}
                if repeatable_rows:
                    names = ', '.join(row['platform'] for row in repeatable_rows)
                    self.log_message.emit(
                        f"找到 {len(repeatable_rows)} 个可重复平台: {names}"
                    )
                    self.log_message.emit("这些平台的记录将被跳过，不参与重复检测")
                else:
                    self.log_message.emit("没有可重复平台")

            self.progress.emit(15)

            # 构建字段列表
            fields_str = ', '.join(self.selected_fields)
            self.log_message.emit(f"检测字段: {fields_str}")

            # 查询所有未删除的注册表记录
            self.log_message.emit("正在查询 ba_ptzcb_register 数据...")
            cursor.execute("""
                SELECT * FROM ba_ptzcb_register
                WHERE delete_time IS NULL
            """)
            all_records = cursor.fetchall()
            total_records = len(all_records)
            self.log_message.emit(f"共查询到 {total_records} 条未删除记录")
            self.progress.emit(30)

            # 过滤：如果选了platform_id，排除可重复平台的记录
            filtered_records = []
            skipped_count = 0
            if 'platform_id' in self.selected_fields and repeatable_platform_ids:
                for rec in all_records:
                    if rec.get('platform_id') in repeatable_platform_ids:
                        skipped_count += 1
                    else:
                        filtered_records.append(rec)
                self.log_message.emit(
                    f"跳过 {skipped_count} 条可重复平台记录，剩余 {len(filtered_records)} 条参与检测"
                )
            else:
                filtered_records = all_records

            self.progress.emit(40)

            # 按选择的字段分组，找出重复的组合
            # 任意一个检测字段为空的记录跳过，不参与重复检测
            self.log_message.emit("正在检测重复数据...")
            groups = {}
            null_skipped = 0
            for rec in filtered_records:
                # 检查是否有任意选择字段为空
                has_null = False
                key_parts = []
                for field in self.selected_fields:
                    val = rec.get(field)
                    if val is None or str(val).strip() == '':
                        has_null = True
                        break
                    key_parts.append(str(val))

                if has_null:
                    null_skipped += 1
                    continue

                key = '||'.join(key_parts)
                if key not in groups:
                    groups[key] = []
                groups[key].append(rec)

            if null_skipped > 0:
                self.log_message.emit(f"跳过 {null_skipped} 条检测字段含空值的记录")

            self.progress.emit(60)

            # 筛选出有重复的组（记录数>1）
            duplicate_groups = {}
            for key, recs in groups.items():
                if len(recs) > 1:
                    duplicate_groups[key] = recs

            if not duplicate_groups:
                self.log_message.emit("未发现重复数据")
                cursor.close()
                connection.close()
                self.progress.emit(100)
                self.finished.emit({
                    'total_records': total_records,
                    'skipped_records': skipped_count,
                    'checked_records': len(filtered_records),
                    'duplicate_groups': 0,
                    'duplicate_records': 0
                })
                return

            # 查询平台名称用于展示
            platform_map = {}
            cursor.execute("SELECT id, platform FROM ba_platform WHERE delete_time IS NULL")
            for row in cursor.fetchall():
                platform_map[row['id']] = row['platform']

            # 查询状态映射
            status_map = {
                0: '待指派', 1: '进行中', 2: '审核中', 3: '银行验证',
                4: '审核失败', 5: '银行卡验证失败', 6: '完成', 7: '失败',
                8: '取消', 9: '暂停', 10: '店铺验证失败', 11: '店铺验证中'
            }

            # 收集状态=完成的记录ID，通过ba_account_mag.register_id关联查询库存信息
            self.log_message.emit("正在查询关联库存信息...")
            completed_reg_ids = set()
            for recs in duplicate_groups.values():
                for rec in recs:
                    if rec.get('status') == 6:
                        completed_reg_ids.add(rec['id'])

            # key=register_id -> ba_account_mag info
            account_mag_by_reg = {}
            if completed_reg_ids:
                id_list = ','.join(str(rid) for rid in completed_reg_ids)
                cursor.execute(f"""
                    SELECT id, register_id, inventory_num, delete_time
                    FROM ba_account_mag
                    WHERE register_id IN ({id_list})
                """)
                for row in cursor.fetchall():
                    dt = row.get('delete_time')
                    if dt:
                        try:
                            dt_str = datetime.fromtimestamp(
                                dt / 1000 if dt > 9999999999 else dt
                            ).strftime('%Y-%m-%d %H:%M:%S')
                        except:
                            dt_str = str(dt)
                    else:
                        dt_str = ''
                    account_mag_by_reg[row['register_id']] = {
                        'mag_id': row['id'],
                        'inventory_num': row.get('inventory_num', ''),
                        'mag_delete_time': dt_str
                    }
                self.log_message.emit(f"查询到 {len(account_mag_by_reg)} 条关联库存记录")

            cursor.close()
            connection.close()
            self.progress.emit(70)

            # 构建导出数据
            export_rows = []
            group_num = 0
            for key, recs in duplicate_groups.items():
                group_num += 1
                for rec in recs:
                    row = {'分组编号': group_num}
                    row['id'] = rec['id']
                    row['project_id'] = rec.get('project_id', '')
                    row['status'] = rec.get('status', '')
                    row['status_text'] = status_map.get(rec.get('status'), '')
                    row['current_progress'] = rec.get('current_progress', '')
                    row['email'] = rec.get('email', '')
                    row['shopindex_id'] = rec.get('shopindex_id', '')
                    row['shudi_id'] = rec.get('shudi_id', '')
                    row['platform_id'] = rec.get('platform_id', '')
                    row['platform_name'] = platform_map.get(rec.get('platform_id'), '')
                    row['company_name'] = rec.get('company_name', '')
                    row['bank_card_type'] = rec.get('bank_card_type', '')
                    row['remark'] = rec.get('remark', '')
                    row['account_mag_id'] = rec.get('account_mag_id', '')

                    # 状态=完成的记录，通过register_id关联补充库存信息
                    reg_id = rec['id']
                    if rec.get('status') == 6 and reg_id in account_mag_by_reg:
                        mag_info = account_mag_by_reg[reg_id]
                        row['库存_id'] = mag_info['mag_id']
                        row['库存_inventory_num'] = mag_info['inventory_num']
                        row['库存_delete_time'] = mag_info['mag_delete_time']
                    else:
                        row['库存_id'] = ''
                        row['库存_inventory_num'] = ''
                        row['库存_delete_time'] = ''

                    # 标记检查字段的值
                    for field in self.selected_fields:
                        row[f'检查_{field}'] = rec.get(field, '')

                    ct = rec.get('create_time')
                    if ct:
                        try:
                            row['create_time'] = datetime.fromtimestamp(
                                ct / 1000 if ct > 9999999999 else ct
                            ).strftime('%Y-%m-%d %H:%M:%S')
                        except:
                            row['create_time'] = str(ct)
                    else:
                        row['create_time'] = ''

                    export_rows.append(row)

            self.progress.emit(85)

            total_dup_records = len(export_rows)
            total_dup_groups = group_num

            self.log_message.emit(f"发现 {total_dup_groups} 组重复数据，共 {total_dup_records} 条记录")

            # 导出Excel
            self.log_message.emit(f"正在导出到 {self.output_file}...")
            df = pd.DataFrame(export_rows)
            df.to_excel(self.output_file, index=False, engine='openpyxl')
            self.progress.emit(100)

            self.log_message.emit(f"导出完成！")
            self.finished.emit({
                'total_records': total_records,
                'skipped_records': skipped_count,
                'checked_records': len(filtered_records),
                'duplicate_groups': total_dup_groups,
                'duplicate_records': total_dup_records
            })

        except Exception as e:
            self.error.emit(str(e))


class RegisterAttrDuplicateCheckWidget(QWidget):
    """注册表属性相同检测界面"""

    # ba_ptzcb_register 可选检测字段（排除id、时间等不适合做重复检测的字段）
    CHECKABLE_FIELDS = [
        ('shopindex_id', '主体编号ID'),
        ('shudi_id', '属地ID'),
        ('platform_id', '平台ID'),
        ('email', '邮箱'),
        ('company_name', '公司名称'),
        ('country', '国家'),
        ('bank_card_type', '银行卡类型'),
        ('currency_id', '货币ID'),
        ('customer_id', '客户档案ID'),
        ('information_id', '日区企业法人ID'),
        ('information_part2_id', '欧美企业法人ID'),
        ('order_id', '订单关联'),
        ('account_mag_id', 'account_mag表ID'),
    ]

    def __init__(self, db_manager=None):
        super().__init__()
        self.db_manager = db_manager or DatabaseManager()
        self.output_file = None
        self.worker = None
        self.field_checkboxes = {}
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

        # 字段选择区域
        field_group = QGroupBox("选择检测字段 (选择至少一个字段进行重复检测)")
        field_scroll = QScrollArea()
        field_scroll.setWidgetResizable(True)
        field_scroll.setMaximumHeight(200)

        field_container = QWidget()
        field_grid = QGridLayout(field_container)
        field_grid.setSpacing(8)

        for i, (field_name, field_label) in enumerate(self.CHECKABLE_FIELDS):
            cb = QCheckBox(f"{field_label} ({field_name})")
            cb.stateChanged.connect(self.update_start_state)
            self.field_checkboxes[field_name] = cb
            row = i // 3
            col = i % 3
            field_grid.addWidget(cb, row, col)

        field_scroll.setWidget(field_container)

        field_btn_layout = QHBoxLayout()
        select_all_btn = QPushButton("全选")
        select_all_btn.clicked.connect(lambda: self._set_all_checkboxes(True))
        field_btn_layout.addWidget(select_all_btn)

        deselect_all_btn = QPushButton("全不选")
        deselect_all_btn.clicked.connect(lambda: self._set_all_checkboxes(False))
        field_btn_layout.addWidget(deselect_all_btn)

        field_btn_layout.addStretch()

        field_inner_layout = QVBoxLayout()
        field_inner_layout.addWidget(field_scroll)
        field_inner_layout.addLayout(field_btn_layout)
        field_group.setLayout(field_inner_layout)
        layout.addWidget(field_group)

        # 提示信息
        tip_label = QLabel(
            "💡 提示: 如果选择了「平台ID」字段，ba_platform 表中 is_repeatable=1 的平台对应的记录会被自动跳过"
        )
        tip_label.setWordWrap(True)
        tip_label.setStyleSheet("color: #666; font-size: 11px; padding: 4px;")
        layout.addWidget(tip_label)

        # 输出文件选择
        file_group = QGroupBox("输出文件")
        file_layout = QFormLayout()

        file_row = QHBoxLayout()
        self.file_label = QLabel("未选择文件")
        self.file_label.setStyleSheet("color: gray;")
        file_row.addWidget(self.file_label)
        self.select_file_btn = QPushButton("选择保存位置")
        self.select_file_btn.clicked.connect(self.select_output_file)
        file_row.addWidget(self.select_file_btn)
        file_layout.addRow("Excel文件:", file_row)

        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # 操作按钮
        btn_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始检测")
        self.start_btn.setEnabled(False)
        self.start_btn.clicked.connect(self.start_check)
        btn_layout.addWidget(self.start_btn)
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
        self.log_text.setMaximumHeight(200)
        log_layout.addWidget(self.log_text)

        clear_btn = QPushButton("清空日志")
        clear_btn.clicked.connect(self.log_text.clear)
        log_layout.addWidget(clear_btn)

        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        layout.addStretch()

    def _set_all_checkboxes(self, checked):
        for cb in self.field_checkboxes.values():
            cb.setChecked(checked)

    def load_datasources(self):
        self.datasource_combo.clear()
        self.datasource_combo.addItem("请选择数据源", None)
        datasources = self.db_manager.get_all_datasources()
        for ds in datasources:
            self.datasource_combo.addItem(f"{ds.name} ({ds.host}:{ds.port})", ds)
        self.update_start_state()

    def on_datasource_changed(self):
        self.test_conn_btn.setEnabled(self.datasource_combo.currentData() is not None)
        self.update_start_state()

    def test_connection(self):
        ds = self.datasource_combo.currentData()
        if not ds:
            return
        ok, msg = self.db_manager.test_connection(ds)
        if ok:
            QMessageBox.information(self, "连接成功", msg)
        else:
            QMessageBox.critical(self, "连接失败", msg)

    def get_selected_fields(self):
        """获取用户选择的字段列表"""
        fields = []
        for field_name, cb in self.field_checkboxes.items():
            if cb.isChecked():
                fields.append(field_name)
        return fields

    def select_output_file(self):
        default_name = f"注册表属性相同检测_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(
            self, "选择保存位置", default_name, "Excel文件 (*.xlsx)"
        )
        if path:
            self.output_file = path
            self.file_label.setText(os.path.basename(path))
            self.file_label.setStyleSheet("color: green;")
            self.update_start_state()

    def update_start_state(self):
        ds_ok = self.datasource_combo.currentData() is not None
        file_ok = self.output_file is not None
        fields_ok = len(self.get_selected_fields()) > 0
        self.start_btn.setEnabled(ds_ok and file_ok and fields_ok)

    def start_check(self):
        ds = self.datasource_combo.currentData()
        selected = self.get_selected_fields()
        if not ds or not self.output_file or not selected:
            QMessageBox.warning(self, "警告", "请选择数据源、检测字段和输出文件！")
            return

        # 构建字段描述
        field_desc = []
        for field_name in selected:
            for fn, fl in self.CHECKABLE_FIELDS:
                if fn == field_name:
                    field_desc.append(f"  - {fl} ({fn})")
                    break

        reply = QMessageBox.question(
            self, "确认检测",
            f"确定要检测 ba_ptzcb_register 表的属性重复吗？\n\n"
            f"数据库: {ds.name}\n"
            f"检测字段:\n" + '\n'.join(field_desc) + '\n\n'
            f"输出文件: {os.path.basename(self.output_file)}\n\n"
            + ("注意: 选择了platform_id，可重复平台(is_repeatable=1)的记录将被跳过\n" if 'platform_id' in selected else ''),
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        self.start_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在检测...")
        self.log_text.clear()

        self.worker = RegisterAttrDuplicateWorker(ds, selected, self.output_file)
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log_message.connect(self.append_log)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def append_log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{ts}] {msg}")

    def on_finished(self, results):
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)

        total = results['total_records']
        skipped = results['skipped_records']
        checked = results['checked_records']
        dup_groups = results['duplicate_groups']
        dup_records = results['duplicate_records']

        if dup_groups == 0:
            text = f"未发现重复数据 (总{total}条, 跳过{skipped}条, 检测{checked}条)"
            self.result_label.setStyleSheet("color: green;")
        else:
            text = f"发现 {dup_groups} 组重复，共 {dup_records} 条 (总{total}, 跳过{skipped}, 检测{checked})"
            self.result_label.setStyleSheet("color: orange;")

        self.result_label.setText(text)
        self.append_log(f"\n=== 检测完成 ===\n{text}")

        if dup_groups > 0:
            QMessageBox.information(
                self, "完成",
                f"检测完成！\n\n{text}\n\n结果已导出到:\n{self.output_file}"
            )

    def on_error(self, error_msg):
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.result_label.setText("检测失败")
        self.result_label.setStyleSheet("color: red;")
        self.append_log(f"错误: {error_msg}")
        QMessageBox.critical(self, "错误", f"检测失败:\n{error_msg}")
