"""
同步库存账号V2功能界面
- 状态映射：用户自定义Excel值到状态码的映射
- 银行：只根据银行卡号查找
- 默认状态：可售(3)
"""
import os
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog,
                             QLineEdit, QTableWidget, QTableWidgetItem, QHeaderView)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from database import DatabaseManager
import pandas as pd
import pymysql


class SyncInventoryV2Worker(QThread):
    """同步库存账号V2工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, datasource, excel_file, ptzcb_column, bank_card_column,
                 status_column, status_mapping, register_department):
        super().__init__()
        self.datasource = datasource
        self.excel_file = excel_file
        self.ptzcb_column = ptzcb_column
        self.bank_card_column = bank_card_column
        self.status_column = status_column
        self.status_mapping = status_mapping  # {excel_value: status_code}
        self.register_department = register_department
        self.results = {
            'total_rows': 0,
            'success_count': 0,
            'failed_count': 0,
            'failed_records': [],
            'completed_data': []
        }
        self.inventory_counters = {}

    def run(self):
        try:
            self.log_message.emit("开始读取Excel文件...")
            self.progress.emit(5)

            df = pd.read_excel(self.excel_file, engine='openpyxl')
            self.log_message.emit(f"成功读取Excel文件，共 {len(df)} 行数据")
            self.results['total_rows'] = len(df)
            self.progress.emit(10)

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
            self.progress.emit(15)

            for index, row in df.iterrows():
                try:
                    inventory_num, account_mag_id = self.process_row(cursor, index + 1, row)
                    self.results['success_count'] += 1

                    completed_row = row.to_dict()
                    completed_row['生成的库存编号'] = inventory_num
                    completed_row['库存账号主键ID'] = account_mag_id
                    completed_row['处理状态'] = '成功'
                    completed_row['处理时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    self.results['completed_data'].append(completed_row)

                    progress = 15 + int((index + 1) / len(df) * 80)
                    self.progress.emit(progress)

                    if (index + 1) % 50 == 0:
                        connection.commit()
                        self.log_message.emit(f"已处理 {index + 1} 行，提交事务")

                except Exception as e:
                    self.results['failed_count'] += 1
                    error_msg = f"第{index + 1}行处理失败: {str(e)}"
                    self.log_message.emit(error_msg)

                    failed_row = row.to_dict()
                    failed_row['生成的库存编号'] = ''
                    failed_row['库存账号主键ID'] = ''
                    failed_row['处理状态'] = '失败'
                    failed_row['失败原因'] = str(e)
                    failed_row['处理时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    self.results['completed_data'].append(failed_row)

                    self.results['failed_records'].append({
                        'row': index + 1,
                        'error': str(e),
                        'data': row.to_dict()
                    })

            connection.commit()
            cursor.close()
            connection.close()

            self.progress.emit(100)
            self.finished.emit(self.results)

        except Exception as e:
            self.error.emit(str(e))

    def process_row(self, cursor, row_num, row):
        """处理单行数据"""
        ptzcb_value = str(row[self.ptzcb_column]).strip() if self.ptzcb_column in row.index else ""
        bank_card_no = str(row[self.bank_card_column]).strip() if self.bank_card_column in row.index else ""
        status_value = str(row[self.status_column]).strip() if self.status_column in row.index else ""

        if not ptzcb_value or ptzcb_value == 'nan':
            raise Exception("PTZCB查找列为空")

        self.log_message.emit(f"第{row_num}行: PTZCB={ptzcb_value}, 银行卡={bank_card_no}, 状态={status_value}")

        # 查找ptzcb记录
        ptzcb_record = self.find_ptzcb_record(cursor, ptzcb_value)
        if not ptzcb_record:
            raise Exception("未找到对应的ba_ptzcb_register记录")

        # 获取平台和属地
        platform_info = self.get_platform_info(cursor, ptzcb_record['platform_id'])
        shudi_info = self.get_shudi_info(cursor, ptzcb_record['shudi_id'])

        # 生成库存编号
        inventory_num = self.generate_inventory_num(platform_info['code'], shudi_info['territory_abbreviation'])

        # 银行：只根据银行卡号查找
        bank_id = self.find_bank_by_card(cursor, bank_card_no, row_num)

        # 货币ID
        currency_id = self.get_currency_id(cursor, bank_id) if bank_id else None

        # 状态转换：用用户自定义映射，默认可售(3)
        account_status = self.status_mapping.get(status_value, 3)

        # 插入
        account_mag_id = self.insert_account_mag(cursor, {
            'inventory_num': inventory_num,
            'shudi_id': ptzcb_record['shudi_id'],
            'platform_id': ptzcb_record['platform_id'],
            'register_id': ptzcb_record['id'],
            'account_status': account_status,
            'shopindex_id': ptzcb_record['shopindex_id'],
            'bank_id': bank_id,
            'order_id': ptzcb_record['order_id'],
            'currency_id': currency_id,
            'legal_id': ptzcb_record['customer_id'],
            'information_id': ptzcb_record['information_id'],
            'information_part2_id': ptzcb_record['information_part2_id'],
            'register_department': self.register_department,
            'source_reg_id': 1
        })

        self.log_message.emit(f"第{row_num}行: 成功 {inventory_num} (ID: {account_mag_id})")
        return inventory_num, account_mag_id

    def find_ptzcb_record(self, cursor, ptzcb_value):
        cursor.execute("""
            SELECT id, shudi_id, platform_id, shopindex_id, order_id, customer_id,
                   information_id, information_part2_id
            FROM ba_ptzcb_register
            WHERE project_id = %s OR id = %s
        """, (ptzcb_value, ptzcb_value))
        return cursor.fetchone()

    def get_platform_info(self, cursor, platform_id):
        cursor.execute("SELECT code FROM ba_platform WHERE id = %s", (platform_id,))
        result = cursor.fetchone()
        if not result:
            raise Exception(f"未找到平台ID {platform_id} 的信息")
        return result

    def get_shudi_info(self, cursor, shudi_id):
        cursor.execute("SELECT territory_abbreviation FROM ba_shudi WHERE id = %s", (shudi_id,))
        result = cursor.fetchone()
        if not result:
            raise Exception(f"未找到属地ID {shudi_id} 的信息")
        return result

    def generate_inventory_num(self, platform_code, territory_abbr):
        today = datetime.now()
        year_month = today.strftime('%y%m')
        group_key = f"{platform_code}-{territory_abbr}-{year_month}"
        if group_key not in self.inventory_counters:
            self.inventory_counters[group_key] = 0
        self.inventory_counters[group_key] += 1
        sequence = str(self.inventory_counters[group_key]).zfill(4)
        return f"{platform_code}-{territory_abbr}-{year_month}-{sequence}"

    def find_bank_by_card(self, cursor, bank_card_no, row_num):
        """只根据银行卡号查找银行ID"""
        if not bank_card_no or bank_card_no.strip() == '' or bank_card_no == 'nan':
            return None
        cursor.execute("""
            SELECT id FROM ba_zhb_bank
            WHERE bank_card_number = %s
            ORDER BY update_time DESC
            LIMIT 1
        """, (bank_card_no,))
        result = cursor.fetchone()
        if result:
            self.log_message.emit(f"  第{row_num}行: 银行卡 {bank_card_no} -> 银行ID: {result['id']}")
        return result['id'] if result else None

    def get_currency_id(self, cursor, bank_id):
        if not bank_id:
            return None
        cursor.execute("SELECT currency_id FROM ba_zhb_bank WHERE id = %s", (bank_id,))
        result = cursor.fetchone()
        return result['currency_id'] if result else None

    def insert_account_mag(self, cursor, data):
        current_time = int(datetime.now().timestamp())
        cursor.execute("""
            INSERT INTO ba_account_mag (
                inventory_num, shudi_id, platform_id, register_id, account_status,
                shopindex_id, bank_id, order_id, currency_id, legal_id,
                information_id, information_part2_id, register_department, source_reg_id,
                remark, create_time, update_time
            ) VALUES (
                %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s
            )
        """, (
            data['inventory_num'], data['shudi_id'], data['platform_id'], data['register_id'],
            data['account_status'], data['shopindex_id'], data['bank_id'], data['order_id'],
            data['currency_id'], data['legal_id'], data['information_id'], data['information_part2_id'],
            data['register_department'], data['source_reg_id'], '系统脚本补全V2', current_time, current_time
        ))
        return cursor.lastrowid


class SyncInventoryV2Widget(QWidget):
    """同步库存账号V2功能组件"""

    STATUS_OPTIONS = [
        (0, '待更换密码'),
        (1, '过审暂停'),
        (2, '出售中'),
        (3, '可售'),
        (4, '已售'),
        (5, '封号'),
    ]

    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.excel_file = None
        self.excel_columns = []
        self.completed_data = []
        self.failed_records = []
        self.status_unique_values = []
        self.init_ui()
        self.load_datasources()

    def init_ui(self):
        layout = QVBoxLayout()

        # 数据源
        ds_group = QGroupBox("选择目标数据库")
        ds_layout = QFormLayout()
        self.datasource_combo = QComboBox()
        self.datasource_combo.currentTextChanged.connect(self.on_datasource_changed)
        ds_layout.addRow("数据源:", self.datasource_combo)
        self.test_conn_btn = QPushButton("测试连接")
        self.test_conn_btn.clicked.connect(self.test_connection)
        self.test_conn_btn.setEnabled(False)
        ds_layout.addRow("", self.test_conn_btn)
        ds_group.setLayout(ds_layout)
        layout.addWidget(ds_group)

        # Excel文件
        file_group = QGroupBox("Excel文件选择")
        file_layout = QFormLayout()
        file_row = QHBoxLayout()
        self.file_label = QLabel("未选择文件")
        self.file_label.setStyleSheet("color: gray;")
        file_row.addWidget(self.file_label)
        self.select_file_btn = QPushButton("选择Excel文件")
        self.select_file_btn.clicked.connect(self.select_excel_file)
        file_row.addWidget(self.select_file_btn)
        file_layout.addRow("Excel文件:", file_row)
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # 列选择
        col_group = QGroupBox("列选择")
        col_layout = QFormLayout()

        self.ptzcb_column_combo = QComboBox()
        self.ptzcb_column_combo.setEnabled(False)
        col_layout.addRow("PTZCB查找列:", self.ptzcb_column_combo)

        self.bank_card_column_combo = QComboBox()
        self.bank_card_column_combo.setEnabled(False)
        col_layout.addRow("银行卡号列:", self.bank_card_column_combo)

        status_col_row = QHBoxLayout()
        self.status_column_combo = QComboBox()
        self.status_column_combo.setEnabled(False)
        self.status_column_combo.currentTextChanged.connect(self.on_status_column_changed)
        status_col_row.addWidget(self.status_column_combo)
        self.load_status_btn = QPushButton("加载唯一值")
        self.load_status_btn.setEnabled(False)
        self.load_status_btn.clicked.connect(self.load_status_values)
        status_col_row.addWidget(self.load_status_btn)
        col_layout.addRow("状态列:", status_col_row)

        col_group.setLayout(col_layout)
        layout.addWidget(col_group)

        # 状态映射表
        status_group = QGroupBox("状态映射（Excel值 -> 账号状态，未配置的默认为 可售(3)）")
        status_layout = QVBoxLayout()

        self.status_table = QTableWidget()
        self.status_table.setColumnCount(2)
        self.status_table.setHorizontalHeaderLabels(["Excel状态值", "映射到状态"])
        self.status_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.status_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.status_table.setMaximumHeight(180)
        status_layout.addWidget(self.status_table)

        status_group.setLayout(status_layout)
        layout.addWidget(status_group)

        # 参数
        param_group = QGroupBox("参数设置")
        param_layout = QFormLayout()
        self.register_dept_edit = QLineEdit()
        self.register_dept_edit.setPlaceholderText("请输入注册部门名称")
        self.register_dept_edit.textChanged.connect(self.update_start_state)
        param_layout.addRow("注册部门:", self.register_dept_edit)
        param_group.setLayout(param_layout)
        layout.addWidget(param_group)

        # 按钮
        btn_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始同步")
        self.start_btn.clicked.connect(self.start_process)
        self.start_btn.setEnabled(False)
        btn_layout.addWidget(self.start_btn)

        self.export_completed_btn = QPushButton("导出完整结果")
        self.export_completed_btn.clicked.connect(self.export_completed_data)
        self.export_completed_btn.setEnabled(False)
        btn_layout.addWidget(self.export_completed_btn)

        self.export_failed_btn = QPushButton("导出失败记录")
        self.export_failed_btn.clicked.connect(self.export_failed_records)
        self.export_failed_btn.setEnabled(False)
        btn_layout.addWidget(self.export_failed_btn)

        btn_layout.addStretch()

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        btn_layout.addWidget(self.progress_bar)
        layout.addLayout(btn_layout)

        # 日志
        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(200)
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        log_btn_layout = QHBoxLayout()
        clear_btn = QPushButton("清空日志")
        clear_btn.clicked.connect(self.log_text.clear)
        log_btn_layout.addWidget(clear_btn)
        log_btn_layout.addStretch()
        log_layout.addLayout(log_btn_layout)
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)

        # 结果
        self.result_label = QLabel("")
        self.result_label.setAlignment(Qt.AlignCenter)
        font = QFont()
        font.setBold(True)
        self.result_label.setFont(font)
        layout.addWidget(self.result_label)

        self.setLayout(layout)

    # === 数据源 ===
    def load_datasources(self):
        self.datasource_combo.clear()
        self.datasource_combo.addItem("请选择数据源", None)
        for ds in self.db_manager.get_all_datasources():
            self.datasource_combo.addItem(f"{ds.name} ({ds.host}:{ds.port})", ds)

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

    # === Excel ===
    def select_excel_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", "", "Excel文件 (*.xlsx *.xls)")
        if not path:
            return
        try:
            df = pd.read_excel(path, engine='openpyxl', nrows=0)
            self.excel_columns = list(df.columns)
            self.excel_file = path
            self.file_label.setText(os.path.basename(path))
            self.file_label.setStyleSheet("color: green;")

            for combo in [self.ptzcb_column_combo, self.bank_card_column_combo, self.status_column_combo]:
                combo.clear()
                combo.addItem("请选择列")
                combo.addItems([str(c) for c in self.excel_columns])
                combo.setEnabled(True)

            self.load_status_btn.setEnabled(False)
            self.status_table.setRowCount(0)
            self.append_log(f"读取Excel表头，共 {len(self.excel_columns)} 列")
            self.update_start_state()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取Excel失败: {str(e)}")
            self.excel_file = None
            self.file_label.setText("未选择文件")
            self.file_label.setStyleSheet("color: gray;")

    # === 状态映射 ===
    def on_status_column_changed(self, text):
        self.load_status_btn.setEnabled(text != "请选择列" and text != "")

    def load_status_values(self):
        """读取Excel状态列的唯一值，填充映射表"""
        col = self.status_column_combo.currentText()
        if col == "请选择列" or not self.excel_file:
            return
        try:
            df = pd.read_excel(self.excel_file, engine='openpyxl', usecols=[col])
            unique_vals = df[col].dropna().astype(str).str.strip().unique()
            unique_vals = sorted(set(v for v in unique_vals if v and v != 'nan'))
            self.status_unique_values = unique_vals

            self.status_table.setRowCount(len(unique_vals))
            for i, val in enumerate(unique_vals):
                # 值列（只读）
                item = QTableWidgetItem(val)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.status_table.setItem(i, 0, item)

                # 状态下拉框
                combo = QComboBox()
                for code, label in self.STATUS_OPTIONS:
                    combo.addItem(f"{code} - {label}", code)
                # 默认选可售(3)
                combo.setCurrentIndex(3)
                self.status_table.setCellWidget(i, 1, combo)

            self.append_log(f"状态列 [{col}] 共 {len(unique_vals)} 个唯一值")
            self.update_start_state()
        except Exception as e:
            QMessageBox.warning(self, "错误", f"读取状态列失败: {str(e)}")

    def get_status_mapping(self):
        """从映射表获取 {excel_value: status_code}"""
        mapping = {}
        for i in range(self.status_table.rowCount()):
            val = self.status_table.item(i, 0).text()
            combo = self.status_table.cellWidget(i, 1)
            mapping[val] = combo.currentData()
        return mapping

    # === 状态 ===
    def update_start_state(self):
        ds_ok = self.datasource_combo.currentData() is not None
        file_ok = self.excel_file is not None
        dept_ok = bool(self.register_dept_edit.text().strip())
        status_ok = self.status_table.rowCount() > 0
        self.start_btn.setEnabled(ds_ok and file_ok and dept_ok and status_ok)

    # === 执行 ===
    def start_process(self):
        ds = self.datasource_combo.currentData()
        if not ds or not self.excel_file:
            QMessageBox.warning(self, "警告", "请选择数据源和Excel文件！")
            return

        ptzcb_col = self.ptzcb_column_combo.currentText()
        bank_card_col = self.bank_card_column_combo.currentText()
        status_col = self.status_column_combo.currentText()
        dept = self.register_dept_edit.text().strip()

        if ptzcb_col == "请选择列" or bank_card_col == "请选择列" or status_col == "请选择列":
            QMessageBox.warning(self, "警告", "请选择所有必要的列！")
            return
        if not dept:
            QMessageBox.warning(self, "警告", "请填写注册部门！")
            return

        status_mapping = self.get_status_mapping()
        mapping_desc = "\n".join(f"  {k} -> {v}" for k, v in status_mapping.items())

        reply = QMessageBox.question(
            self, "确认同步",
            f"确定要同步库存账号吗？\n\n"
            f"数据库: {ds.name}\n"
            f"Excel: {os.path.basename(self.excel_file)}\n"
            f"PTZCB列: {ptzcb_col}\n"
            f"银行卡号列: {bank_card_col}\n"
            f"状态列: {status_col}\n"
            f"注册部门: {dept}\n\n"
            f"状态映射:\n{mapping_desc}\n\n"
            f"操作不可撤销，请确认！",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        self.start_btn.setEnabled(False)
        self.export_completed_btn.setEnabled(False)
        self.export_failed_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在同步中...")
        self.log_text.clear()
        self.completed_data = []
        self.failed_records = []

        self.worker = SyncInventoryV2Worker(
            ds, self.excel_file, ptzcb_col, bank_card_col,
            status_col, status_mapping, dept
        )
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

        self.completed_data = results['completed_data']
        self.failed_records = results['failed_records']

        s = results['success_count']
        f = results['failed_count']
        t = results['total_rows']
        text = f"同步完成！成功: {s}, 失败: {f}, 总计: {t}"
        self.result_label.setText(text)
        self.result_label.setStyleSheet("color: green;" if f == 0 else "color: orange;")

        self.export_completed_btn.setEnabled(len(self.completed_data) > 0)
        self.export_failed_btn.setEnabled(len(self.failed_records) > 0)

        self.log_text.append(f"\n=== 同步完成 ===\n{text}")
        if f > 0:
            self.log_text.append("\n失败记录:")
            for r in results['failed_records'][:10]:
                self.log_text.append(f"第{r['row']}行: {r['error']}")

    def on_error(self, msg):
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.result_label.setText("同步失败！")
        self.result_label.setStyleSheet("color: red;")
        QMessageBox.critical(self, "错误", f"同步出错: {msg}")
        self.log_text.append(f"错误: {msg}")

    # === 导出 ===
    def export_completed_data(self):
        if not self.completed_data:
            QMessageBox.warning(self, "警告", "没有可导出的数据！")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "导出完整结果",
            f"同步库存V2_完整结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel文件 (*.xlsx)"
        )
        if path:
            try:
                df = pd.DataFrame(self.completed_data)
                df.to_excel(path, index=False, engine='openpyxl')
                QMessageBox.information(self, "成功", f"已导出到: {path}")
                self.append_log(f"完整结果已导出: {path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出失败: {str(e)}")

    def export_failed_records(self):
        if not self.failed_records:
            QMessageBox.warning(self, "警告", "没有失败记录！")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "导出失败记录",
            f"同步库存V2_失败记录_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel文件 (*.xlsx)"
        )
        if path:
            try:
                data = []
                for r in self.failed_records:
                    row = r['data'].copy()
                    row['失败行号'] = r['row']
                    row['失败原因'] = r['error']
                    row['处理时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    data.append(row)
                df = pd.DataFrame(data)
                df.to_excel(path, index=False, engine='openpyxl')
                QMessageBox.information(self, "成功", f"已导出到: {path}")
                self.append_log(f"失败记录已导出: {path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出失败: {str(e)}")
