"""
根据Excel补充系统平台注册部表信息功能界面
"""
import os
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QFileDialog,
                             QTableWidget, QTableWidgetItem, QScrollArea)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from database import DatabaseManager
import pandas as pd
import pymysql


# 状态选项：数字 -> 中文
STATUS_OPTIONS = {
    0: "待指派",
    1: "进行中",
    2: "审核中",
    3: "银行验证",
    4: "审核失败",
    5: "银行卡验证失败",
    6: "完成",
    7: "失败",
    8: "取消",
    9: "暂停",
    10: "店铺验证失败",
    11: "店铺验证中"
}

# 中文状态 -> 数字（用于从Excel读取状态）
STATUS_NAME_TO_VALUE = {v: k for k, v in STATUS_OPTIONS.items()}


class PtzcbRegisterWorker(QThread):
    """平台注册部表信息补充工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    
    def __init__(self, datasource, db_manager, excel_file, column_mapping):
        super().__init__()
        self.datasource = datasource
        self.db_manager = db_manager
        self.excel_file = excel_file
        self.column_mapping = column_mapping  # 列映射字典
        self.results = {
            'total_rows': 0,
            'success_count': 0,
            'failed_count': 0,
            'skipped_count': 0,
            'failed_records': [],
            'completed_data': []
        }
        self.project_counters = {}
    
    def get_cell_value(self, row, field_name):
        """根据字段名获取单元格值"""
        col_index = self.column_mapping.get(field_name)
        if col_index is None or col_index < 0:
            return ""
        try:
            value = row.iloc[col_index]
            if pd.isna(value):
                return ""
            if isinstance(value, (int, float)) and value == int(value):
                return str(int(value))
            result = str(value)
            if result.endswith('.0'):
                result = result[:-2]
            return result.strip()
        except Exception:
            return ""
    
    def run(self):
        try:
            self.log_message.emit("开始读取Excel文件...")
            self.progress.emit(5)
            
            try:
                df = pd.read_excel(self.excel_file, engine='openpyxl')
                self.log_message.emit(f"成功读取Excel文件，共 {len(df)} 行数据")
            except Exception as e:
                self.error.emit(f"读取Excel文件失败: {str(e)}")
                return
            
            self.results['total_rows'] = len(df)
            self.progress.emit(10)
            
            self.log_message.emit("开始连接数据库...")
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
            self.progress.emit(15)
            
            for index, row in df.iterrows():
                try:
                    order_value = self.get_cell_value(row, 'order_id')
                    
                    if not order_value:
                        self.log_message.emit(f"第{index + 1}行: 订单编号列为空，跳过处理")
                        self.results['skipped_count'] += 1
                        continue
                    
                    project_id, primary_key_id = self.process_row(cursor, index + 1, row)
                    self.results['success_count'] += 1
                    
                    completed_row = row.to_dict()
                    completed_row['生成的项目编号'] = project_id
                    completed_row['ba_ptzcb_register表主键ID'] = primary_key_id
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
                    failed_row['生成的项目编号'] = ''
                    failed_row['ba_ptzcb_register表主键ID'] = ''
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
        try:
            order_id = self.get_cell_value(row, 'order_id')
            legal_name = self.get_cell_value(row, 'legal_name')
            dispose_username = self.get_cell_value(row, 'dispose_username')
            company_name = self.get_cell_value(row, 'company_name')
            country = self.get_cell_value(row, 'country')
            attribute = self.get_cell_value(row, 'attribute')
            project_prefix = self.get_cell_value(row, 'project_prefix')
            current_progress = self.get_cell_value(row, 'current_progress')
            project_start_time_str = self.get_cell_value(row, 'project_start_time')
            get_store_time_str = self.get_cell_value(row, 'get_store_time')
            register_submit_time_str = self.get_cell_value(row, 'register_submit_time')
            time_consume = self.get_cell_value(row, 'time_consume')
            status_text = self.get_cell_value(row, 'status')
            
            # 根据中文状态转换为数字
            status_value = STATUS_NAME_TO_VALUE.get(status_text, 0)  # 默认为0(待指派)
            if status_text and status_text not in STATUS_NAME_TO_VALUE:
                self.log_message.emit(f"第{row_num}行: 状态'{status_text}'无法识别，使用默认值0(待指派)")
            
            self.log_message.emit(f"第{row_num}行: 处理数据 - 订单ID={order_id}, 法人姓名={legal_name}, 属性={attribute}, 前缀={project_prefix}")
            
            full_prefix = self.generate_project_prefix(project_prefix)
            
            shopindex_info = self.find_shopindex_info(cursor, legal_name, company_name, attribute, country, row_num)
            if not shopindex_info:
                raise Exception("未找到对应的shopindex信息")
            
            shopindex_id, shopindex_data = shopindex_info
            
            project_id = self.generate_project_id(full_prefix, shopindex_data['shopindex_id'])
            
            order_info = self.get_order_info(cursor, order_id, legal_name)
            
            dispose_id = self.get_dispose_id(cursor, dispose_username) if dispose_username else None
            
            primary_key_id = self.insert_ptzcb_register(cursor, project_id, order_id, legal_name, shopindex_data, 
                                                       order_info, dispose_id, company_name, country, current_progress,
                                                       project_start_time_str, get_store_time_str, register_submit_time_str, 
                                                       time_consume, status_value, row_num)
            
            self.log_message.emit(f"第{row_num}行: 成功创建项目 {project_id}，主键ID: {primary_key_id}")
            
            return project_id, primary_key_id
            
        except Exception as e:
            raise e
    
    def generate_project_prefix(self, col_h):
        """生成项目编号前缀: H列值 + PT + 年份缩写 + 月日"""
        today = datetime.now()
        year_short = str(today.year)[2:]
        month_day = today.strftime('%m%d')
        prefix = f"{col_h}PT{year_short}{month_day}"
        return prefix
    
    def find_shopindex_info(self, cursor, legal_name, company_name, attribute, country, row_num):
        """根据属性判断查找shopindex信息"""
        try:
            cursor.execute("""
                SELECT id FROM ba_rlb_customer 
                WHERE legal_name = %s AND (delete_time IS NULL OR delete_time = 0)
            """, (legal_name,))
            customer_result = cursor.fetchone()
            
            if not customer_result:
                raise Exception(f"未找到客户: {legal_name}")
            
            customer_id = customer_result[0]
            self.log_message.emit(f"第{row_num}行: 找到客户ID={customer_id}")
            
            if attribute != '企业法人':
                cursor.execute("""
                    SELECT id, shopindex_id, legal_id FROM ba_shopindex 
                    WHERE legal_id = %s AND belong_information IS NULL 
                    AND (delete_time IS NULL OR delete_time = 0)
                """, (customer_id,))
                shopindex_result = cursor.fetchone()
                
                if not shopindex_result:
                    raise Exception(f"未找到个人shopindex记录 (customer_id: {customer_id})")
                
                shopindex_id, shopindex_code, legal_id = shopindex_result
                return shopindex_id, {
                    'id': shopindex_id,
                    'shopindex_id': shopindex_code,
                    'legal_id': legal_id,
                    'customer_id': customer_id,
                    'information_id': None,
                    'information_part2_id': None
                }
            
            else:
                if country.upper() in ["日本", "JAPAN", "JP"]:
                    cursor.execute("""
                        SELECT id FROM ba_rlb_legal_information 
                        WHERE legal_id = %s AND company_name = %s 
                        AND (delete_time IS NULL OR delete_time = 0)
                    """, (customer_id, company_name))
                    legal_info_result = cursor.fetchone()
                    
                    if not legal_info_result:
                        raise Exception(f"未找到日区企业法人信息 (customer_id: {customer_id}, company: {company_name})")
                    
                    information_id = legal_info_result[0]
                    
                    cursor.execute("""
                        SELECT id, shopindex_id, legal_id FROM ba_shopindex 
                        WHERE legal_id = %s AND belong_information = 1 
                        AND (delete_time IS NULL OR delete_time = 0)
                    """, (information_id,))
                    shopindex_result = cursor.fetchone()
                    
                    if not shopindex_result:
                        raise Exception(f"未找到日区企业法人shopindex记录 (information_id: {information_id})")
                    
                    shopindex_id, shopindex_code, legal_id = shopindex_result
                    return shopindex_id, {
                        'id': shopindex_id,
                        'shopindex_id': shopindex_code,
                        'legal_id': legal_id,
                        'customer_id': customer_id,
                        'information_id': information_id,
                        'information_part2_id': None
                    }
                
                else:
                    cursor.execute("""
                        SELECT id FROM ba_rlb_legal_information_part2 
                        WHERE legal_id = %s AND company_name = %s 
                        AND (delete_time IS NULL OR delete_time = 0)
                    """, (customer_id, company_name))
                    legal_info_result = cursor.fetchone()
                    
                    if not legal_info_result:
                        raise Exception(f"未找到欧美企业法人信息 (customer_id: {customer_id}, company: {company_name})")
                    
                    information_part2_id = legal_info_result[0]
                    
                    cursor.execute("""
                        SELECT id, shopindex_id, legal_id FROM ba_shopindex 
                        WHERE legal_id = %s AND belong_information = 2 
                        AND (delete_time IS NULL OR delete_time = 0)
                    """, (information_part2_id,))
                    shopindex_result = cursor.fetchone()
                    
                    if not shopindex_result:
                        raise Exception(f"未找到欧美企业法人shopindex记录 (information_part2_id: {information_part2_id})")
                    
                    shopindex_id, shopindex_code, legal_id = shopindex_result
                    return shopindex_id, {
                        'id': shopindex_id,
                        'shopindex_id': shopindex_code,
                        'legal_id': legal_id,
                        'customer_id': customer_id,
                        'information_id': None,
                        'information_part2_id': information_part2_id
                    }
                    
        except Exception as e:
            raise Exception(f"查找shopindex信息失败: {str(e)}")
    
    def generate_project_id(self, prefix, shopindex_code):
        """生成完整项目编号: 前缀 + shopindex_code + 序号"""
        base_project_id = f"{prefix}{shopindex_code}"
        
        if base_project_id not in self.project_counters:
            self.project_counters[base_project_id] = 1
        else:
            self.project_counters[base_project_id] += 1
        
        counter = self.project_counters[base_project_id]
        full_project_id = f"{base_project_id}{counter:03d}"
        
        return full_project_id
    
    def get_order_info(self, cursor, order_id_str, legal_name):
        """获取订单信息"""
        try:
            cursor.execute("""
                SELECT id, shudi_id, platform_id, currency_id, bank_card_type 
                FROM ba_order 
                WHERE id = %s AND (delete_time IS NULL OR delete_time = 0)
            """, (order_id_str,))
            order_result = cursor.fetchone()
            
            if not order_result:
                raise Exception(f"未找到订单: {order_id_str}")
            
            order_id, shudi_id, platform_id, currency_id, bank_card_type = order_result
            return {
                'order_id': order_id,
                'shudi_id': shudi_id,
                'platform_id': platform_id,
                'currency_id': currency_id,
                'bank_card_type': bank_card_type
            }
            
        except Exception as e:
            raise Exception(f"获取订单信息失败: {str(e)}")
    
    def get_dispose_id(self, cursor, username):
        """根据用户名获取dispose_id"""
        try:
            cursor.execute("""
                SELECT id FROM ba_admin 
                WHERE (username = %s OR nickname = %s) AND status = 'enable'
            """, (username, username))
            admin_result = cursor.fetchone()
            
            return admin_result[0] if admin_result else None
            
        except Exception as e:
            self.log_message.emit(f"获取dispose_id失败: {str(e)}")
            return None
    
    def parse_datetime(self, date_str):
        """解析日期时间字符串"""
        if not date_str:
            return None
        
        try:
            formats = [
                '%Y-%m-%d %H:%M:%S',
                '%Y-%m-%d',
                '%Y/%m/%d %H:%M:%S',
                '%Y/%m/%d',
                '%m/%d/%Y',
                '%d/%m/%Y'
            ]
            
            for fmt in formats:
                try:
                    return datetime.strptime(str(date_str), fmt)
                except ValueError:
                    continue
            
            return None
            
        except Exception:
            return None
    
    def insert_ptzcb_register(self, cursor, project_id, order_id_str, legal_name, shopindex_data, 
                             order_info, dispose_id, company_name, country, current_progress,
                             project_start_time_str, get_store_time_str, register_submit_time_str, time_consume, status_value, row_num):
        """插入ba_ptzcb_register记录"""
        try:
            create_time = int(datetime(2025, 1, 1, 0, 0, 0).timestamp())
            current_time = int(datetime.now().timestamp())
            
            project_start_time = self.parse_datetime(project_start_time_str)
            get_store_time = self.parse_datetime(get_store_time_str)
            register_submit_time = self.parse_datetime(register_submit_time_str)
            done_time = datetime.now()
            
            insert_data = {
                'project_id': project_id,
                'order_id': order_info['order_id'],
                'status': status_value,  # 使用从Excel读取并转换的状态值
                'current_progress': current_progress if current_progress else '',
                'done_time': done_time if status_value == 6 else None,  # 只有完成状态才设置done_time
                'admin_id': 1,
                'dept_id': 1,
                'remark': '系统脚本自动录入',
                'create_time': create_time,
                'update_time': current_time,
                'shopindex_id': shopindex_data['id'],
                'shudi_id': order_info['shudi_id'],
                'platform_id': order_info['platform_id'],
                'bank_card_type': order_info['bank_card_type'],
                'currency_id': order_info['currency_id'],
                'customer_id': shopindex_data['customer_id'],
                'company_name': company_name if company_name else '',
                'country': country if country else ''
            }
            
            if time_consume:
                try:
                    time_consume_value = float(time_consume)
                    insert_data['time_consume'] = time_consume_value
                except (ValueError, TypeError):
                    self.log_message.emit(f"第{row_num}行: time_consume值'{time_consume}'无法转换为数字，跳过该字段")
            
            if dispose_id:
                insert_data['dispose_id'] = dispose_id
            
            if project_start_time:
                insert_data['project_start_time'] = project_start_time
            
            if get_store_time:
                insert_data['get_store_time'] = get_store_time.date()
            
            if register_submit_time:
                insert_data['register_submit_time'] = register_submit_time
            
            if shopindex_data['information_id']:
                insert_data['information_id'] = shopindex_data['information_id']
            
            if shopindex_data['information_part2_id']:
                insert_data['information_part2_id'] = shopindex_data['information_part2_id']
            
            # 移除值为None的字段
            insert_data = {k: v for k, v in insert_data.items() if v is not None}
            
            fields = list(insert_data.keys())
            placeholders = ['%s'] * len(fields)
            values = list(insert_data.values())
            
            insert_sql = f"""
                INSERT INTO ba_ptzcb_register ({', '.join(fields)}) 
                VALUES ({', '.join(placeholders)})
            """
            
            cursor.execute(insert_sql, values)
            
            primary_key_id = cursor.lastrowid
            
            self.log_message.emit(f"第{row_num}行: 插入项目数据 - 项目编号:{project_id}, shopindex_id:{shopindex_data['legal_id']}, 状态:{status_value}({STATUS_OPTIONS.get(status_value, '未知')})")
            
            return primary_key_id
            
        except Exception as e:
            raise Exception(f"插入ba_ptzcb_register记录失败: {str(e)}")


class PtzcbRegisterWidget(QWidget):
    """根据Excel补充系统平台注册部表信息功能组件"""
    
    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.excel_file = None
        self.excel_data = None
        self.excel_columns = []
        self.failed_records = []
        self.completed_data = []
        self.init_ui()
        self.load_datasources()
    
    def init_ui(self):
        """初始化UI"""
        main_layout = QVBoxLayout()
        
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_widget = QWidget()
        layout = QVBoxLayout()
        
        # 数据源选择区域
        datasource_group = QGroupBox("选择目标数据库")
        datasource_layout = QFormLayout()
        
        self.datasource_combo = QComboBox()
        self.datasource_combo.currentTextChanged.connect(self.on_datasource_changed)
        datasource_layout.addRow("数据源:", self.datasource_combo)
        
        self.test_connection_btn = QPushButton("测试连接")
        self.test_connection_btn.clicked.connect(self.test_connection)
        self.test_connection_btn.setEnabled(False)
        datasource_layout.addRow("", self.test_connection_btn)
        
        datasource_group.setLayout(datasource_layout)
        layout.addWidget(datasource_group)
        
        # Excel文件选择区域
        file_group = QGroupBox("Excel文件选择")
        file_layout = QFormLayout()
        
        file_select_layout = QHBoxLayout()
        self.file_label = QLabel("未选择文件")
        self.file_label.setStyleSheet("color: gray;")
        file_select_layout.addWidget(self.file_label)
        
        self.select_file_btn = QPushButton("选择Excel文件")
        self.select_file_btn.clicked.connect(self.select_excel_file)
        file_select_layout.addWidget(self.select_file_btn)
        
        file_layout.addRow("Excel文件:", file_select_layout)
        
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # 状态说明区域
        status_info_group = QGroupBox("状态对照表")
        status_info_layout = QVBoxLayout()
        
        status_info_label = QLabel(
            "Excel中的状态列会自动转换为数字状态：<br>"
            "<b>待指派→0, 进行中→1, 审核中→2, 银行验证→3, 审核失败→4, 银行卡验证失败→5,<br>"
            "完成→6, 失败→7, 取消→8, 暂停→9, 店铺验证失败→10, 店铺验证中→11</b>"
        )
        status_info_label.setWordWrap(True)
        status_info_layout.addWidget(status_info_label)
        
        status_info_group.setLayout(status_info_layout)
        layout.addWidget(status_info_group)
        
        # 列映射配置区域
        mapping_group = QGroupBox("列映射配置（选择Excel文件后可配置）")
        mapping_layout = QFormLayout()
        
        # 必选列
        self.order_col_combo = QComboBox()
        self.order_col_combo.setEnabled(False)
        self.order_col_combo.currentIndexChanged.connect(self.update_start_button_state)
        mapping_layout.addRow("订单编号列 (必选，用于查找ba_order):", self.order_col_combo)
        
        self.legal_name_col_combo = QComboBox()
        self.legal_name_col_combo.setEnabled(False)
        self.legal_name_col_combo.currentIndexChanged.connect(self.update_start_button_state)
        mapping_layout.addRow("法人姓名列 (必选，用于查找客户):", self.legal_name_col_combo)
        
        self.attribute_col_combo = QComboBox()
        self.attribute_col_combo.setEnabled(False)
        self.attribute_col_combo.currentIndexChanged.connect(self.update_start_button_state)
        mapping_layout.addRow("属性列 (必选，判断是否企业法人):", self.attribute_col_combo)
        
        self.project_prefix_col_combo = QComboBox()
        self.project_prefix_col_combo.setEnabled(False)
        self.project_prefix_col_combo.currentIndexChanged.connect(self.update_start_button_state)
        mapping_layout.addRow("项目前缀列 (必选，如FR、JP等):", self.project_prefix_col_combo)
        
        self.status_col_combo = QComboBox()
        self.status_col_combo.setEnabled(False)
        self.status_col_combo.currentIndexChanged.connect(self.update_start_button_state)
        mapping_layout.addRow("状态列 (必选，中文状态如'完成'、'进行中'):", self.status_col_combo)
        
        # 可选列
        self.dispose_col_combo = QComboBox()
        self.dispose_col_combo.setEnabled(False)
        mapping_layout.addRow("指派人列 (可选，匹配ba_admin):", self.dispose_col_combo)
        
        self.company_col_combo = QComboBox()
        self.company_col_combo.setEnabled(False)
        mapping_layout.addRow("公司名列 (企业法人时必填):", self.company_col_combo)
        
        self.country_col_combo = QComboBox()
        self.country_col_combo.setEnabled(False)
        mapping_layout.addRow("国家列 (企业法人时必填):", self.country_col_combo)
        
        self.progress_col_combo = QComboBox()
        self.progress_col_combo.setEnabled(False)
        mapping_layout.addRow("当前进度列 (可选):", self.progress_col_combo)
        
        self.start_time_col_combo = QComboBox()
        self.start_time_col_combo.setEnabled(False)
        mapping_layout.addRow("立项时间列 (可选):", self.start_time_col_combo)
        
        self.store_time_col_combo = QComboBox()
        self.store_time_col_combo.setEnabled(False)
        mapping_layout.addRow("拿店时间列 (可选):", self.store_time_col_combo)
        
        self.submit_time_col_combo = QComboBox()
        self.submit_time_col_combo.setEnabled(False)
        mapping_layout.addRow("注册提交时间列 (可选):", self.submit_time_col_combo)
        
        self.time_consume_col_combo = QComboBox()
        self.time_consume_col_combo.setEnabled(False)
        mapping_layout.addRow("耗时列 (可选):", self.time_consume_col_combo)
        
        mapping_group.setLayout(mapping_layout)
        layout.addWidget(mapping_group)
        
        # 数据预览区域
        preview_group = QGroupBox("数据预览（前5行）")
        preview_layout = QVBoxLayout()
        
        self.preview_table = QTableWidget()
        self.preview_table.setMaximumHeight(150)
        preview_layout.addWidget(self.preview_table)
        
        preview_group.setLayout(preview_layout)
        layout.addWidget(preview_group)
        
        # 功能说明区域
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        
        info_label = QLabel("""
<b>根据Excel补充系统平台注册部表信息说明:</b><br>
<b>状态说明:</b> 0=待指派, 1=进行中, 2=审核中, 3=银行验证, 4=审核失败, 5=银行卡验证失败, 6=完成, 7=失败, 8=取消, 9=暂停, 10=店铺验证失败, 11=店铺验证中<br><br>
<b>必选列:</b> 订单编号、法人姓名、属性、项目前缀<br>
<b>企业法人必填:</b> 公司名、国家<br><br>
<b>项目编号生成规则:</b> 前缀列值 + PT + 年份缩写 + 月日 + shopindex编号 + 序号<br>
<b>示例:</b> FRPT260129C00987001 = FR(前缀) + PT + 26(年份) + 0129(月日) + C00987(shopindex) + 001(序号)<br><br>
<b>shopindex查找规则:</b><br>
• 非企业法人: 通过法人姓名找customer_id → ba_shopindex(belong_information=null)<br>
• 企业法人+日本: customer_id+公司名 → ba_rlb_legal_information → ba_shopindex(belong_information=1)<br>
• 企业法人+非日本: customer_id+公司名 → ba_rlb_legal_information_part2 → ba_shopindex(belong_information=2)
        """)
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # 操作按钮区域
        button_layout = QHBoxLayout()
        
        self.start_btn = QPushButton("开始补充")
        self.start_btn.clicked.connect(self.start_process)
        self.start_btn.setEnabled(False)
        button_layout.addWidget(self.start_btn)
        
        self.export_completed_btn = QPushButton("导出完整结果")
        self.export_completed_btn.clicked.connect(self.export_completed_data)
        self.export_completed_btn.setEnabled(False)
        button_layout.addWidget(self.export_completed_btn)
        
        self.export_failed_btn = QPushButton("导出失败记录")
        self.export_failed_btn.clicked.connect(self.export_failed_records)
        self.export_failed_btn.setEnabled(False)
        button_layout.addWidget(self.export_failed_btn)
        
        button_layout.addStretch()
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        button_layout.addWidget(self.progress_bar)
        
        layout.addLayout(button_layout)
        
        # 日志区域
        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout()
        
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(150)
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
        
        scroll_widget.setLayout(layout)
        scroll_area.setWidget(scroll_widget)
        main_layout.addWidget(scroll_area)
        
        self.setLayout(main_layout)
    
    def load_datasources(self):
        """加载数据源列表"""
        self.datasource_combo.clear()
        self.datasource_combo.addItem("请选择数据源", None)
        
        datasources = self.db_manager.get_all_datasources()
        for ds in datasources:
            self.datasource_combo.addItem(f"{ds.name} ({ds.host}:{ds.port})", ds)
    
    def on_datasource_changed(self):
        """数据源选择变化"""
        current_data = self.datasource_combo.currentData()
        self.test_connection_btn.setEnabled(current_data is not None)
        self.update_start_button_state()
    
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
    
    def select_excel_file(self):
        """选择Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Excel文件",
            "",
            "Excel文件 (*.xlsx *.xls)"
        )
        
        if file_path:
            try:
                df = pd.read_excel(file_path, engine='openpyxl')
                self.excel_data = df
                self.excel_columns = list(df.columns)
                
                self.excel_file = file_path
                self.file_label.setText(os.path.basename(file_path))
                self.file_label.setStyleSheet("color: green;")
                
                self.populate_column_combos()
                self.auto_match_columns(self.excel_columns)
                self.show_data_preview()
                
                self.log_text.append(f"成功读取Excel文件，共 {len(df)} 行数据，{len(self.excel_columns)} 列")
                self.log_text.append(f"列名: {', '.join(self.excel_columns)}")
                
                self.update_start_button_state()
                
            except Exception as e:
                QMessageBox.critical(self, "错误", f"读取Excel文件失败: {str(e)}")
                self.excel_file = None
                self.excel_data = None
                self.file_label.setText("未选择文件")
                self.file_label.setStyleSheet("color: gray;")
    
    def populate_column_combos(self):
        """填充列选择下拉框"""
        all_combos = [
            self.order_col_combo,
            self.legal_name_col_combo,
            self.attribute_col_combo,
            self.project_prefix_col_combo,
            self.status_col_combo,
            self.dispose_col_combo,
            self.company_col_combo,
            self.country_col_combo,
            self.progress_col_combo,
            self.start_time_col_combo,
            self.store_time_col_combo,
            self.submit_time_col_combo,
            self.time_consume_col_combo
        ]
        
        for combo in all_combos:
            combo.clear()
            combo.addItem("-- 不选择 --", -1)
            for i, col in enumerate(self.excel_columns):
                combo.addItem(f"{i+1}. {col}", i)
            combo.setEnabled(True)
    
    def auto_match_columns(self, columns):
        """自动匹配列"""
        column_keywords = {
            'order_id': ['订单', 'order', '编号'],
            'legal_name': ['法人', '姓名', 'legal'],
            'attribute': ['属性', '类型'],
            'project_prefix': ['前缀', 'prefix'],
            'status': ['状态', 'status'],
            'dispose': ['指派', '担当', 'dispose'],
            'company': ['公司', 'company', '企业'],
            'country': ['国家', '国别', 'country'],
            'progress': ['进度', 'progress'],
            'start_time': ['立项', '开始'],
            'store_time': ['拿店', '店铺时间'],
            'submit_time': ['提交', 'submit'],
            'time_consume': ['耗时', '时长', 'consume']
        }
        
        combo_mapping = {
            'order_id': self.order_col_combo,
            'legal_name': self.legal_name_col_combo,
            'attribute': self.attribute_col_combo,
            'project_prefix': self.project_prefix_col_combo,
            'status': self.status_col_combo,
            'dispose': self.dispose_col_combo,
            'company': self.company_col_combo,
            'country': self.country_col_combo,
            'progress': self.progress_col_combo,
            'start_time': self.start_time_col_combo,
            'store_time': self.store_time_col_combo,
            'submit_time': self.submit_time_col_combo,
            'time_consume': self.time_consume_col_combo
        }
        
        for field, keywords in column_keywords.items():
            combo = combo_mapping[field]
            for i, col in enumerate(columns):
                col_lower = str(col).lower()
                for keyword in keywords:
                    if keyword.lower() in col_lower:
                        combo.setCurrentIndex(i + 1)  # +1 因为第一个是"不选择"
                        break
                else:
                    continue
                break
    
    def show_data_preview(self):
        """显示数据预览"""
        if self.excel_data is None:
            return
        
        preview_df = self.excel_data.head(5)
        
        self.preview_table.clear()
        self.preview_table.setRowCount(len(preview_df))
        self.preview_table.setColumnCount(len(preview_df.columns))
        self.preview_table.setHorizontalHeaderLabels([str(col) for col in preview_df.columns])
        
        for row_idx, (_, row) in enumerate(preview_df.iterrows()):
            for col_idx, value in enumerate(row):
                cell_value = "" if pd.isna(value) else str(value)
                item = QTableWidgetItem(cell_value)
                self.preview_table.setItem(row_idx, col_idx, item)
        
        self.preview_table.resizeColumnsToContents()
    
    def get_column_mapping(self):
        """获取列映射"""
        return {
            'order_id': self.order_col_combo.currentData(),
            'legal_name': self.legal_name_col_combo.currentData(),
            'attribute': self.attribute_col_combo.currentData(),
            'project_prefix': self.project_prefix_col_combo.currentData(),
            'status': self.status_col_combo.currentData(),
            'dispose_username': self.dispose_col_combo.currentData(),
            'company_name': self.company_col_combo.currentData(),
            'country': self.country_col_combo.currentData(),
            'current_progress': self.progress_col_combo.currentData(),
            'project_start_time': self.start_time_col_combo.currentData(),
            'get_store_time': self.store_time_col_combo.currentData(),
            'register_submit_time': self.submit_time_col_combo.currentData(),
            'time_consume': self.time_consume_col_combo.currentData()
        }
    
    def update_start_button_state(self):
        """更新开始按钮状态"""
        datasource_selected = self.datasource_combo.currentData() is not None
        file_selected = self.excel_file is not None
        
        # 检查必选列
        order_selected = self.order_col_combo.currentData() is not None and self.order_col_combo.currentData() >= 0
        legal_selected = self.legal_name_col_combo.currentData() is not None and self.legal_name_col_combo.currentData() >= 0
        attribute_selected = self.attribute_col_combo.currentData() is not None and self.attribute_col_combo.currentData() >= 0
        prefix_selected = self.project_prefix_col_combo.currentData() is not None and self.project_prefix_col_combo.currentData() >= 0
        status_selected = self.status_col_combo.currentData() is not None and self.status_col_combo.currentData() >= 0
        
        all_required = order_selected and legal_selected and attribute_selected and prefix_selected and status_selected
        
        self.start_btn.setEnabled(datasource_selected and file_selected and all_required)
    
    def start_process(self):
        """开始补充"""
        datasource = self.datasource_combo.currentData()
        if not datasource or not self.excel_file:
            QMessageBox.warning(self, "警告", "请选择数据源和Excel文件！")
            return
        
        column_mapping = self.get_column_mapping()
        
        # 构建列映射信息
        mapping_info = []
        mapping_info.append(f"• 订单编号列: {self.order_col_combo.currentText()}")
        mapping_info.append(f"• 法人姓名列: {self.legal_name_col_combo.currentText()}")
        mapping_info.append(f"• 属性列: {self.attribute_col_combo.currentText()}")
        mapping_info.append(f"• 项目前缀列: {self.project_prefix_col_combo.currentText()}")
        mapping_info.append(f"• 状态列: {self.status_col_combo.currentText()}")
        
        if self.dispose_col_combo.currentData() >= 0:
            mapping_info.append(f"• 指派人列: {self.dispose_col_combo.currentText()}")
        if self.company_col_combo.currentData() >= 0:
            mapping_info.append(f"• 公司名列: {self.company_col_combo.currentText()}")
        if self.country_col_combo.currentData() >= 0:
            mapping_info.append(f"• 国家列: {self.country_col_combo.currentText()}")
        
        reply = QMessageBox.question(
            self,
            "确认补充",
            f"确定要补充平台注册部表信息吗？\n\n"
            f"数据库: {datasource.name}\n"
            f"Excel文件: {os.path.basename(self.excel_file)}\n\n"
            f"列映射:\n" + "\n".join(mapping_info) + "\n\n"
            f"状态将根据Excel状态列的中文值自动转换为数字。\n"
            f"此操作将根据Excel数据创建ba_ptzcb_register记录。\n"
            f"只处理订单编号列不为空的数据行。\n"
            f"操作不可撤销，请确认数据正确！",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        self.start_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在补充中...")
        self.log_text.clear()
        self.failed_records = []
        self.completed_data = []
        
        self.worker = PtzcbRegisterWorker(datasource, self.db_manager, self.excel_file, column_mapping)
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log_message.connect(self.append_log)
        self.worker.finished.connect(self.on_process_finished)
        self.worker.error.connect(self.on_process_error)
        self.worker.start()
    
    def append_log(self, message):
        """添加日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
    
    def on_process_finished(self, results):
        """补充完成"""
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.failed_records = results['failed_records']
        self.completed_data = results['completed_data']
        
        result_text = (
            f"补充完成！\n"
            f"总行数：{results['total_rows']} 行\n"
            f"成功：{results['success_count']} 行\n"
            f"跳过（订单为空）：{results['skipped_count']} 行\n"
            f"失败：{results['failed_count']} 行"
        )
        
        self.result_label.setText(result_text)
        if results['failed_count'] > 0:
            self.result_label.setStyleSheet("color: orange;")
            self.export_failed_btn.setEnabled(True)
        else:
            self.result_label.setStyleSheet("color: green;")
        
        self.export_completed_btn.setEnabled(True)
        
        QMessageBox.information(self, "补充完成", result_text)
    
    def on_process_error(self, error_msg):
        """补充错误"""
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.result_label.setText("补充失败！")
        self.result_label.setStyleSheet("color: red;")
        
        QMessageBox.critical(self, "补充错误", f"补充失败：{error_msg}")
    
    def export_completed_data(self):
        """导出完整结果"""
        if not self.completed_data:
            QMessageBox.information(self, "提示", "没有完整结果可导出")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存完整结果",
            f"ptzcb_register_completed_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel文件 (*.xlsx)"
        )
        
        if file_path:
            try:
                df = pd.DataFrame(self.completed_data)
                df.to_excel(file_path, index=False, engine='openpyxl')
                
                QMessageBox.information(self, "导出成功", f"完整结果已导出到：\n{file_path}")
                
            except Exception as e:
                QMessageBox.critical(self, "导出失败", f"导出完整结果时出错：{str(e)}")
    
    def export_failed_records(self):
        """导出失败记录"""
        if not self.failed_records:
            QMessageBox.information(self, "提示", "没有失败记录可导出")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存失败记录",
            f"ptzcb_register_failed_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel文件 (*.xlsx)"
        )
        
        if file_path:
            try:
                export_data = []
                for record in self.failed_records:
                    row_data = record['data'].copy()
                    row_data['错误行号'] = record['row']
                    row_data['错误原因'] = record['error']
                    export_data.append(row_data)
                
                df = pd.DataFrame(export_data)
                df.to_excel(file_path, index=False, engine='openpyxl')
                
                QMessageBox.information(self, "导出成功", f"失败记录已导出到：\n{file_path}")
                
            except Exception as e:
                QMessageBox.critical(self, "导出失败", f"导出失败记录时出错：{str(e)}")
    
    def clear_log(self):
        """清空日志"""
        self.log_text.clear()
