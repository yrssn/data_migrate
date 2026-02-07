"""
根据Excel补充系统平台注册部表信息功能界面
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


class PtzcbRegisterWorker(QThread):
    """平台注册部表信息补充工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    
    def __init__(self, datasource, db_manager, excel_file, order_column, legal_column):
        super().__init__()
        self.datasource = datasource
        self.db_manager = db_manager
        self.excel_file = excel_file
        self.order_column = order_column  # 订单编号列名
        self.legal_column = legal_column  # 法人姓名列名
        self.results = {
            'total_rows': 0,
            'success_count': 0,
            'failed_count': 0,
            'failed_records': [],
            'completed_data': []
        }
        # 项目编号计数器
        self.project_counters = {}
    
    def run(self):
        try:
            self.log_message.emit("开始读取Excel文件...")
            self.progress.emit(5)
            
            # 读取Excel文件
            try:
                df = pd.read_excel(self.excel_file, engine='openpyxl')
                self.log_message.emit(f"成功读取Excel文件，共 {len(df)} 行数据")
            except Exception as e:
                self.error.emit(f"读取Excel文件失败: {str(e)}")
                return
            
            self.results['total_rows'] = len(df)
            self.progress.emit(10)
            
            # 连接数据库
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
            
            # 处理每一行数据
            for index, row in df.iterrows():
                try:
                    # 只处理订单编号列不为空的数据
                    order_value = row[self.order_column] if self.order_column in row.index else ""
                    
                    if not order_value or str(order_value) == 'nan' or str(order_value).strip() == '':
                        self.log_message.emit(f"第{index + 1}行: 订单编号列为空，跳过处理")
                        continue
                    
                    project_id, primary_key_id = self.process_row(cursor, index + 1, row)
                    self.results['success_count'] += 1
                    
                    # 记录完整数据
                    completed_row = row.to_dict()
                    completed_row['生成的项目编号'] = project_id
                    completed_row['ba_ptzcb_register表主键ID'] = primary_key_id  # 明确标识这是ba_ptzcb_register表的主键
                    completed_row['处理状态'] = '成功'
                    completed_row['处理时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    self.results['completed_data'].append(completed_row)
                    
                    # 更新进度
                    progress = 15 + int((index + 1) / len(df) * 80)
                    self.progress.emit(progress)
                    
                    # 每50行提交一次
                    if (index + 1) % 50 == 0:
                        connection.commit()
                        self.log_message.emit(f"已处理 {index + 1} 行，提交事务")
                        
                except Exception as e:
                    self.results['failed_count'] += 1
                    error_msg = f"第{index + 1}行处理失败: {str(e)}"
                    self.log_message.emit(error_msg)
                    
                    # 记录失败数据
                    failed_row = row.to_dict()
                    failed_row['生成的项目编号'] = ''
                    failed_row['ba_ptzcb_register表主键ID'] = ''  # 明确标识这是ba_ptzcb_register表的主键
                    failed_row['处理状态'] = '失败'
                    failed_row['失败原因'] = str(e)
                    failed_row['处理时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    self.results['completed_data'].append(failed_row)
                    
                    self.results['failed_records'].append({
                        'row': index + 1,
                        'error': str(e),
                        'data': row.to_dict()
                    })
            
            # 最终提交
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
            # 获取Excel列数据，处理数字格式
            def clean_cell_value(cell_value):
                """清理单元格值，去除.0后缀"""
                if cell_value is None or str(cell_value) == 'nan':
                    return ""
                try:
                    if isinstance(cell_value, (int, float)) and cell_value == int(cell_value):
                        return str(int(cell_value))
                    else:
                        return str(cell_value).replace('.0', '') if str(cell_value).endswith('.0') else str(cell_value)
                except:
                    return str(cell_value)
            
            col_a = clean_cell_value(row.iloc[0]) if len(row) > 0 else ""  # A列
            col_c = clean_cell_value(row.iloc[2]) if len(row) > 2 else ""  # C列：dispose用户名
            col_e = clean_cell_value(row.iloc[4]) if len(row) > 4 else ""  # E列：company_name
            col_f = clean_cell_value(row.iloc[5]) if len(row) > 5 else ""  # F列：country
            col_g = clean_cell_value(row.iloc[6]) if len(row) > 6 else ""  # G列：属性判断
            col_h = clean_cell_value(row.iloc[7]) if len(row) > 7 else ""  # H列：项目前缀
            col_i = clean_cell_value(row.iloc[8]) if len(row) > 8 else ""  # I列：current_progress
            col_j = clean_cell_value(row.iloc[9]) if len(row) > 9 else ""  # J列：project_start_time
            col_k = clean_cell_value(row.iloc[10]) if len(row) > 10 else ""  # K列：get_store_time
            col_m = clean_cell_value(row.iloc[12]) if len(row) > 12 else ""  # M列：register_submit_time
            col_n = clean_cell_value(row.iloc[13]) if len(row) > 13 else ""  # N列：time_consume
            
            # 使用用户选择的列
            order_id = clean_cell_value(row[self.order_column]) if self.order_column in row.index else ""  # 订单ID
            legal_name = clean_cell_value(row[self.legal_column]) if self.legal_column in row.index else ""  # 法人姓名
            
            self.log_message.emit(f"第{row_num}行: 处理数据 - 订单ID={order_id}, 法人姓名={legal_name}, 属性={col_g}, 前缀={col_h}")
            
            # 第一步：生成项目编号前缀
            project_prefix = self.generate_project_prefix(col_h)
            
            # 第二步：根据属性判断查找shopindex_id
            shopindex_info = self.find_shopindex_info(cursor, legal_name, col_e, col_g, col_f, row_num)
            if not shopindex_info:
                raise Exception("未找到对应的shopindex信息")
            
            shopindex_id, shopindex_data = shopindex_info
            
            # 第三步：生成完整项目编号
            project_id = self.generate_project_id(project_prefix, shopindex_data['shopindex_id'])
            
            # 第四步：查找order_id对应的订单信息
            order_info = self.get_order_info(cursor, order_id, legal_name)
            
            # 第五步：查找dispose_id
            dispose_id = self.get_dispose_id(cursor, col_c) if col_c and col_c != 'nan' else None
            
            # 第六步：插入ba_ptzcb_register记录
            primary_key_id = self.insert_ptzcb_register(cursor, project_id, order_id, legal_name, shopindex_data, 
                                                       order_info, dispose_id, col_e, col_f, col_i, 
                                                       col_j, col_k, col_m, col_n, row_num)
            
            self.log_message.emit(f"第{row_num}行: 成功创建项目 {project_id}，主键ID: {primary_key_id}")
            
            return project_id, primary_key_id
            
        except Exception as e:
            raise e
    
    def generate_project_prefix(self, col_h):
        """生成项目编号前缀: H列值 + PT + 年份缩写 + 月日"""
        today = datetime.now()
        year_short = str(today.year)[2:]  # 2026 -> 26
        month_day = today.strftime('%m%d')  # 0129
        prefix = f"{col_h}PT{year_short}{month_day}"
        return prefix
    
    def find_shopindex_info(self, cursor, legal_name, company_name, attribute, country, row_num):
        """根据属性判断查找shopindex信息"""
        try:
            # 先通过legal_name找到customer_id
            cursor.execute("""
                SELECT id FROM ba_rlb_customer 
                WHERE legal_name = %s AND (delete_time IS NULL OR delete_time = 0)
            """, (legal_name,))
            customer_result = cursor.fetchone()
            
            if not customer_result:
                raise Exception(f"未找到客户: {legal_name}")
            
            customer_id = customer_result[0]
            self.log_message.emit(f"第{row_num}行: 找到客户ID={customer_id}")
            
            # 判断属性类型
            if attribute != '企业法人':
                # 非企业法人：查找belong_information为null的shopindex
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
                    'id': shopindex_id,  # ba_shopindex表的id
                    'shopindex_id': shopindex_code,
                    'legal_id': legal_id,
                    'customer_id': customer_id,
                    'information_id': None,
                    'information_part2_id': None
                }
            
            else:
                # 企业法人：根据F列（国家）判断查找哪个表
                if country.upper() == "日本" or country.upper() == "JAPAN" or country.upper() == "JP":
                    # 日区企业法人：ba_rlb_legal_information
                    cursor.execute("""
                        SELECT id FROM ba_rlb_legal_information 
                        WHERE legal_id = %s AND company_name = %s 
                        AND (delete_time IS NULL OR delete_time = 0)
                    """, (customer_id, company_name))
                    legal_info_result = cursor.fetchone()
                    
                    if not legal_info_result:
                        raise Exception(f"未找到日区企业法人信息 (customer_id: {customer_id}, company: {company_name})")
                    
                    information_id = legal_info_result[0]
                    
                    # 查找对应的shopindex (belong_information = 1)
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
                        'id': shopindex_id,  # ba_shopindex表的id
                        'shopindex_id': shopindex_code,
                        'legal_id': legal_id,
                        'customer_id': customer_id,
                        'information_id': information_id,
                        'information_part2_id': None
                    }
                
                else:
                    # 非日区企业法人：ba_rlb_legal_information_part2
                    cursor.execute("""
                        SELECT id FROM ba_rlb_legal_information_part2 
                        WHERE legal_id = %s AND company_name = %s 
                        AND (delete_time IS NULL OR delete_time = 0)
                    """, (customer_id, company_name))
                    legal_info_result = cursor.fetchone()
                    
                    if not legal_info_result:
                        raise Exception(f"未找到欧美企业法人信息 (customer_id: {customer_id}, company: {company_name})")
                    
                    information_part2_id = legal_info_result[0]
                    
                    # 查找对应的shopindex (belong_information = 2)
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
                        'id': shopindex_id,  # ba_shopindex表的id
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
        # 生成基础项目编号
        base_project_id = f"{prefix}{shopindex_code}"
        
        # 获取或初始化计数器
        if base_project_id not in self.project_counters:
            self.project_counters[base_project_id] = 1
        else:
            self.project_counters[base_project_id] += 1
        
        # 生成完整项目编号
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
        if not date_str or date_str == 'nan':
            return None
        
        try:
            # 尝试多种日期格式
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
            
            # 如果都不匹配，返回None
            return None
            
        except Exception:
            return None
    
    def insert_ptzcb_register(self, cursor, project_id, order_id_str, legal_name, shopindex_data, 
                             order_info, dispose_id, company_name, country, current_progress,
                             project_start_time_str, get_store_time_str, register_submit_time_str, time_consume, row_num):
        """插入ba_ptzcb_register记录"""
        try:
            # 设置创建时间为2025-01-01 00:00:00的时间戳
            create_time = int(datetime(2025, 1, 1, 0, 0, 0).timestamp())
            current_time = int(datetime.now().timestamp())
            
            # 解析时间字段
            project_start_time = self.parse_datetime(project_start_time_str)
            get_store_time = self.parse_datetime(get_store_time_str)
            register_submit_time = self.parse_datetime(register_submit_time_str)
            done_time = datetime.now()
            
            # 构建插入数据
            insert_data = {
                'project_id': project_id,
                'order_id': order_info['order_id'],
                'status': 6,  # 完成
                'current_progress': current_progress if current_progress and current_progress != 'nan' else '',
                'done_time': done_time,
                'admin_id': 1,
                'dept_id': 1,
                'remark': '系统脚本自动录入',
                'create_time': create_time,  # 使用2025-01-01的时间戳
                'update_time': current_time,
                'shopindex_id': shopindex_data['id'],  # ba_shopindex的id
                'shudi_id': order_info['shudi_id'],
                'platform_id': order_info['platform_id'],
                'bank_card_type': order_info['bank_card_type'],
                'currency_id': order_info['currency_id'],
                'customer_id': shopindex_data['customer_id'],
                'company_name': company_name if company_name and company_name != 'nan' else '',
                'country': country if country and country != 'nan' else ''
            }
            
            # 添加time_consume字段（只有有值时才添加）
            if time_consume and time_consume != 'nan' and str(time_consume).strip():
                try:
                    # 尝试转换为数字
                    time_consume_value = float(time_consume)
                    insert_data['time_consume'] = time_consume_value
                except (ValueError, TypeError):
                    # 如果转换失败，记录日志但不添加字段
                    self.log_message.emit(f"第{row_num}行: time_consume值'{time_consume}'无法转换为数字，跳过该字段")
            
            # 添加可选字段
            if dispose_id:
                insert_data['dispose_id'] = dispose_id
            
            if project_start_time:
                insert_data['project_start_time'] = project_start_time
            
            if get_store_time:
                insert_data['get_store_time'] = get_store_time.date()
            
            if register_submit_time:
                insert_data['register_submit_time'] = register_submit_time
            
            # 根据shopindex信息设置information相关字段
            if shopindex_data['information_id']:
                insert_data['information_id'] = shopindex_data['information_id']
            
            if shopindex_data['information_part2_id']:
                insert_data['information_part2_id'] = shopindex_data['information_part2_id']
            
            # 构建SQL
            fields = list(insert_data.keys())
            placeholders = ['%s'] * len(fields)
            values = list(insert_data.values())
            
            insert_sql = f"""
                INSERT INTO ba_ptzcb_register ({', '.join(fields)}) 
                VALUES ({', '.join(placeholders)})
            """
            
            cursor.execute(insert_sql, values)
            
            # 获取插入后的主键ID
            primary_key_id = cursor.lastrowid
            
            self.log_message.emit(f"第{row_num}行: 插入项目数据 - 项目编号:{project_id}, shopindex_id:{shopindex_data['legal_id']}")
            
            return primary_key_id
            
        except Exception as e:
            raise Exception(f"插入ba_ptzcb_register记录失败: {str(e)}")


class PtzcbRegisterWidget(QWidget):
    """根据Excel补充系统平台注册部表信息功能组件"""
    
    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.excel_file = None
        self.excel_columns = []  # 存储Excel列名
        self.failed_records = []
        self.completed_data = []
        self.init_ui()
        self.load_datasources()
    
    def init_ui(self):
        """初始化UI"""
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
        
        # 列选择区域
        column_group = QGroupBox("列选择")
        column_layout = QFormLayout()
        
        self.order_column_combo = QComboBox()
        self.order_column_combo.setEnabled(False)
        column_layout.addRow("订单编号列:", self.order_column_combo)
        
        self.legal_column_combo = QComboBox()
        self.legal_column_combo.setEnabled(False)
        column_layout.addRow("法人姓名列:", self.legal_column_combo)
        
        column_group.setLayout(column_layout)
        layout.addWidget(column_group)
        
        # 功能说明区域
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        
        info_label = QLabel("""
<b>根据Excel补充系统平台注册部表信息说明:</b><br>
<b>使用步骤:</b><br>
1. 选择Excel文件后，系统会自动读取表头<br>
2. 在"列选择"区域选择对应的列：<br>
   • <b>订单编号列</b>: 用于查找ba_order表（必填，只处理此列不为空的数据）<br>
   • <b>法人姓名列</b>: 用于查找客户信息（必填）<br><br>
<b>Excel列要求（按固定位置）:</b><br>
• A列: 数据列<br>
• C列: 指派人用户名<br>
• E列: 公司名称 (企业法人时使用)<br>
• F列: 国家<br>
• G列: 属性判断 (是否企业法人)<br>
• H列: 项目前缀 (如FR、JP等)<br>
• I列: 当前进度<br>
• J列: 立项时间<br>
• K列: 拿店时间<br>
• M列: 注册提交时间<br>
• N列: 耗时信息<br><br>
<b>项目编号生成规则:</b><br>
• 格式: H列值 + PT + 年份缩写 + 月日 + shopindex编号 + 序号<br>
• 示例: FRPT260129C00987001<br>
• FR(H列) + PT + 26(年份) + 0129(月日) + C00987(shopindex) + 001(序号)<br><br>
<b>shopindex查找规则:</b><br>
• <b>非企业法人</b>: 通过法人姓名找customer_id → ba_shopindex(belong_information=null)<br>
• <b>企业法人+日本</b>: customer_id+E列 → ba_rlb_legal_information → ba_shopindex(belong_information=1)<br>
  &nbsp;&nbsp;- F列为"日本"、"JAPAN"或"JP"时<br>
• <b>企业法人+非日本</b>: customer_id+E列 → ba_rlb_legal_information_part2 → ba_shopindex(belong_information=2)<br>
  &nbsp;&nbsp;- F列不是日本时<br><br>
<b>数据补充:</b><br>
• 从ba_order获取: shudi_id, platform_id, currency_id, bank_card_type<br>
• 从ba_admin获取: dispose_id (根据C列用户名)<br>
• 时间字段自动解析多种格式<br>
• 状态固定为6(完成), admin_id和dept_id为1
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
                # 读取Excel表头
                df = pd.read_excel(file_path, engine='openpyxl', nrows=0)
                self.excel_columns = list(df.columns)
                
                # 更新UI
                self.excel_file = file_path
                self.file_label.setText(os.path.basename(file_path))
                self.file_label.setStyleSheet("color: green;")
                
                # 填充列选择下拉框
                self.order_column_combo.clear()
                self.order_column_combo.addItems(self.excel_columns)
                self.order_column_combo.setEnabled(True)
                
                self.legal_column_combo.clear()
                self.legal_column_combo.addItems(self.excel_columns)
                self.legal_column_combo.setEnabled(True)
                
                # 尝试自动选择默认列（如果存在）
                # 订单编号列可能叫"订单编号"、"order"等
                for i, col in enumerate(self.excel_columns):
                    if '订单' in str(col) or 'order' in str(col).lower():
                        self.order_column_combo.setCurrentIndex(i)
                        break
                
                # 法人姓名列可能叫"法人姓名"、"姓名"等
                for i, col in enumerate(self.excel_columns):
                    if '法人' in str(col) or '姓名' in str(col):
                        self.legal_column_combo.setCurrentIndex(i)
                        break
                
                self.log_text.append(f"成功读取Excel表头，共 {len(self.excel_columns)} 列")
                self.log_text.append(f"列名: {', '.join(self.excel_columns)}")
                
                self.update_start_button_state()
                
            except Exception as e:
                QMessageBox.critical(self, "错误", f"读取Excel文件失败: {str(e)}")
                self.excel_file = None
                self.file_label.setText("未选择文件")
                self.file_label.setStyleSheet("color: gray;")
    
    def update_start_button_state(self):
        """更新开始按钮状态"""
        datasource_selected = self.datasource_combo.currentData() is not None
        file_selected = self.excel_file is not None
        self.start_btn.setEnabled(datasource_selected and file_selected)
    
    def start_process(self):
        """开始补充"""
        datasource = self.datasource_combo.currentData()
        if not datasource or not self.excel_file:
            QMessageBox.warning(self, "警告", "请选择数据源和Excel文件！")
            return
        
        # 获取选择的列
        order_column = self.order_column_combo.currentText()
        legal_column = self.legal_column_combo.currentText()
        
        if not order_column or not legal_column:
            QMessageBox.warning(self, "警告", "请选择订单编号列和法人姓名列！")
            return
        
        # 确认对话框
        reply = QMessageBox.question(
            self,
            "确认补充",
            f"确定要补充平台注册部表信息吗？\n\n"
            f"数据库: {datasource.name}\n"
            f"Excel文件: {os.path.basename(self.excel_file)}\n"
            f"订单编号列: {order_column}\n"
            f"法人姓名列: {legal_column}\n\n"
            f"此操作将根据Excel数据创建ba_ptzcb_register记录。\n"
            f"只处理订单编号列不为空的数据行。\n"
            f"操作不可撤销，请确认数据正确！",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        # 禁用按钮，显示进度条
        self.start_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在补充中...")
        self.log_text.clear()
        self.failed_records = []
        self.completed_data = []
        
        # 启动工作线程
        self.worker = PtzcbRegisterWorker(datasource, self.db_manager, self.excel_file, order_column, legal_column)
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
            f"失败：{results['failed_count']} 行"
        )
        
        self.result_label.setText(result_text)
        if results['failed_count'] > 0:
            self.result_label.setStyleSheet("color: orange;")
            self.export_failed_btn.setEnabled(True)
        else:
            self.result_label.setStyleSheet("color: green;")
        
        # 启用导出完整结果按钮
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
                # 创建DataFrame并导出
                df = pd.DataFrame(self.completed_data)
                df.to_excel(file_path, index=False, engine='openpyxl')
                
                QMessageBox.information(self, "导出成功", f"完整结果已导出到：\n{file_path}\n\n包含字段：\n• 原始Excel所有列（AB列保持订单主键ID）\n• 生成的项目编号\n• ba_ptzcb_register表主键ID\n• 处理状态\n• 处理时间\n• 失败原因（如有）")
                
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
                # 准备导出数据
                export_data = []
                for record in self.failed_records:
                    row_data = record['data'].copy()
                    row_data['错误行号'] = record['row']
                    row_data['错误原因'] = record['error']
                    export_data.append(row_data)
                
                # 创建DataFrame并导出
                df = pd.DataFrame(export_data)
                df.to_excel(file_path, index=False, engine='openpyxl')
                
                QMessageBox.information(self, "导出成功", f"失败记录已导出到：\n{file_path}")
                
            except Exception as e:
                QMessageBox.critical(self, "导出失败", f"导出失败记录时出错：{str(e)}")
    
    def clear_log(self):
        """清空日志"""
        self.log_text.clear()
