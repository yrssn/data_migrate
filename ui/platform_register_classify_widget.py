"""
平台注册部表分类功能界面
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


class PlatformRegisterClassifyWorker(QThread):
    """平台注册部表分类工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    
    def __init__(self, datasource, db_manager, excel_file, output_file):
        super().__init__()
        self.datasource = datasource
        self.db_manager = db_manager
        self.excel_file = excel_file
        self.output_file = output_file
        self.results = {
            'total_rows': 0,
            'sheet1_count': 0,  # 平台注册部数据在财务系统中
            'sheet2_count': 0,  # 日联部店铺（数据库有但Excel没有）
            'sheet3_count': 0,  # 平台注册部不在财务系统的（法人未找到）
            'sheet4_count': 0,  # 法人找到但平台未找到
        }
        
        # 存储四个sheet的数据
        self.sheet1_data = []  # 平台注册部数据在财务系统中
        self.sheet2_data = []  # 日联部店铺
        self.sheet3_data = []  # 平台注册部不在财务系统的
        self.sheet4_data = []  # 法人找到但平台未找到
        
        # 用于去重检查
        self.sheet1_keys = set()  # 存储(B列, D列)的组合
        self.sheet2_keys = set()  # 存储(legal_name, site_type_name)的组合
        self.sheet4_keys = set()  # 存储(B列, D列)的组合
        
        # 存储所有数据库中的法人和平台组合
        self.db_legal_platform_pairs = set()  # 存储数据库中所有(legal_name, platform_name)组合
        # 存储Excel中的法人和平台组合
        self.excel_legal_platform_pairs = set()  # 存储Excel中所有(B列, D列)组合
        # 存储Excel中出现的所有法人名称
        self.excel_legal_names = set()  # 存储Excel中所有B列的法人名称
    
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
            cursor = connection.cursor(pymysql.cursors.DictCursor)
            self.log_message.emit("数据库连接成功")
            self.progress.emit(15)
            
            # 第一步：收集Excel中的所有法人和平台组合
            self.log_message.emit("第一步：收集Excel数据...")
            for index, row in df.iterrows():
                b_col = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ""
                d_col = str(row.iloc[3]).strip() if len(row) > 3 and pd.notna(row.iloc[3]) else ""
                if b_col:
                    self.excel_legal_names.add(b_col)  # 收集所有法人名称
                if b_col and d_col:
                    self.excel_legal_platform_pairs.add((b_col, d_col))
            self.log_message.emit(f"Excel中共有 {len(self.excel_legal_names)} 个法人")
            self.log_message.emit(f"Excel中共有 {len(self.excel_legal_platform_pairs)} 个法人-平台组合")
            self.progress.emit(20)
            
            # 第二步：查询数据库中所有的法人-平台组合（在Python中过滤）
            self.log_message.emit("第二步：收集数据库中的法人-平台组合...")
            cursor.execute("""
                SELECT l.name as legal_name, st.name as platform_name, 
                       s.id, s.channel, s.charge, s.referrer,
                       s.storeAcquisitionMonth, s.storeAcquisitionDate, s.storeApprovalMonth,
                       s.storeAcquisitionApprovalId, s.storeApprovalId, s.status
                FROM ea_dy_legal l
                INNER JOIN ea_dy_shop s ON l.id = s.legal_id
                INNER JOIN ea_dy_site_type st ON s.siteType = st.id
                WHERE (l.delete_time IS NULL OR l.delete_time = 0)
                  AND (s.delete_time IS NULL OR s.delete_time = 0)
                  AND (st.delete_time IS NULL OR st.delete_time = 0)
            """)
            all_db_records = cursor.fetchall()
            
            # 在Python中过滤：只保留Excel中出现的法人
            for record in all_db_records:
                legal_name = record['legal_name']
                platform_name = record['platform_name']
                
                # 只保留Excel中出现的法人
                if legal_name in self.excel_legal_names:
                    self.db_legal_platform_pairs.add((legal_name, platform_name))
                    
            self.log_message.emit(f"数据库中Excel法人共有 {len(self.db_legal_platform_pairs)} 个法人-平台组合")
            self.progress.emit(30)
            
            # 第三步：处理Excel中的每一行数据
            self.log_message.emit("第三步：处理Excel数据...")
            for index, row in df.iterrows():
                try:
                    self.process_excel_row(cursor, index + 1, row, df.columns)
                    
                    # 更新进度
                    progress = 30 + int((index + 1) / len(df) * 40)
                    self.progress.emit(progress)
                    
                except Exception as e:
                    error_msg = f"第{index + 1}行处理失败: {str(e)}"
                    self.log_message.emit(error_msg)
                    # 处理失败的数据也放到sheet3
                    self.add_to_sheet3_original(row, df.columns)
            
            # 第四步：找出数据库中有但Excel中没有的数据（Sheet2）
            self.log_message.emit("第四步：查找数据库独有的法人-平台组合...")
            self.find_db_only_records(cursor)
            self.progress.emit(80)
            
            # 关闭数据库连接
            cursor.close()
            connection.close()
            self.log_message.emit("数据库连接已关闭")
            self.progress.emit(85)
            
            # 导出Excel
            self.export_to_excel()
            self.progress.emit(100)
            
            self.results['sheet1_count'] = len(self.sheet1_data)
            self.results['sheet2_count'] = len(self.sheet2_data)
            self.results['sheet3_count'] = len(self.sheet3_data)
            self.results['sheet4_count'] = len(self.sheet4_data)
            
            self.finished.emit(self.results)
            
        except Exception as e:
            self.error.emit(f"处理过程出错: {str(e)}")
    
    def process_excel_row(self, cursor, row_num, row, columns):
        """处理Excel中的单行数据"""
        # 获取B列和D列的值（索引1和3）
        b_col = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ""
        d_col = str(row.iloc[3]).strip() if len(row) > 3 and pd.notna(row.iloc[3]) else ""
        
        if not b_col:
            self.log_message.emit(f"第{row_num}行: B列为空，放入sheet3")
            self.add_to_sheet3_original(row, columns)
            return
        
        # 查询ea_dy_legal表，通过name匹配B列
        cursor.execute("""
            SELECT id, name, status, remark, create_time, update_time, delete_time
            FROM ea_dy_legal
            WHERE name = %s AND (delete_time IS NULL OR delete_time = 0)
        """, (b_col,))
        
        legal_record = cursor.fetchone()
        
        if not legal_record:
            # B列没有匹配到法人，放入sheet3
            self.log_message.emit(f"第{row_num}行: 法人'{b_col}'未在财务系统中找到，放入sheet3")
            self.add_to_sheet3_original(row, columns)
            return
        
        legal_id = legal_record[0]  # id字段
        legal_name = legal_record[1]  # name字段
        
        # 检查这个法人-平台组合是否在数据库中存在
        if not d_col:
            self.log_message.emit(f"第{row_num}行: D列为空，放入sheet3")
            self.add_to_sheet3_original(row, columns)
            return
        
        # 检查数据库中是否有这个法人-平台组合
        if (b_col, d_col) in self.db_legal_platform_pairs:
            # 数据库中有这个组合，放入sheet1
            # 查询具体的店铺和平台信息（取创建时间最新的一条）
            cursor.execute("""
                SELECT s.id, s.name, s.legal_id, s.siteType, s.channel, s.charge, 
                       s.referrer, s.storeAcquisitionMonth, s.storeAcquisitionDate, s.storeApprovalMonth,
                       s.storeAcquisitionApprovalId, s.storeApprovalId, s.status,
                       st.name as platform_name
                FROM ea_dy_shop s
                INNER JOIN ea_dy_site_type st ON s.siteType = st.id
                WHERE s.legal_id = %s AND st.name = %s
                  AND (s.delete_time IS NULL OR s.delete_time = 0)
                  AND (st.delete_time IS NULL OR st.delete_time = 0)
                ORDER BY s.create_time DESC
                LIMIT 1
            """, (legal_id, d_col))
            
            shop_record = cursor.fetchone()
            
            if shop_record:
                key = (b_col, d_col)
                # 添加调试信息：显示查询结果长度和状态值
                self.log_message.emit(f"第{row_num}行: 查询结果字段数={len(shop_record)}")
                if len(shop_record) > 12:
                    status_value = shop_record[12]  # status字段在查询结果中的索引
                    self.log_message.emit(f"第{row_num}行: 查询到状态值={status_value} (类型:{type(status_value)})")
                else:
                    self.log_message.emit(f"第{row_num}行: 查询结果字段不足，无法获取状态值")
                
                if key not in self.sheet1_keys:
                    self.log_message.emit(f"第{row_num}行: 法人'{b_col}'和平台'{d_col}'匹配，放入sheet1")
                    self.add_to_sheet1(row, columns, shop_record)
                    self.sheet1_keys.add(key)
                else:
                    self.log_message.emit(f"第{row_num}行: 法人'{b_col}'和平台'{d_col}'已存在于sheet1，跳过")
        else:
            # 数据库中没有这个法人-平台组合，放入sheet4
            key = (b_col, d_col)
            if key not in self.sheet4_keys:
                self.log_message.emit(f"第{row_num}行: 法人'{b_col}'找到但平台'{d_col}'未找到，放入sheet4")
                self.add_to_sheet4(row, columns)
                self.sheet4_keys.add(key)
            else:
                self.log_message.emit(f"第{row_num}行: 法人'{b_col}'和平台'{d_col}'已存在于sheet4，跳过")
    
    def find_db_only_records(self, cursor):
        """查找数据库中有但Excel中没有的记录（Sheet2）"""
        # 找出数据库独有的法人-平台组合
        db_only_pairs = self.db_legal_platform_pairs - self.excel_legal_platform_pairs
        
        self.log_message.emit(f"发现 {len(db_only_pairs)} 个数据库独有的法人-平台组合")
        
        for legal_name, platform_name in db_only_pairs:
            key = (legal_name, platform_name)
            if key not in self.sheet2_keys:
                # 查询详细信息（取创建时间最新的一条）
                cursor.execute("""
                    SELECT l.id as legal_id, l.name as legal_name,
                           s.id, s.name, s.channel, s.charge, s.referrer,
                           s.storeAcquisitionMonth, s.storeAcquisitionDate, s.storeApprovalMonth,
                           s.storeAcquisitionApprovalId, s.storeApprovalId, s.status,
                           st.name as platform_name
                    FROM ea_dy_legal l
                    INNER JOIN ea_dy_shop s ON l.id = s.legal_id
                    INNER JOIN ea_dy_site_type st ON s.siteType = st.id
                    WHERE l.name = %s AND st.name = %s
                      AND (l.delete_time IS NULL OR l.delete_time = 0)
                      AND (s.delete_time IS NULL OR s.delete_time = 0)
                      AND (st.delete_time IS NULL OR st.delete_time = 0)
                    ORDER BY s.create_time DESC
                    LIMIT 1
                """, (legal_name, platform_name))
                
                record = cursor.fetchone()
                if record:
                    self.add_to_sheet2_from_db(record)
                    self.sheet2_keys.add(key)
                    self.log_message.emit(f"数据库独有: 法人'{legal_name}'和平台'{platform_name}'，放入sheet2")
    
    def add_to_sheet1(self, row, columns, shop):
        """添加数据到sheet1（平台注册部数据在财务系统中）"""
        # 原始Excel数据
        data = {}
        for i, col in enumerate(columns):
            data[col] = row.iloc[i] if i < len(row) else ""
        
        # 添加额外字段（shop是元组，按索引访问）
        data['渠道'] = shop[4] if len(shop) > 4 else ''  # s.channel
        data['担当'] = shop[5] if len(shop) > 5 else ''  # s.charge
        data['介绍人'] = shop[6] if len(shop) > 6 else ''  # s.referrer
        data['拿店月份'] = shop[7] if len(shop) > 7 else ''  # s.storeAcquisitionMonth
        data['拿店日期'] = shop[8] if len(shop) > 8 else ''  # s.storeAcquisitionDate
        data['给店月份'] = shop[9] if len(shop) > 9 else ''  # s.storeApprovalMonth
        data['拿店审批号'] = shop[10] if len(shop) > 10 else ''  # s.storeAcquisitionApprovalId
        data['给店审批号'] = shop[11] if len(shop) > 11 else ''  # s.storeApprovalId
        
        # 状态转换
        status = shop[12] if len(shop) > 12 else 1  # s.status
        status_map = {
            0: '暂停',
            1: '正常',
            2: '已死店'
        }
        converted_status = status_map.get(status, f'未知状态({status})')
        data['店铺状态'] = converted_status
        
        # 添加调试信息
        if status == 3:
            print(f"DEBUG: 发现已退店状态 - 原始值:{status}, 转换后:{converted_status}")
        
        self.sheet1_data.append(data)
    
    def add_to_sheet2_from_db(self, record):
        """从数据库记录添加数据到sheet2（日联部店铺）"""
        data = {
            '法人姓名': record[1] if len(record) > 1 else '',  # l.name as legal_name
            '店铺类型': record[13] if len(record) > 13 else '',  # st.name as platform_name
            '渠道': record[4] if len(record) > 4 else '',  # s.channel
            '担当': record[5] if len(record) > 5 else '',  # s.charge
            '介绍人': record[6] if len(record) > 6 else '',  # s.referrer
            '拿店月份': record[7] if len(record) > 7 else '',  # s.storeAcquisitionMonth
            '拿店日期': record[8] if len(record) > 8 else '',  # s.storeAcquisitionDate
            '给店月份': record[9] if len(record) > 9 else '',  # s.storeApprovalMonth
            '拿店审批号': record[10] if len(record) > 10 else '',  # s.storeAcquisitionApprovalId
            '给店审批号': record[11] if len(record) > 11 else '',  # s.storeApprovalId
        }
        
        # 状态转换
        status = record[12] if len(record) > 12 else 1  # s.status
        status_map = {
            0: '暂停',
            1: '正常',
            2: '已死店'
        }
        data['店铺状态'] = status_map.get(status, f'未知状态({status})')
        
        self.sheet2_data.append(data)
    
    def add_to_sheet3_original(self, row, columns):
        """添加原始数据到sheet3（平台注册部不在财务系统的）"""
        data = {}
        for i, col in enumerate(columns):
            data[col] = row.iloc[i] if i < len(row) else ""
        
        self.sheet3_data.append(data)
    
    def add_to_sheet4(self, row, columns):
        """添加数据到sheet4（法人找到但平台未找到）"""
        data = {}
        for i, col in enumerate(columns):
            data[col] = row.iloc[i] if i < len(row) else ""
        
        self.sheet4_data.append(data)
    
    def export_to_excel(self):
        """导出数据到Excel"""
        self.log_message.emit("开始导出Excel文件...")
        
        with pd.ExcelWriter(self.output_file, engine='openpyxl') as writer:
            # Sheet1: 平台注册部数据在财务系统中
            if self.sheet1_data:
                df1 = pd.DataFrame(self.sheet1_data)
                df1.to_excel(writer, sheet_name='平台注册部数据在财务系统中', index=False)
                self.log_message.emit(f"Sheet1导出完成，共 {len(self.sheet1_data)} 条数据")
            
            # Sheet2: 日联部店铺
            if self.sheet2_data:
                df2 = pd.DataFrame(self.sheet2_data)
                df2.to_excel(writer, sheet_name='日联部店铺', index=False)
                self.log_message.emit(f"Sheet2导出完成，共 {len(self.sheet2_data)} 条数据")
            
            # Sheet3: 平台注册部不在财务系统的
            if self.sheet3_data:
                df3 = pd.DataFrame(self.sheet3_data)
                df3.to_excel(writer, sheet_name='平台注册部不在财务系统的', index=False)
                self.log_message.emit(f"Sheet3导出完成，共 {len(self.sheet3_data)} 条数据")
            
            # Sheet4: 法人找到但平台未找到
            if self.sheet4_data:
                df4 = pd.DataFrame(self.sheet4_data)
                df4.to_excel(writer, sheet_name='法人找到但平台未找到', index=False)
                self.log_message.emit(f"Sheet4导出完成，共 {len(self.sheet4_data)} 条数据")
        
        self.log_message.emit(f"Excel文件导出成功: {self.output_file}")


class PlatformRegisterClassifyWidget(QWidget):
    """平台注册部表分类功能界面"""
    
    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.worker = None
        self.init_ui()
    
    def init_ui(self):
        """初始化UI"""
        layout = QVBoxLayout()
        
        # 说明文字
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        info_label = QLabel(
            "此功能用于对平台注册部Excel数据进行分类：\n"
            "1. 导入Excel文件，B列为法人姓名，D列为平台类型\n"
            "2. 与财务系统中的ea_dy_legal、ea_dy_shop、ea_dy_site_type表进行匹配\n"
            "3. 根据匹配结果分类到四个Sheet：\n"
            "   - Sheet1: 平台注册部数据在财务系统中（B列和D列都匹配）\n"
            "   - Sheet2: 日联部店铺（数据库有但Excel没有的法人-平台组合）\n"
            "   - Sheet3: 平台注册部不在财务系统的（法人未匹配）\n"
            "   - Sheet4: 法人找到但平台未找到（法人匹配但平台不在数据库）"
        )
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # 配置区域
        config_group = QGroupBox("配置")
        config_layout = QFormLayout()
        
        # 数据源选择
        self.datasource_combo = QComboBox()
        self.load_datasources()
        config_layout.addRow("数据源:", self.datasource_combo)
        
        # Excel文件选择
        excel_layout = QHBoxLayout()
        self.excel_path_label = QLabel("未选择文件")
        self.excel_path_label.setStyleSheet("color: gray;")
        excel_layout.addWidget(self.excel_path_label)
        
        self.select_excel_btn = QPushButton("选择Excel文件")
        self.select_excel_btn.clicked.connect(self.select_excel_file)
        excel_layout.addWidget(self.select_excel_btn)
        
        config_layout.addRow("导入文件:", excel_layout)
        
        # 输出文件选择
        output_layout = QHBoxLayout()
        self.output_path_label = QLabel("未选择文件")
        self.output_path_label.setStyleSheet("color: gray;")
        output_layout.addWidget(self.output_path_label)
        
        self.select_output_btn = QPushButton("选择输出文件")
        self.select_output_btn.clicked.connect(self.select_output_file)
        output_layout.addWidget(self.select_output_btn)
        
        config_layout.addRow("输出文件:", output_layout)
        
        config_group.setLayout(config_layout)
        layout.addWidget(config_group)
        
        # 操作按钮
        button_layout = QHBoxLayout()
        
        self.start_btn = QPushButton("开始分类")
        self.start_btn.clicked.connect(self.start_classify)
        self.start_btn.setEnabled(False)
        button_layout.addWidget(self.start_btn)
        
        self.stop_btn = QPushButton("停止")
        self.stop_btn.clicked.connect(self.stop_classify)
        self.stop_btn.setEnabled(False)
        button_layout.addWidget(self.stop_btn)
        
        button_layout.addStretch()
        layout.addLayout(button_layout)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)
        
        # 日志区域
        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout()
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(200)
        log_layout.addWidget(self.log_text)
        
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)
        
        layout.addStretch()
        self.setLayout(layout)
    
    def load_datasources(self):
        """加载数据源列表"""
        self.datasource_combo.clear()
        datasources = self.db_manager.get_all_datasources()
        
        for ds in datasources:
            self.datasource_combo.addItem(ds.name, ds.id)
        
        if datasources:
            self.datasource_combo.setCurrentIndex(0)
    
    def select_excel_file(self):
        """选择Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Excel文件",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        
        if file_path:
            self.excel_path_label.setText(file_path)
            self.excel_path_label.setStyleSheet("color: black;")
            self.check_can_start()
    
    def select_output_file(self):
        """选择输出文件"""
        default_name = f"平台注册部分类结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "选择输出文件",
            default_name,
            "Excel Files (*.xlsx)"
        )
        
        if file_path:
            if not file_path.endswith('.xlsx'):
                file_path += '.xlsx'
            self.output_path_label.setText(file_path)
            self.output_path_label.setStyleSheet("color: black;")
            self.check_can_start()
    
    def check_can_start(self):
        """检查是否可以开始处理"""
        has_datasource = self.datasource_combo.count() > 0
        has_excel = self.excel_path_label.text() != "未选择文件"
        has_output = self.output_path_label.text() != "未选择文件"
        
        self.start_btn.setEnabled(has_datasource and has_excel and has_output)
    
    def start_classify(self):
        """开始分类"""
        # 获取选中的数据源
        datasource_id = self.datasource_combo.currentData()
        if datasource_id is None:
            QMessageBox.warning(self, "警告", "请先选择数据源！")
            return
        
        datasource = self.db_manager.get_datasource(datasource_id)
        if not datasource:
            QMessageBox.warning(self, "警告", "数据源不存在！")
            return
        
        excel_file = self.excel_path_label.text()
        output_file = self.output_path_label.text()
        
        if not os.path.exists(excel_file):
            QMessageBox.warning(self, "警告", "Excel文件不存在！")
            return
        
        # 禁用按钮
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.select_excel_btn.setEnabled(False)
        self.select_output_btn.setEnabled(False)
        self.datasource_combo.setEnabled(False)
        
        # 清空日志
        self.log_text.clear()
        self.progress_bar.setValue(0)
        
        # 创建工作线程
        self.worker = PlatformRegisterClassifyWorker(
            datasource, self.db_manager, excel_file, output_file
        )
        self.worker.progress.connect(self.on_progress)
        self.worker.log_message.connect(self.on_log_message)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()
    
    def stop_classify(self):
        """停止分类"""
        if self.worker and self.worker.isRunning():
            self.worker.terminate()
            self.worker.wait()
            self.log_text.append("处理已停止")
            self.reset_ui()
    
    def on_progress(self, value):
        """更新进度条"""
        self.progress_bar.setValue(value)
    
    def on_log_message(self, message):
        """添加日志消息"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.log_text.append(f"[{timestamp}] {message}")
        # 自动滚动到底部
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
    
    def on_finished(self, results):
        """处理完成"""
        self.log_text.append("=" * 50)
        self.log_text.append("处理完成！")
        self.log_text.append(f"总行数: {results['total_rows']}")
        self.log_text.append(f"Sheet1（平台注册部数据在财务系统中）: {results['sheet1_count']} 条")
        self.log_text.append(f"Sheet2（日联部店铺）: {results['sheet2_count']} 条")
        self.log_text.append(f"Sheet3（平台注册部不在财务系统的）: {results['sheet3_count']} 条")
        self.log_text.append(f"Sheet4（法人找到但平台未找到）: {results['sheet4_count']} 条")
        self.log_text.append("=" * 50)
        
        QMessageBox.information(
            self,
            "完成",
            f"分类完成！\n\n"
            f"总行数: {results['total_rows']}\n"
            f"Sheet1: {results['sheet1_count']} 条\n"
            f"Sheet2: {results['sheet2_count']} 条\n"
            f"Sheet3: {results['sheet3_count']} 条\n"
            f"Sheet4: {results['sheet4_count']} 条"
        )
        
        self.reset_ui()
    
    def on_error(self, error_msg):
        """处理错误"""
        self.log_text.append(f"错误: {error_msg}")
        QMessageBox.critical(self, "错误", error_msg)
        self.reset_ui()
    
    def reset_ui(self):
        """重置UI状态"""
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.select_excel_btn.setEnabled(True)
        self.select_output_btn.setEnabled(True)
        self.datasource_combo.setEnabled(True)
        self.check_can_start()
