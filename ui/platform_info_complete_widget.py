"""
平台详情信息补全功能界面
"""
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QMessageBox, QComboBox,
                             QGroupBox, QProgressBar, QTextEdit, QCheckBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from database import DatabaseManager
import pymysql


class PlatformInfoCompleteWorker(QThread):
    """平台详情信息补全工作线程"""
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal(dict)  # 返回统计结果
    error = pyqtSignal(str)
    
    def __init__(self, datasource, db_manager):
        super().__init__()
        self.datasource = datasource
        self.db_manager = db_manager
        self.results = {
            'shopindex_total': 0,
            'shopindex_processed': 0,
            'shopindex_skipped': 0,
            'platform_info_added': 0,
            'failed': 0
        }
    
    def run(self):
        try:
            self.log_message.emit("开始连接数据库...")
            self.progress.emit(5)
            
            # 连接数据库
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
            self.progress.emit(10)
            
            # 获取基础数据统计
            self.log_message.emit("获取基础数据统计...")
            
            # 获取属地数量
            cursor.execute("SELECT COUNT(*) FROM ba_shudi WHERE status = 1")
            shudi_count = cursor.fetchone()[0]
            
            # 获取平台数量
            cursor.execute("SELECT COUNT(*) FROM ba_platform WHERE status = 1")
            platform_count = cursor.fetchone()[0]
            
            # 每个shopindex应该有的platform_info数量
            expected_per_shopindex = shudi_count * platform_count
            
            self.log_message.emit(f"属地数量: {shudi_count}, 平台数量: {platform_count}")
            self.log_message.emit(f"每个主体编号应该有 {expected_per_shopindex} 条平台详情记录")
            
            if expected_per_shopindex == 0:
                self.error.emit("没有可用的属地或平台数据，无法进行补全")
                return
            
            self.progress.emit(20)
            
            # 获取所有shopindex记录
            cursor.execute("""
                SELECT id, admin_id, admin_dept_id 
                FROM ba_shopindex 
                WHERE status = 1 AND (delete_time IS NULL OR delete_time = 0)
            """)
            shopindex_records = cursor.fetchall()
            
            self.results['shopindex_total'] = len(shopindex_records)
            self.log_message.emit(f"找到 {len(shopindex_records)} 个主体编号记录")
            
            if len(shopindex_records) == 0:
                self.log_message.emit("没有找到可处理的主体编号记录")
                self.finished.emit(self.results)
                return
            
            self.progress.emit(30)
            
            # 批量检查和处理
            current_time = int(datetime.now().timestamp())
            batch_size = 100  # 批量处理大小
            
            for i in range(0, len(shopindex_records), batch_size):
                batch = shopindex_records[i:i + batch_size]
                self.process_batch(cursor, batch, expected_per_shopindex, current_time)
                
                # 更新进度
                progress = 30 + int((i + len(batch)) / len(shopindex_records) * 60)
                self.progress.emit(progress)
                
                # 提交批次
                connection.commit()
            
            cursor.close()
            connection.close()
            
            self.progress.emit(100)
            self.finished.emit(self.results)
            
        except Exception as e:
            self.error.emit(str(e))
    
    def process_batch(self, cursor, batch, expected_count, current_time):
        """批量处理shopindex记录"""
        try:
            # 构建shopindex_id列表用于批量查询
            shopindex_ids = [str(record[0]) for record in batch]
            shopindex_ids_str = ','.join(shopindex_ids)
            
            # 批量查询现有的platform_info数量
            cursor.execute(f"""
                SELECT shopindex_id, COUNT(*) as count
                FROM ba_platform_info 
                WHERE shopindex_id IN ({shopindex_ids_str}) AND (delete_time IS NULL OR delete_time = 0)
                GROUP BY shopindex_id
            """)
            existing_counts = dict(cursor.fetchall())
            
            # 获取所有属地和平台的组合
            cursor.execute("""
                SELECT s.id as shudi_id, p.id as platform_id
                FROM ba_shudi s
                CROSS JOIN ba_platform p
                WHERE s.status = 1 AND p.status = 1
                ORDER BY s.id, p.id
            """)
            shudi_platform_combinations = cursor.fetchall()
            
            # 处理每个shopindex
            for shopindex_record in batch:
                shopindex_id, admin_id, admin_dept_id = shopindex_record
                
                try:
                    # 检查现有数量
                    existing_count = existing_counts.get(shopindex_id, 0)
                    
                    if existing_count >= expected_count:
                        self.results['shopindex_skipped'] += 1
                        self.log_message.emit(f"主体编号ID {shopindex_id}: 已有 {existing_count} 条记录，跳过")
                        continue
                    
                    # 需要补充的数量
                    need_to_add = expected_count - existing_count
                    
                    # 获取已存在的组合
                    cursor.execute("""
                        SELECT shudi_id, platform_id
                        FROM ba_platform_info
                        WHERE shopindex_id = %s AND (delete_time IS NULL OR delete_time = 0)
                    """, (shopindex_id,))
                    existing_combinations = set(cursor.fetchall())
                    
                    # 准备批量插入数据
                    insert_data = []
                    added_count = 0
                    
                    for shudi_id, platform_id in shudi_platform_combinations:
                        if (shudi_id, platform_id) not in existing_combinations:
                            insert_data.append((
                                platform_id,      # platform_id
                                shudi_id,         # shudi_id
                                1,                # status (可用)
                                current_time,     # create_time
                                current_time,     # update_time
                                shopindex_id,     # shopindex_id
                                admin_id or 1,    # admin_id
                                admin_dept_id or 1  # admin_dept_id
                            ))
                            added_count += 1
                            
                            if added_count >= need_to_add:
                                break
                    
                    # 批量插入
                    if insert_data:
                        cursor.executemany("""
                            INSERT INTO ba_platform_info 
                            (platform_id, shudi_id, status, create_time, update_time, shopindex_id, admin_id, admin_dept_id)
                            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                        """, insert_data)
                        
                        self.results['platform_info_added'] += len(insert_data)
                        self.results['shopindex_processed'] += 1
                        self.log_message.emit(f"主体编号ID {shopindex_id}: 补充了 {len(insert_data)} 条记录")
                    else:
                        self.results['shopindex_skipped'] += 1
                        self.log_message.emit(f"主体编号ID {shopindex_id}: 无需补充")
                        
                except Exception as e:
                    self.results['failed'] += 1
                    self.log_message.emit(f"处理主体编号ID {shopindex_id} 时出错: {str(e)}")
                    
        except Exception as e:
            self.log_message.emit(f"批量处理出错: {str(e)}")


class PlatformInfoCompleteWidget(QWidget):
    """平台详情信息补全功能组件"""
    
    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
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
        
        # 功能说明区域
        info_group = QGroupBox("功能说明")
        info_layout = QVBoxLayout()
        
        info_label = QLabel("""
<b>平台详情信息补全说明:</b><br>
• 为每个主体编号(ba_shopindex)补全平台详情信息(ba_platform_info)<br>
• 补全规则: 每个主体编号 × 所有属地 × 所有平台 = 完整的平台详情记录<br>
• 智能检查: 如果某个主体编号的记录数已经等于 属地数×平台数，则跳过<br>
• 高效处理: 使用批量查询和批量插入，避免逐条处理的低效问题<br><br>
<b>字段填充:</b><br>
• platform_id, shudi_id: 所有可能的组合<br>
• status: 1 (可用状态)<br>
• admin_id, admin_dept_id: 来自对应的ba_shopindex记录<br>
• create_time, update_time: 当前时间戳<br><br>
<b>性能优化:</b><br>
• 批量处理，避免单条插入<br>
• 智能跳过已完整的记录<br>
• 事务提交优化
        """)
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # 操作按钮区域
        button_layout = QHBoxLayout()
        
        self.start_btn = QPushButton("开始补全")
        self.start_btn.clicked.connect(self.start_complete)
        self.start_btn.setEnabled(False)
        button_layout.addWidget(self.start_btn)
        
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
        self.start_btn.setEnabled(current_data is not None)
    
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
    
    def start_complete(self):
        """开始补全"""
        datasource = self.datasource_combo.currentData()
        if not datasource:
            QMessageBox.warning(self, "警告", "请选择数据源！")
            return
        
        # 确认对话框
        reply = QMessageBox.question(
            self,
            "确认补全",
            f"确定要补全平台详情信息吗？\n\n"
            f"数据库: {datasource.name}\n\n"
            f"此操作将为所有主体编号补全缺失的平台详情记录。\n"
            f"系统会智能检查已有记录，避免重复处理。\n\n"
            f"注意：如果数据量很大，可能需要较长时间。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        # 禁用按钮，显示进度条
        self.start_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.result_label.setText("正在补全中...")
        self.log_text.clear()
        
        # 启动工作线程
        self.worker = PlatformInfoCompleteWorker(datasource, self.db_manager)
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log_message.connect(self.append_log)
        self.worker.finished.connect(self.on_complete_finished)
        self.worker.error.connect(self.on_complete_error)
        self.worker.start()
    
    def append_log(self, message):
        """添加日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
    
    def on_complete_finished(self, results):
        """补全完成"""
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        result_text = (
            f"补全完成！\n"
            f"主体编号总数：{results['shopindex_total']} 个\n"
            f"已处理：{results['shopindex_processed']} 个\n"
            f"跳过(已完整)：{results['shopindex_skipped']} 个\n"
            f"新增平台详情记录：{results['platform_info_added']} 条\n"
            f"失败：{results['failed']} 个"
        )
        
        self.result_label.setText(result_text)
        self.result_label.setStyleSheet("color: green;")
        
        QMessageBox.information(self, "补全完成", result_text)
    
    def on_complete_error(self, error_msg):
        """补全错误"""
        self.start_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.result_label.setText("补全失败！")
        self.result_label.setStyleSheet("color: red;")
        
        QMessageBox.critical(self, "补全错误", f"补全失败：{error_msg}")
    
    def clear_log(self):
        """清空日志"""
        self.log_text.clear()
