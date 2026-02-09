"""
主窗口
"""
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QTableWidget, QTableWidgetItem,
                             QMessageBox, QHeaderView, QMenu, QAction, QLabel,
                             QTabWidget, QGroupBox, QDialog, QComboBox, QStackedWidget)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
from database import DatabaseManager
from models import DataSource
from ui.datasource_dialog import DataSourceDialog
from ui.excel_dedup_widget import ExcelDedupWidget
from ui.excel_to_db_widget import ExcelToDbWidget
from ui.legal_info_import_widget import LegalInfoImportWidget
from ui.shopindex_complete_widget import ShopIndexCompleteWidget
from ui.platform_info_complete_widget import PlatformInfoCompleteWidget
from ui.platform_status_update_widget import PlatformStatusUpdateWidget
from ui.bank_info_sync_widget import BankInfoSyncWidget
from ui.inventory_bank_complete_widget import InventoryBankCompleteWidget
from ui.order_pool_init_widget import OrderPoolInitWidget
from ui.ptzcb_register_widget import PtzcbRegisterWidget
from ui.platform_info_status_widget import PlatformInfoStatusWidget
from ui.platform_register_classify_widget import PlatformRegisterClassifyWidget
from ui.excel_compare_widget import ExcelCompareWidget
from .sync_inventory_widget import SyncInventoryWidget
from .table_merge_widget import TableMergeWidget
from .inventory_update_widget import InventoryUpdateWidget
from .platform_filter_widget import PlatformFilterWidget
from .rlb_company_info_widget import RlbCompanyInfoWidget
from .customer_info_complete_widget import CustomerInfoCompleteWidget
from .customer_status_fix_widget import CustomerStatusFixWidget
from .inventory_status_update_widget import InventoryStatusUpdateWidget
from .excel_json_format_widget import ExcelJsonFormatWidget


class MainWindow(QMainWindow):
    """主窗口"""
    
    def __init__(self):
        super().__init__()
        self.db_manager = DatabaseManager()
        self.init_ui()
        self.load_datasources()
    
    def init_ui(self):
        """初始化UI"""
        self.setWindowTitle("数据库迁移工具")
        self.setMinimumSize(1000, 600)
        
        # 中心部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout()
        
        # 创建标签页
        self.tab_widget = QTabWidget()
        
        # 数据源管理标签页
        datasource_tab = self.create_datasource_tab()
        self.tab_widget.addTab(datasource_tab, "数据源管理")
        
        # 迁移任务标签页（占位）
        migration_tab = self.create_migration_tab()
        self.tab_widget.addTab(migration_tab, "迁移任务")
        
        main_layout.addWidget(self.tab_widget)
        
        central_widget.setLayout(main_layout)
        
        # 状态栏
        self.statusBar().showMessage("就绪")
    
    def create_datasource_tab(self) -> QWidget:
        """创建数据源管理标签页"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        # 顶部按钮栏
        button_layout = QHBoxLayout()
        
        self.add_btn = QPushButton("添加数据源")
        self.add_btn.clicked.connect(self.add_datasource)
        button_layout.addWidget(self.add_btn)
        
        self.edit_btn = QPushButton("编辑")
        self.edit_btn.clicked.connect(self.edit_datasource)
        self.edit_btn.setEnabled(False)
        button_layout.addWidget(self.edit_btn)
        
        self.delete_btn = QPushButton("删除")
        self.delete_btn.clicked.connect(self.delete_datasource)
        self.delete_btn.setEnabled(False)
        button_layout.addWidget(self.delete_btn)
        
        self.refresh_btn = QPushButton("刷新")
        self.refresh_btn.clicked.connect(self.load_datasources)
        button_layout.addWidget(self.refresh_btn)
        
        self.test_btn = QPushButton("测试连接")
        self.test_btn.clicked.connect(self.test_selected_datasource)
        self.test_btn.setEnabled(False)
        button_layout.addWidget(self.test_btn)
        
        button_layout.addStretch()
        
        layout.addLayout(button_layout)
        
        # 数据源表格
        self.datasource_table = QTableWidget()
        self.datasource_table.setColumnCount(7)
        self.datasource_table.setHorizontalHeaderLabels([
            "ID", "名称", "主机", "端口", "数据库", "用户名", "描述"
        ])
        
        # 设置表格属性
        self.datasource_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.datasource_table.setSelectionMode(QTableWidget.SingleSelection)
        self.datasource_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.datasource_table.setAlternatingRowColors(True)
        
        # 设置列宽
        header = self.datasource_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.Stretch)
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(6, QHeaderView.Stretch)
        
        # 连接选择信号
        self.datasource_table.itemSelectionChanged.connect(self.on_datasource_selection_changed)
        
        # 双击编辑
        self.datasource_table.doubleClicked.connect(self.edit_datasource)
        
        # 右键菜单
        self.datasource_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.datasource_table.customContextMenuRequested.connect(self.show_context_menu)
        
        layout.addWidget(self.datasource_table)
        
        widget.setLayout(layout)
        return widget
    
    def create_migration_tab(self) -> QWidget:
        """创建迁移任务标签页"""
        widget = QWidget()
        layout = QVBoxLayout()
        
        # 功能选择区域
        function_group = QGroupBox("选择迁移功能")
        function_layout = QHBoxLayout()
        
        function_label = QLabel("功能类型:")
        function_layout.addWidget(function_label)
        
        self.function_combo = QComboBox()
        self.function_combo.addItem("Excel表格去重", "excel_dedup")
        self.function_combo.addItem("Excel导入客户表", "excel_to_db")
        self.function_combo.addItem("Excel导入法人信息", "legal_info_import")
        self.function_combo.addItem("补全主体编号", "shopindex_complete")
        self.function_combo.addItem("补全平台详情信息", "platform_info_complete")
        self.function_combo.addItem("修改平台详情状态", "platform_status_update")
        self.function_combo.addItem("现有客户银行信息补全", "bank_info_sync")
        self.function_combo.addItem("补全现有库存关联银行卡信息", "inventory_bank_complete")
        self.function_combo.addItem("导入平台注册部Excel初始化订单池", "order_pool_init")
        self.function_combo.addItem("根据Excel补充系统平台注册部表信息", "ptzcb_register")
        self.function_combo.addItem("过滤掉平台为银行的数据", "platform_filter")
        self.function_combo.addItem("验证后的非平台注册部数据平台详情补全", "platform_info_status")
        self.function_combo.addItem("平台注册部表分类", "platform_register_classify")
        self.function_combo.addItem("Excel表对比", "excel_compare")
        self.function_combo.addItem("同步库存账号", "sync_inventory")
        self.function_combo.addItem("表格对比合并", "table_merge")
        self.function_combo.addItem("库存内容补充", "inventory_update")
        self.function_combo.addItem("补全日联部企业基础信息表", "rlb_company_info")
        self.function_combo.addItem("补全客户信息字段初始值", "customer_info_complete")
        self.function_combo.addItem("修正is_customer字段", "customer_status_fix")
        self.function_combo.addItem("按库存主键更新状态", "inventory_status_update")
        self.function_combo.addItem("Excel JSON格式转换", "excel_json_format")
        self.function_combo.addItem("数据库表迁移（待开发）", "db_migrate")
        self.function_combo.addItem("数据格式转换（待开发）", "format_convert")
        self.function_combo.currentTextChanged.connect(self.on_function_changed)
        function_layout.addWidget(self.function_combo)
        
        function_layout.addStretch()
        function_group.setLayout(function_layout)
        layout.addWidget(function_group)
        
        # 功能内容区域（使用堆叠窗口）
        self.function_stack = QStackedWidget()
        
        # Excel去重功能
        self.excel_dedup_widget = ExcelDedupWidget()
        self.function_stack.addWidget(self.excel_dedup_widget)
        
        # Excel导入客户表功能
        self.excel_to_db_widget = ExcelToDbWidget(self.db_manager)
        self.function_stack.addWidget(self.excel_to_db_widget)
        
        # Excel导入法人信息功能
        self.legal_info_import_widget = LegalInfoImportWidget(self.db_manager)
        self.function_stack.addWidget(self.legal_info_import_widget)
        
        # 补全主体编号功能
        self.shopindex_complete_widget = ShopIndexCompleteWidget(self.db_manager)
        self.function_stack.addWidget(self.shopindex_complete_widget)
        
        # 补全平台详情信息功能
        self.platform_info_complete_widget = PlatformInfoCompleteWidget(self.db_manager)
        self.function_stack.addWidget(self.platform_info_complete_widget)
        
        # 修改平台详情状态功能
        self.platform_status_update_widget = PlatformStatusUpdateWidget(self.db_manager)
        self.function_stack.addWidget(self.platform_status_update_widget)
        
        # 现有客户银行信息补全功能
        self.bank_info_sync_widget = BankInfoSyncWidget(self.db_manager)
        self.function_stack.addWidget(self.bank_info_sync_widget)
        
        # 补全现有库存关联银行卡信息功能
        self.inventory_bank_complete_widget = InventoryBankCompleteWidget(self.db_manager)
        self.function_stack.addWidget(self.inventory_bank_complete_widget)
        
        # 导入平台注册部Excel初始化订单池功能
        self.order_pool_init_widget = OrderPoolInitWidget(self.db_manager)
        self.function_stack.addWidget(self.order_pool_init_widget)
        
        # 根据Excel补充系统平台注册部表信息功能
        self.ptzcb_register_widget = PtzcbRegisterWidget(self.db_manager)
        self.function_stack.addWidget(self.ptzcb_register_widget)
        
        # 过滤掉平台为银行的数据功能
        self.platform_filter_widget = PlatformFilterWidget(self.db_manager)
        self.function_stack.addWidget(self.platform_filter_widget)
        
        # 验证后的非平台注册部数据平台详情补全功能
        self.platform_info_status_widget = PlatformInfoStatusWidget(self.db_manager)
        self.function_stack.addWidget(self.platform_info_status_widget)
        
        # 平台注册部表分类功能
        self.platform_register_classify_widget = PlatformRegisterClassifyWidget(self.db_manager)
        self.function_stack.addWidget(self.platform_register_classify_widget)
        
        # Excel表对比功能
        self.excel_compare_widget = ExcelCompareWidget()
        self.function_stack.addWidget(self.excel_compare_widget)
        
        # 同步库存账号功能
        self.sync_inventory_widget = SyncInventoryWidget(self.db_manager)
        self.function_stack.addWidget(self.sync_inventory_widget)
        
        # 表格对比合并功能
        self.table_merge_widget = TableMergeWidget()
        self.function_stack.addWidget(self.table_merge_widget)
        
        # 库存内容补充功能
        self.inventory_update_widget = InventoryUpdateWidget(self.db_manager)
        self.function_stack.addWidget(self.inventory_update_widget)
        
        # 补全日联部企业基础信息表功能
        self.rlb_company_info_widget = RlbCompanyInfoWidget(self.db_manager)
        self.function_stack.addWidget(self.rlb_company_info_widget)
        
        # 补全客户信息字段初始值功能
        self.customer_info_complete_widget = CustomerInfoCompleteWidget(self.db_manager)
        self.function_stack.addWidget(self.customer_info_complete_widget)
        
        # 修正is_customer字段功能
        self.customer_status_fix_widget = CustomerStatusFixWidget(self.db_manager)
        self.function_stack.addWidget(self.customer_status_fix_widget)
        
        # 按库存主键更新状态功能
        self.inventory_status_update_widget = InventoryStatusUpdateWidget(self.db_manager)
        self.function_stack.addWidget(self.inventory_status_update_widget)
        
        # Excel JSON格式转换功能
        self.excel_json_format_widget = ExcelJsonFormatWidget(self.db_manager)
        self.function_stack.addWidget(self.excel_json_format_widget)
        
        # 占位页面（其他功能）
        placeholder_widget = QWidget()
        placeholder_layout = QVBoxLayout()
        placeholder_label = QLabel("该功能正在开发中...")
        placeholder_label.setAlignment(Qt.AlignCenter)
        placeholder_label.setStyleSheet("font-size: 16px; color: gray;")
        placeholder_layout.addWidget(placeholder_label)
        placeholder_widget.setLayout(placeholder_layout)
        self.function_stack.addWidget(placeholder_widget)
        
        layout.addWidget(self.function_stack)
        
        widget.setLayout(layout)
        return widget
    
    def on_function_changed(self, text):
        """功能选择变化"""
        if "Excel表格去重" in text:
            self.function_stack.setCurrentIndex(0)
        elif "Excel导入客户表" in text:
            self.function_stack.setCurrentIndex(1)
        elif "Excel导入法人信息" in text:
            self.function_stack.setCurrentIndex(2)
        elif "补全主体编号" in text:
            self.function_stack.setCurrentIndex(3)
        elif "补全平台详情信息" in text:
            self.function_stack.setCurrentIndex(4)
        elif "修改平台详情状态" in text:
            self.function_stack.setCurrentIndex(5)
        elif "现有客户银行信息补全" in text:
            self.function_stack.setCurrentIndex(6)
        elif "补全现有库存关联银行卡信息" in text:
            self.function_stack.setCurrentIndex(7)
        elif "导入平台注册部Excel初始化订单池" in text:
            self.function_stack.setCurrentIndex(8)
        elif "根据Excel补充系统平台注册部表信息" in text:
            self.function_stack.setCurrentIndex(9)
        elif "过滤掉平台为银行的数据" in text:
            self.function_stack.setCurrentIndex(10)
        elif "验证后的非平台注册部数据平台详情补全" in text:
            self.function_stack.setCurrentIndex(11)
        elif "平台注册部表分类" in text:
            self.function_stack.setCurrentIndex(12)
        elif "Excel表对比" in text:
            self.function_stack.setCurrentIndex(13)
        elif "同步库存账号" in text:
            self.function_stack.setCurrentIndex(14)
        elif "表格对比合并" in text:
            self.function_stack.setCurrentIndex(15)
        elif "库存内容补充" in text:
            self.function_stack.setCurrentIndex(16)
        elif "补全日联部企业基础信息表" in text:
            self.function_stack.setCurrentIndex(17)
        elif "补全客户信息字段初始值" in text:
            self.function_stack.setCurrentIndex(18)
        elif "修正is_customer字段" in text:
            self.function_stack.setCurrentIndex(19)
        elif "按库存主键更新状态" in text:
            self.function_stack.setCurrentIndex(20)
        elif "Excel JSON格式转换" in text:
            self.function_stack.setCurrentIndex(21)
        else:
            self.function_stack.setCurrentIndex(22)
    
    def load_datasources(self):
        """加载数据源列表"""
        self.datasource_table.setRowCount(0)
        
        datasources = self.db_manager.get_all_datasources()
        
        for datasource in datasources:
            row = self.datasource_table.rowCount()
            self.datasource_table.insertRow(row)
            
            self.datasource_table.setItem(row, 0, QTableWidgetItem(str(datasource.id)))
            self.datasource_table.setItem(row, 1, QTableWidgetItem(datasource.name))
            self.datasource_table.setItem(row, 2, QTableWidgetItem(datasource.host))
            self.datasource_table.setItem(row, 3, QTableWidgetItem(str(datasource.port)))
            self.datasource_table.setItem(row, 4, QTableWidgetItem(datasource.database))
            self.datasource_table.setItem(row, 5, QTableWidgetItem(datasource.username))
            self.datasource_table.setItem(row, 6, QTableWidgetItem(datasource.description))
        
        self.statusBar().showMessage(f"共 {len(datasources)} 个数据源")
    
    def on_datasource_selection_changed(self):
        """数据源选择变化"""
        has_selection = len(self.datasource_table.selectedItems()) > 0
        self.edit_btn.setEnabled(has_selection)
        self.delete_btn.setEnabled(has_selection)
        self.test_btn.setEnabled(has_selection)
    
    def get_selected_datasource_id(self) -> int:
        """获取选中的数据源ID"""
        selected_rows = self.datasource_table.selectedItems()
        if selected_rows:
            row = selected_rows[0].row()
            return int(self.datasource_table.item(row, 0).text())
        return None
    
    def add_datasource(self):
        """添加数据源"""
        dialog = DataSourceDialog(self, db_manager=self.db_manager)
        if dialog.exec_() == QDialog.Accepted:
            self.load_datasources()
    
    def edit_datasource(self):
        """编辑数据源"""
        datasource_id = self.get_selected_datasource_id()
        if datasource_id is None:
            return
        
        datasource = self.db_manager.get_datasource(datasource_id)
        if datasource:
            dialog = DataSourceDialog(self, datasource=datasource, db_manager=self.db_manager)
            if dialog.exec_() == QDialog.Accepted:
                self.load_datasources()
    
    def delete_datasource(self):
        """删除数据源"""
        datasource_id = self.get_selected_datasource_id()
        if datasource_id is None:
            return
        
        reply = QMessageBox.question(
            self,
            "确认删除",
            "确定要删除选中的数据源吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            success = self.db_manager.delete_datasource(datasource_id)
            if success:
                QMessageBox.information(self, "成功", "数据源删除成功！")
                self.load_datasources()
            else:
                QMessageBox.warning(self, "失败", "数据源删除失败！")
    
    def test_selected_datasource(self):
        """测试选中的数据源连接"""
        datasource_id = self.get_selected_datasource_id()
        if datasource_id is None:
            return
        
        datasource = self.db_manager.get_datasource(datasource_id)
        if datasource:
            self.statusBar().showMessage("正在测试连接...")
            success, message = self.db_manager.test_connection(datasource)
            
            if success:
                QMessageBox.information(self, "连接成功", message)
                self.statusBar().showMessage("连接测试成功")
            else:
                QMessageBox.critical(self, "连接失败", message)
                self.statusBar().showMessage("连接测试失败")
    
    def show_context_menu(self, pos):
        """显示右键菜单"""
        if not self.datasource_table.itemAt(pos):
            return
        
        menu = QMenu(self)
        
        edit_action = QAction("编辑", self)
        edit_action.triggered.connect(self.edit_datasource)
        menu.addAction(edit_action)
        
        test_action = QAction("测试连接", self)
        test_action.triggered.connect(self.test_selected_datasource)
        menu.addAction(test_action)
        
        menu.addSeparator()
        
        delete_action = QAction("删除", self)
        delete_action.triggered.connect(self.delete_datasource)
        menu.addAction(delete_action)
        
        menu.exec_(self.datasource_table.viewport().mapToGlobal(pos))
