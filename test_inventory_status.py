"""
测试库存状态更新功能
"""
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from PyQt5.QtWidgets import QApplication
from database import DatabaseManager
from ui.inventory_status_update_widget import InventoryStatusUpdateWidget

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # 创建数据库管理器
    db_manager = DatabaseManager()
    
    # 创建窗口
    window = InventoryStatusUpdateWidget(db_manager)
    window.show()
    
    sys.exit(app.exec_())
