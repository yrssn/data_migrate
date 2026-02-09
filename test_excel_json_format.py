"""
测试Excel JSON格式转换功能
"""
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from PyQt5.QtWidgets import QApplication
from database import DatabaseManager
from ui.excel_json_format_widget import ExcelJsonFormatWidget

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # 创建数据库管理器
    db_manager = DatabaseManager()
    
    # 创建Excel JSON格式转换界面
    widget = ExcelJsonFormatWidget(db_manager)
    widget.setWindowTitle("Excel JSON格式转换测试")
    widget.resize(1200, 800)
    widget.show()
    
    sys.exit(app.exec_())
