"""
Tester Tool - 试验数据处理工具
主入口
"""

import sys
from pathlib import Path


def main():
    """主入口函数"""
    # 确保依赖可用
    try:
        from PyQt6.QtWidgets import QApplication
    except ImportError:
        print("错误: 缺少 PyQt6 依赖")
        print("请运行: pip install PyQt6")
        sys.exit(1)
    
    try:
        from openpyxl import load_workbook
    except ImportError:
        print("错误: 缺少 openpyxl 依赖")
        print("请运行: pip install openpyxl")
        sys.exit(1)
    
    # 启动 GUI
    from tester.gui import TesterApp
    
    app = QApplication(sys.argv)
    app.setApplicationName("试验数据处理工具")
    app.setOrganizationName("TesterTool")
    
    window = TesterApp()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
