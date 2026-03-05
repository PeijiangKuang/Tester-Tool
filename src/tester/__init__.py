"""
Tester Tool - 试验数据处理工具
主入口模块
"""

import sys
from pathlib import Path

# 确保 src 目录在 Python 路径中
src_path = Path(__file__).parent.parent / "src"
if str(src_path) not in sys.path:
    sys.path.insert(0, str(src_path))

from tester.gui import main

if __name__ == "__main__":
    main()
