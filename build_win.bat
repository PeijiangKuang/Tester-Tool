@echo off
echo ========================================
echo   试验数据处理工具 - Windows 打包
echo ========================================
echo.

echo [1/4] 安装依赖...
pip install pyinstaller PyQt6 openpyxl pandas -q

echo [2/4] 打包程序...
pyinstaller --onefile --windowed --name "TesterTool" src\tester\__main__.py

echo [3/4] 清理...
rmdir /s /q build 2>nul
del *.spec 2>nul

echo [4/4] 完成!
echo.
echo ========================================
echo   ✅ 打包成功！
echo ========================================
echo.
echo 输出文件: dist\TesterTool.exe
echo.
echo 使用方法:
echo   1. 双击运行 TesterTool.exe
echo.
pause
