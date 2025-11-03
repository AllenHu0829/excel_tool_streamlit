@echo off
chcp 65001 >nul
echo ========================================
echo Excel工具打包脚本
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python，请先安装Python
    pause
    exit /b 1
)

echo 正在检查并安装依赖...
python -m pip install --upgrade pip
python -m pip install pyinstaller pandas openpyxl

echo.
echo 开始打包...
echo.

REM 使用PyInstaller打包
pyinstaller --name=Excel拆分合并工具 --onefile --windowed --clean excel_tool_gui.py

if errorlevel 1 (
    echo.
    echo 打包失败！
    pause
    exit /b 1
)

echo.
echo ========================================
echo 打包完成！
echo ========================================
echo.
echo exe文件位置: dist\Excel拆分合并工具.exe
echo.
echo 提示：
echo 1. 可以将此exe文件复制到其他Windows电脑上直接运行
echo 2. 不需要安装Python环境
echo.
pause


