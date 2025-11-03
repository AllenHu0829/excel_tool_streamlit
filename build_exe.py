"""
打包Excel工具为exe可执行文件的脚本
使用 PyInstaller 进行打包
"""

import subprocess
import sys
import os

def install_pyinstaller():
    """安装PyInstaller如果未安装"""
    try:
        import PyInstaller
        print("PyInstaller 已安装")
    except ImportError:
        print("正在安装 PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("PyInstaller 安装完成")

def build_exe():
    """打包为exe文件"""
    # 检查并安装PyInstaller
    install_pyinstaller()
    
    # PyInstaller命令参数
    cmd = [
        "pyinstaller",
        "--name=Excel拆分合并工具",  # exe文件名
        "--onefile",  # 打包为单个exe文件
        "--windowed",  # 不显示控制台窗口（GUI应用）
        "--icon=NONE",  # 可以指定图标文件
        "--clean",  # 清理临时文件
        "excel_tool_gui.py"
    ]
    
    print("开始打包...")
    print(f"执行命令: {' '.join(cmd)}")
    
    try:
        subprocess.check_call(cmd)
        print("\n" + "="*50)
        print("打包完成！")
        print("="*50)
        print(f"exe文件位置: {os.path.abspath('dist/Excel拆分合并工具.exe')}")
        print("\n提示：")
        print("1. exe文件位于 'dist' 文件夹中")
        print("2. 可以将此exe文件复制到其他Windows电脑上直接运行")
        print("3. 不需要安装Python环境")
    except subprocess.CalledProcessError as e:
        print(f"打包失败: {e}")
        sys.exit(1)

if __name__ == "__main__":
    build_exe()


