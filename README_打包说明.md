# Excel工具打包说明

## 方法一：使用批处理文件（推荐，最简单）

1. 双击运行 `build_exe.bat`
2. 等待打包完成
3. 在 `dist` 文件夹中找到 `Excel拆分合并工具.exe`

## 方法二：使用Python脚本

1. 运行命令：
   ```bash
   python build_exe.py
   ```

## 方法三：手动使用PyInstaller

1. 安装PyInstaller（如果未安装）：
   ```bash
   pip install pyinstaller
   ```

2. 执行打包命令：
   ```bash
   pyinstaller --name=Excel拆分合并工具 --onefile --windowed --clean excel_tool_gui.py
   ```

3. 打包完成后，exe文件位于 `dist` 文件夹中

## 打包参数说明

- `--name=Excel拆分合并工具`: 生成的exe文件名
- `--onefile`: 打包为单个exe文件（推荐，方便分发）
- `--windowed`: 不显示控制台窗口（GUI应用）
- `--clean`: 清理临时文件

## 注意事项

1. 首次打包可能需要较长时间（下载依赖）
2. 打包生成的exe文件可能较大（包含Python解释器和所有依赖）
3. exe文件可以在任何Windows系统上运行，无需安装Python
4. 如果被杀毒软件误报，可以添加到白名单

## 依赖说明

程序依赖以下Python库：
- pandas: 用于Excel文件读写和合并
- openpyxl: 用于Excel文件操作和样式设置
- tkinter: GUI界面（Python标准库）

这些依赖会在打包时自动包含在exe文件中。


