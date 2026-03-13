@echo off
chcp 65001 >nul
echo ========================================
echo  视频查询工具 - Windows 自动打包脚本
echo ========================================
echo.

echo [1/4] 检查 Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo 未找到 Python！请先安装 Python 3.10：
    echo https://www.python.org/downloads/release/python-31011/
    echo 安装时勾选 "Add Python to PATH"
    pause
    exit /b 1
)
python --version

echo.
echo [2/4] 安装依赖...
python -m pip install --upgrade pip -q
python -m pip install pyqt6 pandas openpyxl requests websockets pyinstaller -q
if errorlevel 1 (
    echo 安装依赖失败！
    pause
    exit /b 1
)
echo 依赖安装完成

echo.
echo [3/4] 开始打包...
python -m PyInstaller ^
    --noconfirm ^
    --onefile ^
    --windowed ^
    --name "视频查询工具" ^
    --add-data "api_core.py;." ^
    gui_app.py

if errorlevel 1 (
    echo 打包失败！
    pause
    exit /b 1
)

echo.
echo [4/4] 完成！
echo exe 文件在 dist\ 目录下：视频查询工具.exe
echo.
pause
