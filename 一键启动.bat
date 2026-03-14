@echo off
chcp 65001 >nul
echo ========================================
echo   视频素材批量查询工具 - 一键启动
echo ========================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo 未检测到 Python，请先安装 Python 3.10：
    echo https://www.python.org/downloads/release/python-31011/
    echo 安装时请勾选 "Add Python to PATH"
    pause
    exit /b 1
)

echo 正在检查并安装依赖...
python -m pip install pyqt6 pandas openpyxl xlrd requests websockets numbers-parser -q
if errorlevel 1 (
    echo 依赖安装失败，请检查网络或手动运行：
    echo pip install pyqt6 pandas openpyxl xlrd requests websockets numbers-parser
    pause
    exit /b 1
)

echo 启动中...
python gui_app.py
