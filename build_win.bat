@echo off
echo Checking Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo Python not found! Please install Python 3.10:
    echo https://www.python.org/downloads/release/python-31011/
    echo Make sure to check "Add Python to PATH"
    pause
    exit /b 1
)
python --version

echo.
echo Installing dependencies...
python -m pip install --upgrade pip -q
python -m pip install pyqt6 pandas openpyxl requests websockets pyinstaller -q
if errorlevel 1 (
    echo Failed to install dependencies!
    pause
    exit /b 1
)
echo Dependencies installed OK

echo.
echo Building exe...
python -m PyInstaller --noconfirm --onefile --windowed --name "VideoQueryTool" --add-data "api_core.py;." gui_app.py
if errorlevel 1 (
    echo Build failed!
    pause
    exit /b 1
)

echo.
echo Done! exe is in dist\VideoQueryTool.exe
pause
