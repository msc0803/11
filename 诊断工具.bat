@echo off
chcp 65001 >nul
echo ===== 浏览器诊断工具 =====
echo.

echo [1] 检查 Edge 路径...
if exist "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" (
    echo   找到: C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe
) else if exist "C:\Program Files\Microsoft\Edge\Application\msedge.exe" (
    echo   找到: C:\Program Files\Microsoft\Edge\Application\msedge.exe
) else (
    echo   未找到 Edge
)

echo.
echo [2] 检查 Chrome 路径...
if exist "C:\Program Files\Google\Chrome\Application\chrome.exe" (
    echo   找到: C:\Program Files\Google\Chrome\Application\chrome.exe
) else if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" (
    echo   找到: C:\Program Files (x86)\Google\Chrome\Application\chrome.exe
) else (
    echo   未找到 Chrome
)

echo.
echo [3] 检查 9222 端口...
netstat -an | findstr "9222"
if errorlevel 1 echo   9222 端口未占用

echo.
echo [4] 检查运行中的浏览器进程...
tasklist | findstr /i "chrome msedge"

echo.
echo ===== 诊断完成，请截图发给开发者 =====
pause
