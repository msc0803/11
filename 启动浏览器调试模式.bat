@echo off
chcp 65001 >nul
echo 正在以调试模式启动浏览器...
echo 启动后请登录 sucaiwang.zhishangsoft.com，然后再打开查询工具。
echo.

set EDGE1="C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
set EDGE2="C:\Program Files\Microsoft\Edge\Application\msedge.exe"
set CHROME1="C:\Program Files\Google\Chrome\Application\chrome.exe"
set CHROME2="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

if exist %EDGE1% (
    start "" %EDGE1% --remote-debugging-port=9222 https://sucaiwang.zhishangsoft.com
    echo 已使用 Microsoft Edge 启动
    goto done
)
if exist %EDGE2% (
    start "" %EDGE2% --remote-debugging-port=9222 https://sucaiwang.zhishangsoft.com
    echo 已使用 Microsoft Edge 启动
    goto done
)
if exist %CHROME1% (
    start "" %CHROME1% --remote-debugging-port=9222 https://sucaiwang.zhishangsoft.com
    echo 已使用 Google Chrome 启动
    goto done
)
if exist %CHROME2% (
    start "" %CHROME2% --remote-debugging-port=9222 https://sucaiwang.zhishangsoft.com
    echo 已使用 Google Chrome 启动
    goto done
)

echo 未找到 Edge 或 Chrome，请手动安装其中一个浏览器。
pause
exit /b 1

:done
echo 浏览器已启动！请登录后，再打开视频查询工具。
timeout /t 3 >nul
