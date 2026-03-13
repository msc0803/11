@echo off
chcp 65001 >nul
echo 正在以调试模式启动 Chrome...
echo 启动后请登录 sucaiwang.zhishangsoft.com，然后再打开查询工具。
echo.

set CHROME1="C:\Program Files\Google\Chrome\Application\chrome.exe"
set CHROME2="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

if exist %CHROME1% (
    start "" %CHROME1% --remote-debugging-port=9222
    goto done
)
if exist %CHROME2% (
    start "" %CHROME2% --remote-debugging-port=9222
    goto done
)

echo 未找到 Chrome，请手动指定路径。
pause
exit /b 1

:done
echo Chrome 已启动！请在浏览器登录后，再打开视频查询工具。
timeout /t 3 >nul
