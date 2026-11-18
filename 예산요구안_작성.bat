@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
set URL=http://127.0.0.1:8789

rem If the server is already running, just open the browser.
netstat -ano | findstr "127.0.0.1:8789" | findstr "LISTENING" >nul
if not errorlevel 1 goto open_only

rem Otherwise open the browser after 2 seconds and run the server in this window.
start "" /b powershell -NoProfile -WindowStyle Hidden -Command "Start-Sleep -Seconds 2; Start-Process %URL%"
echo Closing this window stops the server.
python "%~dp0budget_app\server.py"
pause
goto :eof

:open_only
start "" "%URL%"
