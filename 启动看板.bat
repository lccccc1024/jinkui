@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ================================
echo   金匮 · 个人财务看板
echo ================================
echo.
echo 正在启动服务...
start "" http://127.0.0.1:5000
python app.py
pause
