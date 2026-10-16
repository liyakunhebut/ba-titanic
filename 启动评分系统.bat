@echo off
title 在线评分系统 - 商务大数据分析课程竞赛
cd /d "%~dp0"
echo ============================================
echo   商务大数据分析课程竞赛 - 在线评分系统
echo ============================================
echo 正在启动服务，请勿关闭本窗口（停止服务请按 Ctrl+C 或直接关闭窗口）
echo.
where python >nul 2>nul
if not %errorlevel%==0 goto fallback
python app.py
goto done
:fallback
if exist "D:\anaconda3\python.exe" goto anaconda
echo [错误] 未找到 Python，请先安装 Python 或 Anaconda。
pause
exit /b 1
:anaconda
"D:\anaconda3\python.exe" app.py
:done
echo.
echo 服务已停止。
pause
