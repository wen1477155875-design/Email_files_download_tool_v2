@echo off
setlocal
cd /d "%~dp0"
if exist ".venv\Scripts\pythonw.exe" goto use_venv
if exist "python_env.txt" goto use_env
where pythonw >nul 2>nul
if not errorlevel 1 goto use_path
echo [提示] 未找到 Python 环境，请先运行 安装依赖.bat
pause
exit /b 1

:use_venv
start "" ".venv\Scripts\pythonw.exe" ui.py
exit /b 0

:use_env
set /p PYW=<python_env.txt
rem 去掉可能存在的引号，统一按"是否为盘符路径"决定怎么启动
set "PYW=%PYW:"=%"
echo %PYW%| findstr /i /r "^[A-Za-z]:" >nul
if not errorlevel 1 goto env_quoted
start "" %PYW% ui.py
exit /b 0

:env_quoted
start "" "%PYW%" ui.py
exit /b 0

:use_path
start "" pythonw ui.py
exit /b 0
