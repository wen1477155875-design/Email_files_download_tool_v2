@echo off
setlocal
cd /d "%~dp0"
echo ============================================================
echo  邮件附件下载工具 - 依赖安装（在线多镜像 + 离线兜底）
echo ============================================================
echo.

if exist ".venv" rmdir /s /q ".venv"

set "PYCMD="

call :try_interpreter py -3
if defined PYCMD goto got_py
call :try_interpreter python
if defined PYCMD goto got_py

call :try_path "C://Program Files//Python313//python.exe"
call :try_path "C://Program Files//Python312//python.exe"
call :try_path "C://Program Files//Python311//python.exe"
call :try_path "C://Program Files//Python310//python.exe"
call :try_path "C://Program Files//Python314//python.exe"
call :try_path "%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
call :try_path "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
call :try_path "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
call :try_path "%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
call :try_path "C://Python313//python.exe"
call :try_path "C://Python312//python.exe"
call :try_path "C://Python311//python.exe"
call :try_path "C://Python310//python.exe"
if defined PYCMD goto got_py

echo [失败] 没有找到"能实际运行"的 Python。
echo 可能原因：电脑没装 Python，或被公司策略拦截（AppLocker）。
echo 可在 cmd 里输入 python --version 试试，把结果发给管理员排查。
pause
exit /b 1

:got_py
echo [1/5] 找到可用解释器：%PYCMD%
%PYCMD% -c "import sys;print(sys.version)"

rem 记录 pythonw 路径（带引号），给 启动界面.bat 用
%PYCMD% -c "import sys,os;d=os.path.dirname(sys.executable);w=os.path.join(d,'pythonw.exe');q=chr(34);p=w if os.path.exists(w) else sys.executable;print(q+p+q)" > "python_env.txt"
if errorlevel 1 goto pyw_fallback
for %%S in ("python_env.txt") do if %%~zS==0 goto pyw_fallback
goto write_ok
:pyw_fallback
>"python_env.txt" echo %PYCMD%
:write_ok
echo.

echo [2/5] 安装依赖 pywin32 / PyYAML 到用户目录（--user，无需管理员）...
echo   尝试 1/4：pip 当前默认源（可能为公司内网源）
%PYCMD% -m pip install --user -r requirements.txt --timeout 15 --retries 1 --disable-pip-version-check
if not errorlevel 1 goto pip_ok

set "PIPDONE="
for %%U in ("https://pypi.tuna.tsinghua.edu.cn/simple" "https://mirrors.aliyun.com/pypi/simple" "https://mirrors.cloud.tencent.com/pypi/simple") do (
  if not defined PIPDONE (
    echo   尝试镜像源：%%~U
    %PYCMD% -m pip install --user -r requirements.txt -i %%~U --timeout 15 --retries 1 --disable-pip-version-check
    if not errorlevel 1 set "PIPDONE=1"
  )
)
if defined PIPDONE goto pip_ok

echo   在线安装全部失败，补装 pip 后再试一轮...
%PYCMD% -m ensurepip --upgrade >nul 2>nul
set "PIPDONE="
for %%U in ("https://pypi.tuna.tsinghua.edu.cn/simple" "https://mirrors.aliyun.com/pypi/simple") do (
  if not defined PIPDONE (
    echo   尝试镜像源：%%~U
    %PYCMD% -m pip install --user -r requirements.txt -i %%~U --timeout 15 --retries 1 --disable-pip-version-check
    if not errorlevel 1 set "PIPDONE=1"
  )
)
if defined PIPDONE goto pip_ok

if not exist "wheels" goto fail_pip
echo [3/5] 在线安装失败，改用包内自带的离线安装包（wheels 目录）...
%PYCMD% -m pip install --user --no-index --find-links wheels -r requirements.txt --disable-pip-version-check
if errorlevel 1 goto fail_pip

:pip_ok
echo.
echo [4/5] 依赖安装完成。运行离线自检（不连真实 Outlook）...
chcp 65001 >nul
%PYCMD% smoke_test.py
set "TEST_RC=%errorlevel%"
chcp 936 >nul
if not "%TEST_RC%"=="0" goto fail_test

echo.
echo ============================================================
echo  [5/5] 部署完成！以后双击 启动界面.bat 打开图形界面
echo  如需卸载依赖：%PYCMD% -m pip uninstall -y pywin32 pyyaml
echo ============================================================
pause
exit /b 0

:fail_pip
echo [错误] 依赖安装失败。
echo 已尝试：默认源 / 清华 / 阿里 / 腾讯镜像 / 离线 wheels。
echo 请把上面的完整报错截图发回来排查（注意看 [1/5] 显示的 Python 版本）。
pause
exit /b 1

:fail_test
echo [错误] 自检未通过，请把上面的输出发回来排查
pause
exit /b 1

:try_interpreter
if defined PYCMD goto :eof
%* -c "import sys" >nul 2>nul
if not errorlevel 1 set "PYCMD=%*"
goto :eof

:try_path
if defined PYCMD goto :eof
if not exist %1 goto :eof
%1 -c "import sys" >nul 2>nul
if not errorlevel 1 set "PYCMD=%~1"
goto :eof
