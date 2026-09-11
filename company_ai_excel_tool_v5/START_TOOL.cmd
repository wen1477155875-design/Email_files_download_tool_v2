@echo off
setlocal
cd /d "%~dp0"
title Company AI Excel Batch Tool

rem =====================================================================
rem  Step 1: locate a usable Python 3 interpreter.
rem =====================================================================
set "PYCMD="

py -3 -c "import sys" >nul 2>nul && set "PYCMD=py -3"
if not defined PYCMD (
    python -c "import sys" >nul 2>nul && set "PYCMD=python"
)

if not defined PYCMD (
    echo.
    echo [ERROR] Python 3 was not found.
    echo         Please install Python 3 and make sure "py" or "python"
    echo         works in a Command Prompt.
    echo.
    pause
    exit /b 1
)

rem =====================================================================
rem  Step 2: make sure the required packages are available.
rem =====================================================================
%PYCMD% -c "import openpyxl, requests" >nul 2>nul
if errorlevel 1 (
    echo.
    echo [INFO] Required packages are missing.
    echo [INFO] Installing openpyxl and requests, please wait...
    echo.
    %PYCMD% -m pip install --disable-pip-version-check openpyxl requests
    %PYCMD% -c "import openpyxl, requests" >nul 2>nul
    if errorlevel 1 (
        echo.
        echo [ERROR] Could not install the required packages.
        echo         Please run this command manually, then retry:
        echo             %PYCMD% -m pip install openpyxl requests
        echo.
        pause
        exit /b 1
    )
)

rem =====================================================================
rem  Step 3: launch the tool. Keep the window open if it exits with error.
rem =====================================================================
echo [INFO] Starting with: %PYCMD%
echo.
%PYCMD% company_ai_excel_tool.py
set "RC=%errorlevel%"

if not "%RC%"=="0" (
    echo.
    echo [ERROR] The tool exited with code %RC%.
    echo         Please read the messages above for the reason.
    echo.
    pause
    exit /b %RC%
)

endlocal
