@echo off
chcp 65001 >nul
cd /d "%~dp0"
setlocal enabledelayedexpansion
title JinKui Dashboard

echo ================================
echo   JinKui - Personal Finance
echo ================================
echo.

:: Step 1 - Check Python
set PYTHON_CMD=python
where python >nul 2>&1
if errorlevel 1 (
    where py >nul 2>&1
    if errorlevel 1 (
        echo [!!] Python not found
        echo     Please install Python 3.9+ from:
        echo     https://www.python.org/downloads/
        echo.
        echo     Check "Add Python.exe to PATH" during setup.
        pause
        exit /b 1
    )
    set PYTHON_CMD=py
)

%PYTHON_CMD% -c "import sys; sys.exit(0 if sys.version_info >= (3,9) else 1)" 2>nul
if errorlevel 1 (
    echo [!!] Python version too low
    %PYTHON_CMD% --version
    echo     Need Python 3.9+
    pause
    exit /b 1
)

%PYTHON_CMD% --version 2>nul
echo [OK] Python ready

:: Step 2 - Install dependencies
%PYTHON_CMD% -c "import flask, openpyxl" 2>nul
if errorlevel 1 (
    echo.
    echo [..] Installing dependencies...
    %PYTHON_CMD% -m pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo [!!] Install failed
        echo     Try: %PYTHON_CMD% -m pip install -r requirements.txt
        pause
        exit /b 1
    )
    echo [OK] Dependencies installed
) else (
    echo [OK] Dependencies ready
)

:: Step 3 - Config / Excel path prompt
if not exist config.json (
    echo.
    echo ======================================
    echo   First run: point to your Excel file
    echo ======================================
    echo.
    echo  Drag your Excel file here, or type the full path:
    echo.
    set /p "EXCEL_PATH=>> "
    echo.

    if not "!EXCEL_PATH!"=="" (
        set "EXCEL_PATH=!EXCEL_PATH:"=!"
        set "EXCEL_PATH=!EXCEL_PATH:\=/!"

        if exist "!EXCEL_PATH!" (
            if exist config.example.json (
                copy /Y config.example.json config.json >nul
            )
            %PYTHON_CMD% -c "import json; cfg=json.load(open('config.json')); cfg['excel_path']='!EXCEL_PATH!'; json.dump(cfg, open('config.json','w',encoding='utf-8'), indent=4, ensure_ascii=False)"
            echo [OK] Config saved
            echo      Path: !EXCEL_PATH!
        ) else (
            echo [!!] File not found - check path and retry
            pause
            exit /b 1
        )
    ) else (
        if exist config.example.json (
            copy /Y config.example.json config.json >nul
        )
        echo.
        echo  [i] No path entered. Will try default:
        echo      project-dir/asset-liability.xlsx
        echo.
        echo  If your file is elsewhere, edit config.json
        echo  after startup, or set JINKUI_EXCEL_PATH.
    )
    echo.
)

:: Step 4 - Launch
echo.
echo ================================
echo   Starting server ...
echo   Open http://127.0.0.1:5000
echo ================================
echo.

start "" http://127.0.0.1:5000
%PYTHON_CMD% app.py

if errorlevel 1 (
    echo.
    echo [!!] Server exited with an error
    echo.
    echo  Common causes:
    echo   1. Excel path wrong - check config.json
    echo   2. Excel format - see README.md
    echo   3. Port 5000 in use by another app
    pause
)

endlocal
