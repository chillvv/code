@echo off
chcp 65001 >nul
title Meal Counter WeChat Sender

echo ================================================================
echo                   Meal Counter WeChat Sender
echo ================================================================
echo.

:: Set encoding environment variables
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
set LANG=zh_CN.UTF-8
set LC_ALL=zh_CN.UTF-8

:: Check Python installation
echo Checking Python environment...
python --version >nul 2>&1
if errorlevel 1 (
    echo Trying py command...
    py -3 --version >nul 2>&1
    if errorlevel 1 (
        echo Error: Python not found!
        echo Please install Python 3.8+
        echo Download: https://www.python.org/downloads/
        pause
        exit /b 1
    ) else (
        set PYTHON_CMD=py -3
        for /f "tokens=2" %%i in ('py -3 --version 2^>^&1') do set PYTHON_VERSION=%%i
    )
) else (
    set PYTHON_CMD=python
    for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
)

echo Python version: %PYTHON_VERSION%
echo.

:: Check dependencies
echo Checking dependencies...
%PYTHON_CMD% -c "import PyQt5, pandas, openpyxl, uiautomation, pyautogui, pyperclip, pywin32" >nul 2>&1
if errorlevel 1 (
    echo Missing dependencies detected, installing...
    echo.
    %PYTHON_CMD% -m pip install --upgrade pip
    %PYTHON_CMD% -m pip install -r requirements.txt
    echo.
    echo Dependencies installed!
) else (
    echo Dependencies check passed!
)
echo.

:: Start program
echo Starting Meal Counter WeChat Sender...
echo.
%PYTHON_CMD% main.py

if errorlevel 1 (
    echo.
    echo Program failed to start!
    pause
    exit /b 1
)

echo.
echo Program closed
pause