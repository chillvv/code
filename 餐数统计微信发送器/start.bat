@echo off
cd /d "%~dp0"

python --version >nul 2>&1
if errorlevel 1 (
    py -3 --version >nul 2>&1
    if errorlevel 1 (
        echo Python not found! Please install Python 3.8+
        pause
        exit /b 1
    ) else (
        py -3 -m pip install -r requirements.txt >nul 2>&1
        py -3 main.py
    )
) else (
    python -m pip install -r requirements.txt >nul 2>&1
    python main.py
)

pause

