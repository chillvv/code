@echo off
setlocal EnableExtensions EnableDelayedExpansion

REM Use UTF-8 code page to reduce garbled output
chcp 65001 >nul

REM Resolve script directory
set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

REM Pick Python launcher
where py >nul 2>nul
if %errorlevel%==0 (
  set "PYLAUNCH=py -3"
) else (
  set "PYLAUNCH=python"
)

REM Create virtual environment
if not exist ".venv" (
  %PYLAUNCH% -m venv .venv
)

call ".venv\Scripts\activate.bat"

REM Upgrade pip toolchain
python -m pip install --upgrade pip setuptools wheel

REM Install dependencies with retries
set RETRIES=3
for /l %%i in (1,1,%RETRIES%) do (
  python -m pip install -r requirements.txt && goto deps_ok
  echo Install failed, retry %%i/%RETRIES%...
  timeout /t 2 >nul
)
echo Dependencies installation failed.
pause
goto :eof

:deps_ok
echo Dependencies installed.

REM Ensure wxauto and pywin32 present (redundant if requirements ok)
python -m pip install wxauto pywin32

set PYTHONUTF8=1
python -X utf8 -m wx_order_sender.main

echo.
pause

endlocal
