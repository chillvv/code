@echo off
setlocal EnableExtensions EnableDelayedExpansion

REM Force UTF-8 to avoid Chinese garbled text
chcp 65001 >nul

REM Resolve script directory
set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"

REM Create virtual environment
if not exist .venv (
  py -3 -m venv .venv
)

call .venv\Scripts\activate.bat

REM Upgrade pip and wheel
python -m pip install --upgrade pip wheel setuptools

REM Install dependencies with retries
set RETRIES=3
for /l %%i in (1,1,%RETRIES%) do (
  python -m pip install -r requirements.txt && goto :deps_ok
  echo 安装依赖失败，正在重试 (%%i/%RETRIES%)...
  timeout /t 2 >nul
)
echo 依赖安装失败，请检查网络或代理设置。
pause
goto :eof

:deps_ok
echo 依赖安装成功。

REM Launch app
set PYTHONUTF8=1
python -X utf8 -m wx_order_sender.main

REM Keep window open after exit
echo .
pause

endlocal
