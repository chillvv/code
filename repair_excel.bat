@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

REM Reuse main venv under user profile
set "DEFAULT_VENV=%USERPROFILE%\.wxorder_venv"
if not defined USERPROFILE set "DEFAULT_VENV=C:\.wxorder_venv"
set "VENV_DIR=%DEFAULT_VENV%"

if not exist "%VENV_DIR%" (
  py -3 -m venv "%VENV_DIR%"
)
call "%VENV_DIR%\Scripts\activate.bat"

set "ALIYUN_INDEX=https://mirrors.aliyun.com/pypi/simple/"
set "ALIYUN_HOST=mirrors.aliyun.com"
python -m pip install --upgrade pip setuptools wheel -i %ALIYUN_INDEX% --trusted-host %ALIYUN_HOST%
python -m pip install openpyxl xlrd -i %ALIYUN_INDEX% --trusted-host %ALIYUN_HOST%

if "%~1"=="" (
  echo Usage: repair_excel.bat ^<input-path^> [output-path]
  pause
  goto :eof
)

set "IN=%~1"
set "OUT=%~2"
if "%OUT%"=="" (
  python -X utf8 -m tools.repair_excel "%IN%"
) else (
  python -X utf8 -m tools.repair_excel "%IN%" "%OUT%"
)

echo.
pause

endlocal
