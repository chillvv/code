@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

set "DEFAULT_VENV=%USERPROFILE%\.wxorder_venv"
if not defined USERPROFILE set "DEFAULT_VENV=C:\.wxorder_venv"
set "VENV_DIR=%DEFAULT_VENV%"

where py >nul 2>nul
if %errorlevel%==0 ( set "PYLAUNCH=py -3" ) else ( set "PYLAUNCH=python" )

if not exist "%VENV_DIR%" (
  %PYLAUNCH% -m venv "%VENV_DIR%"
)
set "PYEXE=%VENV_DIR%\Scripts\python.exe"
if not exist "%PYEXE%" (
  echo 虚拟环境创建失败。
  pause
  exit /b 1
)

set "ALIYUN_INDEX=https://mirrors.aliyun.com/pypi/simple/"
set "ALIYUN_HOST=mirrors.aliyun.com"
"%PYEXE%" -m pip config set global.index-url %ALIYUN_INDEX% >nul 2>nul
"%PYEXE%" -m pip config set install.trusted-host %ALIYUN_HOST% >nul 2>nul
echo 升级 pip...
"%PYEXE%" -m pip install --upgrade pip setuptools wheel -i %ALIYUN_INDEX% --trusted-host %ALIYUN_HOST%
echo 安装依赖...
"%PYEXE%" -m pip install -r requirements.txt -i %ALIYUN_INDEX% --trusted-host %ALIYUN_HOST%
if not %errorlevel%==0 (
  echo 依赖安装失败。
  pause
  exit /b 1
)

set PYTHONUTF8=1
echo 启动应用...
"%PYEXE%" -X utf8 -m py_wechat_sender.main
set EXIT=%errorlevel%
echo 应用退出，代码：%EXIT%
pause
endlocal
