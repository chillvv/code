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

REM Use short venv path to avoid Windows long path issues
set "DEFAULT_VENV=%USERPROFILE%\.wxorder_venv"
if not defined USERPROFILE set "DEFAULT_VENV=C:\\.wxorder_venv"
set "VENV_DIR=%DEFAULT_VENV%"

if not exist "%VENV_DIR%" (
  %PYLAUNCH% -m venv "%VENV_DIR%"
)

call "%VENV_DIR%\Scripts\activate.bat"

REM Upgrade pip toolchain
set "ALIYUN_INDEX=https://mirrors.aliyun.com/pypi/simple/"
set "ALIYUN_HOST=mirrors.aliyun.com"
set "PIP_INDEX_URL=%ALIYUN_INDEX%"
set "PIP_TRUSTED_HOST=%ALIYUN_HOST%"
set "PIP_DISABLE_PIP_VERSION_CHECK=1"

python -m pip config set global.index-url %ALIYUN_INDEX% >nul 2>nul
python -m pip config set install.trusted-host %ALIYUN_HOST% >nul 2>nul

python -m pip install --upgrade pip setuptools wheel -i %ALIYUN_INDEX% --trusted-host %ALIYUN_HOST%

REM Install dependencies with retries
set RETRIES=3
for /l %%i in (1,1,%RETRIES%) do (
  python -m pip install -r requirements.txt -i %ALIYUN_INDEX% --trusted-host %ALIYUN_HOST% --timeout 120 && goto deps_ok
  echo Install failed, retry %%i/%RETRIES%...
  timeout /t 2 >nul
)
echo Dependencies installation failed.
pause
goto :eof

:deps_ok
echo Dependencies installed.

REM Prepare Node.js bot (Wechaty)
where node >nul 2>nul
set NODE_OK=%errorlevel%

if %NODE_OK%==0 (
  pushd node_bot
  if not exist ".npmrc" (
    echo registry=https://registry.npmmirror.com> .npmrc
  )
  set "npm_config_registry=https://registry.npmmirror.com"
  if exist package-lock.json (
    npm ci --no-fund --no-audit --registry https://registry.npmmirror.com
  ) else (
    npm install --no-fund --no-audit --registry https://registry.npmmirror.com
  )
  if not %errorlevel%==0 (
    echo NPM install failed. Bot may not start.
    pause
  )
  set WX_BOT_PORT=8788
  start "wechaty-bot" /min node server.mjs
  popd
  rem Wait for bot to be ready (up to ~20s)
  for /l %%s in (1,1,10) do (
    powershell -Command "try { iwr -UseBasicParsing http://127.0.0.1:8788/status -TimeoutSec 2 ^| Out-Null; exit 0 } catch { exit 1 }" >nul 2>nul && goto bot_ready
    timeout /t 2 >nul
  )
  :bot_ready
) else (
  echo Node.js not found. UI will start without bot. Please install Node 18+.
)

set PYTHONUTF8=1
start "wx-order-ui" python -X utf8 -m wx_order_sender.main

echo.
echo Setup finished. If UI did not appear, check errors above.
pause

endlocal
