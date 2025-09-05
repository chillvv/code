@echo off
setlocal EnableExtensions EnableDelayedExpansion

REM Use UTF-8 to reduce garbled output
chcp 65001 >nul

REM Resolve project directories
set "SCRIPT_DIR=%~dp0"
pushd "%SCRIPT_DIR%" >nul
set "APP_DIR=%SCRIPT_DIR%electron-app"

REM If running from inside electron-app, fallback to current folder
if not exist "%APP_DIR%\package.json" (
  if exist "%SCRIPT_DIR%package.json" (
    set "APP_DIR=%SCRIPT_DIR%"
  ) else (
    echo 错误：未找到 electron-app\package.json
    echo 请确认已解压完整项目，并从项目根目录运行此脚本。
    echo 当前目录：%SCRIPT_DIR%
    pause
    exit /b 1
  )
)

pushd "%APP_DIR%" >nul
echo 使用应用目录：%APP_DIR%
if not exist "%APP_DIR%\package.json" (
  echo 警告：未发现 package.json，将尝试直接用主文件启动。
)
if not exist "%APP_DIR%\src\main.js" (
  echo 错误：未找到主入口：%APP_DIR%\src\main.js
  dir "%APP_DIR%\src" /b
  pause
  exit /b 1
)

REM Force China mirrors for npm and binaries (Electron, sharp, node-sass, etc.)
if not exist ".npmrc" (
  > .npmrc echo registry=https://registry.npmmirror.com
  >> .npmrc echo strict-ssl=false
  >> .npmrc echo fund=false
  >> .npmrc echo audit=false
  >> .npmrc echo disturl=https://npmmirror.com/mirrors/node
  >> .npmrc echo electron_mirror=https://npmmirror.com/mirrors/electron/
  >> .npmrc echo sass_binary_site=https://npmmirror.com/mirrors/node-sass
  >> .npmrc echo sharp_dist_base_url=https://npmmirror.com/mirrors/sharp
)

REM Check Node.js
where node >nul 2>nul
if %errorlevel% neq 0 (
  echo 需要安装 Node.js 18+，请先安装后重试：https://nodejs.org/
  pause
  exit /b 1
)

REM Install dependencies (will use .npmrc registry)
echo 正在安装依赖，请稍候...
set "npm_config_registry=https://registry.npmmirror.com"
if exist package-lock.json (
  npm ci --no-fund --no-audit --registry https://registry.npmmirror.com
) else (
  npm install --no-fund --no-audit --registry https://registry.npmmirror.com
)
if %errorlevel% neq 0 (
  echo 依赖安装失败，请检查网络或代理设置。
  pause
  exit /b 1
)

REM Prepare logs folder
if not exist "%APP_DIR%\logs" (
  mkdir "%APP_DIR%\logs" >nul 2>nul
)

REM Launch Electron (quote paths with spaces/parentheses)
set "ELECTRON_BIN=%APP_DIR%\node_modules\.bin\electron.cmd"
set "ELECTRON_ENABLE_LOGGING=1"
set "ELECTRON_DEBUG_NOTIFICATIONS=1"
set "ELECTRON_LOG_FILE=%APP_DIR%\logs\electron.log"
set "ELECTRON_DISABLE_SECURITY_WARNINGS=1"
set "MAIN_FILE=%APP_DIR%\src\main.js"

if exist "%ELECTRON_BIN%" (
  echo 启动应用中...
  call "%ELECTRON_BIN%" "%MAIN_FILE%" --no-sandbox --disable-gpu --enable-logging 1>>"%APP_DIR%\logs\stdout.log" 2>>"%APP_DIR%\logs\stderr.log"
) else (
  echo 本地 electron 未找到，尝试使用 npx 启动...
  npx --yes --registry https://registry.npmmirror.com electron "%MAIN_FILE%" --no-sandbox --disable-gpu --enable-logging 1>>"%APP_DIR%\logs\stdout.log" 2>>"%APP_DIR%\logs\stderr.log"
)

set EXITCODE=%errorlevel%
if %EXITCODE% neq 0 (
  echo Electron 退出，错误码：%EXITCODE%
) else (
  echo Electron 已退出。
)

popd >nul
popd >nul
echo.
pause

endlocal
