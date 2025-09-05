$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Resolve directories
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$appDir = Join-Path $scriptDir 'electron-app'
if (-not (Test-Path (Join-Path $appDir 'package.json'))) {
  if (Test-Path (Join-Path $scriptDir 'package.json')) {
    $appDir = $scriptDir
  } else {
    Write-Host "错误：未找到 electron-app/package.json 或当前目录 package.json" -ForegroundColor Red
    Write-Host "当前目录：$scriptDir"
    Read-Host '按回车退出'
    exit 1
  }
}
Write-Host "使用应用目录：$appDir"

Push-Location $appDir

# Ensure npm mirrors
$npmrc = Join-Path $appDir '.npmrc'
if (-not (Test-Path $npmrc)) {
  @(
    'registry=https://registry.npmmirror.com'
    'strict-ssl=false'
    'fund=false'
    'audit=false'
    'disturl=https://npmmirror.com/mirrors/node'
    'electron_mirror=https://npmmirror.com/mirrors/electron/'
    'sass_binary_site=https://npmmirror.com/mirrors/node-sass'
    'sharp_dist_base_url=https://npmmirror.com/mirrors/sharp'
  ) | Set-Content -Encoding ASCII $npmrc
}

# Check Node
if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
  Write-Host '需要安装 Node.js 18+，请先安装后重试：https://nodejs.org/' -ForegroundColor Yellow
  Read-Host '按回车退出'
  exit 1
}

# Install deps
Write-Host '正在安装依赖，请稍候...'
$env:npm_config_registry = 'https://registry.npmmirror.com'
if (Test-Path (Join-Path $appDir 'package-lock.json')) {
  npm ci --no-fund --no-audit --registry https://registry.npmmirror.com
} else {
  npm install --no-fund --no-audit --registry https://registry.npmmirror.com
}

# Prepare logs
$logs = Join-Path $appDir 'logs'
if (-not (Test-Path $logs)) { New-Item -Force -ItemType Directory -Path $logs | Out-Null }

# Launch Electron
$electronBin = Join-Path $appDir 'node_modules/.bin/electron.cmd'
$mainFile = Join-Path $appDir 'src/main.js'
if (-not (Test-Path $mainFile)) {
  Write-Host "错误：未找到主入口：$mainFile" -ForegroundColor Red
  Get-ChildItem -Name (Join-Path $appDir 'src') -ErrorAction SilentlyContinue
  Read-Host '按回车退出'
  exit 1
}

$env:ELECTRON_ENABLE_LOGGING = '1'
$env:ELECTRON_DISABLE_SECURITY_WARNINGS = '1'

Write-Host '启动应用中...'
if (Test-Path $electronBin) {
  & $electronBin $mainFile --no-sandbox --disable-gpu --enable-logging 2>> (Join-Path $logs 'stderr.log') | Out-Null
} else {
  npx --yes --registry https://registry.npmmirror.com electron $mainFile --no-sandbox --disable-gpu --enable-logging 2>> (Join-Path $logs 'stderr.log') | Out-Null
}

Pop-Location
Write-Host ''
Read-Host '程序已退出，按回车关闭窗口'

