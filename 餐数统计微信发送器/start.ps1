# 餐数统计微信发送器启动脚本
# PowerShell版本

Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "                   餐数统计微信发送器" -ForegroundColor Yellow
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

# 设置编码
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"

# 检查Python
Write-Host "检查Python环境..." -ForegroundColor Green

$pythonCmd = $null
try {
    $pythonVersion = python --version 2>&1
    if ($LASTEXITCODE -eq 0) {
        $pythonCmd = "python"
        Write-Host "找到Python: $pythonVersion" -ForegroundColor Green
    }
} catch {
    try {
        $pythonVersion = py -3 --version 2>&1
        if ($LASTEXITCODE -eq 0) {
            $pythonCmd = "py -3"
            Write-Host "找到Python: $pythonVersion" -ForegroundColor Green
        }
    } catch {
        Write-Host "错误：未找到Python！" -ForegroundColor Red
        Write-Host "请先安装Python 3.8+" -ForegroundColor Yellow
        Write-Host "下载地址：https://www.python.org/downloads/" -ForegroundColor Cyan
        Read-Host "按回车键退出"
        exit 1
    }
}

if (-not $pythonCmd) {
    Write-Host "错误：Python未正确安装！" -ForegroundColor Red
    Read-Host "按回车键退出"
    exit 1
}

Write-Host ""

# 检查依赖
Write-Host "检查依赖包..." -ForegroundColor Green
$checkDeps = Invoke-Expression "$pythonCmd -c `"import PyQt5, pandas, openpyxl, uiautomation, pyautogui, pyperclip, pywin32`" 2>&1"

if ($LASTEXITCODE -ne 0) {
    Write-Host "检测到缺少依赖包，正在安装..." -ForegroundColor Yellow
    Write-Host ""
    
    Invoke-Expression "$pythonCmd -m pip install --upgrade pip"
    Invoke-Expression "$pythonCmd -m pip install -r requirements.txt"
    
    Write-Host ""
    Write-Host "依赖安装完成！" -ForegroundColor Green
} else {
    Write-Host "依赖检查通过！" -ForegroundColor Green
}

Write-Host ""

# 启动程序
Write-Host "启动餐数统计微信发送器..." -ForegroundColor Green
Write-Host ""

try {
    Invoke-Expression "$pythonCmd main.py"
} catch {
    Write-Host ""
    Write-Host "程序启动失败！" -ForegroundColor Red
    Write-Host "错误信息：$_" -ForegroundColor Red
}

Write-Host ""
Write-Host "程序已关闭" -ForegroundColor Yellow
Read-Host "按回车键退出"

