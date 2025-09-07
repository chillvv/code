@echo off
setlocal EnableExtensions EnableDelayedExpansion
title 微信订单发送器 - 改进版

REM 使用UTF-8编码减少乱码
chcp 65001 >nul

echo ========================================
echo    微信订单发送器 - 改进版
echo ========================================
echo.

REM 获取脚本所在目录
set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

REM 检测Python环境
set "PYTHON_FOUND=0"
set "PYLAUNCH="

python --version >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_FOUND=1"
    set "PYLAUNCH=python"
    for /f "tokens=*" %%i in ('python --version 2^>^&1') do set "PY_VERSION=%%i"
    echo Python检测成功: !PY_VERSION!
    goto python_ok
)

where py >nul 2>nul
if %errorlevel%==0 (
    py -3 --version >nul 2>nul
    if !errorlevel!==0 (
        set "PYTHON_FOUND=1"
        set "PYLAUNCH=py -3"
        for /f "tokens=*" %%i in ('py -3 --version 2^>^&1') do set "PY_VERSION=%%i"
        echo Python启动器检测成功: !PY_VERSION!
        goto python_ok
    )
)

if %PYTHON_FOUND%==0 (
    echo 错误：未找到Python环境！
    echo 请先安装Python 3.8+
    pause
    exit /b 1
)

:python_ok

echo.
echo 🔧 检查依赖包...
"%PYLAUNCH%" -c "import pandas, PyQt5, uiautomation, pyperclip, pyautogui; print('✅ 依赖包检查通过')" 2>nul
if %errorlevel% neq 0 (
    echo 正在安装缺失的依赖包...
    "%PYLAUNCH%" -m pip install PyQt5 pandas openpyxl uiautomation pyperclip pyautogui pywin32 --quiet --disable-pip-version-check
)

echo.
echo 🚀 启动改进版订单发送器...
echo.
echo 💡 改进功能:
echo   ✅ 更准确的微信窗口识别
echo   ✅ 智能群聊搜索和验证
echo   ✅ 双重发送方式（UI自动化 + 热键）
echo   ✅ 详细的发送进度提示
echo   ✅ 测试模式安全保护
echo.
echo ⚠️ 使用前请确保：
echo   1. 微信PC版已登录
echo   2. 群聊名称准确无误
echo   3. 建议先使用测试模式
echo.

REM 设置环境变量
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

"%PYLAUNCH%" -X utf8 "终极微信发送器.py"

echo.
if %errorlevel%==0 (
    echo 程序正常退出
) else (
    echo 程序异常退出，错误代码: %errorlevel%
)
pause
endlocal
