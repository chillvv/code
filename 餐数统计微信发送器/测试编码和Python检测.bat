@echo off
setlocal EnableExtensions EnableDelayedExpansion
title 测试Python检测和中文编码

REM 使用UTF-8编码
chcp 65001 >nul

echo ========================================
echo    测试Python检测和中文编码功能
echo ========================================
echo.

REM 获取脚本所在目录
set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

echo 1. 测试中文显示...
echo 中文字符测试：你好世界！简知轻食餐数统计发送器
echo 特殊字符测试：📊🍽️💰😊✅❌⚠️
echo.

echo 2. 测试Python检测逻辑...
REM 检测Python启动器
set "PYTHON_FOUND=0"
set "PYLAUNCH="
set "PY_VERSION="

REM 首先尝试python命令
python --version >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_FOUND=1"
    set "PYLAUNCH=python"
    for /f "tokens=*" %%i in ('python --version 2^>^&1') do set "PY_VERSION=%%i"
    echo ✅ Python检测成功: !PY_VERSION!
    goto python_ok
)

REM 尝试py启动器
where py >nul 2>nul
if %errorlevel%==0 (
    py -3 --version >nul 2>nul
    if !errorlevel!==0 (
        set "PYTHON_FOUND=1"
        set "PYLAUNCH=py -3"
        for /f "tokens=*" %%i in ('py -3 --version 2^>^&1') do set "PY_VERSION=%%i"
        echo ✅ Python启动器检测成功: !PY_VERSION!
        goto python_ok
    )
)

REM 如果都没找到
if %PYTHON_FOUND%==0 (
    echo ❌ 错误：未找到可用的Python环境！
    pause
    exit /b 1
)

:python_ok

echo.
echo 3. 测试Python中文编码...
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

"%PYLAUNCH%" -X utf8 -c "print('Python中文测试：你好世界！')"
"%PYLAUNCH%" -X utf8 -c "print('默认编码:', __import__('sys').getdefaultencoding())"
"%PYLAUNCH%" -X utf8 -c "print('文件系统编码:', __import__('sys').getfilesystemencoding())"

echo.
echo 4. 测试Excel文件读取...
if exist "简知轻食.xlsx" (
    "%PYLAUNCH%" -X utf8 -c "import pandas as pd; df = pd.read_excel('简知轻食.xlsx', sheet_name='会员9月扣餐表'); print('✅ Excel文件读取成功，数据行数:', len(df)); print('✅ 包含会员:', df['会员姓名'].dropna().iloc[0] if len(df) > 0 else '无数据')"
) else (
    echo ⚠️ 未找到简知轻食.xlsx测试文件
)

echo.
echo ========================================
echo 测试完成！
echo ========================================
echo.
echo 如果以上测试都显示✅，说明环境配置正确
echo 如果有问题或乱码，请检查：
echo 1. Windows系统语言设置
echo 2. 控制台字体设置（建议使用Consolas或微软雅黑）
echo 3. Python安装是否完整
echo.
pause
