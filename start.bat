@echo off
chcp 65001 >nul
:: Windows 双击启动脚本
:: 功能：自动检查 Python → 安装依赖 → 运行报表生成

set "DIR=%~dp0"
cd /d "%DIR%"

echo ======================================
echo   富途资本利得报告生成器
echo ======================================
echo.

:: 1. 检查 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ⚠️  未检测到 Python
    echo.
    echo 请先安装 Python 3.10 或更高版本：
    echo.
    echo   方法 1：访问 https://www.python.org/downloads/
    echo     下载 Windows installer，双击安装
    echo     ⚠️ 安装时务必勾选 "Add Python to PATH"
    echo.
    echo   方法 2：Microsoft Store 搜索 "Python" 安装
    echo.
    echo 安装完成后，重新双击本文件。
    echo.
    pause
    exit /b 1
)

for /f "tokens=2" %%a in ('python --version 2^>^&1') do set PYTHON_VER=%%a
echo ✅ Python 版本: %PYTHON_VER%

:: 2. 检查并安装 openpyxl
echo.
echo 正在检查依赖...
python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo 📦 正在安装 openpyxl（约需 10-30 秒）...
    python -m pip install openpyxl --user
    if errorlevel 1 (
        echo.
        echo ⚠️  依赖安装失败，请尝试手动运行：
        echo     python -m pip install openpyxl --user
        echo.
        pause
        exit /b 1
    )
)
echo ✅ 依赖已就绪

:: 3. 检查账单文件
echo.
set BILL_COUNT=0
for %%f in (*_年度账单_*.xlsx) do set /a BILL_COUNT+=1

if %BILL_COUNT%==0 (
    echo ⚠️  当前目录没有找到年度账单 xlsx 文件
    echo.
    echo 请将富途导出的账单文件（如 2025_年度账单_12345678.xlsx）
    echo 放到本文件夹中，然后重新运行。
    echo.
    pause
    exit /b 1
)
echo ✅ 发现 %BILL_COUNT% 个年度账单文件

:: 4. 运行脚本
echo.
echo ======================================
echo   正在生成报告...
echo ======================================
echo.

python calc_capital_gains_v2.py "%DIR%"

if errorlevel 1 (
    echo.
    echo ======================================
    echo   ❌ 生成失败，请查看上方错误信息
    echo ======================================
    echo.
    pause
    exit /b 1
)

echo.
echo ======================================
echo   ✅ 报告生成成功！
echo ======================================
echo.
echo 生成的文件：
echo   📄 %DIR%capital_gains_report_v2.html
echo   📊 %DIR%capital_gains_report_v2.json
echo.
echo 用浏览器打开 .html 文件即可查看报告。
echo.
pause
