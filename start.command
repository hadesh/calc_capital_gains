#!/bin/bash
# macOS 双击启动脚本
# 功能：自动检查 Python → 安装依赖 → 运行报表生成

# 获取脚本所在目录
DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$DIR"

echo "======================================"
echo "  富途资本利得报告生成器"
echo "======================================"
echo ""

# 1. 检查 Python3
if ! command -v python3 &> /dev/null; then
    echo "⚠️  未检测到 Python3"
    echo ""
    echo "请先安装 Python 3.10 或更高版本："
    echo ""
    echo "  方法 1（推荐）：打开终端，运行："
    echo "    brew install python"
    echo ""
    echo "  方法 2：访问 https://www.python.org/downloads/macos/"
    echo "    下载安装包，双击安装"
    echo ""
    echo "安装完成后，重新双击本文件。"
    echo ""
    read -n 1 -s -r -p "按任意键关闭..."
    exit 1
fi

PYTHON_VER=$(python3 --version 2>&1 | awk '{print $2}')
echo "✅ Python 版本: $PYTHON_VER"

# 2. 检查并安装 openpyxl
echo ""
echo "正在检查依赖..."
if ! python3 -c "import openpyxl" 2>/dev/null; then
    echo "📦 正在安装 openpyxl（约需 10-30 秒）..."
    python3 -m pip install openpyxl --user
    if [ $? -ne 0 ]; then
        echo ""
        echo "⚠️  依赖安装失败，请尝试手动运行："
        echo "    python3 -m pip install openpyxl --user"
        echo ""
        read -n 1 -s -r -p "按任意键关闭..."
        exit 1
    fi
fi
echo "✅ 依赖已就绪"

# 3. 检查账单文件
echo ""
BILL_COUNT=$(ls -1 *_年度账单_*.xlsx 2>/dev/null | wc -l)
if [ "$BILL_COUNT" -eq 0 ]; then
    echo "⚠️  当前目录没有找到年度账单 xlsx 文件"
    echo ""
    echo "请将富途导出的账单文件（如 2025_年度账单_12345678.xlsx）"
    echo "放到本文件夹中，然后重新运行。"
    echo ""
    read -n 1 -s -r -p "按任意键关闭..."
    exit 1
fi
echo "✅ 发现 $BILL_COUNT 个年度账单文件"

# 4. 运行脚本
echo ""
echo "======================================"
echo "  正在生成报告..."
echo "======================================"
echo ""

python3 calc_capital_gains_v2.py "$DIR"

EXIT_CODE=$?
echo ""

if [ $EXIT_CODE -eq 0 ]; then
    echo "======================================"
    echo "  ✅ 报告生成成功！"
    echo "======================================"
    echo ""
    echo "生成的文件："
    echo "  📄 $DIR/capital_gains_report_v2.html"
    echo "  📊 $DIR/capital_gains_report_v2.json"
    echo ""
    echo "用浏览器打开 .html 文件即可查看报告。"
else
    echo "======================================"
    echo "  ❌ 生成失败，请查看上方错误信息"
    echo "======================================"
fi

echo ""
read -n 1 -s -r -p "按任意键关闭..."
