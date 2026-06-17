#!/usr/bin/env bash
# 金匮 · 个人财务看板 — 启动脚本 (macOS / Linux)
set -e
cd "$(dirname "$0")"

echo "================================"
echo "  金匮 · 个人财务看板"
echo "================================"
echo ""

# 检查 Python
if ! command -v python3 &>/dev/null && ! command -v python &>/dev/null; then
    echo "❌ 未找到 Python，请先安装 Python 3.9+"
    exit 1
fi
PYTHON=$(command -v python3 || command -v python)

# 检查依赖
if ! $PYTHON -c "import flask" 2>/dev/null; then
    echo "📦 正在安装依赖..."
    $PYTHON -m pip install -r requirements.txt
fi

echo "🚀 正在启动服务..."
echo "   浏览器访问 http://127.0.0.1:5000"
echo ""

# 打开浏览器
if command -v open &>/dev/null; then
    open http://127.0.0.1:5000
elif command -v xdg-open &>/dev/null; then
    xdg-open http://127.0.0.1:5000
fi

$PYTHON app.py
