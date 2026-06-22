#!/usr/bin/env bash
# 金匮 · 个人财务看板 — 启动脚本 (macOS / Linux)
set -u
cd "$(dirname "$0")" || exit 1

echo "================================"
echo "  金匮 · 个人财务看板"
echo "================================"
echo ""

# ═══════════════════════════════════
#  Step 1 — 检查 Python 环境
# ═══════════════════════════════════

PYTHON=""
if command -v python3 &>/dev/null; then
    PYTHON=$(command -v python3)
elif command -v python &>/dev/null; then
    PYTHON=$(command -v python)
else
    echo "[!!] 未检测到 Python"
    echo "     本应用需要 Python 3.9 或更高版本"
    echo "     安装方式："
    echo "       macOS:  brew install python"
    echo "       Ubuntu: sudo apt install python3"
    echo "       Fedora: sudo dnf install python3"
    exit 1
fi

$PYTHON -c "import sys; sys.exit(0 if sys.version_info >= (3,9) else 1)" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "[!!] Python 版本过低"
    $PYTHON --version
    echo "     本应用需要 Python 3.9 或更高版本"
    exit 1
fi

echo "[OK] $($PYTHON --version 2>/dev/null)"

# ═══════════════════════════════════
#  Step 2 — 检查 / 安装依赖
# ═══════════════════════════════════

$PYTHON -c "import flask, openpyxl" 2>/dev/null
if [ $? -ne 0 ]; then
    echo ""
    echo "[..] 正在安装依赖（首次运行需要下载，请稍候）..."
    $PYTHON -m pip install -r requirements.txt
    if [ $? -ne 0 ]; then
        echo ""
        echo "[!!] 依赖安装失败"
        echo "     请尝试手动执行："
        echo "     $PYTHON -m pip install -r requirements.txt"
        exit 1
    fi
    echo "[OK] 依赖安装完成"
else
    echo "[OK] 依赖已就绪"
fi

# ═══════════════════════════════════
#  Step 3 — 引导 Excel 路径配置
# ═══════════════════════════════════

if [ ! -f config.json ]; then
    echo ""
    echo "======================================"
    echo "  首次使用：指定 Excel 数据文件路径"
    echo "======================================"
    echo ""
    echo "  输入你的资产负债 Excel 文件完整路径："
    echo ""
    printf ">> "
    read EXCEL_PATH
    echo ""

    if [ -n "$EXCEL_PATH" ]; then
        # 去除引号，转换反斜杠（Windows 路径兼容）
        EXCEL_PATH="${EXCEL_PATH%\"}"
        EXCEL_PATH="${EXCEL_PATH#\"}"
        EXCEL_PATH="${EXCEL_PATH//\\/\/}"

        if [ -f "$EXCEL_PATH" ]; then
            if [ -f config.example.json ]; then
                cp config.example.json config.json
            fi
            $PYTHON <<ENDSCRIPT >/dev/null
import json
cfg = json.load(open('config.json'))
cfg['excel_path'] = "$EXCEL_PATH"
json.dump(cfg, open('config.json', 'w', encoding='utf-8'), indent=4, ensure_ascii=False)
ENDSCRIPT
            echo "[OK] 配置已保存"
            echo "     数据文件：$EXCEL_PATH"
        else
            echo "[!!] 未找到文件，请检查路径后重新启动本脚本"
            exit 1
        fi
    else
        if [ -f config.example.json ]; then
            cp config.example.json config.json
        fi
        echo "[..] 已跳过，将使用项目目录下的 资产负债.xlsx"
        echo "     或编辑 config.json 指定路径"
    fi
    echo ""
fi

# ═══════════════════════════════════
#  Step 4 — 启动服务
# ═══════════════════════════════════

echo ""
echo "================================"
echo "  正在启动服务..."
echo "  浏览器访问：http://127.0.0.1:5000"
echo "================================"
echo ""

# 打开浏览器
if command -v open &>/dev/null; then
    open http://127.0.0.1:5000
elif command -v xdg-open &>/dev/null; then
    xdg-open http://127.0.0.1:5000
fi

$PYTHON app.py
EXIT_CODE=$?

if [ $EXIT_CODE -ne 0 ]; then
    echo ""
    echo "[!!] 服务异常退出，请检查上方错误信息"
    echo ""
    echo "  常见原因："
    echo "  1. Excel 文件路径不对 — 检查 config.json 或环境变量 JINKUI_EXCEL_PATH"
    echo "  2. Excel 格式不符合要求 — 参考 README.md 中的格式说明"
    echo "  3. 端口 5000 被占用 — 关闭其他程序后重试"
fi

exit $EXIT_CODE
