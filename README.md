# 金匮 · 个人财务看板

> 金匮石室，藏珍蓄宝 — 你的个人财务管理工具

## 功能

| 模块 | 说明 |
|------|------|
| 📊 **看板** | 净资产趋势、资产构成、收支分析、月度变化 |
| 🎯 **目标** | 设定财务目标（净资产 / 总资产 / 总负债），自动追踪进度 |
| 📋 **报表** | 月度 & 年度财务报告，同比对比，账户排名 |

## 快速开始

### 1. 准备数据

将资产负债表 Excel 文件放在项目目录下，命名为 `资产负债.xlsx`（或通过配置指定路径）。

Excel 格式要求：
- 每年一个工作表，表名为年份（如 `2025`）
- 首行为月份表头（如 `1月`、`2月`）
- 包含资产、负债、净资产、收支分析区域

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

### 3. 启动

**Windows:**
```bash
双击 启动看板.bat
```

**macOS / Linux:**
```bash
python app.py
# 或
bash start.sh
```

浏览器访问 `http://127.0.0.1:5000`

## 配置

### Excel 文件路径

金匮按以下优先级查找数据文件：

1. **环境变量** `JINKUI_EXCEL_PATH`
2. **config.json** 中的 `excel_path` 字段
3. **项目目录下的** `资产负债.xlsx`

新建其他设备时，复制 `config.example.json` → `config.json`，填入你的 Excel 绝对路径：

```json
{
    "excel_path": "C:/Users/你的用户名/Documents/资产负债.xlsx"
}
```

或设置环境变量（推荐）：

```bash
# Windows PowerShell
$env:JINKUI_EXCEL_PATH = "D:\财务\资产负债.xlsx"

# macOS / Linux
export JINKUI_EXCEL_PATH="/Users/you/Documents/资产负债.xlsx"
```

## 技术栈

- **后端**: Flask + openpyxl
- **前端**: 原生 HTML/CSS/JS + ECharts
- **存储**: Excel (只读) + JSON (目标数据)
- **字体**: Inter + JetBrains Mono

## 许可证

MIT
