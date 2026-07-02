"""
个人财务看板 — Flask 后端
"""
from flask import Flask, render_template, jsonify, request
import parse_excel
import goals

app = Flask(__name__)


@app.route("/")
def index():
    """主看板页面"""
    return render_template("index.html")


@app.route("/api/summary")
def api_summary():
    """最新月份摘要"""
    try:
        data = parse_excel.get_summary()
        return jsonify({"ok": True, "data": data})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/trend")
def api_trend():
    """净资产趋势"""
    try:
        years = request.args.get("years", None)
        if years:
            years = [y.strip() for y in years.split(",") if y.strip()]
        data = parse_excel.get_trend(years)
        return jsonify({"ok": True, "data": data})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/breakdown")
def api_breakdown():
    """资产/负债构成"""
    try:
        latest_only = request.args.get("latest_only", "1") == "1"
        data = parse_excel.get_asset_breakdown(latest_only)
        return jsonify({"ok": True, "data": data})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/monthly")
def api_monthly():
    """指定月份明细"""
    try:
        year = request.args.get("year", type=int)
        month = request.args.get("month", type=int)
        if not year or not month:
            return jsonify({"ok": False, "error": "need year and month"}), 400
        data = parse_excel.get_monthly_detail(year, month)
        if data is None:
            return jsonify({"ok": False, "error": "not found"}), 404
        return jsonify({"ok": True, "data": data})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/years")
def api_years():
    """可用年份"""
    try:
        data = parse_excel.load_data()
        year_months = {}
        for m in data["monthly"]:
            y = str(m["year"])
            if y not in year_months:
                year_months[y] = []
            year_months[y].append(m["month"])
        return jsonify({"ok": True, "years": data["years"], "year_months": year_months})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/full_data")
def api_full_data():
    """一次性返回全量数据（供前端缓存）"""
    try:
        data = parse_excel.load_data()
        # Convert to JSON-safe format
        return jsonify({
            "ok": True,
            "data": {
                "years": data["years"],
                "months": data["month_labels"],
                "net_worth": data["net_worth"],
                "total_assets": data["total_assets"],
                "total_liabilities": data["total_liabilities"],
                "income": data["income"],
                "actual_expense": data["actual_expense"],
                "planned_expense": data["planned_expense"],
                "net_worth_change": data["net_worth_change"],
                "monthly": data["monthly"],
                "accounts": {k: v for k, v in data["accounts"].items()
                            if k not in ("合计", "合計", "资产", "資產", "负债", "負債")},
                "latest_idx": data["latest_idx"],
            }
        })
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


# ═══════════════════════════════════════════
#  财务目标 API
# ═══════════════════════════════════════════

@app.route("/api/goals", methods=["GET"])
def api_goals_list():
    """获取所有目标（含自动进度）"""
    try:
        data = goals.get_goals()
        return jsonify({"ok": True, "data": data})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/goals", methods=["POST"])
def api_goals_create():
    """创建新目标"""
    try:
        body = request.get_json(silent=True)
        if body is None:
            return jsonify({"ok": False, "error": "请求体为空或 JSON 格式无效"}), 400
        result = goals.create_goal(body)
        if "error" in result:
            return jsonify({"ok": False, "error": result["error"]}), 400
        return jsonify({"ok": True, "data": result})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/goals/<goal_id>", methods=["PUT"])
def api_goals_update(goal_id):
    """更新目标"""
    try:
        body = request.get_json(silent=True)
        if body is None:
            return jsonify({"ok": False, "error": "请求体为空或 JSON 格式无效"}), 400
        result = goals.update_goal(goal_id, body)
        if result is None:
            return jsonify({"ok": False, "error": "目标不存在"}), 404
        return jsonify({"ok": True, "data": result})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/goals/<goal_id>", methods=["DELETE"])
def api_goals_delete(goal_id):
    """删除目标"""
    try:
        ok = goals.delete_goal(goal_id)
        if not ok:
            return jsonify({"ok": False, "error": "目标不存在"}), 404
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


# ═══════════════════════════════════════════
#  报表 API
# ═══════════════════════════════════════════

@app.route("/api/report/monthly")
def api_report_monthly():
    """月度财务报告"""
    try:
        year = request.args.get("year", type=int)
        month = request.args.get("month", type=int)
        if not year or not month:
            return jsonify({"ok": False, "error": "need year and month"}), 400
        data = parse_excel.get_monthly_report(year, month)
        if data is None:
            return jsonify({"ok": False, "error": "not found"}), 404
        return jsonify({"ok": True, "data": data})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/report/annual")
def api_report_annual():
    """年度财务报告"""
    try:
        year = request.args.get("year", type=int)
        if not year:
            return jsonify({"ok": False, "error": "need year"}), 400
        data = parse_excel.get_annual_report(year)
        if data is None:
            return jsonify({"ok": False, "error": "not found"}), 404
        return jsonify({"ok": True, "data": data})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/report/yoy")
def api_report_yoy():
    """多年同比数据"""
    try:
        data = parse_excel.get_yoy_data()
        return jsonify({"ok": True, "data": data})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/refresh", methods=["POST"])
def api_refresh():
    """强制刷新数据缓存（Excel 更新后调用）"""
    try:
        parse_excel.refresh_data()
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
