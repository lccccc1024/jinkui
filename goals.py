"""
财务目标管理模块

数据存储: goals.json
目标类型: net_worth（净资产）, total_assets（总资产）, custom（自定义）
进度自动从 parse_excel 最新数据计算
"""
import json
import os
from datetime import datetime

GOALS_FILE = os.path.join(os.path.dirname(__file__), "goals.json")

_DEFAULT_GOALS = []


def _load_raw() -> list[dict]:
    """读取原始目标列表"""
    if not os.path.exists(GOALS_FILE):
        return []
    try:
        with open(GOALS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data if isinstance(data, list) else []
    except (json.JSONDecodeError, OSError):
        return []


def _save_raw(goals: list[dict]):
    """写入目标列表"""
    with open(GOALS_FILE, "w", encoding="utf-8") as f:
        json.dump(goals, f, ensure_ascii=False, indent=2)


def get_goals() -> list[dict]:
    """获取所有目标（含自动计算的当前值）"""
    import parse_excel
    data = parse_excel.load_data()
    idx = data["latest_idx"]
    latest = data["monthly"][idx] if idx >= 0 and idx < len(data["monthly"]) else {}

    goals = _load_raw()
    for g in goals:
        if g["metric"] == "net_worth":
            g["current_value"] = latest.get("net_worth", 0)
        elif g["metric"] == "total_assets":
            g["current_value"] = latest.get("total_assets", 0)
        elif g["metric"] == "total_liabilities":
            g["current_value"] = latest.get("total_liabilities", 0)
        else:
            g.setdefault("current_value", 0)
        # 计算进度
        target = g.get("target_value", 1)
        g["progress"] = round(min(g["current_value"] / target, 1) * 100, 1) if target > 0 else 0
        g["achieved"] = g["progress"] >= 100
    return goals


def create_goal(data: dict) -> dict:
    """创建新目标"""
    goals = _load_raw()
    raw_target = data.get("target_value")
    if raw_target is None:
        return {"error": "需要目标金额"}
    try:
        target_value = float(raw_target)
    except (ValueError, TypeError):
        return {"error": "目标金额格式无效"}
    goal = {
        "id": f"g{os.urandom(4).hex()}",
        "name": data.get("name", "").strip(),
        "metric": data.get("metric", "net_worth"),
        "target_value": target_value,
        "target_date": data.get("target_date", ""),
        "note": data.get("note", ""),
        "created_at": datetime.now().strftime("%Y-%m-%d"),
    }
    if not goal["name"] or goal["target_value"] <= 0:
        return {"error": "需要名称和目标金额"}
    goals.append(goal)
    _save_raw(goals)
    return goal


def update_goal(gid: str, data: dict) -> dict | None:
    """更新目标"""
    goals = _load_raw()
    for g in goals:
        if g["id"] == gid:
            if "name" in data:
                name = data["name"].strip() if isinstance(data["name"], str) else ""
                if not name:
                    continue  # 跳过空名称
                g["name"] = name
            if "metric" in data:
                g["metric"] = data["metric"]
            if "target_value" in data:
                raw = data["target_value"]
                if raw is None:
                    continue
                try:
                    tv = float(raw)
                except (ValueError, TypeError):
                    continue
                if tv <= 0:
                    continue
                g["target_value"] = tv
            if "target_date" in data:
                g["target_date"] = data["target_date"]
            if "note" in data:
                g["note"] = data["note"]
            _save_raw(goals)
            return g
    return None


def delete_goal(gid: str) -> bool:
    """删除目标"""
    goals = _load_raw()
    new_goals = [g for g in goals if g["id"] != gid]
    if len(new_goals) == len(goals):
        return False
    _save_raw(new_goals)
    return True
