"""Planner：列出到期日接近、高優先順序的任務。"""
from datetime import datetime, timezone, timedelta
from typing import Any, Optional
import os
import requests
from dotenv import load_dotenv
from pathlib import Path

load_dotenv(Path(__file__).resolve().parents[1] / ".env")
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
PLAN_ID = os.environ.get("PLAN_ID", "")
GROUP_ID = os.environ.get("GROUP_ID", "")  # Planner 網址的 tid= 即群組 ID，可填此取代 PLAN_ID
DUE_WITHIN_DAYS = int(os.environ.get("DUE_WITHIN_DAYS", "7"))
HIGH_PRIORITY_THRESHOLD = int(os.environ.get("HIGH_PRIORITY_THRESHOLD", "7"))


def get_plan_id_from_group(access_token: str, group_id: str) -> str:
    """用群組 ID 取得該群組的 Planner 計畫 ID（網址中 tid= 即群組 ID）。"""
    url = f"{GRAPH_BASE}/groups/{group_id}/planner/plans"
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    plans = r.json().get("value", [])
    if not plans:
        raise ValueError(f"此群組下沒有 Planner 計畫，請確認 GROUP_ID={group_id} 是否正確")
    return plans[0]["id"]


def resolve_plan_id(access_token: str) -> str:
    """若 .env 有 PLAN_ID 用 PLAN_ID，否則用 GROUP_ID 查詢計畫 ID。"""
    if PLAN_ID:
        return PLAN_ID
    if GROUP_ID:
        return get_plan_id_from_group(access_token, GROUP_ID)
    raise ValueError("請在 .env 設定 PLAN_ID 或 GROUP_ID（Planner 網址的 tid= 即群組 ID）")


def list_planner_tasks(
    access_token: str,
    plan_id: Optional[str] = None,
    due_within_days: Optional[int] = None,
    min_priority: Optional[int] = None,
    bucket_names: Optional[list[str]] = None,
) -> list[dict[str, Any]]:
    if plan_id is None:
        plan_id = resolve_plan_id(access_token)
    due_within_days = due_within_days if due_within_days is not None else DUE_WITHIN_DAYS
    min_priority = min_priority if min_priority is not None else HIGH_PRIORITY_THRESHOLD
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}

    buckets_url = f"{GRAPH_BASE}/planner/plans/{plan_id}/buckets"
    br = requests.get(buckets_url, headers=headers)
    br.raise_for_status()
    buckets = {b["id"]: b["name"] for b in br.json().get("value", [])}
    if bucket_names:
        bucket_id_by_name = {name: bid for bid, name in buckets.items()}
        allowed_bucket_ids = {bucket_id_by_name[n] for n in bucket_names if n in bucket_id_by_name}
    else:
        allowed_bucket_ids = set(buckets.keys())

    tasks_url = f"{GRAPH_BASE}/planner/plans/{plan_id}/tasks"
    tr = requests.get(tasks_url, headers=headers)
    tr.raise_for_status()
    all_tasks = tr.json().get("value", [])

    now = datetime.now(timezone.utc)
    end_date = now + timedelta(days=due_within_days)
    result = []
    for t in all_tasks:
        if t.get("bucketId") not in allowed_bucket_ids:
            continue
        due = t.get("dueDateTime")
        if due:
            try:
                due_dt = datetime.fromisoformat(due.replace("Z", "+00:00"))
            except Exception:
                due_dt = None
            if due_dt and due_dt > end_date:
                continue
        if (t.get("priority") or 0) < min_priority:
            continue
        result.append({
            "id": t["id"],
            "title": t.get("title", ""),
            "dueDateTime": due,
            "priority": t.get("priority", 0),
            "bucketId": t.get("bucketId"),
            "bucketName": buckets.get(t.get("bucketId"), ""),
        })
    result.sort(key=lambda x: (x.get("dueDateTime") or "", -x.get("priority", 0)))
    return result


def format_tasks_for_display(tasks: list[dict[str, Any]]) -> str:
    if not tasks:
        return "（無符合條件的任務）"
    lines = []
    for i, t in enumerate(tasks, 1):
        lines.append(f"{i}. [{t.get('bucketName', '')}] {t.get('title', '')} | 到期: {t.get('dueDateTime', '無')} | 優先: {t.get('priority', 0)}")
    return "\n".join(lines)
