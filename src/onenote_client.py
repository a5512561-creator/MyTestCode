"""OneNote：建立週報頁面。"""
from datetime import datetime, timedelta, timezone
from typing import Any, Optional
import os
import requests
from dotenv import load_dotenv
from pathlib import Path

load_dotenv(Path(__file__).resolve().parents[1] / ".env")
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
ONENOTE_SECTION_ID = os.environ.get("ONENOTE_SECTION_ID", "")


def create_weekly_status_page(
    access_token: str,
    section_id: Optional[str] = None,
    title: Optional[str] = None,
    planner_summary: Optional[str] = None,
    calendar_summary: Optional[str] = None,
    user_id: str = "me",
) -> dict[str, Any]:
    section_id = section_id or ONENOTE_SECTION_ID
    if not section_id:
        raise ValueError("請設定 ONENOTE_SECTION_ID 或傳入 section_id")
    now = datetime.now(timezone.utc)
    week_start = now - timedelta(days=now.weekday())
    week_end = week_start + timedelta(days=6)
    title = title or f"Weekly Status {week_start.date()} ~ {week_end.date()}"
    html_parts = [
        "<!DOCTYPE html><html><head><meta charset='utf-8'/></head><body>",
        f"<h1>{title}</h1>",
        f"<p>產生時間: {now.isoformat()}</p>",
    ]
    if planner_summary:
        html_parts.append("<h2>本週 Planner 工作進展</h2>")
        html_parts.append(planner_summary)
    if calendar_summary:
        html_parts.append("<h2>本週行事曆摘要</h2>")
        html_parts.append(calendar_summary)
    html_parts.append("</body></html>")
    content = "\n".join(html_parts)
    url = f"{GRAPH_BASE}/me/onenote/sections/{section_id}/pages" if user_id == "me" else f"{GRAPH_BASE}/users/{user_id}/onenote/sections/{section_id}/pages"
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/xhtml+xml"}
    resp = requests.post(url, headers=headers, data=content.encode("utf-8"))
    resp.raise_for_status()
    return resp.json()


def build_planner_summary_html(access_token: str, plan_id: str) -> str:
    from .planner_client import list_planner_tasks
    now = datetime.now(timezone.utc)
    week_end = now - timedelta(days=now.weekday()) + timedelta(days=6)
    tasks = list_planner_tasks(access_token, plan_id=plan_id, due_within_days=(week_end - now).days + 1, min_priority=0)
    if not tasks:
        return "<p>本週無符合條件的 Planner 任務。</p>"
    lines = ["<ul>"]
    for t in tasks:
        lines.append(f"<li><strong>{t.get('title', '')}</strong> [{t.get('bucketName', '')}] 到期: {t.get('dueDateTime', '無')} 優先: {t.get('priority', 0)}</li>")
    lines.append("</ul>")
    return "\n".join(lines)


def build_calendar_summary_html(access_token: str, user_id: str = "me") -> str:
    now = datetime.now(timezone.utc)
    week_start = now - timedelta(days=now.weekday())
    week_end = week_start + timedelta(days=6)
    url = f"{GRAPH_BASE}/me/calendarView" if user_id == "me" else f"{GRAPH_BASE}/users/{user_id}/calendarView"
    params = {"startDateTime": week_start.isoformat(), "endDateTime": week_end.isoformat()}
    headers = {"Authorization": f"Bearer {access_token}"}
    resp = requests.get(url, headers=headers, params=params)
    resp.raise_for_status()
    events = resp.json().get("value", [])
    if not events:
        return "<p>本週無行事曆事件。</p>"
    lines = ["<ul>"]
    for e in events:
        start = e.get("start", {}).get("dateTime", "")
        lines.append(f"<li>{start[:16]} - {e.get('subject', '(無標題)')}</li>")
    lines.append("</ul>")
    return "\n".join(lines)
