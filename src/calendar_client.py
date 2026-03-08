"""Outlook 行事曆：取得空檔、在空檔建立事件。"""
from datetime import datetime, timedelta, timezone
from typing import Any, Optional
import os
import requests
from dotenv import load_dotenv
from pathlib import Path

load_dotenv(Path(__file__).resolve().parents[1] / ".env")
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
DEFAULT_DURATION_MIN = int(os.environ.get("DEFAULT_TASK_DURATION_MINUTES", "60"))


def get_free_slots(
    access_token: str,
    user_id: str = "me",
    start: Optional[datetime] = None,
    end: Optional[datetime] = None,
    slot_interval_minutes: int = 30,
) -> list[dict[str, Any]]:
    start = start or datetime.now(timezone.utc)
    end = end or (start + timedelta(days=7))
    if start.tzinfo is None:
        start = start.replace(tzinfo=timezone.utc)
    if end.tzinfo is None:
        end = end.replace(tzinfo=timezone.utc)
    url = f"{GRAPH_BASE}/users/{user_id}/calendar/getSchedule"
    if user_id == "me":
        url = f"{GRAPH_BASE}/me/calendar/getSchedule"
    body = {
        "startTime": {"dateTime": start.isoformat(), "timeZone": "UTC"},
        "endTime": {"dateTime": end.isoformat(), "timeZone": "UTC"},
        "availabilityViewInterval": slot_interval_minutes,
    }
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    if user_id == "me":
        me_resp = requests.get(f"{GRAPH_BASE}/me", headers={"Authorization": f"Bearer {access_token}"})
        me_resp.raise_for_status()
        body["schedules"] = [me_resp.json().get("userPrincipalName")]
    else:
        body["schedules"] = [user_id]

    resp = requests.post(url, headers=headers, json=body)
    resp.raise_for_status()
    free_slots = []
    for sched in resp.json().get("value", []):
        for item in sched.get("scheduleItems", []):
            if (item.get("isFree") is True or (item.get("status") or "").lower() == "free"):
                free_slots.append({"start": item.get("start", {}).get("dateTime"), "end": item.get("end", {}).get("dateTime")})
    return free_slots


def find_slot_for_duration(free_slots: list[dict], duration_minutes: int) -> Optional[tuple[str, str]]:
    duration_delta = timedelta(minutes=duration_minutes)
    for slot in free_slots:
        start_s, end_s = slot.get("start"), slot.get("end")
        if not start_s or not end_s:
            continue
        try:
            st = datetime.fromisoformat(start_s.replace("Z", "+00:00"))
            et = datetime.fromisoformat(end_s.replace("Z", "+00:00"))
        except Exception:
            continue
        if (et - st) >= duration_delta:
            end_booking = st + duration_delta
            return start_s, end_booking.isoformat().replace("+00:00", "Z")
    return None


def create_calendar_event(
    access_token: str,
    subject: str,
    start: str,
    end: str,
    body: Optional[str] = None,
    user_id: str = "me",
) -> dict[str, Any]:
    url = f"{GRAPH_BASE}/me/calendar/events" if user_id == "me" else f"{GRAPH_BASE}/users/{user_id}/calendar/events"
    payload = {"subject": subject, "start": {"dateTime": start, "timeZone": "UTC"}, "end": {"dateTime": end, "timeZone": "UTC"}}
    if body:
        payload["body"] = {"contentType": "HTML", "content": body}
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    resp = requests.post(url, headers=headers, json=payload)
    resp.raise_for_status()
    return resp.json()


def schedule_tasks_in_free_slots(
    access_token: str,
    tasks: list[dict[str, Any]],
    duration_minutes: Optional[int] = None,
    user_id: str = "me",
) -> dict:
    duration_minutes = duration_minutes or DEFAULT_DURATION_MIN
    now = datetime.now(timezone.utc)
    week_end = now + timedelta(days=7)
    free_slots = get_free_slots(access_token, user_id=user_id, start=now, end=week_end)
    free_slots.sort(key=lambda x: (x.get("start") or ""))
    created, failed = [], []
    for task in tasks:
        title = task.get("title", "Planner 任務")
        slot = find_slot_for_duration(free_slots, duration_minutes)
        if not slot:
            failed.append(task)
            continue
        start_iso, end_iso = slot
        try:
            event = create_calendar_event(access_token, subject=title, start=start_iso, end=end_iso, body=f"Planner 任務 ID: {task.get('id', '')}", user_id=user_id)
            created.append({"task": task, "event": event})
            free_slots = [s for s in free_slots if s.get("start") != start_iso]
        except Exception as e:
            failed.append({**task, "error": str(e)})
    return {"created": created, "failed": failed}
