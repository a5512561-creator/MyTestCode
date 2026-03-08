"""
Power Automate 後端：透過 HTTP 觸發 Flow 取得 Planner 任務、建立行事曆事件、建立 OneNote 週報。
當 .env 設 USE_POWER_AUTOMATE=true 時，Agent 改由此模組呼叫，不需 Graph token。
手動觸發時：Flow 將結果寫入檔案，Agent 用 TASKS_INPUT_FILE 讀取並整理顯示。
"""
from datetime import datetime, timedelta, timezone
from typing import Any, Optional
import json
import os
import re
import requests
from dotenv import load_dotenv
from pathlib import Path

try:
    from zoneinfo import ZoneInfo
except ImportError:
    ZoneInfo = None  # type: ignore

load_dotenv(Path(__file__).resolve().parents[1] / ".env")

FLOW_PLANNER_URL = os.environ.get("FLOW_PLANNER_URL", "").strip()
FLOW_CALENDAR_URL = os.environ.get("FLOW_CALENDAR_URL", "").strip()
FLOW_ONENOTE_URL = os.environ.get("FLOW_ONENOTE_URL", "").strip()
TASKS_INPUT_FILE = os.environ.get("TASKS_INPUT_FILE", "").strip()
DEFAULT_DURATION_MIN = int(os.environ.get("DEFAULT_TASK_DURATION_MINUTES", "60"))
DUE_WITHIN_DAYS = int(os.environ.get("DUE_WITHIN_DAYS", "7"))
HIGH_PRIORITY_THRESHOLD = int(os.environ.get("HIGH_PRIORITY_THRESHOLD", "7"))
LOCAL_TIMEZONE = os.environ.get("LOCAL_TIMEZONE", "Asia/Taipei").strip()
NEXT_WEEK_BUCKETS_STR = os.environ.get("NEXT_WEEK_BUCKETS", "本週工作,進行中").strip()
NOT_STARTED_BUCKET = os.environ.get("NOT_STARTED_BUCKET", "未開始").strip()


def _parse_and_filter_tasks(
    raw_tasks: list,
    raw_buckets: list,
    due_within_days: int,
    min_priority: int,
    bucket_names: Optional[list[str]] = None,
) -> list[dict[str, Any]]:
    """共用的解析與篩選邏輯（到期日、優先順序、bucket）。"""
    if isinstance(raw_tasks, dict) and "value" in raw_tasks:
        raw_tasks = raw_tasks["value"]
    if isinstance(raw_buckets, dict) and "value" in raw_buckets:
        raw_buckets = raw_buckets["value"]
    buckets = {b.get("id"): b.get("name", "") for b in (raw_buckets or []) if isinstance(b, dict) and b.get("id")}
    now = datetime.now(timezone.utc)
    end_date = now + timedelta(days=due_within_days)
    if bucket_names:
        bucket_id_by_name = {name: bid for bid, name in buckets.items()}
        allowed_bucket_ids = {bucket_id_by_name[n] for n in bucket_names if n in bucket_id_by_name}
    else:
        allowed_bucket_ids = set(buckets.keys()) if buckets else set()

    result = []
    for t in raw_tasks or []:
        if not isinstance(t, dict):
            continue
        bid = t.get("bucketId")
        if allowed_bucket_ids and bid not in allowed_bucket_ids:
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
            "id": t.get("id", ""),
            "title": t.get("title", ""),
            "dueDateTime": due,
            "priority": t.get("priority", 0),
            "bucketId": bid,
            "bucketName": buckets.get(bid, ""),
            "percentComplete": t.get("percentComplete"),
            "completedDateTime": t.get("completedDateTime"),
        })
    result.sort(key=lambda x: (x.get("dueDateTime") or "", -x.get("priority", 0)))
    return result


def _is_task_incomplete(task: dict[str, Any]) -> bool:
    """任務未完成：percentComplete < 100 或 completedDateTime 為空。"""
    if task.get("completedDateTime"):
        return False
    pct = task.get("percentComplete")
    if pct is not None:
        return int(pct) < 100
    return True


def parse_duration_minutes_from_title(title: str) -> Optional[int]:
    """
    從標題解析估時（分鐘）。支援 [2h], (90m), 2.5h 等格式。
    """
    if not title or not isinstance(title, str):
        return None
    # [2h], [1.5h], [90m]
    m = re.search(r"\[(\d+(?:\.\d+)?)\s*(h|m|hr|min)\s*\]", title, re.I)
    if m:
        val, unit = float(m.group(1)), m.group(2).lower()
        if unit in ("h", "hr"):
            return int(round(val * 60))
        return int(round(val))
    # (90m), (2h)
    m = re.search(r"\((\d+(?:\.\d+)?)\s*(h|m|hr|min)\s*\)", title, re.I)
    if m:
        val, unit = float(m.group(1)), m.group(2).lower()
        if unit in ("h", "hr"):
            return int(round(val * 60))
        return int(round(val))
    # 2.5h, 90m 前綴或獨立
    m = re.search(r"(\d+(?:\.\d+)?)\s*(h|hr|m|min)\b", title, re.I)
    if m:
        val, unit = float(m.group(1)), m.group(2).lower()
        if unit in ("h", "hr"):
            return int(round(val * 60))
        return int(round(val))
    return None


def get_next_week_range(tz_name: str) -> tuple[datetime, datetime]:
    """下週一 00:00 ~ 下週日 23:59:59.999999 於指定時區。無 zoneinfo 時用 UTC。"""
    if ZoneInfo is None:
        tz = timezone.utc
    else:
        try:
            tz = ZoneInfo(tz_name)
        except Exception:
            tz = timezone.utc
    now = datetime.now(tz)
    # 下週一：今天 weekday() 0=Mon .. 6=Sun，下週一 = 現在 + (7 - weekday) 天，再設 00:00
    days_until_next_monday = (7 - now.weekday()) % 7
    if days_until_next_monday == 0:
        days_until_next_monday = 7
    start = (now + timedelta(days=days_until_next_monday)).replace(
        hour=0, minute=0, second=0, microsecond=0
    )
    end = start + timedelta(days=6, hours=23, minutes=59, seconds=59, microseconds=999999)
    return start, end


def load_raw_tasks_and_buckets_from_file(
    file_path: Optional[str] = None,
) -> tuple[list[dict[str, Any]], dict[str, str]]:
    """從 TASKS_INPUT_FILE 讀取原始 tasks 與 buckets（不做篩選），供下週報表使用。"""
    path = file_path or TASKS_INPUT_FILE
    if not path:
        raise ValueError("請在 .env 設定 TASKS_INPUT_FILE")
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(
            f"找不到檔案: {p}\n請先手動執行 GotPlannerTasks Flow，並確認 Flow 已寫入此檔案。"
        )
    with open(p, "r", encoding="utf-8") as f:
        data = json.load(f)
    raw_tasks = data.get("tasks") or data.get("value") or []
    raw_buckets = data.get("buckets") or []
    if isinstance(raw_tasks, dict) and "value" in raw_tasks:
        raw_tasks = raw_tasks["value"]
    if isinstance(raw_buckets, dict) and "value" in raw_buckets:
        raw_buckets = raw_buckets["value"]
    buckets = {
        b.get("id"): b.get("name", "")
        for b in (raw_buckets or [])
        if isinstance(b, dict) and b.get("id")
    }
    return raw_tasks, buckets


def build_next_week_report(
    file_path: Optional[str] = None,
    tz_name: Optional[str] = None,
    carryover_bucket_names: Optional[list[str]] = None,
    not_started_bucket_name: Optional[str] = None,
) -> dict[str, Any]:
    """
    產出下週待排清單：Carryover（本週工作/進行中未完成）、UpcomingDue（未開始且下週到期）、Triage（缺估時/缺到期等）。
    回傳 dict: carryover, upcoming_due, triage, total_estimated_minutes, missing_duration_titles
    """
    raw_tasks, buckets = load_raw_tasks_and_buckets_from_file(file_path)
    tz_name = tz_name or LOCAL_TIMEZONE
    carryover_bucket_names = carryover_bucket_names or [
        s.strip() for s in NEXT_WEEK_BUCKETS_STR.split(",") if s.strip()
    ]
    not_started_bucket_name = not_started_bucket_name or NOT_STARTED_BUCKET
    bucket_id_by_name = {name: bid for bid, name in buckets.items()}
    carryover_ids = {
        bucket_id_by_name[n]
        for n in carryover_bucket_names
        if n in bucket_id_by_name
    }
    not_started_id = bucket_id_by_name.get(not_started_bucket_name)
    start, end = get_next_week_range(tz_name)
    # due 多為 UTC (Z)，轉成與 start/end 相同時區再比較
    def due_in_range(due_str: Optional[str]) -> bool:
        if not due_str:
            return False
        try:
            due_dt = datetime.fromisoformat(due_str.replace("Z", "+00:00"))
            if due_dt.tzinfo is None:
                due_dt = due_dt.replace(tzinfo=timezone.utc)
            if start.tzinfo is not None and due_dt.tzinfo != start.tzinfo:
                due_dt = due_dt.astimezone(start.tzinfo)
            return start <= due_dt <= end
        except Exception:
            return False

    carryover: list[dict[str, Any]] = []
    upcoming_due: list[dict[str, Any]] = []
    triage_missing_duration: list[dict[str, Any]] = []
    triage_missing_due: list[dict[str, Any]] = []
    triage_other_bucket: list[dict[str, Any]] = []

    for t in raw_tasks or []:
        if not isinstance(t, dict):
            continue
        bid = t.get("bucketId")
        bucket_name = buckets.get(bid, "")
        due = t.get("dueDateTime")
        title = t.get("title", "")
        task_obj = {
            "id": t.get("id", ""),
            "title": title,
            "dueDateTime": due,
            "priority": t.get("priority", 0),
            "bucketId": bid,
            "bucketName": bucket_name,
            "percentComplete": t.get("percentComplete"),
            "completedDateTime": t.get("completedDateTime"),
        }
        duration_min = parse_duration_minutes_from_title(title)
        task_obj["estimatedMinutes"] = duration_min

        if bid in carryover_ids and _is_task_incomplete(task_obj):
            carryover.append(task_obj)
        if not_started_id and bid == not_started_id and due_in_range(due):
            upcoming_due.append(task_obj)
        if bid in carryover_ids or (not_started_id and bid == not_started_id):
            if duration_min is None:
                triage_missing_duration.append(task_obj)
            if not due and bid in carryover_ids:
                triage_missing_due.append(task_obj)
        elif bucket_name and bid not in carryover_ids and (not_started_id and bid != not_started_id):
            triage_other_bucket.append(task_obj)

    carryover.sort(key=lambda x: (x.get("dueDateTime") or "", -x.get("priority", 0)))
    upcoming_due.sort(key=lambda x: (x.get("dueDateTime") or "", -x.get("priority", 0)))
    total_min = sum(
        x.get("estimatedMinutes") or 0
        for x in carryover + upcoming_due
        if x.get("estimatedMinutes") is not None
    )
    missing_duration_titles = [
        x.get("title", "") for x in carryover + upcoming_due if x.get("estimatedMinutes") is None
    ]

    return {
        "carryover": carryover,
        "upcoming_due": upcoming_due,
        "triage": {
            "missing_duration": triage_missing_duration,
            "missing_due": triage_missing_due,
            "other_bucket": triage_other_bucket,
        },
        "total_estimated_minutes": total_min,
        "missing_duration_titles": missing_duration_titles,
        "next_week_start": start,
        "next_week_end": end,
    }


def format_next_week_report_text(report: dict[str, Any]) -> str:
    """將 build_next_week_report 的結果轉成可列印文字。"""
    lines = []
    start = report.get("next_week_start")
    end = report.get("next_week_end")
    if start and end:
        lines.append(f"下週區間：{start.date()} ~ {end.date()}\n")
    carryover = report.get("carryover", [])
    upcoming = report.get("upcoming_due", [])
    triage = report.get("triage", {})
    total_min = report.get("total_estimated_minutes", 0)
    missing = report.get("missing_duration_titles", [])

    lines.append("=== 延續到下週（本週工作 / 進行中，未完成）===")
    if not carryover:
        lines.append("（無）")
    else:
        for i, t in enumerate(carryover, 1):
            est = t.get("estimatedMinutes")
            est_str = f" 估時 {est} 分" if est is not None else " [缺估時]"
            lines.append(
                f"  {i}. [{t.get('bucketName', '')}] {t.get('title', '')} | "
                f"到期: {t.get('dueDateTime') or '無'}{est_str}"
            )
    lines.append("")

    lines.append("=== 下週到期需準備（未開始，到期日在下週）===")
    if not upcoming:
        lines.append("（無）")
    else:
        for i, t in enumerate(upcoming, 1):
            est = t.get("estimatedMinutes")
            est_str = f" 估時 {est} 分" if est is not None else " [缺估時]"
            lines.append(
                f"  {i}. {t.get('title', '')} | 到期: {t.get('dueDateTime')}{est_str}"
            )
    lines.append("")

    lines.append("=== 待補齊資訊（Triage）===")
    missing_dur = triage.get("missing_duration", [])
    missing_due_list = triage.get("missing_due", [])
    other = triage.get("other_bucket", [])
    if missing_dur:
        lines.append("  缺估時（請在標題加 [2h] / (90m) 等）：")
        for t in missing_dur[:20]:
            lines.append(f"    - {t.get('title', '')}")
        if len(missing_dur) > 20:
            lines.append(f"    ... 共 {len(missing_dur)} 筆")
    if missing_due_list:
        lines.append("  缺到期日（本週工作/進行中）：")
        for t in missing_due_list[:10]:
            lines.append(f"    - {t.get('title', '')}")
        if len(missing_due_list) > 10:
            lines.append(f"    ... 共 {len(missing_due_list)} 筆")
    if not missing_dur and not missing_due_list:
        lines.append("  （無）")
    lines.append("")

    lines.append("---")
    lines.append(f"下週候選總估時：{total_min} 分鐘（{total_min / 60:.1f} 小時）")
    if missing:
        lines.append(f"缺估時任務數：{len(missing)}（排入行事曆前請補上）")
    return "\n".join(lines)


def list_planner_tasks_from_file(
    file_path: Optional[str] = None,
    due_within_days: Optional[int] = None,
    min_priority: Optional[int] = None,
    bucket_names: Optional[list[str]] = None,
) -> list[dict[str, Any]]:
    """從 Flow 輸出的 JSON 檔案讀取任務並篩選（手動觸發 Flow 後使用）。"""
    path = file_path or TASKS_INPUT_FILE
    if not path:
        raise ValueError("請在 .env 設定 TASKS_INPUT_FILE（Flow 寫入的 JSON 檔案路徑）")
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"找不到檔案: {p}\n請先手動執行 GotPlannerTasks Flow，並確認 Flow 已寫入此檔案。")
    with open(p, "r", encoding="utf-8") as f:
        data = json.load(f)
    raw_tasks = data.get("tasks") or data.get("value") or []
    raw_buckets = data.get("buckets") or []
    due_within_days = due_within_days if due_within_days is not None else DUE_WITHIN_DAYS
    min_priority = min_priority if min_priority is not None else HIGH_PRIORITY_THRESHOLD
    return _parse_and_filter_tasks(raw_tasks, raw_buckets, due_within_days, min_priority, bucket_names)


def list_planner_tasks(
    plan_id: Optional[str] = None,
    due_within_days: Optional[int] = None,
    min_priority: Optional[int] = None,
    bucket_names: Optional[list[str]] = None,
) -> list[dict[str, Any]]:
    """呼叫 Flow 取得任務，在 Python 篩選與排序（與 planner_client 相同介面，無 token）。"""
    if not FLOW_PLANNER_URL:
        raise ValueError("請在 .env 設定 FLOW_PLANNER_URL（Flow 1 的 HTTP POST URL）")
    due_within_days = due_within_days if due_within_days is not None else DUE_WITHIN_DAYS
    min_priority = min_priority if min_priority is not None else HIGH_PRIORITY_THRESHOLD

    r = requests.post(FLOW_PLANNER_URL, json={}, timeout=60)
    r.raise_for_status()
    data = r.json()
    raw_tasks = data.get("tasks") or data.get("value") or []
    raw_buckets = data.get("buckets") or []
    due_within_days = due_within_days if due_within_days is not None else DUE_WITHIN_DAYS
    min_priority = min_priority if min_priority is not None else HIGH_PRIORITY_THRESHOLD
    return _parse_and_filter_tasks(raw_tasks, raw_buckets, due_within_days, min_priority, bucket_names)


def format_tasks_for_display(tasks: list[dict[str, Any]]) -> str:
    """與 planner_client 相同輸出格式。"""
    if not tasks:
        return "（無符合條件的任務）"
    lines = []
    for i, t in enumerate(tasks, 1):
        lines.append(
            f"{i}. [{t.get('bucketName', '')}] {t.get('title', '')} | "
            f"到期: {t.get('dueDateTime', '無')} | 優先: {t.get('priority', 0)}"
        )
    return "\n".join(lines)


def create_calendar_event(
    subject: str,
    start: str,
    end: str,
    body: Optional[str] = None,
) -> dict[str, Any]:
    """呼叫 Flow 建立行事曆事件。start/end 為 ISO 8601 字串。"""
    if not FLOW_CALENDAR_URL:
        raise ValueError("請在 .env 設定 FLOW_CALENDAR_URL")
    payload = {"subject": subject, "start": start, "end": end}
    if body:
        payload["body"] = body
    r = requests.post(FLOW_CALENDAR_URL, json=payload, timeout=60)
    r.raise_for_status()
    return r.json()


def schedule_tasks_in_free_slots(
    tasks: list[dict[str, Any]],
    duration_minutes: Optional[int] = None,
    start_from: Optional[datetime] = None,
) -> dict[str, Any]:
    """
    將任務排進行事曆。每筆任務優先使用 task.estimatedMinutes，否則用 duration_minutes 或預設 60 分。
    start_from 若提供則從該時刻開始排；否則從「明天 9:00」起依序排入。
    """
    default_dur = duration_minutes or DEFAULT_DURATION_MIN
    tz = timezone.utc
    now = datetime.now(tz)
    if start_from is not None:
        start = start_from
        if start.tzinfo is None:
            start = start.replace(tzinfo=tz)
    else:
        start = now.replace(hour=9, minute=0, second=0, microsecond=0)
        if start <= now:
            start = start + timedelta(days=1)
        else:
            start = start + timedelta(days=1) if (start - now).days >= 1 else start

    created = []
    failed = []
    for task in tasks:
        mins = task.get("estimatedMinutes")
        if mins is None or mins <= 0:
            mins = default_dur
        title = task.get("title", "Planner 任務")
        end = start + timedelta(minutes=mins)
        start_iso = start.isoformat().replace("+00:00", "Z")
        end_iso = end.isoformat().replace("+00:00", "Z")
        try:
            ev = create_calendar_event(
                subject=title,
                start=start_iso,
                end=end_iso,
                body=f"Planner 任務 ID: {task.get('id', '')}",
            )
            created.append({"task": task, "event": ev})
            start = end
        except Exception as e:
            failed.append({**task, "error": str(e)})
    return {"created": created, "failed": failed}


def schedule_next_week_to_calendar(
    file_path: Optional[str] = None,
    tz_name: Optional[str] = None,
    default_duration_minutes: Optional[int] = None,
    include_tasks_without_estimate: bool = True,
) -> dict[str, Any]:
    """
    將「延續到下週」與「下週到期需準備」排入 Outlook 行事曆。
    從下週一 9:00（當地時區）起依序排入，每筆任務使用標題估時或 default_duration_minutes。
    """
    report = build_next_week_report(file_path=file_path, tz_name=tz_name)
    carryover = report.get("carryover", [])
    upcoming = report.get("upcoming_due", [])
    tasks = carryover + upcoming
    if not tasks:
        return {"created": [], "failed": [], "message": "沒有下週待排任務。"}

    default_dur = default_duration_minutes or DEFAULT_DURATION_MIN
    if not include_tasks_without_estimate:
        tasks = [t for t in tasks if t.get("estimatedMinutes") is not None and t.get("estimatedMinutes", 0) > 0]
        if not tasks:
            return {"created": [], "failed": [], "message": "所有任務皆缺估時，未建立事件。請在標題加 [2h] 等後再執行。"}

    start_monday = report.get("next_week_start")
    if start_monday is None:
        start_monday = datetime.now(timezone.utc)
    # 下週一 9:00
    first_slot = start_monday.replace(hour=9, minute=0, second=0, microsecond=0)
    if first_slot.tzinfo is None:
        first_slot = first_slot.replace(tzinfo=timezone.utc)

    return schedule_tasks_in_free_slots(
        tasks,
        duration_minutes=default_dur,
        start_from=first_slot,
    )


def create_weekly_status_page(
    title: Optional[str] = None,
    planner_summary: Optional[str] = None,
    calendar_summary: Optional[str] = None,
    section_id: Optional[str] = None,
) -> dict[str, Any]:
    """組裝週報 HTML 並呼叫 Flow 建立 OneNote 頁面。"""
    if not FLOW_ONENOTE_URL:
        raise ValueError("請在 .env 設定 FLOW_ONENOTE_URL")
    section_id = section_id or os.environ.get("ONENOTE_SECTION_ID", "")
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

    payload = {"title": title, "content": content}
    if section_id:
        payload["sectionId"] = section_id
    r = requests.post(FLOW_ONENOTE_URL, json=payload, timeout=60)
    r.raise_for_status()
    return r.json()


def build_planner_summary_html(
    due_within_days: int = 7,
) -> str:
    """從 Flow 取得任務並產出週報用的 HTML 片段。"""
    tasks = list_planner_tasks(due_within_days=due_within_days, min_priority=0)
    if not tasks:
        return "<p>本週無符合條件的 Planner 任務。</p>"
    lines = ["<ul>"]
    for t in tasks:
        lines.append(
            f"<li><strong>{t.get('title', '')}</strong> [{t.get('bucketName', '')}] "
            f"到期: {t.get('dueDateTime', '無')} 優先: {t.get('priority', 0)}</li>"
        )
    lines.append("</ul>")
    return "\n".join(lines)


def build_calendar_summary_html() -> str:
    """方案 B 未實作取得行事曆事件，回傳說明文字。"""
    return "<p>本週行事曆摘要（方案 B 未建「取得行事曆」Flow 時不顯示）。</p>"
