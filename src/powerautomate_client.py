"""
Power Automate 後端：透過 HTTP 觸發 Flow 取得 Planner 任務、建立行事曆事件、建立 OneNote 週報。
當 .env 設 USE_POWER_AUTOMATE=true 時，Agent 改由此模組呼叫，不需 Graph token。
手動觸發時：Flow 將結果寫入檔案，Agent 用 TASKS_INPUT_FILE 讀取並整理顯示。
"""
from datetime import datetime, timedelta, timezone
from typing import Any, Optional
import html
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
SCHEDULE_OUTPUT_FILE = os.environ.get("SCHEDULE_OUTPUT_FILE", "").strip()
SCHEDULE_LIMIT = int(os.environ.get("SCHEDULE_LIMIT", "0") or "0")  # 寫入 schedule_requests.json 的筆數上限，0=不限制（測試可設 2）
DEFAULT_DURATION_MIN = int(os.environ.get("DEFAULT_TASK_DURATION_MINUTES", "60"))
DUE_WITHIN_DAYS = int(os.environ.get("DUE_WITHIN_DAYS", "7"))
HIGH_PRIORITY_THRESHOLD = int(os.environ.get("HIGH_PRIORITY_THRESHOLD", "7"))
LOCAL_TIMEZONE = os.environ.get("LOCAL_TIMEZONE", "Asia/Taipei").strip()
NEXT_WEEK_BUCKETS_STR = os.environ.get("NEXT_WEEK_BUCKETS", "本週工作,進行中").strip()
NOT_STARTED_BUCKET = os.environ.get("NOT_STARTED_BUCKET", "未開始").strip()
# 工作時段（避開會議用）：當地時間，格式 HH:MM；WORK_DAYS 為 0=一 … 6=日，逗號分隔
WORK_DAY_START = os.environ.get("WORK_DAY_START", "09:00").strip()
WORK_DAY_END = os.environ.get("WORK_DAY_END", "18:00").strip()
LUNCH_START = os.environ.get("LUNCH_START", "12:00").strip()
LUNCH_END = os.environ.get("LUNCH_END", "13:00").strip()
DINNER_START = os.environ.get("DINNER_START", "18:00").strip()
DINNER_END = os.environ.get("DINNER_END", "19:00").strip()
_WORK_DAYS_STR = os.environ.get("WORK_DAYS", "0,1,2,3,4").strip()
WORK_DAYS = set(int(x.strip()) for x in _WORK_DAYS_STR.split(",") if x.strip().isdigit()) if _WORK_DAYS_STR else {0, 1, 2, 3, 4}


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


def get_last_week_range(tz_name: str) -> tuple[datetime, datetime]:
    """上週一 00:00 ~ 上週日 23:59:59（與 get_next_week_range 同一套週定義，往前 7 天）。"""
    start, end = get_next_week_range(tz_name)
    return start - timedelta(days=7), end - timedelta(days=7)


def _parse_time(s: str) -> tuple[int, int]:
    """Parse 'HH:MM' or 'H:MM' to (hour, minute)."""
    if not s or ":" not in s:
        return 9, 0
    parts = s.strip().split(":", 1)
    try:
        h, m = int(parts[0].strip()), int(parts[1].strip()) if len(parts) > 1 else 0
        return max(0, min(23, h)), max(0, min(59, m))
    except ValueError:
        return 9, 0


def build_work_windows(
    week_start: datetime,
    week_end: datetime,
    tz: Any,
) -> list[tuple[datetime, datetime]]:
    """
    在 week_start..week_end 內，依 WORK_DAY_* 與 WORK_DAYS 產出可排工作的時段（每個工作日扣除午休、晚餐）。
    六日不排：WORK_DAYS 預設 0,1,2,3,4（一～五）。
    """
    if week_start.tzinfo is None:
        week_start = week_start.replace(tzinfo=timezone.utc)
    if week_end.tzinfo is None:
        week_end = week_end.replace(tzinfo=timezone.utc)
    if tz is not None and week_start.tzinfo != tz:
        week_start = week_start.astimezone(tz)
        week_end = week_end.astimezone(tz)
    sh, sm = _parse_time(WORK_DAY_START)
    eh, em = _parse_time(WORK_DAY_END)
    lsh, lsm = _parse_time(LUNCH_START)
    leh, lem = _parse_time(LUNCH_END)
    dsh, dsm = _parse_time(DINNER_START)
    deh, dem = _parse_time(DINNER_END)
    windows: list[tuple[datetime, datetime]] = []
    current = week_start.replace(hour=0, minute=0, second=0, microsecond=0)
    one_day = timedelta(days=1)
    while current <= week_end:
        if current.weekday() not in WORK_DAYS:
            current += one_day
            continue
        # [WORK_DAY_START, LUNCH_START)
        w1_s = current.replace(hour=sh, minute=sm, second=0, microsecond=0)
        w1_e = current.replace(hour=lsh, minute=lsm, second=0, microsecond=0)
        if w1_s < w1_e:
            windows.append((w1_s, w1_e))
        # [LUNCH_END, DINNER_START)
        w2_s = current.replace(hour=leh, minute=lem, second=0, microsecond=0)
        w2_e = current.replace(hour=dsh, minute=dsm, second=0, microsecond=0)
        if w2_s < w2_e:
            windows.append((w2_s, w2_e))
        # [DINNER_END, WORK_DAY_END)
        w3_s = current.replace(hour=deh, minute=dem, second=0, microsecond=0)
        w3_e = current.replace(hour=eh, minute=em, second=0, microsecond=0)
        if w3_s < w3_e:
            windows.append((w3_s, w3_e))
        current += one_day
    return windows


def _event_to_interval(ev: dict) -> Optional[tuple[datetime, datetime]]:
    """從 Outlook/Graph 事件取出 start/end 轉成 (start_dt, end_dt)。支援 start/end 字串或 start.dateTime。"""
    start_val = ev.get("start") or ev.get("Start")
    end_val = ev.get("end") or ev.get("End")
    if isinstance(start_val, dict):
        start_val = start_val.get("dateTime") or start_val.get("DateTime")
    if isinstance(end_val, dict):
        end_val = end_val.get("dateTime") or end_val.get("DateTime")
    if not start_val or not end_val:
        return None
    try:
        s = start_val.replace("Z", "+00:00") if isinstance(start_val, str) else str(start_val)
        e = end_val.replace("Z", "+00:00") if isinstance(end_val, str) else str(end_val)
        start_dt = datetime.fromisoformat(s)
        end_dt = datetime.fromisoformat(e)
        if start_dt.tzinfo is None:
            start_dt = start_dt.replace(tzinfo=timezone.utc)
        if end_dt.tzinfo is None:
            end_dt = end_dt.replace(tzinfo=timezone.utc)
        return (start_dt, end_dt)
    except Exception:
        return None


def busy_to_free(
    work_windows: list[tuple[datetime, datetime]],
    busy_list: list[tuple[datetime, datetime]],
) -> list[tuple[datetime, datetime]]:
    """從可工作時段扣掉 busy 區間，得到空檔（已排序）。"""
    free: list[tuple[datetime, datetime]] = []
    for ws, we in work_windows:
        gaps = [(ws, we)]
        for bs, be in busy_list:
            new_gaps: list[tuple[datetime, datetime]] = []
            for gs, ge in gaps:
                if be <= gs or bs >= ge:
                    new_gaps.append((gs, ge))
                    continue
                if bs > gs:
                    new_gaps.append((gs, min(ge, bs)))
                if be < ge:
                    new_gaps.append((max(gs, be), ge))
            gaps = [g for g in new_gaps if g[1] > g[0]]
        free.extend(gaps)
    free.sort(key=lambda x: x[0])
    return free


def find_first_slot_for_duration(
    free_slots: list[tuple[datetime, datetime]],
    duration_minutes: int,
) -> Optional[tuple[datetime, datetime]]:
    """回傳 (slot_start, slot_end)，並從 free_slots 中扣除已使用的區間。"""
    delta = timedelta(minutes=duration_minutes)
    for i, (fs, fe) in enumerate(free_slots):
        if (fe - fs) >= delta:
            end_booking = fs + delta
            if end_booking < fe:
                free_slots[i] = (end_booking, fe)
            else:
                free_slots.pop(i)
            return (fs, end_booking)
    return None


def _is_all_day_or_reminder_event(ev: dict) -> bool:
    """全天或跨日提醒類事件不納入忙碌時段（不擋排程）。"""
    if ev.get("isAllDay") is True or ev.get("IsAllDay") is True:
        return True
    start_val = ev.get("start") or ev.get("Start")
    end_val = ev.get("end") or ev.get("End")
    if isinstance(start_val, dict):
        start_val = start_val.get("dateTime") or start_val.get("date") or ""
    if isinstance(end_val, dict):
        end_val = end_val.get("dateTime") or end_val.get("date") or ""
    s = str(start_val) if start_val else ""
    e = str(end_val) if end_val else ""
    if len(s) <= 10 and len(e) <= 10:
        return True
    if "T" not in s or "T" not in e:
        return True
    try:
        start_dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
        end_dt = datetime.fromisoformat(e.replace("Z", "+00:00"))
        if start_dt.tzinfo is None:
            start_dt = start_dt.replace(tzinfo=timezone.utc)
        if end_dt.tzinfo is None:
            end_dt = end_dt.replace(tzinfo=timezone.utc)
        if (end_dt - start_dt).total_seconds() >= 24 * 3600:
            return True
    except Exception:
        pass
    return False


def load_calendar_events_from_file(file_path: Optional[str] = None) -> list[tuple[datetime, datetime]]:
    """從 TASKS_INPUT_FILE 的 JSON 讀取 calendarEvents，轉成 (start, end) 區間列表。全天／跨日提醒事件會略過。"""
    path = file_path or TASKS_INPUT_FILE
    if not path or not Path(path).exists():
        return []
    with open(Path(path), "r", encoding="utf-8") as f:
        data = json.load(f)
    raw = data.get("calendarEvents") or data.get("calendar_events") or []
    if isinstance(raw, dict) and "value" in raw:
        raw = raw["value"]
    intervals: list[tuple[datetime, datetime]] = []
    for ev in raw or []:
        if not isinstance(ev, dict):
            continue
        if _is_all_day_or_reminder_event(ev):
            continue
        interval = _event_to_interval(ev)
        if interval:
            intervals.append(interval)
    intervals.sort(key=lambda x: x[0])
    return intervals


def load_calendar_events_raw_from_file(
    file_path: Optional[str] = None,
) -> list[dict[str, Any]]:
    """從 TASKS_INPUT_FILE 讀取 calendarEvents，回傳具 subject/start/end/body 的列表（排除全天）；start/end 為 datetime。"""
    path = file_path or TASKS_INPUT_FILE
    if not path or not Path(path).exists():
        return []
    with open(Path(path), "r", encoding="utf-8") as f:
        data = json.load(f)
    raw = data.get("calendarEvents") or data.get("calendar_events") or []
    if isinstance(raw, dict) and "value" in raw:
        raw = raw["value"]
    result: list[dict[str, Any]] = []
    for ev in raw or []:
        if not isinstance(ev, dict):
            continue
        if _is_all_day_or_reminder_event(ev):
            continue
        interval = _event_to_interval(ev)
        if not interval:
            continue
        start_dt, end_dt = interval
        result.append({
            "subject": ev.get("subject") or ev.get("Subject") or "",
            "start": start_dt,
            "end": end_dt,
            "body": ev.get("body") or ev.get("Body") or "",
        })
    result.sort(key=lambda x: x["start"])
    return result


def get_last_week_meetings_and_stats(
    file_path: Optional[str] = None,
    tz_name: Optional[str] = None,
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    """
    從行事曆取得上週會議（排除 Planner 排程的工作），並計算每日會議時數與總時數。
    回傳 (meetings_list, stats)。meetings_list 每項為 { subject, start, end, duration_minutes }。
    stats 為 { "by_day": { "yyyy-MM-dd": minutes }, "total_minutes": int }。
    """
    path = file_path or TASKS_INPUT_FILE
    tz_name = tz_name or LOCAL_TIMEZONE
    tz = timezone.utc
    if ZoneInfo and tz_name:
        try:
            tz = ZoneInfo(tz_name)
        except Exception:
            pass
    last_start, last_end = get_last_week_range(tz_name)
    raw = load_calendar_events_raw_from_file(path)
    planner_marker = "Planner 任務 ID"
    meetings: list[dict[str, Any]] = []
    by_day: dict[str, int] = {}
    total_minutes = 0
    for ev in raw:
        start_dt, end_dt = ev["start"], ev["end"]
        body = str(ev.get("body") or "")
        if planner_marker in body:
            continue
        if end_dt <= last_start or start_dt >= last_end:
            continue
        mins = int((end_dt - start_dt).total_seconds() / 60)
        meetings.append({
            "subject": ev.get("subject", ""),
            "start": start_dt,
            "end": end_dt,
            "duration_minutes": mins,
        })
        total_minutes += mins
        day_key = start_dt.astimezone(tz).strftime("%Y-%m-%d")
        by_day[day_key] = by_day.get(day_key, 0) + mins
    return meetings, {"by_day": by_day, "total_minutes": total_minutes}


def get_last_week_report_tasks(
    file_path: Optional[str] = None,
    tz_name: Optional[str] = None,
) -> tuple[list[dict[str, Any]], datetime, datetime]:
    """
    取得「上週」due 或 completed 的 Planner 任務，供週報表格使用。
    回傳 (tasks, last_week_start, last_week_end)。每個 task 含 id, title, priority, percentComplete, startDateTime, dueDateTime 等。
    """
    raw_tasks, buckets = load_raw_tasks_and_buckets_from_file(file_path)
    tz_name = tz_name or LOCAL_TIMEZONE
    last_start, last_end = get_last_week_range(tz_name)
    buckets_by_id = buckets

    def dt_in_last_week(dt_str: Optional[str]) -> bool:
        if not dt_str:
            return False
        try:
            dt = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            if last_start.tzinfo and dt.tzinfo != last_start.tzinfo:
                dt = dt.astimezone(last_start.tzinfo)
            return last_start <= dt <= last_end
        except Exception:
            return False

    out: list[dict[str, Any]] = []
    for t in raw_tasks or []:
        if not isinstance(t, dict):
            continue
        due = t.get("dueDateTime")
        completed = t.get("completedDateTime")
        if not dt_in_last_week(due) and not dt_in_last_week(completed):
            continue
        bid = t.get("bucketId")
        out.append({
            "id": t.get("id", ""),
            "title": t.get("title", ""),
            "priority": t.get("priority", 0),
            "percentComplete": t.get("percentComplete"),
            "startDateTime": t.get("startDateTime"),
            "dueDateTime": due,
            "bucketName": buckets_by_id.get(bid, ""),
        })
    out.sort(key=lambda x: (x.get("dueDateTime") or "", -x.get("priority", 0)))
    return out, last_start, last_end


def _fmt_date_only(dt_str: Optional[str], tz: Any) -> str:
    """將 ISO 日期字串轉成 yyyy/MM/dd 顯示。"""
    if not dt_str:
        return ""
    try:
        dt = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        if tz and dt.tzinfo != tz:
            dt = dt.astimezone(tz)
        return dt.strftime("%Y/%m/%d")
    except Exception:
        return str(dt_str)[:10] if dt_str else ""


def build_weekly_report_html(
    tasks: list[dict[str, Any]],
    meetings: list[dict[str, Any]],
    stats: dict[str, Any],
    page_title: str = "",
    tz_name: Optional[str] = None,
) -> str:
    """
    產出週報 HTML：任務表格（Priority, Item, Prog %, Start, End, BW spent %, 過去一週報告）、上週會議列表、會議時間統計。
    BW spent % 與 過去一週報告 留白。
    """
    tz = timezone.utc
    if ZoneInfo and (tz_name or LOCAL_TIMEZONE):
        try:
            tz = ZoneInfo(tz_name or LOCAL_TIMEZONE)
        except Exception:
            pass
    parts = [
        "<!DOCTYPE html><html><head><meta charset='utf-8'/></head><body>",
        f"<h1>{html.escape(page_title)}</h1>",
        "<table border='1' cellpadding='4' cellspacing='0'><thead><tr>",
        "<th>Priority</th><th>Item</th><th>Prog %</th><th>Start</th><th>End</th><th>BW spent %</th><th>過去一週報告</th>",
        "</tr></thead><tbody>",
    ]
    for t in tasks:
        pct = t.get("percentComplete")
        pct_str = f"{int(pct)}%" if pct is not None else "-"
        start_str = _fmt_date_only(t.get("startDateTime"), tz)
        end_str = _fmt_date_only(t.get("dueDateTime"), tz)
        title = html.escape(str(t.get("title") or ""))
        parts.append(
            f"<tr><td>{t.get('priority', '')}</td><td>{title}</td><td>{pct_str}</td>"
            f"<td>{start_str}</td><td>{end_str}</td><td></td><td></td></tr>"
        )
    parts.append("</tbody></table>")
    parts.append("<h2>上週會議</h2><ul>")
    for m in meetings:
        subj = html.escape(str(m.get("subject") or ""))
        mins = m.get("duration_minutes", 0)
        parts.append(f"<li>{subj}（{mins} 分鐘）</li>")
    parts.append("</ul>")
    parts.append("<h2>會議時間統計</h2><ul>")
    by_day = stats.get("by_day") or {}
    for day in sorted(by_day.keys()):
        parts.append(f"<li>{day}：{by_day[day]} 分鐘</li>")
    total = stats.get("total_minutes") or 0
    parts.append(f"<li><strong>總會議時間：{total} 分鐘</strong></li>")
    parts.append("</ul></body></html>")
    return "\n".join(parts)


def write_weekly_report_to_file(
    file_path: Optional[str] = None,
    section_id: Optional[str] = None,
    page_title: Optional[str] = None,
    tz_name: Optional[str] = None,
) -> dict[str, Any]:
    """
    產出上週週報（任務 + 會議 + 統計）並寫入 weekly_report.json（sectionId, title, content），
    與 SCHEDULE_OUTPUT_FILE 同資料夾，供 CreateOutlookEvent Flow 在建立事件後讀取並建立 OneNote 頁面。
    """
    section_id = section_id or os.environ.get("ONENOTE_SECTION_ID", "").strip()
    if not section_id:
        return {
            "written_file": None,
            "message": "請在 .env 設定 ONENOTE_SECTION_ID（要新增頁面的 OneNote section ID）。",
        }
    tasks, _last_start, _last_end = get_last_week_report_tasks(file_path, tz_name)
    meetings, stats = get_last_week_meetings_and_stats(file_path, tz_name)
    now = datetime.now(timezone.utc)
    if ZoneInfo and (tz_name or LOCAL_TIMEZONE):
        try:
            tz = ZoneInfo(tz_name or LOCAL_TIMEZONE)
            now = now.astimezone(tz)
        except Exception:
            pass
    title_str = page_title or now.strftime("%Y/%m/%d")
    html_content = build_weekly_report_html(tasks, meetings, stats, page_title=title_str, tz_name=tz_name)
    out_dir = Path(SCHEDULE_OUTPUT_FILE).parent if SCHEDULE_OUTPUT_FILE else Path.cwd()
    out_path = out_dir / "weekly_report.json"
    payload = {
        "sectionId": section_id,
        "title": title_str,
        "content": html_content,
    }
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    return {"written_file": str(out_path), "tasks_count": len(tasks), "meetings_count": len(meetings)}


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


def write_schedule_requests_to_file(
    file_path: Optional[str] = None,
    tz_name: Optional[str] = None,
    default_duration_minutes: Optional[int] = None,
    include_tasks_without_estimate: bool = True,
) -> dict[str, Any]:
    """
    將「延續到下週」與「下週到期需準備」的排程寫入 JSON 檔，供手動觸發的 Flow 讀取並建立 Outlook 事件（無 HTTP 觸發時使用）。
    若有 calendarEvents 與工作時段設定，僅排入空檔（避開會議與午休/晚餐）；否則從下週一 9:00 起依序排。
    回傳含 events、not_scheduled（無法排入的任務）。
    檔案格式：{"events": [{"subject","start","end","body"}, ...]}
    """
    report = build_next_week_report(file_path=file_path, tz_name=tz_name)
    carryover = report.get("carryover", [])
    upcoming = report.get("upcoming_due", [])
    tasks = carryover + upcoming
    if not tasks:
        return {"written_file": None, "events_count": 0, "events": [], "not_scheduled": [], "message": "沒有下週待排任務。", "week_start_date": None, "week_end_date": None}

    default_dur = default_duration_minutes or DEFAULT_DURATION_MIN
    if not include_tasks_without_estimate:
        tasks = [t for t in tasks if t.get("estimatedMinutes") is not None and t.get("estimatedMinutes", 0) > 0]
        if not tasks:
            return {"written_file": None, "events_count": 0, "events": [], "not_scheduled": [], "message": "所有任務皆缺估時。請在標題加 [2h] 等後再執行。", "week_start_date": None, "week_end_date": None}

    tz = timezone.utc
    if ZoneInfo and (tz_name or LOCAL_TIMEZONE):
        try:
            tz = ZoneInfo(tz_name or LOCAL_TIMEZONE)
        except Exception:
            pass
    week_start = report.get("next_week_start")
    week_end = report.get("next_week_end")
    if week_start is None:
        week_start = datetime.now(tz)
    if week_end is None:
        week_end = week_start + timedelta(days=6, hours=23, minutes=59, seconds=59, microseconds=999999)
    if week_start.tzinfo is None:
        week_start = week_start.replace(tzinfo=tz)
    if week_end.tzinfo is None:
        week_end = week_end.replace(tzinfo=tz)

    # 週一執行時改排入「本週」（與行事曆同一週），其餘日子仍排「下週」
    now_local = datetime.now(tz)
    if now_local.weekday() == 0:
        week_start = now_local.replace(hour=0, minute=0, second=0, microsecond=0)
        week_end = week_start + timedelta(days=6, hours=23, minutes=59, seconds=59, microseconds=999999)

    work_windows = build_work_windows(week_start, week_end, tz)
    busy_list = load_calendar_events_from_file(file_path or TASKS_INPUT_FILE)
    free_slots = busy_to_free(work_windows, busy_list)

    events: list[dict[str, Any]] = []
    not_scheduled: list[dict[str, Any]] = []
    for i, task in enumerate(tasks):
        mins = task.get("estimatedMinutes") or default_dur
        if mins <= 0:
            mins = default_dur
        title = task.get("title", "Planner 任務")
        slot = find_first_slot_for_duration(free_slots, mins)
        if slot:
            start_dt, end_dt = slot
            start_iso = start_dt.isoformat().replace("+00:00", "Z")
            end_iso = end_dt.isoformat().replace("+00:00", "Z")
            events.append({
                "subject": title,
                "start": start_iso,
                "end": end_iso,
                "body": f"Planner 任務 ID: {task.get('id', '')}",
            })
            if SCHEDULE_LIMIT > 0 and len(events) >= SCHEDULE_LIMIT:
                not_scheduled.extend(tasks[i + 1 :])
                break
        else:
            not_scheduled.append(task)

    out_path = SCHEDULE_OUTPUT_FILE
    if not out_path:
        return {"written_file": None, "events_count": len(events), "events": events, "not_scheduled": not_scheduled, "message": "請在 .env 設定 SCHEDULE_OUTPUT_FILE（寫入路徑，例如與 TASKS_INPUT_FILE 同資料夾的 schedule_requests.json）。", "week_start_date": week_start.date() if week_start else None, "week_end_date": week_end.date() if week_end else None}

    p = Path(out_path)
    p.parent.mkdir(parents=True, exist_ok=True)
    payload = {"events": events}
    with open(p, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    return {
        "written_file": str(p),
        "events_count": len(events),
        "events": events,
        "not_scheduled": not_scheduled,
        "week_start_date": week_start.date() if week_start else None,
        "week_end_date": week_end.date() if week_end else None,
    }


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
