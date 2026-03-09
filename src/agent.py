"""工作管理 AI Agent 主程式。"""
import argparse
import os
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

# 先載入 .env，否則 USE_POWER_AUTOMATE 等變數會是空的，誤走 Graph 登入
from dotenv import load_dotenv
load_dotenv(Path(__file__).resolve().parents[1] / ".env")

USE_PA = os.environ.get("USE_POWER_AUTOMATE", "").strip().lower() in ("1", "true", "yes")

if USE_PA:
    from src.powerautomate_client import (
        list_planner_tasks,
        list_planner_tasks_from_file,
        format_tasks_for_display,
        schedule_tasks_in_free_slots,
        schedule_next_week_to_calendar,
        write_schedule_requests_to_file,
        write_weekly_report_to_file,
        create_weekly_status_page,
        build_planner_summary_html,
        build_calendar_summary_html,
        build_next_week_report,
        format_next_week_report_text,
    )
else:
    from src.auth import get_access_token
    from src.planner_client import list_planner_tasks, format_tasks_for_display, resolve_plan_id
    from src.calendar_client import schedule_tasks_in_free_slots
    from src.onenote_client import (
        create_weekly_status_page,
        build_planner_summary_html,
        build_calendar_summary_html,
    )


def run_planner_summary(bucket_names=None):
    bucket_names = bucket_names or ["本週工作", "進行中"]
    if USE_PA:
        flow_url = os.getenv("FLOW_PLANNER_URL", "").strip()
        tasks_file = os.getenv("TASKS_INPUT_FILE", "").strip()
        if flow_url:
            tasks = list_planner_tasks(bucket_names=bucket_names)
        elif tasks_file:
            try:
                tasks = list_planner_tasks_from_file(bucket_names=bucket_names)
            except FileNotFoundError as e:
                print(e)
                print("請先手動執行 GotPlannerTasks Flow，並確認 Flow 已將結果寫入 TASKS_INPUT_FILE 路徑。")
                return
            except Exception as e:
                print("讀取任務檔案失敗：", e)
                return
        else:
            print("請在 .env 設定 FLOW_PLANNER_URL（HTTP 觸發）或 TASKS_INPUT_FILE（手動觸發 Flow 寫入的檔案路徑）。")
            return
    else:
        token = get_access_token()
        if not os.getenv("PLAN_ID") and not os.getenv("GROUP_ID"):
            print("請在 .env 設定 PLAN_ID 或 GROUP_ID（Planner 網址中的 tid= 即群組 ID）")
            return
        tasks = list_planner_tasks(token, bucket_names=bucket_names)
    print("=== 到期日接近且高優先的任務 ===\n")
    print(format_tasks_for_display(tasks))
    return tasks


def run_schedule_to_calendar(tasks=None, duration_minutes=None):
    if tasks is None:
        tasks = run_planner_summary()
        if not tasks:
            print("沒有可排程的任務。")
            return
    if USE_PA:
        if not os.getenv("FLOW_CALENDAR_URL"):
            print("請在 .env 設定 FLOW_CALENDAR_URL")
            return
        result = schedule_tasks_in_free_slots(tasks, duration_minutes=duration_minutes)
    else:
        token = get_access_token()
        result = schedule_tasks_in_free_slots(token, tasks, duration_minutes=duration_minutes)
    print("\n=== 已建立行事曆事件 ===")
    for item in result["created"]:
        ev = item["event"]
        if USE_PA:
            print(f"  - {ev.get('subject', item.get('task', {}).get('title'))} 已建立")
        else:
            print(f"  - {ev.get('subject')} @ {ev.get('start', {}).get('dateTime')}")
    if result["failed"]:
        print("\n無法排入的任務：", [t.get("title") for t in result["failed"]])


def run_weekly_status_to_onenote():
    if USE_PA:
        if not os.getenv("FLOW_ONENOTE_URL"):
            print("請在 .env 設定 FLOW_ONENOTE_URL")
            return
        planner_html = build_planner_summary_html()
        calendar_html = build_calendar_summary_html()
        page = create_weekly_status_page(
            planner_summary=planner_html,
            calendar_summary=calendar_html,
        )
    else:
        token = get_access_token()
        section_id = os.getenv("ONENOTE_SECTION_ID")
        if not section_id:
            print("請設定環境變數 ONENOTE_SECTION_ID")
            return
        try:
            plan_id = resolve_plan_id(token)
            planner_html = build_planner_summary_html(token, plan_id)
        except Exception:
            planner_html = "<p>未設定 PLAN_ID/GROUP_ID 或無法取得計畫</p>"
        calendar_html = build_calendar_summary_html(token)
        page = create_weekly_status_page(
            token,
            planner_summary=planner_html,
            calendar_summary=calendar_html,
        )
    print("OneNote 週報頁面已建立：", page.get("contentUrl", page.get("id", "")))


def run_schedule_nextweek():
    """將「延續到下週」與「下週到期需準備」排入 Outlook 行事曆。有 FLOW_CALENDAR_URL 則 HTTP 觸發；否則寫入 SCHEDULE_OUTPUT_FILE 供手動觸發 Flow 讀取。"""
    if not USE_PA:
        print("排入下週行事曆目前僅支援 Power Automate 模式，請設 USE_POWER_AUTOMATE=true 與 TASKS_INPUT_FILE。")
        return
    flow_url = os.getenv("FLOW_CALENDAR_URL", "").strip()
    schedule_file = os.getenv("SCHEDULE_OUTPUT_FILE", "").strip()
    if flow_url:
        try:
            result = schedule_next_week_to_calendar(include_tasks_without_estimate=True)
        except FileNotFoundError as e:
            print(e)
            print("請先手動執行 GotPlannerTasks Flow，並確認已寫入 TASKS_INPUT_FILE。")
            return
        except Exception as e:
            print("排入行事曆失敗：", e)
            return
        if result.get("message"):
            print(result["message"])
            return
        print("\n=== 已建立行事曆事件（下週一 9:00 起）===")
        for item in result.get("created", []):
            task = item.get("task", {})
            ev = item.get("event", {})
            mins = task.get("estimatedMinutes")
            dur = f" {mins} 分" if mins else ""
            print(f"  - {task.get('title', ev.get('subject', ''))}{dur}")
        if result.get("failed"):
            print("\n無法排入的任務：", [t.get("title") for t in result["failed"]])
        return
    if schedule_file:
        try:
            result = write_schedule_requests_to_file(include_tasks_without_estimate=True)
        except FileNotFoundError as e:
            print(e)
            print("請先手動執行 GotPlannerTasks Flow，並確認已寫入 TASKS_INPUT_FILE。")
            return
        except Exception as e:
            print("寫入排程檔失敗：", e)
            return
        if result.get("message"):
            print(result["message"])
            if not result.get("written_file"):
                return
        path = result.get("written_file", "")
        n = result.get("events_count", 0)
        ws = result.get("week_start_date")
        we = result.get("week_end_date")
        week_label = f"（排入 {ws} ~ {we}）" if ws and we else ""
        not_scheduled = result.get("not_scheduled", [])
        n_not = len(not_scheduled)
        print(f"\n已將 {n} 筆行事曆事件寫入：{path} {week_label}")
        print(f"  已排入行事曆：{n} 筆；無法排入空檔：{n_not} 筆")
        if n > 0:
            print("\n【已排入行事曆】下列任務已寫入 schedule_requests.json，執行 Flow 後會建立 Outlook 事件：")
            for i, ev in enumerate(result.get("events", []), 1):
                subj = ev.get("subject", "")
                start = ev.get("start", "")[:16] if ev.get("start") else ""
                print(f"  {i}. {subj}  {start}")
        if not_scheduled:
            print("\n【無法排入空檔】下列任務未寫入（缺估時或無空檔），請手動安排或補上 [xhr] 後重跑：")
            for i, t in enumerate(not_scheduled, 1):
                print(f"  {i}. {t.get('title', '')}")
        print("\n請依序完成：")
        print("  1. 等待 OneDrive 同步該檔案（若路徑在 OneDrive 資料夾內）。")
        print("  2. 到 Power Automate 手動執行「從檔案建立 Outlook 事件」Flow。")
        print("  3. 到 Outlook 行事曆確認事件已建立。")
        return
    print("請在 .env 設定 FLOW_CALENDAR_URL（HTTP 觸發）或 SCHEDULE_OUTPUT_FILE（手動觸發時 Agent 寫入的排程檔路徑）。")


def run_nextweek():
    """產出下週待排清單：Carryover、UpcomingDue、Triage 與總估時。"""
    if not USE_PA:
        print("下週待排清單（nextweek）目前僅支援 Power Automate 讀檔模式，請設 USE_POWER_AUTOMATE=true 與 TASKS_INPUT_FILE。")
        return
    try:
        report = build_next_week_report()
        print(format_next_week_report_text(report))
    except FileNotFoundError as e:
        print(e)
        print("請先手動執行 GotPlannerTasks Flow，並確認 Flow 已將結果寫入 TASKS_INPUT_FILE 路徑。")
    except Exception as e:
        print("產生下週報表失敗：", e)


def run_guide():
    """手動觸發 Flow 時的指引：Agent 印出待辦順序與提醒。"""
    from pathlib import Path
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).resolve().parents[1] / ".env")
    print("=== 工作管理：手動觸發 Flow 指引 ===\n")
    print("你目前使用「手動觸發」Power Automate Flow（無 Premium / 無 HTTP 觸發）。")
    print("請依以下順序在 Power Automate 中手動執行你的 Flow：\n")
    print("1. 【整理本週任務】")
    print("   開啟 Power Automate → 執行 Flow「GotPlannerTasks」（或你建立的「列出 Planner 任務」Flow）。")
    print("   若 Flow 有設定把結果寫到 OneDrive/Email，可查看任務清單。\n")
    print("2. 【排入行事曆】")
    print("   執行你建立的「建立行事曆事件」Flow（需在 Flow 內選好要排的任務或時間）。")
    print("   或手動將任務複製到 Outlook 行事曆。\n")
    print("3. 【每週三：寫入 OneNote 週報】")
    print("   執行「建立 OneNote 週報」Flow，或手動整理本週進展到 OneNote。\n")
    print("---")
    print("之後若要改為「Agent 自動呼叫 Flow」，需 Power Automate Premium（HTTP 觸發）或改回方案 A（管理員核准 App）。")
    print("執行本指引： py -m src.agent guide\n")


def run_weekly_report():
    """產出上週週報（任務 + 會議 + 統計）寫入 weekly_report.json，由 CreateOutlookEvent Flow 建立事件後讀取並建立 OneNote 頁面。"""
    if not USE_PA:
        print("週報產出目前僅支援 Power Automate 模式，請設 USE_POWER_AUTOMATE=true 與 TASKS_INPUT_FILE。")
        return
    try:
        result = write_weekly_report_to_file()
    except FileNotFoundError as e:
        print(e)
        print("請先手動執行 GotPlannerTasks Flow，並確認已寫入 TASKS_INPUT_FILE（且含上週行事曆）。")
        return
    except Exception as e:
        print("產出週報失敗：", e)
        return
    if result.get("message"):
        print(result["message"])
        return
    path = result.get("written_file", "")
    n_tasks = result.get("tasks_count", 0)
    n_meetings = result.get("meetings_count", 0)
    n_next = result.get("next_week_tasks_count", 0)
    print(f"\n已產出上週週報：{path}")
    print(f"  上一週任務 {n_tasks} 筆、上週會議 {n_meetings} 筆、下一週任務 {n_next} 筆")
    last_titles = result.get("last_week_task_titles") or []
    if last_titles:
        print("\n【上一週任務回顧】列入週報的任務：")
        for i, title in enumerate(last_titles, 1):
            short = (title[:60] + "…") if len(title) > 60 else title
            print(f"  {i}. {short}")
    meetings_list = result.get("meeting_subjects") or []
    if meetings_list:
        print("\n【上週會議】列入週報的會議：")
        for i, subj in enumerate(meetings_list, 1):
            short = (subj[:60] + "…") if len(subj) > 60 else subj
            print(f"  {i}. {short}")
    next_titles = result.get("next_week_task_titles") or []
    if next_titles:
        print("\n【下一週任務安排】列入週報的任務：")
        for i, title in enumerate(next_titles, 1):
            short = (title[:60] + "…") if len(title) > 60 else title
            print(f"  {i}. {short}")
    print("\n請依序完成：")
    print("  1. 等待 OneDrive 同步該檔案（若路徑在 OneDrive 資料夾內）。")
    print("  2. 到 Power Automate 手動執行「從檔案建立 Outlook 事件」Flow（Flow 會先建立 Outlook 事件，再讀取 weekly_report.json 於 OneNote 指定 section 新增週報頁面）。")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("action", choices=["planner", "schedule", "onenote", "all", "guide", "nextweek", "schedule-nextweek", "weekly-report"])
    parser.add_argument("--bucket", nargs="*", default=None)
    parser.add_argument("--duration", type=int, default=None)
    args = parser.parse_args()
    if args.action == "guide":
        run_guide()
        return
    if args.action == "nextweek":
        run_nextweek()
        return
    if args.action == "schedule-nextweek":
        run_schedule_nextweek()
        return
    if args.action == "weekly-report":
        run_weekly_report()
        return
    if args.action == "planner":
        run_planner_summary(bucket_names=args.bucket)
    elif args.action == "schedule":
        run_schedule_to_calendar(duration_minutes=args.duration)
    elif args.action == "onenote":
        run_weekly_status_to_onenote()
    elif args.action == "all":
        tasks = run_planner_summary(bucket_names=args.bucket)
        if tasks:
            run_schedule_to_calendar(tasks=tasks, duration_minutes=args.duration)
        print("\n若要寫入 OneNote 週報，請執行: python -m src.agent onenote")


if __name__ == "__main__":
    main()
