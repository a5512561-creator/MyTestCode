"""
逐步測試：認證 → Planner → Outlook → OneNote
執行方式（在專案根目錄）：
  python test_step_by_step.py step1   # 只測認證
  python test_step_by_step.py step2   # 只測 Planner
  python test_step_by_step.py step3   # 只測 Outlook 空檔
  python test_step_by_step.py step4   # 只測 OneNote（需 ONENOTE_SECTION_ID）
  python test_step_by_step.py all     # 依序執行 step1 ~ step4
"""
import os
import sys
from pathlib import Path

# 專案根目錄
ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

def load_env():
    from dotenv import load_dotenv
    load_dotenv(ROOT / ".env")

def use_power_automate():
    load_env()
    return os.getenv("USE_POWER_AUTOMATE", "").strip().lower() in ("1", "true", "yes")

def check_env():
    """檢查必要環境變數（Graph 後端時）。"""
    load_env()
    if use_power_automate():
        return True
    needed = ["CLIENT_ID", "TENANT_ID"]
    missing = [n for n in needed if not os.getenv(n)]
    if missing:
        print(f"缺少環境變數: {missing}")
        print("請複製 config.example.env 為 .env 並填入 Azure 應用程式資訊。")
        return False
    return True

def step1_auth():
    """步驟 1：測試取得 Microsoft Graph access token。"""
    print("=== 步驟 1：測試認證 ===\n")
    load_env()
    if use_power_automate():
        print("使用 Power Automate 後端，不需 Graph 認證。可略過 step1，直接測 step2。\n")
        return True
    if not check_env():
        return False
    try:
        from src.auth import get_access_token
        token = get_access_token()
        print("OK：已取得 access token（長度", len(token), "）\n")
        return True
    except Exception as e:
        print("失敗：", e)
        return False

def step2_planner():
    """步驟 2：測試 Planner 列出任務。"""
    print("=== 步驟 2：測試 Planner 列出任務 ===\n")
    load_env()
    if use_power_automate():
        flow_url = os.getenv("FLOW_PLANNER_URL", "").strip()
        tasks_file = os.getenv("TASKS_INPUT_FILE", "").strip()
        if not flow_url and not tasks_file:
            print("請在 .env 設定 FLOW_PLANNER_URL（HTTP 觸發）或 TASKS_INPUT_FILE（手動觸發 Flow 寫入的檔案路徑）。\n")
            return False
        try:
            from src.powerautomate_client import list_planner_tasks, list_planner_tasks_from_file, format_tasks_for_display
            if flow_url:
                tasks = list_planner_tasks(bucket_names=["本週工作", "進行中"])
            else:
                tasks = list_planner_tasks_from_file(bucket_names=["本週工作", "進行中"])
            print(format_tasks_for_display(tasks))
            print("\nOK：Planner 查詢成功（Power Automate）。\n")
            return True
        except FileNotFoundError as e:
            print("失敗：", e)
            print("請先手動執行 GotPlannerTasks Flow，並確認 Flow 已寫入 TASKS_INPUT_FILE 路徑。\n")
            return False
        except Exception as e:
            print("失敗：", e)
            return False
    if not check_env():
        return False
    plan_id = os.getenv("PLAN_ID")
    group_id = os.getenv("GROUP_ID")
    if (not plan_id or plan_id == "your-plan-id") and (not group_id or group_id == "your-group-id-from-url-tid"):
        print("請在 .env 設定 PLAN_ID 或 GROUP_ID（Planner 網址中的 tid= 即群組 ID）後再執行。\n")
        return False
    try:
        from src.auth import get_access_token
        from src.planner_client import list_planner_tasks, format_tasks_for_display
        token = get_access_token()
        tasks = list_planner_tasks(token, bucket_names=["本週工作", "進行中"])
        print(format_tasks_for_display(tasks))
        print("\nOK：Planner 查詢成功。\n")
        return True
    except Exception as e:
        print("失敗：", e)
        return False

def step3_calendar():
    """步驟 3：測試 Outlook 行事曆（Graph=空檔查詢；Power Automate=建立事件）。"""
    print("=== 步驟 3：測試 Outlook 行事曆 ===\n")
    load_env()
    if use_power_automate():
        if not os.getenv("FLOW_CALENDAR_URL"):
            print("請在 .env 設定 FLOW_CALENDAR_URL。方案 B 無空檔 API，請用 py -m src.agent schedule 測試排程。\n")
            return False
        print("方案 B：請用 py -m src.agent schedule 測試建立行事曆事件。\n")
        return True
    if not check_env():
        return False
    try:
        from src.auth import get_access_token
        from src.calendar_client import get_free_slots
        token = get_access_token()
        slots = get_free_slots(token, user_id="me")
        print(f"本週空檔數量: {len(slots)}")
        for s in slots[:5]:
            print(" ", s.get("start"), "~", s.get("end"))
        if len(slots) > 5:
            print(" ...")
        print("\nOK：行事曆空檔查詢成功。\n")
        return True
    except Exception as e:
        print("失敗：", e)
        return False

def step4_onenote():
    """步驟 4：測試 OneNote 建立週報頁面。"""
    print("=== 步驟 4：測試 OneNote 週報 ===\n")
    load_env()
    if use_power_automate():
        if not os.getenv("FLOW_ONENOTE_URL"):
            print("請在 .env 設定 FLOW_ONENOTE_URL。\n")
            return False
        try:
            from src.powerautomate_client import create_weekly_status_page, build_planner_summary_html, build_calendar_summary_html
            planner_html = build_planner_summary_html()
            calendar_html = build_calendar_summary_html()
            page = create_weekly_status_page(planner_summary=planner_html, calendar_summary=calendar_html)
            print("OK：OneNote 頁面已建立（Power Automate）。", page.get("contentUrl", page.get("id", "")), "\n")
            return True
        except Exception as e:
            print("失敗：", e)
            return False
    if not check_env():
        return False
    section_id = os.getenv("ONENOTE_SECTION_ID")
    if not section_id or section_id == "your-onenote-section-id":
        print("請在 .env 設定 ONENOTE_SECTION_ID 後再執行。\n")
        return False
    try:
        from src.auth import get_access_token
        from src.onenote_client import create_weekly_status_page, build_planner_summary_html, build_calendar_summary_html
        token = get_access_token()
        plan_id = os.getenv("PLAN_ID")
        planner_html = build_planner_summary_html(token, plan_id) if plan_id and plan_id != "your-plan-id" else "<p>未設定 PLAN_ID</p>"
        calendar_html = build_calendar_summary_html(token)
        page = create_weekly_status_page(token, planner_summary=planner_html, calendar_summary=calendar_html)
        print("OK：OneNote 頁面已建立。", page.get("contentUrl", page.get("id", "")), "\n")
        return True
    except Exception as e:
        print("失敗：", e)
        return False

def main():
    step = (sys.argv[1] if len(sys.argv) > 1 else "all").lower()
    steps = {"step1": step1_auth, "step2": step2_planner, "step3": step3_calendar, "step4": step4_onenote}
    if step == "all":
        for name, fn in steps.items():
            if not fn():
                print(f"請修正後再執行。可單獨測試: python test_step_by_step.py {name}")
                sys.exit(1)
        print("全部步驟通過。")
        return
    if step not in steps:
        print("用法: python test_step_by_step.py [step1|step2|step3|step4|all]")
        sys.exit(1)
    if not steps[step]():
        sys.exit(1)

if __name__ == "__main__":
    main()
