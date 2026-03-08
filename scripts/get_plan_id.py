"""
取得「你有權限看到的」Planner 計畫 ID，只需 Tasks.ReadWrite，不需 Group.Read.All。

使用方式（在專案根目錄）：
  py scripts/get_plan_id.py

會用「你的帳號」裝置碼登入，從「分配給你的任務」反查計畫，列出計畫 ID。
請把要用的 PLAN_ID 複製到 .env 的 PLAN_ID= 後方。
"""
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from dotenv import load_dotenv
load_dotenv(ROOT / ".env")

# 使用裝置碼委派登入（你的帳號）
os.environ["DEVICE_CODE_FLOW"] = "true"
from src.auth import get_access_token
import requests

GRAPH = "https://graph.microsoft.com/v1.0"

def main():
    print("正在取得 token（請在瀏覽器完成登入）...\n")
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}

    # 只靠 Tasks.ReadWrite：取得「分配給你的」所有任務，從中取出 planId
    r = requests.get(f"{GRAPH}/me/planner/tasks", headers=headers)
    if r.status_code != 200:
        print(f"查詢任務失敗: {r.status_code}")
        print(r.text)
        sys.exit(1)

    tasks = r.json().get("value", [])
    plan_ids = list({t.get("planId") for t in tasks if t.get("planId")})

    if not plan_ids:
        print("目前沒有「分配給你」的 Planner 任務，無法反查計畫。")
        print("請在 Planner 中至少有一項任務是指派給你的，再執行此腳本。")
        sys.exit(1)

    # 用 planId 查計畫標題（GET /planner/plans/{id} 只需 Tasks.Read）
    plans = []
    for pid in plan_ids:
        r2 = requests.get(f"{GRAPH}/planner/plans/{pid}", headers=headers)
        if r2.status_code == 200:
            p = r2.json()
            plans.append({"id": p["id"], "title": p.get("title", "(無標題)")})
        else:
            plans.append({"id": pid, "title": "(無法取得標題)"})

    print("你有權限的 Planner 計畫（請任選一個 PLAN_ID 貼到 .env）：\n")
    for p in plans:
        print(f"  標題: {p['title']}")
        print(f"  PLAN_ID={p['id']}\n")
    print("請將上方的 PLAN_ID= 整行複製到 .env，並可註解 # GROUP_ID。")

if __name__ == "__main__":
    main()
