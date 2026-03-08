"""
讀取 TASKS_INPUT_FILE，顯示原始任務／bucket 與篩選條件，協助排查「無符合條件的任務」。
執行：py -X utf8 scripts/debug_tasks_file.py
"""
import json
import os
import sys
from pathlib import Path

# 載入專案根目錄的 .env
root = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(root))
from dotenv import load_dotenv
load_dotenv(root / ".env")

TASKS_INPUT_FILE = os.environ.get("TASKS_INPUT_FILE", "").strip()
DUE_WITHIN_DAYS = int(os.environ.get("DUE_WITHIN_DAYS", "7"))
HIGH_PRIORITY_THRESHOLD = int(os.environ.get("HIGH_PRIORITY_THRESHOLD", "7"))
DEFAULT_BUCKETS = ["本週工作", "進行中"]

def main():
    if not TASKS_INPUT_FILE:
        print("請在 .env 設定 TASKS_INPUT_FILE")
        return
    p = Path(TASKS_INPUT_FILE)
    if not p.exists():
        print(f"找不到檔案: {p}")
        return
    with open(p, "r", encoding="utf-8") as f:
        data = json.load(f)
    print("=== 檔案結構 ===")
    print("JSON 最上層 key:", list(data.keys()))
    raw_tasks = data.get("tasks") or data.get("value") or []
    raw_buckets = data.get("buckets") or []
    if isinstance(raw_tasks, dict) and "value" in raw_tasks:
        raw_tasks = raw_tasks["value"]
    if isinstance(raw_buckets, dict) and "value" in raw_buckets:
        raw_buckets = raw_buckets["value"]
    buckets = {b.get("id"): b.get("name", "") for b in (raw_buckets or []) if isinstance(b, dict) and b.get("id")}
    print(f"原始任務數: {len(raw_tasks)}")
    print(f"原始 bucket 數: {len(buckets)}")
    print("\n=== Bucket 名稱（Agent 預設只顯示「本週工作」「進行中」）===")
    for bid, name in buckets.items():
        print(f"  id={bid[:20]}...  name=\"{name}\"")
    if not buckets:
        print("  （無 bucket，或 Flow 未輸出 buckets）")
    print("\n=== 目前篩選條件 ===")
    print(f"  DUE_WITHIN_DAYS = {DUE_WITHIN_DAYS}")
    print(f"  HIGH_PRIORITY_THRESHOLD = {HIGH_PRIORITY_THRESHOLD}（只顯示 priority >= 此值，Planner: 1=低 5=中 9=高 10=緊急）")
    print(f"  bucket_names = {DEFAULT_BUCKETS}")
    bucket_id_by_name = {name: bid for bid, name in buckets.items()}
    allowed = {bucket_id_by_name[n] for n in DEFAULT_BUCKETS if n in bucket_id_by_name}
    print(f"  符合的 bucket id 數量: {len(allowed)}")
    if not allowed and buckets:
        print("  → 你的 Planner 欄位名稱與預設「本週工作」「進行中」不符，可改 .env 或執行時加 --bucket 指定欄位名稱。")
    print("\n=== 每筆任務（標題 / bucket / 優先順序 / 到期日）===")
    for i, t in enumerate(raw_tasks or [], 1):
        if not isinstance(t, dict):
            continue
        bid = t.get("bucketId", "")
        name = buckets.get(bid, "")
        pri = t.get("priority", 0)
        due = t.get("dueDateTime", "無")
        in_bucket = "✓" if (not allowed or bid in allowed) else "✗ bucket 不符"
        pri_ok = "✓" if pri >= HIGH_PRIORITY_THRESHOLD else f"✗ 優先<{HIGH_PRIORITY_THRESHOLD}"
        print(f"  {i}. [{name}] {t.get('title', '')} | 優先={pri} {pri_ok} | 到期={due} | {in_bucket}")
    print("\n若多數為 ✗，請調整 .env 的 HIGH_PRIORITY_THRESHOLD（例如改為 0 顯示全部）或確認 Planner 欄位是否為「本週工作」「進行中」。")

if __name__ == "__main__":
    main()
