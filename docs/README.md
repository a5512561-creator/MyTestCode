# 文件索引

本專案以 **Power Automate 手動觸發 + 讀寫檔案** 方式串接 Planner 與 Outlook，無需 HTTP 觸發（Premium）。

---

## 使用流程

| 文件 | 說明 |
|------|------|
| [下一步_使用指引.md](下一步_使用指引.md) | 每週建議流程、指令、SCHEDULE_LIMIT 與進階選項 |
| [下週待排清單_nextweek.md](下週待排清單_nextweek.md) | `nextweek`、`schedule-nextweek` 指令與下週報表說明 |

---

## Power Automate Flow 設定

| 文件 | 說明 |
|------|------|
| [Flow1_手動觸發寫入檔案.md](Flow1_手動觸發寫入檔案.md) | Flow 1：手動觸發，寫入 tasks_output.json（中文） |
| [Flow1_StepByStep_EN.md](Flow1_StepByStep_EN.md) | Flow 1：同上，英文介面逐步版 |
| [Flow2_手動觸發_從檔案建立事件.md](Flow2_手動觸發_從檔案建立事件.md) | Flow 2：讀取 schedule_requests.json，在 Outlook 建立事件 |
| [避開會議與工作時段_實作指引.md](避開會議與工作時段_實作指引.md) | Part A：Flow 1 加入下週行事曆；Part B：Agent 只排入空檔 |

---

## 參考

| 文件 | 說明 |
|------|------|
| [無管理員核准時的替代方案.md](無管理員核准時的替代方案.md) | 為何改用 Power Automate、方案 A/B 簡述 |
| [請IT協助取得計畫ID.txt](請IT協助取得計畫ID.txt) | 向 IT 索取 Graph 用計畫 ID 的範本 |
| [TEST_STEPS.md](TEST_STEPS.md) | 逐步測試 auth / Planner / 讀檔 |
| [GITHUB_上傳步驟.md](GITHUB_上傳步驟.md) | 上傳專案到 GitHub 的步驟 |
