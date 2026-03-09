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
| [Flow1.md](Flow1.md) | Flow 1：手動觸發，寫入 tasks_output.json（Planner + 行事曆上週/下週） |
| [Flow2.md](Flow2.md) | Flow 2：讀 schedule_requests 建立 Outlook 事件；擴充 OneNote 週報 |
| [避開會議與工作時段_實作指引.md](避開會議與工作時段_實作指引.md) | Part A 見 Flow1；Part B Agent 端工作時段與空檔 |

---

## OneNote 週報

| 文件 | 說明 |
|------|------|
| [OneNote_週報自動化_設計討論.md](OneNote_週報自動化_設計討論.md) | 週報要寫哪些任務、資料來源、表格格式、Agent 流程與實作階段 |

---

## 參考

| 文件 | 說明 |
|------|------|
| [TEST_STEPS.md](TEST_STEPS.md) | 逐步測試 auth / Planner / 讀檔 |
| [GITHUB_上傳步驟.md](GITHUB_上傳步驟.md) | 上傳專案到 GitHub 的步驟 |
