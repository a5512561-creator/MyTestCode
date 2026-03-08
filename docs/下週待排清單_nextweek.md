# 下週待排清單（nextweek）

在 Power Automate 手動觸發 GotPlannerTasks 並寫入 `TASKS_INPUT_FILE` 後，可執行：

```bash
py -X utf8 src/agent.py nextweek
```

## 輸出內容

- **延續到下週（Carryover）**：bucket 為「本週工作」「進行中」且**未完成**的任務（`percentComplete < 100` 或無 `completedDateTime`）。
- **下週到期需準備（UpcomingDue）**：bucket 為「未開始」且**到期日落在下週一～下週日**的任務。
- **待補齊資訊（Triage）**：缺估時、缺到期日的任務清單，方便 GTD 整理。
- **下週候選總估時**：有在標題標註估時的任務，其分鐘數加總。

## 估時格式（標題）

在 Planner 任務標題中可標註時間，Agent 會解析並用於總估時與排程：

- `[2h]`、`[90m]`、`(1.5h)`、`2.5h` 等，支援 `h`/`hr`/`m`/`min`。
- 未標註的任務會出現在「缺估時」清單，排入行事曆前建議補上。

## 設定（.env 選填）

| 變數 | 說明 | 預設 |
|------|------|------|
| `LOCAL_TIMEZONE` | 下週區間與到期日比對的時區 | `Asia/Taipei` |
| `NEXT_WEEK_BUCKETS` | 延續到下週的 bucket 名稱（逗號分隔） | `本週工作,進行中` |
| `NOT_STARTED_BUCKET` | 用來篩選「下週到期」的 bucket 名稱 | `未開始` |

## 下週定義

- 下週一 00:00 ～ 下週日 23:59（依 `LOCAL_TIMEZONE`）。
- 需 Python 3.9+ 的 `zoneinfo`；若無則以 UTC 計算。

---

## 排入 Outlook 行事曆

在 `.env` 設定 **FLOW_CALENDAR_URL**（Power Automate「建立 Outlook 事件」Flow 的 HTTP POST URL）後，可執行：

```bash
py -X utf8 src/agent.py schedule-nextweek
```

- 會將 **延續到下週** 與 **下週到期需準備** 的任務，從 **下週一 9:00**（當地時區）起依序建立行事曆事件。
- 每筆任務的長度使用標題中的估時（如 `[2h]`）；無估時則使用 `DEFAULT_TASK_DURATION_MINUTES`（預設 60 分）。
- Flow 2 設定步驟請見 [Power_Automate_Flow_設定步驟.md](Power_Automate_Flow_設定步驟.md) 的「Flow 2：建立 Outlook 行事曆事件」。
