# Flow 1（GotPlannerTasks）：手動觸發、寫入 tasks_output.json

手動觸發流程，從 Planner 取得「指給我的」任務與 buckets，並取得 Outlook 行事曆（上週+下週），組裝成 **tasks_output.json**（含 `tasks`、`buckets`、`calendarEvents`）寫入 OneDrive。Agent 讀取該檔用於 `planner` / `nextweek` / `schedule-nextweek` 與週報。

---

## 一、流程概覽

1. 手動觸發 → List my tasks（Planner）→ List buckets for a plan → 取得行事曆檢視中的事件（上週+下週）→ Compose（tasks + buckets + calendarEvents）→ Create file（tasks_output.json）。
2. 本機 .env 設定 **TASKS_INPUT_FILE** 為同步後檔案路徑。
3. 執行 Flow 後等同步，再執行 `py -m src.agent planner` 或 `nextweek` 等指令。

---

## 二、步驟 1：觸發

- **觸發**：**Manually trigger a flow**（手動觸發流程）
- 不需填任何參數，儲存即可。

---

## 三、步驟 2：只列「我的」任務

- 新增步驟 → **Action** → 搜尋 **Planner** → 選 **Planner**。
- 動作選 **「List my tasks」**（或 **List tasks** 底下有「我的任務」的選項）。
  - 若畫面上沒有「List my tasks」，可用 **「List tasks for a plan」**，再在 Agent 端篩選；但建議優先找 **List my tasks**。
- **參數**：
  - **List my tasks** 通常**不需要**填 Plan id，會自動回傳「指派給目前使用者」的任務。
  - 若有欄位問 Group / Plan，可留空或選你的計畫（依畫面為準）。
- 此步驟輸出會有一個 **value**（陣列），裡面是任務清單。

---

## 四、步驟 3：列 bucket（為了顯示「本週工作」「進行中」等名稱）

- 新增步驟 → **Action** → **Planner**。
- 動作選 **「List buckets for a plan」**。
- **參數**：
  - **Plan id**：填你的計畫 ID，例如 `_YIMOMgqSEiTZd0FdWnnMcgACG9P`（與你 Create a task 用的相同）。
- 此步驟輸出的 **value** 是 bucket 清單（id、name），供 Agent 對照顯示欄位名稱。

---

## 五、步驟 4：取得行事曆（上週 + 下週）

在 **List buckets for a plan** 與 **Compose** 之間新增步驟，一次取得上週與下週的會議（供排程避開會議 + 週報會議列表／統計）。

1. 搜尋 **Office 365 Outlook**（或你用的 Outlook 連接器）→ 選 **取得行事曆檢視中的事件**（Get calendar view of events）或 **List events in a calendar view**。
2. **Calendar**：選 **Calendar**（主要行事曆）。
3. **Start time**、**End time** 使用下方 fx 運算式。

**週的定義**：上週一 00:00 ～ 下週日 23:59（UTC）。若你的行事曆是台灣時間，Outlook 會依你的時區顯示，Agent 端也會依 `LOCAL_TIMEZONE` 解讀。

### Start time（上週一 00:00 UTC）

在 **Start time** 欄位右側點 **fx**（Expression），依序試以下運算式（任一個通過即可）。

**寫法 A（用 startOfWeek，多數環境可用）：**

```text
concat(formatDateTime(addDays(startOfWeek(utcNow(), 'Monday'), -7), 'yyyy-MM-dd'), 'T00:00:00Z')
```

**寫法 B（若 A 報錯，改用 sub 取代負數）：**

```text
concat(formatDateTime(addDays(utcNow(), sub(0, add(dayOfWeek(utcNow()), 6))), 'yyyy-MM-dd'), 'T00:00:00Z')
```

**寫法 C（暫時固定「7 天前」做測試）：**

```text
concat(formatDateTime(addDays(utcNow(), -7), 'yyyy-MM-dd'), 'T00:00:00Z')
```

### End time（下週日 23:59:59 UTC）

在 **End time** 欄位右側點 **fx**，貼上以下運算式。

**請直接使用：**

```text
concat(formatDateTime(addDays(startOfWeek(utcNow(), 'Sunday'), 7), 'yyyy-MM-dd'), 'T23:59:59Z')
```

**若仍報錯，可改試（不含 Z）：**

```text
concat(formatDateTime(addDays(startOfWeek(utcNow(), 'Sunday'), 7), 'yyyy-MM-dd'), 'T23:59:59')
```

### 若 Start / End 只接受「日期」、不要時間

部分連接器只給日期時，可改用：

**Start time（上週一）：**

```text
formatDateTime(addDays(startOfWeek(utcNow(), 'Monday'), -7), 'yyyy-MM-dd')
```

**End time（下週日，當日結束）：**

```text
formatDateTime(addDays(startOfWeek(utcNow(), 'Sunday'), 7), 'yyyy-MM-dd')
```

連接器若會把 End 解讀為「該日結束」，用上面即可；若需明確 23:59:59，再改用上方字串寫法。

### 輸出與後續

- 此步驟輸出會有一串事件（陣列），每筆通常有 **start**、**end**（或 **Start**、**End**）。記下此步驟名稱，例如 **Get_calendar_view_of_events**。
- 此輸出要放進下一步 **Compose** 的 **calendarEvents**，寫入 **tasks_output.json**。
- Agent 會依事件 **start / end** 判斷落在「上週」或「下週」：**schedule-nextweek** 只取下週事件當忙碌時段；**OneNote 週報** 只取上週事件當會議列表與統計。無需拆成兩個「Get calendar view」步驟。

---

## 六、步驟 5：Compose（組裝 tasks、buckets、calendarEvents）

- 新增步驟 → 搜尋 **Compose** → 選 **Compose**（Built-in）。
- 點 **Inputs** 欄位（或「輸入」）。

**做法 A：用鍵值 + 動態內容（推薦）**

1. 在 Inputs 區塊點 **「Add new parameter」** 或 **「新增參數」** → 選 **Inputs**（若已是 Inputs 則略過）。
2. 若出現 **鍵 / 值** 兩欄，新增**三組**：
   - 第一組：**鍵** 填 `tasks`，**值** 點框內 → **Dynamic content** → **List my tasks**（或你步驟 2 的名稱）→ 選 **value**。
   - 第二組：**鍵** 填 `buckets`，**值** 選 **List buckets for a plan** 的 **value**。
   - 第三組：**鍵** 填 `calendarEvents`，**值** 選 **Get calendar view of events**（或你步驟 4 的名稱）→ **value**（若沒有 **value** 改選 **body** 或代表事件陣列的欄位）。
3. 儲存後，Compose 的 **Output** 會是 `{ "tasks": [...], "buckets": [...], "calendarEvents": [...] }`。

**做法 B：用 Expression（若沒有鍵值可選）**

1. 在 **Inputs** 右側點 **fx**（Expression）。
2. 貼上以下其中一行（依你步驟名稱改單引號裡的字；步驟 4 名稱例如 `Get_calendar_view_of_events`）：
   - 若步驟 2 叫 **List my tasks**：
     ```text
     json(concat('{"tasks":', string(body('List_my_tasks')?['value']), ',"buckets":', string(body('List_buckets_for_a_plan')?['value']), ',"calendarEvents":', string(body('Get_calendar_view_of_events')?['value']), '}'))
     ```
   - 若步驟 2 叫 **List tasks**：
     ```text
     json(concat('{"tasks":', string(body('List_tasks')?['value']), ',"buckets":', string(body('List_buckets_for_a_plan')?['value']), ',"calendarEvents":', string(body('Get_calendar_view_of_events')?['value']), '}'))
     ```
3. 按 **OK**。  
   （步驟名稱有空格時，在運算式裡通常會變成底線。不確定時可在 fx 裡打 `body('` 看自動完成的選單。若事件在 **body** 或別欄，改 `['value']` 為對應欄位。）

---

## 七、步驟 6：寫入 OneDrive 檔案

- 新增步驟 → 搜尋 **OneDrive** 或 **OneDrive for Business**。
  - **個人 OneDrive**：選 **OneDrive**（或 **OneDrive (Consumer)**）。
  - **公司 OneDrive**：選 **OneDrive for Business**。
- 動作選 **Create file**（建立檔案）。

**參數建議：**

| 參數 | 要填什麼 |
|------|----------|
| **Site Address** | 僅 OneDrive for Business 有；個人 OneDrive 通常無此欄。 |
| **Folder Path** | 雲端路徑，例如：`/PlannerAgent` 或 `/工作管理/GotPlannerTask`。 |
| **File Name** | `tasks_output.json`。 |
| **File Content** | 選上一步 **Compose** 的 **Output**。若要求「文字」：Expression 填 `string(outputs('Compose'))`。 |

- 儲存 Flow 後手動執行一次，確認 OneDrive 雲端出現檔案，本機同步後再填 .env（見下方）。

---

## 八、本機 .env 設定

**TASKS_INPUT_FILE** 要填的是「本機電腦上」那個檔案的完整路徑（OneDrive 同步下來後的實際路徑），不是雲端網址。

### 個人 OneDrive（非 OneDrive for Business）

- 路徑範例：`C:\Users\你的使用者名稱\OneDrive\PlannerAgent\tasks_output.json` 或 `...\OneDrive - Personal\PlannerAgent\tasks_output.json`。
- 在檔案總管找到該檔案 → **Shift + 右鍵** → **複製為路徑**，貼到 .env。

**.env 範例（個人 OneDrive）：**

```env
USE_POWER_AUTOMATE=true
TASKS_INPUT_FILE=C:\Users\paddychen\OneDrive\PlannerAgent\tasks_output.json
```

- 路徑不要加引號；不要設 **FLOW_PLANNER_URL**（手動觸發不需 HTTP URL）。

### OneDrive for Business（公司用）

- 本機路徑長相通常為：`C:\Users\你的使用者名稱\OneDrive - 公司名\工作管理\GotPlannerTask\tasks_output.json`。
- Flow 步驟 6 的 **Folder Path** 請填雲端路徑，例如：`/工作管理/GotPlannerTask`，**File Name** 填 `tasks_output.json`。

**.env 範例（OneDrive for Business）：**

```env
USE_POWER_AUTOMATE=true
TASKS_INPUT_FILE=C:\Users\paddychen\OneDrive - Realtek\工作管理\GotPlannerTask\tasks_output.json
```

---

## 九、使用流程

1. 在 Power Automate 手動執行此 Flow，讓它寫入 `tasks_output.json`。
2. 等幾秒讓 OneDrive 同步到本機。
3. 在專案目錄執行：`py -m src.agent planner` 或 `py -m src.agent nextweek` 等。  
   Agent 會讀 **TASKS_INPUT_FILE**，顯示任務並依 bucket、到期日與行事曆處理。

---

## 十、若沒有「List my tasks」或「List buckets」

- **若沒有 List my tasks**：改用 **List tasks for a plan**（Plan id 填你的計畫 ID），會得到計畫內全部任務；Agent 目前不會依指派篩選。有 **List my tasks** 時優先使用。
- **若沒有 List buckets for a plan**：可略過步驟 3，在 Compose 裡把 **buckets** 設為空陣列，例如運算式用 `"[]"` 當 buckets 部分：  
  `json(concat('{"tasks":', string(body('List_my_tasks')?['value']), ',"buckets":[],"calendarEvents":', string(body('Get_calendar_view_of_events')?['value']), '}'))`。Agent 仍可執行；bucket 名稱可能顯示為空白。

---

## 十一、檢查輸出格式

手動執行 Flow 一次，到本機打開同步後的 **tasks_output.json**，確認有：

- `tasks`：陣列  
- `buckets`：陣列  
- `calendarEvents`：陣列，每筆事件至少要有 **start**／**end**（或 **Start**／**End**，Agent 可對應）

---

## 附錄：English step-by-step（精簡）

- **Trigger**: Manually trigger a flow.
- **Step 2**: Planner → **List my tasks**（或 List tasks for a plan + Plan id）。
- **Step 3**: Planner → **List buckets for a plan** → Plan id。
- **Step 4**: Office 365 Outlook → **Get calendar view of events** → Start/End 用上方的 fx 運算式。
- **Step 5**: **Compose** → Inputs 為 `tasks`、`buckets`、`calendarEvents`（鍵值 + Dynamic content 或 Expression）。
- **Step 6**: OneDrive → **Create file** → Folder Path、File Name `tasks_output.json`、File Content = Compose Output。
- **.env**: `TASKS_INPUT_FILE` = 本機同步後之完整路徑；不要設 FLOW_PLANNER_URL。
- **Run**: Power Automate 執行 Flow → 同步後執行 `py -X utf8 src/agent.py planner` 或 `nextweek`。
