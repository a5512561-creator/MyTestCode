# Power Automate Flow 設定步驟（方案 B 後端）

請在 Power Automate 建立以下 **3 個 Flow**，每個 Flow 使用 **「當收到 HTTP 要求時」** 觸發。建立完成後，把各 Flow 的 **HTTP POST URL** 複製到 `.env`。

---

## 前置：取得 HTTP 觸發 URL

每個 Flow 儲存後，在觸發步驟 **「當收到 HTTP 要求時」** 會出現 **「HTTP POST URL」**，複製該網址即為該 Flow 的呼叫網址。

---

## Flow 1：取得 Planner 任務

**用途**：回傳指定計畫的任務與 bucket 清單，供 Agent 篩選與顯示。

### 步驟

1. 新增即時雲端 Flow，觸發選 **「當收到 HTTP 要求時」**。
2. 在觸發的 **「要求本文 JSON 結構描述」** 可留白，或貼上以下（方便測試）：
   ```json
   {}
   ```
3. 新增動作 **「Planner」**（或 **Microsoft To Do**）→ **「列出計畫的 bucket」**（List buckets for a plan）。
   - **Plan id**：輸入你的計畫 ID，例如 `_YIMOMgqSEiTZd0FdWnnMcgACG9P`（與 Power Automate 的 Create a task 相同格式）。
4. 新增動作 **「Planner」** → **「列出計畫的任務」**（List tasks for a plan 或 List tasks）。
   - **Plan id**：同上，或選上一步的輸出。
5. 新增動作 **「回應」**（Response）。
   - **本文**：在「回應」的 **本文** 中，用「從動態內容加入」或手動輸入運算式，讓回傳為一個 JSON 物件，且包含兩個屬性：
     - **tasks**：來自「列出計畫的任務」該動作的輸出。若輸出為 `{ "value": [ ... ] }`，請用 `body('列出計畫的任務')?['value']`（將 `列出計畫的任務` 改為你畫面上該動作的實際名稱）。
     - **buckets**：來自「列出計畫的 bucket」該動作的輸出，同樣取 `value` 或整個陣列，例如 `body('列出計畫的_bucket')?['value']`。
   - 動作名稱可能是英文如 `List_tasks_for_a_plan`，請在動態內容裡點選該動作的輸出，再選 `value` 即可。
6. 儲存 Flow，複製 **HTTP POST URL**，貼到 `.env` 的 **FLOW_PLANNER_URL**。

### 若沒有「列出 bucket」動作

只做「列出任務」也可。在 **回應** 的本文改為只回傳任務，例如：
```json
{
  "tasks": @{body('List_tasks')?['value']},
  "buckets": []
}
```
Agent 會依 `bucketId` 顯示，bucket 名稱可能顯示為空白。

---

## Flow 2：建立 Outlook 行事曆事件

**用途**：在指定時間建立一筆行事曆事件。

### 步驟

1. 新增即時雲端 Flow，觸發選 **「當收到 HTTP 要求時」**。
2. 在觸發的 **「要求本文 JSON 結構描述」** 貼上：
   ```json
   {
     "type": "object",
     "properties": {
       "subject": { "type": "string" },
       "start": { "type": "string" },
       "end": { "type": "string" },
       "body": { "type": "string" }
     }
   }
   ```
3. 新增動作 **「Office 365 Outlook」**（或 **Outlook.com**）→ **「建立事件」**（Create event）或 **「排程會議」**。
   - **主旨**：`@triggerBody()?['subject']`
   - **開始**：`@triggerBody()?['start']`
   - **結束**：`@triggerBody()?['end']`
   - **本文**：`@triggerBody()?['body']`（可選）
4. 新增動作 **「回應」**：
   - **狀態碼**：200
   - **本文**：
     ```json
     {
       "success": true,
       "id": "@{body('Create_event')?['Id']}"
     }
     ```
     （將 `Create_event` 改為實際「建立事件」動作名稱；若沒有 Id，可回傳 `"id": ""`。）
5. 儲存 Flow，複製 **HTTP POST URL**，貼到 `.env` 的 **FLOW_CALENDAR_URL**。

**與 Agent 的搭配**：執行 `py -X utf8 src/agent.py schedule-nextweek` 時，Agent 會將「延續到下週」與「下週到期需準備」的任務依序 POST 到本 Flow；`start`、`end` 為 ISO 8601 字串（含時區，例如 `2026-03-09T09:00:00+08:00` 或 `...Z`）。

---

## Flow 3：建立 OneNote 週報頁面

**用途**：在指定 section 建立一頁週報（HTML 內容）。

### 步驟

1. 新增即時雲端 Flow，觸發選 **「當收到 HTTP 要求時」**。
2. 在觸發的 **「要求本文 JSON 結構描述」** 貼上：
   ```json
   {
     "type": "object",
     "properties": {
       "title": { "type": "string" },
       "content": { "type": "string" },
       "sectionId": { "type": "string" }
     }
   }
   ```
3. 新增動作 **「OneNote」** → **「建立頁面」**（Create page）。
   - **Notebook**、**Section**：若連接器要選「筆記本」與「區段」，請選你要放週報的筆記本與區段（Section）。  
   - 若連接器支援 **Section id**，可用 `@triggerBody()?['sectionId']`；否則在 Flow 裡固定選一個 Section。
   - **Title**：`@triggerBody()?['title']`
   - **Content** 或 **Page content**：`@triggerBody()?['content']`（HTML）
4. 新增動作 **「回應」**：
   - **本文**：
     ```json
     {
       "success": true,
       "contentUrl": "@{body('Create_page')?['contentUrl']}"
     }
     ```
     （將 `Create_page` 改為實際「建立頁面」動作名稱。）
5. 儲存 Flow，複製 **HTTP POST URL**，貼到 `.env` 的 **FLOW_ONENOTE_URL**。

---

## .env 設定

建立好 3 個 Flow 後，在專案根目錄的 `.env` 新增：

```env
USE_POWER_AUTOMATE=true
FLOW_PLANNER_URL=https://prod-xx.eastasia.logic.azure.com:443/workflows/xxx/triggers/manual/paths/invoke?...
FLOW_CALENDAR_URL=https://prod-xx.eastasia.logic.azure.com:443/workflows/xxx/triggers/manual/paths/invoke?...
FLOW_ONENOTE_URL=https://prod-xx.eastasia.logic.azure.com:443/workflows/xxx/triggers/manual/paths/invoke?...
```

（上述 URL 為範例，請改為你實際複製的 HTTP POST URL。）

若 OneNote Flow 在連接器裡固定選了 Section、無法從 request 傳 sectionId，可留空 `ONENOTE_SECTION_ID`，並在 Flow 內固定使用同一個 Section。

---

## 測試

在專案根目錄執行：

```powershell
py test_step_by_step.py step2
py -m src.agent planner
```

若回傳 200 且內容正確，表示 Flow 與 Agent 串接成功。
