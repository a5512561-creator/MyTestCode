# OneNote 週報自動化：設計討論

目標：在安排下週任務到 Outlook 的同時，把**上週**的工作記錄整理到 OneNote，格式對齊你目前手寫的「Switch-DD member weekly」週報（依同仁分頁、日期分頁、任務表格欄位順序：Priority / Item / Prog % / Start / End / BW spent % / 過去一週報告）；並在 Item 表格之後列出**上週會議**與**會議時間統計**（每日會議時數、總會議時間，不計 Planner 排程的工作）。

---

## 一、要寫進週報的任務範圍（你的原則）

候選任務應包含：

| 類型 | 說明 |
|------|------|
| **1. 被安排在上週的工作** | 上週排程時有被寫入 `schedule_requests.json` 的任務 |
| 1.1 | 上週**已完成**的任務 |
| 1.2 | 上週**未完成**的任務（進行中） |
| **2. Due day 在上週** | 到期日落在上週，即使當時沒被排進 Outlook 的任務 |

也就是：**「上週有排到的」+「到期日在上週的」** 都要進週報，並標示完成／未完成與進度。

---

## 二、資料從哪裡來？（採用做法 B）

- **Planner 任務**：來自現有 Flow 1 輸出的 `tasks_output.json`（含 `percentComplete`、`completedDateTime`、`dueDateTime`、`bucketName` 等）。
- **週報候選任務**（已選定 **做法 B**）：
  - 不存歷史排程，只依「上週日期區間」從 Planner 篩選：
    - **due 在上週** 的任務，或
    - **completedDateTime 在上週** 的任務（表示上週有完成）
  - 以上兩類合併為週報候選，不區分「當時有沒有被排進 Outlook」；邏輯簡單，實作只讀 `tasks_output.json` 即可。
**結論**：實作以做法 B 為準，週報任務 = 上週區間內「due 或 completed」的 Planner 任務。**不參考上週 OneNote**（任務若切得夠細，不需讀取上週週報，流程較單純）。

---

## 三、「上週」的定義

與現有排程一致，用同一套週區間：

- 若已設定 `WEEK_START_DAY` / `WEEK_END_DAY`（例如週四～隔週三），則「上週」= 上一個週四 00:00 ～ 上一個週三 23:59。
- 若未設定，則「上週」= 上週一 00:00 ～ 上週日 23:59。

Agent 依此算出 `last_week_start`、`last_week_end`，用來篩選 due / completed 是否落在上週。

---

## 四、週報內容格式（維持現有格式）

- **筆記本**：Switch-DD member weekly（SharePoint）
- **標籤頁**：同仁姓名（例如 Paddy）
- **頁面**：日期 YYYY/MM/DD（例如 2026/03/05）
- **表格欄位順序（照原先序）**：**Priority** | **Item** | **Prog %** | **Start** | **End** | **BW spent %** | **過去一週報告**

自動化時欄位對應（表頭依上列順序輸出）：

| 欄位 | 來源 | 說明 |
|------|------|------|
| **Priority** | Planner 優先順序（1/5/9/10 等）或自訂排序 | 可依優先順序數字或「先完成、後未完成」排序 |
| **Item** | Planner 任務標題 `title` | 直接帶入 |
| **Prog %** | Planner `percentComplete` | 0–100，未設則顯示「-」或「0%」 |
| **Start** | Planner `startDateTime` 或 空白 | 有則格式化成 YYYY/MM/DD |
| **End** | Planner `dueDateTime` | 有則 YYYY/MM/DD；若有多階段（如 Review / Trial run），可之後從備註或自訂欄位解析 |
| **BW spent %** | 留白 | **Bandwidth spent**：投入多少工作時間 %，供人判斷；第一版留白，由同仁手動填寫。 |
| **過去一週報告** | 留白 | 標題維持「過去一週報告」，**內容留白**，讓人判斷要寫什麼內容。 |

輸出格式：**HTML 表格**，貼到 OneNote 頁面 body（與現有 `create_weekly_status_page` 相同方式，改為產生表格而非條列）。

---

### 4.1 上週會議列表（Item 之後）

在 **Item 任務表格的最後一列之後**，另起一段落，列出**上週所有會議**及其所花時間：

- **資料來源**：行事曆中**上週**區間內的事件。若 Flow 1 目前只抓「下週」行事曆供排程用，需改為同時輸出**上週**的 calendarEvents（例如在 `tasks_output.json` 增加 `calendarEventsLastWeek`），或由另一 Flow 寫入上週行事曆供週報讀取。
- **排除**：**不計入 Planner 排程的工作項目**。即由 Agent/Flow 寫入的「工作時段」要排除（可依事件 body 含 `Planner 任務 ID` 或 主旨與當週排程任務一致者辨識並排除）；只保留行事曆中**原本就有的會議**。
- **呈現**：每個會議一列（或條列），包含：會議名稱（subject）、開始～結束時間、花費時間（例如 1h、30m）。  
- 方便在週報中一眼看出上週開了哪些會、各佔多少時間。

---

### 4.2 會議時間統計（週報結尾）

在會議列表之後，加上**統計區塊**，資料來源同樣為**行事曆**，**不計入 Planner 安排的工作**：

| 統計項目 | 說明 |
|----------|------|
| **1. 工作日中每天的會議時間** | 每個工作日（依 WORK_DAYS）當天的會議總時長，例如：3/9 2h、3/10 1.5h、3/11 3h… |
| **2. 總會議時間** | 上週所有會議時間加總（僅行事曆會議，不含 Planner 排入的工作時段）。 |

實作時：從上週 calendarEvents 篩出「非 Planner 工作」的事件，依日彙總時長 → 輸出每日會議時數與總計。

---

## 五、Agent 流程建議（不參考上週 OneNote）

1. **決定上週區間**  
   用 `WEEK_START_DAY` / `WEEK_END_DAY`（或預設 Mon–Sun）算出 `last_week_start`、`last_week_end`，以及「週報頁面日期」→ 頁面標題用 `YYYY/MM/DD`。

2. **從 Planner 篩選候選任務**  
   讀取 `tasks_output.json`，篩出 due 或 completed 在上週的任務。

3. **組任務表格**  
   每個任務一列，欄位順序照原先：Priority, Item, Prog %, Start, End, BW spent %（留白）, 過去一週報告（留白）。

4. **從行事曆篩選上週會議（排除 Planner 工作）**  
   讀取上週行事曆事件（來自 `tasks_output.json` 的 上週 calendar 欄位，見 4.1）；排除 body 含 `Planner 任務 ID` 或可辨識為「排程工作」的事件，得到**純會議**清單。

5. **產出會議列表與統計**  
   - 在任務表格後列出：上週會議名稱 + 各場花費時間。  
   - 統計：工作日每日會議時間、總會議時間。

6. **寫入 OneNote**  
   Agent 將「任務表格 + 會議列表 + 會議統計」組成一頁 HTML，與 section ID、頁面標題一併寫入檔案；由現有 **CreateOutlookEvent** Flow 在建立 Outlook 事件後讀取該檔，於指定 section 內新增 OneNote 頁面（見下方 4）。

---

## 六、已決定的設定

1. **Section 對應**  
   Section 已事先建立好，**不需在 Flow 裡新建 section**。用 **.env** 的 **`ONENOTE_SECTION_ID`** 決定要在哪個 section 內新增頁面。  
   Agent 產出週報時會讀取 `ONENOTE_SECTION_ID`，並把此 section ID 寫入要給 Flow 讀的 payload 檔（例如 `weekly_report.json` 內含 `sectionId`、`title`、`content`），Flow 讀檔後用其中的 section ID 呼叫 OneNote 建立頁面。

2. **頁面標題**  
   寫的是「上週的週報」，標題用 **今天的日期**（撰寫週報當天，例如 `2026/03/05`）。Agent 以 `formatDateTime(now, 'yyyy/MM/dd')` 產出標題字串。

3. **「過去一週報告」與「BW spent %」**  
   兩欄**內容皆留白**，只輸出欄位標題，讓人判斷要寫什麼。

4. **不另建 OneNote Flow，併入現有 CreateOutlookEvent**  
   因公司權限無法用 Graph、HTTP 觸發也需更高權限，**不**單獨建 OneNote 專用 Flow，改為在現有的 **CreateOutlookEvent**（從檔案建立 Outlook 事件）Flow 中，在 **Create Outlook event** 之後**同時新增 OneNote 週報頁面**：  
   - 流程：手動觸發 → 讀取排程檔（如 `schedule_requests.json`）→ 建立 Outlook 事件 → **再讀取週報檔**（如 `weekly_report.json`，內含 `sectionId`、`title`、`content`）→ 呼叫 **OneNote** 連接器在該 section 建立頁面。  
   - Agent 端：執行週報指令時產出 **weekly_report.json**（含 `sectionId`、`title`、`content`）寫入與 `schedule_requests.json` **同資料夾**（可由 .env 的 `SCHEDULE_OUTPUT_FILE` 推得資料夾，或另設 `WEEKLY_REPORT_OUTPUT_FILE`），Flow 於建立完 Outlook 事件後讀取該檔並建立 OneNote 頁面。

---

## 七、建議實作順序

1. **Phase 1**  
   - 用「上週」區間 + Planner（`tasks_output.json`）篩選「due 或 completed 在上週」的任務。  
   - 產出**任務表格**（Priority, Item, Prog %, Start, End, BW spent %, 過去一週報告；後兩欄留白）。  
   - 在 Item 表格之後產出**上週會議列表**（行事曆會議，排除 Planner 工作）及**會議時間統計**（每日會議時數、總會議時間）。  
   - Agent 將「任務表格 + 會議列表 + 統計」寫入 **weekly_report.json**（含 `sectionId` 來自 .env、`title` = 今天日期、`content` = HTML），與 schedule_requests 同資料夾；由現有 **CreateOutlookEvent** Flow 在建立 Outlook 事件後讀取該檔並在指定 section 新增 OneNote 頁面（頁面標題 = 今天日期）。  
   - 不參考上週 OneNote。上週行事曆需由 Flow 1 一併提供。

2. **Phase 2（可選）**  
   - 每週存一份 `schedule_requests_YYYY-Www.json`，週報時用來標示「1.1 上週完成 / 1.2 上週未完成」。  
   - 或多同仁、多 section 支援。

3. **Phase 3（可選）**  
   - 支援多同仁、多 section（依參數或設定選 section）。  
   - 若 OneNote API 支援，可依「筆記本名 + Section 名」解析，減少手動查 section ID。

---

## Phase 1 已完成：Agent 指令與後續 Flow 擴充

**Agent 指令**：執行以下即會產出 `weekly_report.json`（與 `schedule_requests.json` 同資料夾）：

```powershell
py -X utf8 src/agent.py weekly-report
```

產出內容：上週任務表格（Priority, Item, Prog %, Start, End, BW spent %, 過去一週報告）、上週會議列表、會議時間統計（每日 + 總計）。需先有 `tasks_output.json`（含上週與下週行事曆）與 `.env` 的 `ONENOTE_SECTION_ID`。

**Flow 擴充**：在現有 **CreateOutlookEvent**（從檔案建立 Outlook 事件）Flow 中，在「建立 Outlook 事件」之後新增：

1. **取得檔案**（或 **Get file content**）：讀取同資料夾的 **weekly_report.json**（若檔案不存在可略過或加條件）。
2. **Parse JSON**：解析出 `sectionId`、`title`、`content`。
3. **OneNote** 連接器：**Create page**（或等同動作），在 **Section** = `sectionId`、**Page title** = `title`、**Page content** = `content`（HTML）建立新頁面。

完成後，手動執行該 Flow 會先建立 Outlook 事件，再於 OneNote 指定 section 新增週報頁面（標題 = 今天日期）。
