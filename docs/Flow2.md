# Flow 2：從檔案建立 Outlook 事件 + OneNote 週報

手動觸發流程：讀取 **schedule_requests.json**，在 Outlook 建立行事曆事件；可選擴充：讀取 **weekly_report.json**，在 OneNote 指定 section 建立週報頁面。無需 HTTP 觸發（Premium）。

---

## 流程概覽

**Part A（必做）**

1. 本機執行 `py -X utf8 src/agent.py schedule-nextweek` → Agent 寫入 **schedule_requests.json**（路徑由 .env 的 `SCHEDULE_OUTPUT_FILE` 決定）。
2. 若路徑在 OneDrive 內，等同步完成。
3. 在 Power Automate 手動執行本 Flow → Flow 讀取該檔，依序在 Outlook 建立事件。
4. 到 Outlook 行事曆確認。

**Part B（擴充，可選）**

- 在「建立 Outlook 事件」之後：讀取同資料夾的 **weekly_report.json** → 解析後在 OneNote 指定 section 建立週報頁面（需先執行 `py -m src.agent weekly-report` 產出該檔）。

---

# Part A：從 schedule_requests.json 建立 Outlook 事件

## 步驟 1：觸發

1. 新增即時雲端流程，名稱可設為「從檔案建立 Outlook 事件」。
2. 觸發選 **手動觸發流程**（Manually trigger a flow）。
3. 不需新增輸入參數。

---

## 步驟 2：取得檔案內容（schedule_requests.json）

1. 新增步驟 → **動作** → 搜尋 **OneDrive for Business**（或你使用的 OneDrive 連接器）。
2. 選擇 **取得檔案內容**（Get file content）。
   - 若沒有「取得檔案內容」，可用 **列出資料夾中的檔案** 指定資料夾後，再用 **取得檔案內容** 並選取檔名（見下方替代作法）。
3. 參數：
   - **檔案識別碼**（File Identifier）：若連接器要求的是「識別碼」，需先在前一步用 **列出資料夾中的檔案** 取得該檔的識別碼，再選動態內容。
   - 或若連接器有 **「依路徑取得檔案內容」**／**「Get file content using path」**：
     - **Site Address**：留預設或你的 SharePoint 網站。
     - **File Path**：雲端路徑，須與 Agent 寫入的檔案一致。例如與 `tasks_output.json` 同資料夾時：`/工作管理/GotPlannerTask/schedule_requests.json`（依你在 OneDrive 的實際路徑調整）。

**若只有「列出資料夾中的檔案」+「取得檔案內容」：**

1. 先加 **列出資料夾中的檔案**（List files in folder）：**Folder** = 你的 OneDrive 資料夾（例如 `工作管理/GotPlannerTask`）。
2. 再加 **取得檔案內容**：**File** = 從動態內容選該檔的 **Id**（若需依檔名找檔案，可在「列出資料夾中的檔案」後加 **篩選陣列**，條件為 `item()?['Name']` 等於 `schedule_requests.json`，再對篩選結果用「套用到每個」跑「取得檔案內容」）。

---

## 步驟 3a：Compose － 將檔案內容轉成字串（避免 Parse JSON 報錯）

**錯誤說明**：若 Parse JSON 出現「content must be of type JSON, but was of type 'application/octet-stream'」，代表「取得檔案內容」回傳的是二進位，Parse JSON 需要**字串**。請在「取得檔案內容」與「Parse JSON」之間加一個 **Compose**，把內容轉成字串。

1. 在 **取得檔案內容** 與 **Parse JSON** 之間，新增步驟 → **Compose**。
2. 點 **Inputs** 右側的 **fx**（Expression），貼上以下其中一個運算式（依你的「取得檔案內容」動作名稱調整單引號內名稱，空格通常會變成底線，例如 `Get file content` → `Get_file_content`）：

   **若 OneDrive 輸出為 Base64：**
   ```text
   base64ToString(body('Get_file_content')?['$content'])
   ```

   **若沒有 `$content`，可試：**
   ```text
   base64ToString(outputs('Get_file_content')?['body']?['$content'])
   ```

   **若仍失敗，改試直接轉字串：**
   ```text
   string(body('Get_file_content'))
   ```

3. 儲存後，下一步 **Parse JSON** 的 **內容** 改為選這個 **Compose** 的 **Output**（不要再用「取得檔案內容」的檔案內容）。

**若仍出現 application/octet-stream 錯誤：**

- **Parse JSON** 的 **Content** 必須選 **Compose**（步驟 3a）的 **Output**，不可選「取得檔案內容」的 File content。
- **Compose** 的 **Inputs** 必須是**運算式**（如上），不能直接選「File content」；否則 Compose 會把二進位原樣傳出，Parse JSON 仍會報錯。

---

## 步驟 3b：剖析 JSON（取得 events 陣列）

1. 新增步驟 → **動作** → 搜尋 **Parse JSON** 或 **剖析 JSON**。
2. **內容**（Content）：**必須**選 **Compose**（步驟 3a）的 **Output**。
3. **結構描述**（Schema）：貼上以下 JSON（或點「使用範例產生結構描述」後貼上整段 `{"events":[...]}` 範例）：

```json
{
  "type": "object",
  "properties": {
    "events": {
      "type": "array",
      "items": {
        "type": "object",
        "properties": {
          "subject": { "type": "string" },
          "start": { "type": "string" },
          "end": { "type": "string" },
          "body": { "type": "string" }
        }
      }
    }
  }
}
```

4. 儲存後，後續步驟可使用 **events** 陣列。

---

## 步驟 4：套用到每個事件並建立 Outlook 事件

1. 新增步驟 → **套用到每個**（Apply to each）。
2. **選取輸出從先前的步驟**：選 **Parse JSON**（步驟 3b）的 **events**。
3. 在 **套用到每個** 迴圈內新增動作 → **Office 365 Outlook** → **建立事件**（Create an event）。
4. 對應欄位（從 **套用到每個** 的目前項目取値）：
   - **主旨**：`subject`
   - **開始**：`start`
   - **結束**：`end`
   - **本文**：`body`

5. 儲存 Flow。

---

## 可選：僅有「列出資料夾中的檔案」時的作法

若 OneDrive 連接器沒有「依路徑取得檔案內容」：

1. **列出資料夾中的檔案**：Folder = 你存 schedule_requests.json 的資料夾。
2. **套用到每個**（外層）：選「列出資料夾中的檔案」的 **value**。
3. **條件**：目前項目的 **Name** 等於 `schedule_requests.json`（若不符合則略過）。
4. **若為是**：**取得檔案內容**，File = 目前項目的 **Id**。
5. 再接 **Compose 轉字串** → **Parse JSON** → 內層 **套用到每個** events → **建立事件**（同上）。

---

# Part B：擴充 — 讀取 weekly_report.json 建立 OneNote 週報

在「套用到每個（建立 Outlook 事件）」**之後**新增以下步驟。需先在本機執行 `py -m src.agent weekly-report` 產出 **weekly_report.json**（與 schedule_requests.json 同資料夾），並在 .env 設定 **ONENOTE_SECTION_ID**。

---

## 前置：取得 weekly_report.json 的檔案內容

1. **取得檔案內容**（Get file content）：與讀取 `schedule_requests.json` 的作法相同，改為讀取**同資料夾**的 **`weekly_report.json`**。
   - 若用「依路徑取得」：**File Path** 改為同資料夾的 `/工作管理/GotPlannerTask/weekly_report.json`（與你實際路徑一致）。
   - 若用「列出資料夾中的檔案」+ 篩選：篩選條件改為 `item()?['Name']` 等於 `weekly_report.json`，再取得該檔內容。

2. **Compose（轉成字串，必做）**  
   - 在「取得檔案內容（weekly_report）」與「Parse JSON」之間**必須**加一個 **Compose**，否則 Parse JSON 會收到 `application/octet-stream` 而報錯。  
   - **Inputs** 點 **fx**，貼上（將 `Get_file_content_weekly_report` 改為你「取得 weekly_report 檔案內容」的步驟名稱，空格改底線）：  
   ```text
   base64ToString(body('Get_file_content_weekly_report')?['$content'])
   ```  
   - 若沒有 `$content` 可試：`string(body('Get_file_content_weekly_report'))`  
   - **Parse JSON** 的 **Content** 務必選此 **Compose** 的 **Output**，不要選 Get file content 的輸出。

**若已加 Compose 且 Parse JSON 已選 Compose 的 Output 仍報錯：** Compose 的 Inputs **不能**用動態內容選「Get file content」的 File content／Body（那樣只是原樣轉傳）。請在 Compose 的 **Inputs** 用 **fx** 貼上上述運算式。

---

## Step 2：Parse JSON（weekly_report）

1. 新增步驟 → **動作** → 搜尋 **Parse JSON**。
2. **Content**：選上一步 **Compose** 的 **Output**（weekly_report 的字串內容）。
3. **Schema**：直接在 **結構描述** 貼上以下 JSON：

```json
{
    "type": "object",
    "properties": {
        "sectionId": {
            "type": "string"
        },
        "title": {
            "type": "string"
        },
        "content": {
            "type": "string"
        }
    }
}
```

4. 儲存後，後續步驟可用動態內容 **Parse JSON** → **sectionId**、**title**、**content**。

---

## Step 3：OneNote 建立頁面（Create page in a section）

### 欄位對應

| 參數（Power Automate 畫面） | 值（選動態內容） |
|-----------------------------|------------------|
| **Section** 或 **Section Id** | **Parse JSON** → **sectionId** |
| **Page Name** 或 **Title** 或 **Page title** | **Parse JSON** → **title** |
| **Page content** 或 **Content** 或 **HTML Content** | **Parse JSON** → **content** |

若連接器有 **Notebook** / **Section** 分開選：**Section** 填 **Parse JSON** → **sectionId**。

### 頁面沒有標題（頂部標題區空白）時

- 若 **Create page in a section** 有 **Page Name**／**Title** 欄位：請選動態內容 **Parse JSON** → **title**（例如 `2026/03/09`）。
- **若畫面上沒有 Page name／Title 可填**（只有 Notebook Key、Notebook section、Page Content）：表示此連接器不支援設定頁面標題。可採：
  1. **手動補標題**：頁面建立後，在 OneNote 該頁最上方「標題」區手動輸入日期。
  2. **以 content 當視覺標題**：週報 HTML 開頭已是 `<h1>2026/03/09</h1>`，日期會顯示在內容第一行；側邊列頁名可能為「未命名」或第一行文字，視 OneNote 版本而定。
  3. 點該動作的 **設定**（Settings）或 **顯示進階選項** 看是否出現 Title／Page name 欄位。

### Page content 與格式

- **Page content** 欄位：直接選動態內容 **Parse JSON** → **content**，不需再轉換。Agent 產出的 **content** 已是完整 HTML（含表格、標題、條列）。若畫面上有 **Content type**／**Format** 可選，請選 **HTML**。
- 若整頁只看到 HTML 原始碼：表示連接器把 content 當成純文字，請在該動作的設定或進階選項中查看是否有「Content type = HTML」。若表格有斷行或版面錯亂，多為 OneNote 對 HTML 的轉譯差異，可之後微調 Agent 產出的 HTML。

---

## 「Create page in a section」出現 BadGateway 或逾時

**BadGateway（502）** 多數是 OneNote／後端暫時性問題或請求逾時，可依序嘗試：

1. **稍後重跑**：過 5～10 分鐘再手動執行一次 Flow。
2. **逾時設定**：點開 **Create page in a section** → **設定**（Settings）→ 逾時（Timeout）可改為較長（例如 PT5M）。
3. **內容過大**：若任務與會議很多，content 很長，可先手動把 **Page content** 改成短 HTML（例如 `<p>test</p>`）測試；若這樣能成功，代表是內容大小或複雜度導致，可考慮在 Agent 端精簡 HTML 或分多頁。
4. **網路／權限**：確認 OneNote 連接器帳號對該 Notebook / Section 有寫入權限；公司 Proxy 有時也會造成閘道逾時。

---

## 條件：僅在 weekly_report.json 存在時執行

若希望沒有週報檔時不要報錯：

- 在「取得檔案內容（weekly_report）」前加 **條件**：僅當「列出資料夾中的檔案」結果裡存在 `Name` 等於 `weekly_report.json` 的項目時，才執行取得檔案內容 → Compose → Parse JSON → OneNote 建立頁面。
- 或將「取得檔案內容（weekly_report）」設為可選／略過失敗，後續 Parse JSON 與 OneNote 放在「若上一步成功」再執行。

---

## 流程順序摘要（含 Part B）

1. 手動觸發  
2. 取得 schedule_requests.json（或列出資料夾 → 篩選 → 取得內容）  
3. Compose 轉字串  
4. Parse JSON（schedule_requests，取得 `events` 陣列）  
5. 套用到每個 events → 建立 Outlook 事件  
6. **取得檔案內容** → **weekly_report.json**（同資料夾）  
7. **Compose** → 將內容轉成字串（見上方運算式）  
8. **Parse JSON** → 結構描述貼上 sectionId / title / content schema → 得到 **sectionId**、**title**、**content**  
9. **OneNote** → **建立頁面** → Section = sectionId、Page title = title、Content = content  

---

# 本機 .env 設定

Agent 寫入的檔案路徑由 **SCHEDULE_OUTPUT_FILE** 決定，建議與 **TASKS_INPUT_FILE** 同資料夾、檔名為 `schedule_requests.json`，這樣 OneDrive 同步後 Flow 讀同一資料夾即可。若要做週報，需設定 **ONENOTE_SECTION_ID**（要新增頁面的 OneNote section ID）。

**範例（OneDrive for Business，與 tasks_output.json 同資料夾）：**

```env
USE_POWER_AUTOMATE=true
TASKS_INPUT_FILE=C:\Users\你的帳號\OneDrive - 公司名\工作管理\GotPlannerTask\tasks_output.json
SCHEDULE_OUTPUT_FILE=C:\Users\你的帳號\OneDrive - 公司名\工作管理\GotPlannerTask\schedule_requests.json
ONENOTE_SECTION_ID=你的 OneNote section ID
```

**不要**設 `FLOW_CALENDAR_URL`（手動觸發版不需 HTTP URL）。

---

# 使用流程與注意事項

1. 執行：`py -X utf8 src/agent.py schedule-nextweek`（若要產出週報：先執行 `py -m src.agent weekly-report`，再執行 schedule-nextweek，或單獨執行 weekly-report 後手動執行 Flow）。
2. 確認輸出顯示「已將 N 筆行事曆事件寫入：...」。
3. 等 OneDrive 同步（若路徑在 OneDrive）。
4. 到 Power Automate 開啟本 Flow，點 **執行**／**Test**。
5. 到 Outlook 行事曆確認事件；若啟用 Part B，到 OneNote 確認週報頁面。

**注意：**

- 每次執行 `schedule-nextweek` 會**覆寫** schedule_requests.json；手動執行 Flow 會依**目前檔案內容**建立事件，不會自動刪除檔案。若怕重複建立，可在 Flow 最後加「刪除檔案」或改檔名（需自行在 Flow 內設計）。
- 若 **取得檔案內容** 失敗（檔案尚未同步或路徑錯誤），Flow 會報錯，請確認路徑與同步狀態。
