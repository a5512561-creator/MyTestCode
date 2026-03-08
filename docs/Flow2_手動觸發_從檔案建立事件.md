# Flow 2（手動觸發版）：從檔案建立 Outlook 行事曆事件

公司 Power Automate **沒有 HTTP 觸發**（需 Premium）時，改用此方式：Agent 把要建立的事件寫成 JSON 檔，你手動執行此 Flow，Flow 讀取該檔後依序在 Outlook 建立事件。

---

## 流程概覽

1. 本機執行 `py -X utf8 src/agent.py schedule-nextweek` → Agent 寫入 **schedule_requests.json**（路徑由 .env 的 `SCHEDULE_OUTPUT_FILE` 決定）。
2. 若該路徑在 OneDrive 資料夾內，等同步完成。
3. 在 Power Automate **手動執行**本 Flow → Flow 從 OneDrive 讀取該檔，依序「建立事件」。
4. 到 Outlook 行事曆確認。

---

## 步驟 1：觸發

1. 新增即時雲端流程，名稱可設為「從檔案建立 Outlook 事件」。
2. 觸發選 **手動觸發流程**（Manually trigger a flow）。
3. 不需新增輸入參數。

---

## 步驟 2：取得檔案內容（OneDrive）

1. 新增步驟 → **動作** → 搜尋 **OneDrive for Business**（或你使用的 OneDrive 連接器）。
2. 選擇 **取得檔案內容**（Get file content）。
   - 若沒有「取得檔案內容」，可用 **列出資料夾中的檔案** 指定資料夾後，再用 **取得檔案內容** 並選取檔名（見下方替代作法）。
3. 參數：
   - **檔案識別碼**（File Identifier）：若連接器要求的是「識別碼」，需先在前一步用 **列出資料夾中的檔案** 取得該檔的識別碼，再選動態內容。
   - 或若連接器有 **「依路徑取得檔案內容」**／**「Get file content using path」**：
     - **Site Address**：留預設或你的 SharePoint 網站。
     - **File Path**：雲端路徑，須與 Agent 寫入的檔案一致。例如與 `tasks_output.json` 同資料夾時：  
       ` /工作管理/GotPlannerTask/schedule_requests.json`  
       （依你在 OneDrive 的實際路徑調整。）

**若只有「列出資料夾中的檔案」+「取得檔案內容」：**

1. 先加 **列出資料夾中的檔案**（List files in folder）：
   - **Folder**：選你的 OneDrive 資料夾，例如 `工作管理/GotPlannerTask` 或從動態內容選。
2. 再加 **取得檔案內容**：
   - **File**：從動態內容選 **列出資料夾中的檔案** 的 **Id**（需用「套用到每個」見下方），或若可選 **Name** 且能指定檔名則選 `schedule_requests.json`。

**若需依「檔名」找檔案：** 在「列出資料夾中的檔案」後加 **篩選陣列**（Filter array），條件為 `item()?['Name']` 等於 `schedule_requests.json`，再對篩選結果用「套用到每個」跑「取得檔案內容」。

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

   **若仍失敗，改試直接轉字串（部分連接器會回傳可轉字串的內容）：**
   ```text
   string(body('Get_file_content'))
   ```

3. 儲存後，下一步 **Parse JSON** 的 **內容** 改為選這個 **Compose** 的 **Output**（不要再用「取得檔案內容」的檔案內容）。

**若仍出現「content must be of type JSON, but was of type 'application/octet-stream'」：**

- 表示 **Parse JSON** 的 **Content** 仍接到二進位。請點開 **Parse JSON**，在 **Content** 欄位：
  - **刪掉**目前的值（例如「取得檔案內容」的「File content」）。
  - 改選 **動態內容** → 找到 **Compose**（步驟 3a）→ 選 **Output**。
- 確認 **Compose** 的 **Inputs** 是 **運算式** `base64ToString(body('Get_file_content')?['$content'])`，不是直接選「File content」；若直接選 File content，Compose 會把二進位原樣傳出，Parse JSON 仍會報錯。

---

## 步驟 3b：剖析 JSON（取得 events 陣列）

1. 新增步驟 → **動作** → 搜尋 **Parse JSON** 或 **剖析 JSON**。
2. **內容**（Content）：**必須**選 **Compose**（步驟 3a）的 **Output**。不可選「取得檔案內容」的「File content」，否則會出現 application/octet-stream 錯誤。
3. **結構描述**（Schema）：在該欄位貼上以下 JSON（或點「使用範例產生結構描述」後貼上整段 `{"events":[...]}` 範例讓系統自動產生）：

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
   - **主旨**：`subject`（目前項目的 subject）。
   - **開始**：`start`。
   - **結束**：`end`。
   - **本文**：`body`。

   （若在迴圈內，動態內容會顯示「目前項目」→ subject、start、end、body。）

5. 儲存 Flow。

---

## 步驟 5：處理「列出資料夾」+「取得檔案內容」的作法（無「依路徑取得」時）

若你的 OneDrive 連接器沒有「依路徑取得檔案內容」，可用：

1. **列出資料夾中的檔案**：Folder = `工作管理/GotPlannerTask`（或你存 schedule_requests.json 的資料夾）。
2. **套用到每個**（外層）：選「列出資料夾中的檔案」的 **value**。
3. **條件**：目前項目的 **Name** 等於 `schedule_requests.json`（若不符合則略過）。
4. **若為是**：**取得檔案內容**，File = 目前項目的 **Id**（或能指向該檔的欄位）。
5. 再接 **剖析 JSON** → 內層 **套用到每個** events → **建立事件**（同上）。

---

## 本機 .env 設定

Agent 寫入的檔案路徑由 **SCHEDULE_OUTPUT_FILE** 決定，建議與 **TASKS_INPUT_FILE** 同資料夾、檔名為 `schedule_requests.json`，這樣 OneDrive 同步後 Flow 讀同一資料夾即可。

**範例（OneDrive for Business，與 tasks_output.json 同資料夾）：**

```env
USE_POWER_AUTOMATE=true
TASKS_INPUT_FILE=C:\Users\你的帳號\OneDrive - 公司名\工作管理\GotPlannerTask\tasks_output.json
SCHEDULE_OUTPUT_FILE=C:\Users\你的帳號\OneDrive - 公司名\工作管理\GotPlannerTask\schedule_requests.json
```

**不要**設 `FLOW_CALENDAR_URL`（手動觸發版不需 HTTP URL）。

---

## 使用流程

1. 執行：`py -X utf8 src/agent.py schedule-nextweek`
2. 確認輸出顯示「已將 N 筆行事曆事件寫入：...」
3. 等 OneDrive 同步（若路徑在 OneDrive）。
4. 到 Power Automate 開啟本 Flow，點 **執行**／**Test**。
5. 到 Outlook 行事曆確認事件是否建立。

---

## 注意事項

- 每次執行 `schedule-nextweek` 會**覆寫** schedule_requests.json；手動執行 Flow 會依**目前檔案內容**建立事件，不會自動刪除檔案。若怕重複建立，可在 Flow 最後加「刪除檔案」或改檔名（需自行在 Flow 內設計）。
- 若 **取得檔案內容** 失敗（例如檔案尚未同步或路徑錯誤），Flow 會報錯，請確認路徑與同步狀態。
