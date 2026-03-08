# Flow 2：建立 Outlook 行事曆事件 – 逐步設定

此 Flow 讓 Agent 在執行 `schedule-nextweek` 時，能將 Planner 任務一筆一筆寫入你的 Outlook 行事曆。  
**觸發方式**：收到 HTTP 要求（Agent 用 POST 傳入主旨、開始、結束時間與內文）。  
**不需使用 Response 動作**（付費功能）；Flow 只做「建立事件」後結束即可，Agent 會以 HTTP 狀態碼 200/202 視為成功。

---

## 步驟 1：建立新 Flow、選觸發

1. 開啟 **Power Automate** → **建立** → **即時雲端流程** → **從空白開始**。
2. 流程名稱可設為：`CreateOutlookEvent` 或 `建立行事曆事件`。
3. 在觸發條件搜尋 **request** 或 **HTTP**。
4. 選擇 **「當收到 HTTP 要求時」**（When a HTTP request is received）。
   - 若此觸發也顯示鑽石（付費），表示你的方案無法使用 HTTP 觸發，需改用其他方式（例如手動觸發 + 從檔案讀取要建立的事件）；目前先以「有 HTTP 觸發」為準。

---

## 步驟 2：設定要求本文（Request Body）的 JSON 結構

1. 點開觸發 **「當收到 HTTP 要求時」**。
2. 找到 **「要求本文 JSON 結構描述」**（Request body JSON schema）。
3. 貼上以下 JSON（讓 Flow 知道會收到 `subject`、`start`、`end`、`body`）：

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

4. 儲存後，下方會出現對應的 **body** 動態內容（subject、start、end、body），待會建立事件時會用到。

---

## 步驟 3：新增「建立事件」動作

1. 點 **+ 新增步驟**（或「+ New step」）。
2. 選擇 **動作**（Action）。
3. 搜尋 **Outlook** 或 **Office 365 Outlook**。
4. 選 **Office 365 Outlook**（或你使用的 Outlook 連接器）。
5. 在動作清單中選 **「建立事件」**（Create an event）或 **「排程會議」**（Schedule a meeting）。  
   - 若只有「排程會議」，也可用；差別在於會議可能會有會議連結，不影響 Agent 排工作時段。

---

## 步驟 4：對應觸發內容到行事曆欄位

在「建立事件」動作中，把 HTTP 要求傳進來的值填到對應欄位：

| 欄位（中文／英文） | 要填的值 |
|-------------------|----------|
| **主旨**（Subject） | 點選動態內容 **subject**（來自「當收到 HTTP 要求時」的 body），或 Expression：`triggerBody()?['subject']` |
| **開始**（Start） | 動態內容 **start**，或 `triggerBody()?['start']` |
| **結束**（End） | 動態內容 **end**，或 `triggerBody()?['end']` |
| **本文／內文**（Body） | 動態內容 **body**，或 `triggerBody()?['body']`（可選；Agent 會傳入 Planner 任務 ID 等說明） |

- **開始**、**結束**請維持 Agent 傳來的 **ISO 8601** 格式（例如 `2026-03-09T09:00:00+08:00` 或 `...Z`），Outlook 連接器通常可接受。
- 若畫面上有「時區」欄位，可選你的時區（例如 UTC+8）或留預設；開始/結束若已含時區，多數情況會正確解讀。

---

## 步驟 5：不要加「回應」動作

- **不要**新增 **Response**（回應）動作。  
- Flow 在「建立事件」完成後直接結束即可。  
- Agent 端會以 HTTP 狀態碼 **200** 或 **202** 視為成功；若你之後發現 Agent 報錯（例如解析回應本文失敗），可再調整 Agent 改為「只檢查狀態碼、不解析 JSON」。

---

## 步驟 6：儲存並取得 HTTP POST URL

1. 點右上角 **儲存**。
2. 儲存後，點回第一個步驟 **「當收到 HTTP 要求時」**。
3. 在畫面上會出現 **HTTP POST URL**（或「在收到 HTTP 要求時觸發」底下的 URL）。
4. **複製整個 URL**，準備貼到 `.env`。

---

## 步驟 7：設定 .env

在專案目錄的 `.env` 中新增或修改：

```env
FLOW_CALENDAR_URL=https://prod-xx.region.logic.azure.com:443/workflows/xxx/triggers/manual/paths/invoke?api-version=...
```

把上面的網址換成你剛複製的 **HTTP POST URL**。

- 若已有 `USE_POWER_AUTOMATE=true` 和 `TASKS_INPUT_FILE`，保留即可。
- 存檔後，Agent 在執行 `schedule-nextweek` 時會依序 POST 每筆任務到這個 URL，在 Outlook 建立對應事件。

---

## 步驟 8：測試

1. 在 Power Automate 可先用 **「以按鈕測試」** 或 **「執行流程」** 手動跑一次（需提供範例 body；若沒有手動測試，也可直接由 Agent 觸發）。
2. 在專案目錄執行：
   ```bash
   py -X utf8 src/agent.py schedule-nextweek
   ```
3. 到 Outlook 行事曆確認是否有新事件（會從你設定的「下週一」或下週第一個工作日起排）。

---

## 常見問題

**Q：觸發「當收到 HTTP 要求時」有鑽石，無法選？**  
A：表示你的 Power Automate 方案不支援 HTTP 觸發，需 Premium 或改用手動觸發 + 從檔案讀取要建立的事件清單（需另設計 Flow 與 Agent 寫檔格式）。

**Q：建立事件時開始/結束格式錯誤？**  
A：Agent 傳的是 ISO 8601（含 `Z` 或 `+08:00`）。若 Outlook 連接器要求不同格式，可在 Flow 裡用 **Expression** 轉換，例如 `formatDateTime(triggerBody()?['start'], 'yyyy-MM-ddTHH:mm:ssZ')`。

**Q：沒有 Office 365 Outlook，只有 Outlook.com？**  
A：可改用 **Outlook.com** 連接器的「建立事件」，欄位對應方式相同（subject、start、end、body）。
