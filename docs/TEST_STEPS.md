# 逐步測試說明

## 測試前準備

1. **複製設定檔**（在專案根目錄執行）
   ```powershell
   Copy-Item config.example.env .env
   ```
2. **編輯 `.env`**，至少填入：
   - `TENANT_ID`、`CLIENT_ID`（Azure 應用程式註冊）
   - `CLIENT_SECRET` 或改用 `DEVICE_CODE_FLOW=true`（本機用裝置碼登入）
   - 測試 Planner 時需要 `PLAN_ID`
   - 測試 OneNote 時需要 `ONENOTE_SECTION_ID`

3. **安裝依賴**
   ```powershell
   py -m pip install -r requirements.txt
   ```

4. **執行逐步測試**
   ```powershell
   py test_step_by_step.py step1   # 認證
   py test_step_by_step.py step2   # Planner（需 PLAN_ID）
   py test_step_by_step.py step3   # Outlook 空檔
   py test_step_by_step.py step4   # OneNote（需 ONENOTE_SECTION_ID）
   py test_step_by_step.py all     # 依序執行
   ```

---

## 步驟 1：測試認證

確認能取得 Microsoft Graph 的 access token。

```bash
python test_step_by_step.py step1
```

- **成功**：顯示「已取得 access token」。
- **失敗**：檢查 Azure 應用程式註冊、redirect URI、API 權限（User.Read, Tasks.ReadWrite, Calendars.ReadWrite, Notes.ReadWrite）。

---

## 步驟 2：測試 Planner

確認能讀取計畫並列出「本週工作」「進行中」的任務。

```bash
python test_step_by_step.py step2
```

- 需先在 `.env` 設定 `PLAN_ID`（可從 Planner URL 或 Graph 查詢取得）。
- **成功**：列出符合條件的任務。
- **失敗**：檢查 PLAN_ID、權限 Tasks.Read。

---

## 步驟 3：測試 Outlook 空檔

只查詢行事曆空檔，不建立事件。

```bash
python test_step_by_step.py step3
```

- **成功**：顯示本週空檔數量與前幾筆時段。
- **失敗**：檢查權限 Calendars.Read 或 Calendars.ReadWrite。

---

## 步驟 4：測試 OneNote 週報

在指定 section 建立一頁週報。

```bash
python test_step_by_step.py step4
```

- 需在 `.env` 設定 `ONENOTE_SECTION_ID`（從 Graph 查 notebook/sections 取得）。
- **成功**：顯示新頁面 URL。
- **失敗**：檢查 ONENOTE_SECTION_ID、權限 Notes.ReadWrite。

---

## 一次跑完全部

```bash
python test_step_by_step.py all
```

會依序執行 step1 → step2 → step3 → step4，任一步失敗即停止。

---

## 使用 Agent 主程式

通過上述測試後，可改用主程式：

```bash
python -m src.agent planner    # 列出任務
python -m src.agent schedule   # 排入行事曆
python -m src.agent onenote    # 寫入週報
```
