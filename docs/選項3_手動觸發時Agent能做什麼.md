# 選項 3：手動觸發 Flow 時，Agent 能做什麼？

當你使用**手動觸發**（Button flow）或**排程觸發**、沒有 HTTP 觸發（無 Premium）時，Python Agent **無法直接呼叫** Flow，但可以擔任以下角色。

---

## 1. 提醒與指引（已內建）

執行：

```powershell
py -m src.agent guide
```

Agent 會印出**建議順序**與待辦提醒，例如：

1. 先執行「GotPlannerTasks」Flow，查看本週任務。
2. 再執行「排入行事曆」Flow 或手動排程。
3. 每週三執行「OneNote 週報」Flow 或手動整理。

你只要照著順序在 Power Automate 裡手動跑對應的 Flow 即可。Agent 的角色是**檢查清單 + 提醒**。

---

## 2. 把多個步驟包成「一鍵提醒」

可以在桌面或快速啟動建立一個捷徑／批次檔，內容就是：

```powershell
py -m src.agent guide
```

每次雙擊就打開終端機並顯示「今天要做的三步驟」，你再依序去 Power Automate 執行 Flow。

---

## 3. 若 Flow 會把結果寫到檔案（進階）

若你調整 Flow，讓它把手動觸發後的結果寫到某個地方，Agent 可以再處理：

- **寫到 OneDrive 檔案**：例如 Flow 把「本週任務」寫成 `tasks.json` 到 OneDrive，你本機同步後，可寫一支小腳本（或擴充 Agent）讀取該檔案，在終端機顯示、過濾或排序。Agent 就負責「讀取 + 顯示/建議」。
- **寫到 SharePoint 清單**：同上，若 Flow 寫入清單，之後可考慮用 SharePoint API 或匯出檔讓 Agent 讀取並做整理/報表。

這部分需要你在 Flow 裡先加上「寫入檔案/清單」的步驟，再在專案裡加對應的讀取邏輯。

---

## 4. 小結

| 方式 | Agent 做的事 |
|------|----------------|
| **guide** | 印出「先跑哪個 Flow、再跑哪個」的順序與提醒（已實作）。 |
| 捷徑／批次 | 一鍵執行 `agent guide`，當成每日/每週檢查清單。 |
| Flow 寫出檔案 | 未來可擴充：Agent 讀取 Flow 產生的檔案並顯示或過濾。 |

目前你可以直接使用：

```powershell
py -m src.agent guide
```

做為手動觸發時的「要做什麼、依什麼順序」指引。
