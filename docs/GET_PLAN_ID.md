# 如何取得 Planner 計畫 ID

---

## 若貴公司要求管理員同意才能用 Graph（例如 realtek.com）

在這種情況下，**不要用 API／腳本**，改用下面任一方式取得計畫 ID 後，手動填到 `.env` 的 `PLAN_ID=`。

### 方式 A：從 tasks.office.com 網址取得（建議先試）

1. 用瀏覽器開啟 **https://tasks.office.com**
2. 登入後，從左側或「我的 day」進入 **Planner**，點進你要用的那一個計畫。
3. 看瀏覽器網址列，若網址裡出現 **`planId=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`**，`=` 後面那一串就是 **PLAN_ID**，複製貼到 `.env` 即可。
4. 若 tasks.office.com 只看到 tid、沒有 planId，或會自動跳轉到 planner.cloud.microsoft，請改用方式 B。

### 方式 B：從瀏覽器開發者工具取得（不需 IT）

Planner 網頁載入時會向 Graph 發請求，回應裡常有計畫的 **GUID**。用你的帳號打開計畫頁面即可，不需管理員權限。

1. 用 Chrome 或 Edge 開啟 **https://planner.cloud.microsoft**，登入後點進你要用的那個計畫（看板畫面）。
2. 按 **F12** 打開開發者工具，切到 **Network**（網路）分頁。
3. 在篩選框輸入 **`planner`** 或 **`graph`**，只顯示相關請求。
4. 重新整理頁面（F5），或切換到其他分頁再切回看板，讓頁面發送請求。
5. 在請求清單中點選看起來像 **plans** 或 **planner** 的請求（URL 可能含 `planner/plans` 或 `planner/tasks`）。
6. 在右側點 **Response**（回應）或 **Preview**，在 JSON 裡找 **`"id"`**，值為 **xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx**（GUID 格式）的那個就是 **PLAN_ID**。  
   - 若回應是陣列（`value: [...]`），就從陣列裡某個物件的 `id` 複製。
7. 把該 GUID 貼到 `.env` 的 `PLAN_ID=`。

若找不到含 plan 的請求，可試著在 Network 篩選輸入 **`v1.0`**，再逐一點開檢查回應內容是否有 `"id"`（GUID）。

### 方式 C：用 Power Automate 取得 GUID（不需 IT）

你已有 Power Automate 與「Create a task」，可用類似方式取得 **Graph 用的計畫 ID**：

1. 新增一個 **手動觸發** 的 Flow。
2. 新增動作，選 **Planner** 連接器。
3. 若有 **「List plans for a group」** 或 **「列出群組的計畫」** 之類的動作，選它；**Group id** 填你已知的 `0faa7f6e-651a-49a9-a711-e984c482aff0`。
4. 儲存並執行一次該 Flow。
5. 在執行結果中點開該動作的 **輸出**，查看 JSON；其中應有 **`id`** 欄位，值為 GUID（如 `a1b2c3d4-e5f6-7890-abcd-ef1234567890`），即為 **PLAN_ID**，複製貼到 `.env` 的 `PLAN_ID=`。

若沒有「List plans for a group」，可試 **「Get plan」** 或 **「取得計畫」**（若有欄位要選計畫，就選你的計畫），從輸出裡找 `id`（GUID）。

### 方式 D：請 IT／管理員代查一次

若上述方式 B、C 都無法取得，再使用。已備好可轉寄的說明檔：**`docs/請IT協助取得計畫ID.txt`**（複製內容轉寄給 IT 即可）。

請有 **Group.Read.All** 或該群組管理員權限的同事幫忙：

1. 在 Graph 總管或 Postman 呼叫（需他們的帳號或應用程式 token）：
   ```http
   GET https://graph.microsoft.com/v1.0/groups/611bbe5b-46a9-4bb3-9580-026cd900033f/planner/plans
   ```
2. 回應裡會有一個 `"id"`，就是該群組底下的 **計畫 ID**，請他們把這串 ID 給你。
3. 你貼到 `.env` 的 `PLAN_ID=`，之後 Agent 只會用這個 ID 讀寫任務，**不需要** 再要任何 Group 權限。

（Power Automate 作法已併入上方方式 C。）

---

## 方法一：用專案腳本取得（僅在「不需管理員同意」的環境）

若貴公司 **沒有** 要求管理員同意 Graph 存取 Planner，才可用此腳本。

在專案根目錄執行：

```powershell
py scripts/get_plan_id.py
```

1. 會用「你的帳號」裝置碼登入。
2. 呼叫 `GET /me/planner/tasks`（只需 Tasks.ReadWrite），從「分配給你的任務」反查所屬的計畫。
3. 列出你有權限的計畫與 **PLAN_ID**，任選一個複製到 `.env` 的 `PLAN_ID=` 即可。

**注意**：Planner 裡至少要有一項任務是「分配給你」的，腳本才能反查到該計畫；若完全沒有，請先在 Planner 指派一項任務給自己再執行。

---

## 方法二：用 Microsoft Graph 總管（你的帳號登入）

1. **開啟 Graph 總管**  
   https://developer.microsoft.com/zh-tw/graph/graph-explorer

2. **登入**  
   點右上角 **「登入」**，用你的 Microsoft 365 帳號登入（要有該 Planner 的存取權）。

3. **若出現 400 Bad Request：先新增權限**  
   點 **「Modify permissions」**（修改權限）分頁 → 勾選 **Tasks.ReadWrite**、**Group.Read.All** → 按 **Consent**／同意，再回到查詢重試。

4. **查詢你的 Planner 計畫**  
   在請求 URL 輸入（二選一）：

   **選項 A：列出「與我共用的」所有計畫**（較簡單）
   ```
   https://graph.microsoft.com/v1.0/me/planner/plans
   ```
   按 **「執行查詢」**。

   **選項 B：若你知道群組 ID（Planner 網址的 tid=）**
   ```
   https://graph.microsoft.com/v1.0/groups/你的群組ID/planner/plans
   ```
   例：`https://graph.microsoft.com/v1.0/groups/611bbe5b-46a9-4bb3-9580-026cd900033f/planner/plans`  
   按 **「執行查詢」**。  
   （若出現 403，表示此帳號仍需 Group 權限，請用選項 A。）

5. **從回應取得計畫 ID**  
   回應為 JSON，例如：
   ```json
   {
     "value": [
       {
         "id": "xYz123AbC-1234-5678-9abc-def012345678",
         "title": "我的工作計畫",
         ...
       }
     ]
   }
   ```
   複製 **`id`** 的值（一串 GUID），那就是 **計畫 ID**。

6. **寫入 .env**  
   在專案根目錄的 `.env` 中設定：
   ```env
   PLAN_ID=xYz123AbC-1234-5678-9abc-def012345678
   ```
   （可註解或刪除 GROUP_ID，避免程式再去查群組。）

7. **再跑 step2**  
   ```powershell
   py test_step_by_step.py step2
   ```

---

## 若選項 A 回傳空陣列

表示 Graph 總管目前沒有列出任何「與你共用的」計畫。可以：

- 確認已用「有該 Planner 存取權」的帳號登入 Graph 總管。  
- 或在 Planner 網頁確認該計畫是否與你的帳號共用／你是成員。  
- 若選項 B 可成功，就直接用選項 B 的回應裡的 `id` 當 PLAN_ID。
