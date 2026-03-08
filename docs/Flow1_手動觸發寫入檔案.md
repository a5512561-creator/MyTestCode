# Flow 1 手動觸發 + 寫入檔案（只列「指給我的」任務）

只關心**指派給你的**任務時，用 **List my tasks**；Flow 把結果寫入 OneDrive，Agent 讀取並整理顯示。

---

## 一、Power Automate 步驟與參數（逐項設定）

### 步驟 1：觸發

- **觸發**：**Manually trigger a flow**（手動觸發流程）
- 不需填任何參數，儲存即可。

---

### 步驟 2：只列「我的」任務

- 新增步驟 → **Action** → 搜尋 **Planner** → 選 **Planner**。
- 動作選 **「List my tasks」**（或 **List tasks** 底下有「我的任務」的選項）。
  - 若畫面上沒有「List my tasks」，可用 **「List tasks for a plan」**，再在 Agent 端篩選；但建議優先找 **List my tasks**。
- **參數**：
  - **List my tasks** 通常**不需要**填 Plan id，會自動回傳「指派給目前使用者」的任務。
  - 若有欄位問 Group / Plan，可留空或選你的計畫（依畫面為準）。
- 此步驟輸出會有一個 **value**（陣列），裡面是任務清單。

---

### 步驟 3：列 bucket（為了顯示「本週工作」「進行中」等名稱）

- 新增步驟 → **Action** → **Planner**。
- 動作選 **「List buckets for a plan」**。
- **參數**：
  - **Plan id**：填你的計畫 ID，例如 `_YIMOMgqSEiTZd0FdWnnMcgACG9P`（與你 Create a task 用的相同）。
- 此步驟輸出的 **value** 是 bucket 清單（id、name），供 Agent 對照顯示欄位名稱。

---

### 步驟 4：Compose（組裝成一個 JSON 給檔案用）

- 新增步驟 → 搜尋 **Compose** → 選 **Compose**（Built-in）。
- 點 **Inputs** 欄位（或「輸入」）。

**做法 A：用鍵值 + 動態內容（推薦）**

1. 在 Inputs 區塊點 **「Add new parameter」** 或 **「新增參數」** → 選 **Inputs**（若已是 Inputs 則略過）。
2. 若出現 **鍵 / 值** 兩欄：
   - 第一組：**鍵** 填 `tasks`，**值** 點框內 → **Dynamic content** → 找到 **List my tasks**（或你步驟 2 的名稱）→ 選 **value**（整包陣列）。
   - 第二組：再 **Add new parameter** → **鍵** 填 `buckets`，**值** 選 **List buckets for a plan** 的 **value**。
3. 儲存後，Compose 的 **Output** 會是 `{ "tasks": [...], "buckets": [...] }`。

**做法 B：用 Expression（若沒有鍵值可選）**

1. 在 **Inputs** 右側點 **fx**（Expression）。
2. 貼上以下其中一行（依你步驟名稱改單引號裡的字）：
   - 若步驟 2 叫 **List my tasks**：
     ```text
     json(concat('{"tasks":', string(body('List_my_tasks')?['value']), ',"buckets":', string(body('List_buckets_for_a_plan')?['value']), '}'))
     ```
   - 若步驟 2 叫 **List tasks**：
     ```text
     json(concat('{"tasks":', string(body('List_tasks')?['value']), ',"buckets":', string(body('List_buckets_for_a_plan')?['value']), '}'))
     ```
3. 按 **OK**。  
   （步驟名稱有空格時，在運算式裡通常會變成底線，例如 `List buckets for a plan` → `List_buckets_for_a_plan`。不確定時可在 fx 裡打 `body('` 看自動完成的選單。）

---

### 步驟 5：寫入 OneDrive 檔案

- 新增步驟 → 搜尋 **OneDrive** 或 **OneDrive for Business**。
  - **個人 OneDrive**：選 **OneDrive**（或 **OneDrive (Consumer)**）。
  - **公司 OneDrive**：選 **OneDrive for Business**。
- 動作選 **Create file**（建立檔案）。

**參數建議：**

| 參數 | 要填什麼 |
|------|----------|
| **Site Address** | 僅 OneDrive for Business 有；個人 OneDrive 通常無此欄。 |
| **Folder Path** | 雲端路徑，例如：`/PlannerAgent`（在 OneDrive 根目錄下建立 PlannerAgent）。 |
| **File Name** | `tasks_output.json`。 |
| **File Content** | 選上一步 **Compose** 的 **Output**。若要求「文字」：Expression 填 `string(outputs('Compose'))`。 |

- 儲存 Flow 後手動執行一次，確認 OneDrive 雲端出現 `PlannerAgent\tasks_output.json`，本機同步後再填 .env（見下方）。

---

## 二、本機 .env 設定

**TASKS_INPUT_FILE** 要填的是「本機電腦上」那個檔案的完整路徑（OneDrive 同步下來後的實際路徑），不是雲端網址。

### 個人 OneDrive（非 OneDrive for Business）

個人 OneDrive 在本機的同步資料夾通常會是以下其中一種（依你當初安裝時的名稱而定）：

- `C:\Users\你的使用者名稱\OneDrive\PlannerAgent\tasks_output.json`
- `C:\Users\你的使用者名稱\OneDrive - Personal\PlannerAgent\tasks_output.json`

**如何確認你的路徑：**

1. 在 **檔案總管** 左側或「本機」底下找到 **OneDrive**（或 **OneDrive - Personal**）。
2. 點進去，確認有 **PlannerAgent** 資料夾與 **tasks_output.json**（需先跑過一次 Flow 才會出現）。
3. 在 **tasks_output.json** 上按 **右鍵** → **內容** → 看 **「位置」** 或 **「一般」** 裡的完整路徑；或按住 **Shift** 再右鍵該檔案 → **複製為路徑**，貼到記事本即為完整路徑。
4. 把該路徑填到 .env 的 **TASKS_INPUT_FILE**。

**.env 範例（個人 OneDrive）：**

```env
USE_POWER_AUTOMATE=true
TASKS_INPUT_FILE=C:\Users\paddychen\OneDrive\PlannerAgent\tasks_output.json
```

若你的 OneDrive 資料夾叫 **OneDrive - Personal**，則改成：

```env
TASKS_INPUT_FILE=C:\Users\paddychen\OneDrive - Personal\PlannerAgent\tasks_output.json
```

- 路徑不要加引號；反斜線 `\` 保留即可（或改為 `/` 也可）。
- 不要設 **FLOW_PLANNER_URL**（手動觸發不需 HTTP URL）。

### OneDrive for Business（公司用）

若 Flow 用的是 **OneDrive for Business**，你在網頁上建立的資料夾（例如 **我的檔案 > 工作管理 > GotPlannerTask**）會同步到本機的「OneDrive - 公司名」資料夾。

**本機路徑長相通常為：**

```text
C:\Users\你的使用者名稱\OneDrive - Realtek\工作管理\GotPlannerTask\tasks_output.json
```

（把 **Realtek** 改成你公司 OneDrive 顯示的名稱，例如 `OneDrive - 公司名`。）

**如何確認：**

1. 在 **檔案總管** 左側找到 **OneDrive - Realtek**（或你的公司 OneDrive 名稱）。
2. 點進去，依序進入 **工作管理** → **GotPlannerTask**。
3. 等 Flow 寫入後會出現 **tasks_output.json**；對該檔案 **Shift + 右鍵** → **複製為路徑**，即為 TASKS_INPUT_FILE 要填的值。

**.env 範例（OneDrive for Business，資料夾 工作管理 > GotPlannerTask）：**

```env
USE_POWER_AUTOMATE=true
TASKS_INPUT_FILE=C:\Users\paddychen\OneDrive - Realtek\工作管理\GotPlannerTask\tasks_output.json
```

**Flow 步驟 5 的 Folder Path：** 在 OneDrive for Business 建立檔案時，**Folder Path** 請填雲端路徑，例如：`/工作管理/GotPlannerTask`，**File Name** 填 `tasks_output.json`，這樣檔案會出現在「工作管理 > GotPlannerTask」底下，同步到本機後路徑即為上式。

---

## 三、使用流程

1. 在 Power Automate 手動執行此 Flow，讓它寫入 `tasks_output.json`。
2. 等幾秒讓 OneDrive 同步到本機。
3. 在專案目錄執行：`py -m src.agent planner`  
   Agent 會讀 **TASKS_INPUT_FILE**，只顯示「指給你的」任務，並依到期日、優先順序與 bucket（本週工作、進行中）整理。

---

## 四、若沒有「List my tasks」

- 若你的 Planner 連接器只有 **「List tasks for a plan」**：
  - 仍用 **List tasks for a plan**（Plan id 填你的計畫 ID），會得到「計畫內全部任務」。
  - Agent 目前不會依「指派給誰」篩選，你會看到計畫裡所有任務；之後若要只顯示「我的」，可再在 Agent 加篩選（需任務裡有 assignments 欄位）。
- 有 **List my tasks** 時，優先使用，才會只關心指給你的工作。
