# Flow 1 (GotPlannerTasks) – Step-by-step (English UI)

**Manual trigger + write to file** – no HTTP trigger and no Response action (works without Power Automate Premium).

Use this flow to export **tasks assigned to you** from Planner into a JSON file on OneDrive. The Agent reads that file via `TASKS_INPUT_FILE`.

---

## Step 1: Trigger

1. Create a **new flow** (Automate → My flows → New flow → Instant cloud flow).
2. For the trigger, search for **Manually trigger a flow**.
3. Select **Manually trigger a flow** (under **Built-in** or **Trigger**).
4. You do not need to add any trigger parameters. Click **+ New step** to continue.

---

## Step 2: List your tasks (Planner)

1. Click **+ New step**.
2. Choose **Action**.
3. Search for **Planner** (or **Microsoft Planner**) and select the **Planner** connector.
4. Select the action **List my tasks** (only tasks assigned to you).
   - If you do not see **List my tasks**, use **List tasks for a plan** instead and fill **Plan id** in the next step.
5. **List my tasks** usually does **not** require a Plan id; it returns tasks assigned to the current user.
   - If there is a Group/Plan field, leave it empty or pick your plan as needed.
6. This step outputs a **value** (array) – that is your task list. Remember the step name (e.g. **List_my_tasks**).

---

## Step 3: List buckets (for column names like “本週工作”, “進行中”)

1. Click **+ New step**.
2. **Action** → search **Planner** → **Planner**.
3. Select **List buckets for a plan**.
4. **Plan id**: Enter your plan ID (e.g. `_YIMOMgqSEiTZd0FdWnnMcgACG9P` – the same value you use in “Create a task” for this plan).
5. This step outputs **value** (array of buckets with id and name). Remember the step name (e.g. **List_buckets_for_a_plan**).

---

## Step 4: Compose – build JSON for the file

1. Click **+ New step**.
2. Search for **Compose** and select **Compose** (Built-in).
3. Click in the **Inputs** field.

**Option A: Add parameters + Dynamic content (recommended)**

1. In the Inputs box, click **Add new parameter** (or **Add dynamic content**).
2. Add two parameters:
   - First: **Key** = `tasks`, **Value** = open **Dynamic content** → find **List my tasks** (or your Step 2 name) → select **value** (the whole array).
   - Second: **Add new parameter** again → **Key** = `buckets`, **Value** = **List buckets for a plan** → **value**.
3. The Compose output will be `{ "tasks": [...], "buckets": [...] }`.

**Option B: Expression (if key/value is not available)**

1. Click the **fx** (Expression) icon next to **Inputs**.
2. Paste one of these (replace the step names in quotes if yours are different):
   - If Step 2 is **List my tasks**:
     ```text
     json(concat('{"tasks":', string(body('List_my_tasks')?['value']), ',"buckets":', string(body('List_buckets_for_a_plan')?['value']), '}'))
     ```
   - If Step 2 is **List tasks for a plan**:
     ```text
     json(concat('{"tasks":', string(body('List_tasks_for_a_plan')?['value']), ',"buckets":', string(body('List_buckets_for_a_plan')?['value']), '}'))
     ```
3. Click **OK**.  
   (Step names with spaces usually become underscores in expressions, e.g. `List buckets for a plan` → `List_buckets_for_a_plan`. Type `body('` in the expression box to see autocomplete.)

---

## Step 5: Create file in OneDrive

1. Click **+ New step**.
2. Search for **OneDrive** or **OneDrive for Business**.
   - Personal: **OneDrive** (or **OneDrive (Consumer)**).
   - Work: **OneDrive for Business**.
3. Select the action **Create file**.

**Parameters:**

| Parameter        | What to enter |
|------------------|----------------|
| **Site Address** | Only for OneDrive for Business; leave default or as required. |
| **Folder Path**  | Cloud folder path, e.g. `/PlannerAgent` or `/工作管理/GotPlannerTask`. |
| **File Name**    | `tasks_output.json` |
| **File Content** | Select the **Output** of the **Compose** step. If it asks for text, use Expression: `string(outputs('Compose'))`. |

4. Click **Save** (top right).

---

## Step 6: Test the flow

1. Click **Test** → **Manually** → **Run flow**.
2. When it finishes, open OneDrive (web or synced folder) and confirm the file exists, e.g. `PlannerAgent\tasks_output.json` or `工作管理\GotPlannerTask\tasks_output.json`.
3. Wait for sync to your PC if you use the OneDrive desktop app.

---

## Step 7: Set .env on your PC

**TASKS_INPUT_FILE** must be the **local path** to that file (after OneDrive sync), not a cloud URL.

**OneDrive for Business example** (folder 工作管理 > GotPlannerTask):

```env
USE_POWER_AUTOMATE=true
TASKS_INPUT_FILE=C:\Users\YourUsername\OneDrive - YourCompany\工作管理\GotPlannerTask\tasks_output.json
```

**Personal OneDrive example:**

```env
USE_POWER_AUTOMATE=true
TASKS_INPUT_FILE=C:\Users\YourUsername\OneDrive\PlannerAgent\tasks_output.json
```

- Do **not** set **FLOW_PLANNER_URL** when using manual trigger (Agent will read from the file).
- To get the exact path: in File Explorer, go to the file → **Shift + right‑click** → **Copy as path**, then paste into .env.

---

## Step 8: Run the Agent

1. In Power Automate, run this flow so it writes the latest `tasks_output.json`.
2. After sync, in your project folder run:

```powershell
py -X utf8 src/agent.py planner
```

Or for the next-week report:

```powershell
py -X utf8 src/agent.py nextweek
```

The Agent reads **TASKS_INPUT_FILE** and shows/filters tasks by bucket (e.g. 本週工作, 進行中) and due date.

---

## If you don’t have “List my tasks”

- Use **List tasks for a plan** in Step 2 and enter your **Plan id**. You will get all tasks in the plan; the Agent does not filter by assignee in that case.
- Prefer **List my tasks** when available so only your assigned tasks are exported.

---

## If you don’t have “List buckets for a plan”

- Skip Step 3.
- In Step 4 (Compose), set **buckets** to an empty array: in the expression, use `"[]"` for the buckets part, e.g.  
  `json(concat('{"tasks":', string(body('List_my_tasks')?['value']), ',"buckets":[]}'))`.  
  The Agent will still run; bucket names may show as blank.
