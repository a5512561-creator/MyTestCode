# 從方式 C 的 JSON 取 PLAN_ID，以及 403 的解法

## 哪一個才是 PLAN_ID？

在方式 C（Power Automate「列出群組的計畫」）的回應裡：

- **`body.value[0].id`** = `"_YIMOMgqSEiTZd0FdWnnMcgACG9P"` → 這才是 **計畫 ID**（Power Automate / 內部 API 用的格式）。
- `request-id`、`client-request-id` 是當次請求的追蹤用 ID，不是計畫 ID。
- `owner`、`container.containerId` = `0faa7f6e-651a-49a9-a711-e984c482aff0` 是 **群組 ID**，不是計畫 ID。

所以請在 `.env` 使用：

```env
PLAN_ID=_YIMOMgqSEiTZd0FdWnnMcgACG9P
```

---

## 為什麼會 403？

你用 **應用程式 (client secret)** 登入時，是「應用程式」在呼叫 Graph，應用程式本身沒有被授權存取該 Planner 計畫，所以會 403。

Power Automate 是以 **你的帳號** 去呼叫，所以能成功；Agent 若也改成 **用你的帳號登入**，有機會用同一個計畫 ID 通過。

---

## 作法：改用「你的帳號」登入（裝置碼）

1. 在 `.env` 設定：
   ```env
   PLAN_ID=_YIMOMgqSEiTZd0FdWnnMcgACG9P
   DEVICE_CODE_FLOW=true
   ```
2. 把 **CLIENT_SECRET** 註解掉（前面加 `#`），讓程式改用裝置碼、以你的帳號取 token：
   ```env
   # CLIENT_SECRET=原本的值
   ```
3. 存檔後執行：
   ```powershell
   py test_step_by_step.py step2
   ```
4. 依畫面在瀏覽器完成登入，再看是否還會 403。

若這樣就通過，之後跑 Planner 相關功能時都維持 **DEVICE_CODE_FLOW=true** 且註解 CLIENT_SECRET 即可。若仍 403，就需請 IT 用 Graph 總管查一次，取得 **GUID 格式** 的計畫 ID 再填回 `.env`。
