# 將專案上傳到 GitHub（版本控管）

目前本機若尚未安裝 Git 或尚未初始化儲存庫，請依下列步驟操作。

---

## 1. 安裝 Git（若尚未安裝）

1. 前往 https://git-scm.com/download/win 下載 Windows 版 Git。
2. 安裝時可保留預設選項（建議勾選「Add Git to PATH」）。
3. 安裝完成後**重新開啟**終端機或 Cursor。

---

## 2. 在 GitHub 建立新儲存庫

1. 登入 https://github.com 。
2. 點右上 **+** → **New repository**。
3. **Repository name**：例如 `Cursor1stLook` 或 `PlannerAgent`。
4. 選擇 **Public**，**不要**勾選 "Add a README"（專案已有檔案）。
5. 點 **Create repository**。
6. 建立後畫面上會顯示 **HTTPS** 或 **SSH** 網址，例如：  
   `https://github.com/你的帳號/Cursor1stLook.git`  
   先複製此網址，後面會用到。

---

## 3. 在本機專案目錄執行 Git 指令

在 **PowerShell** 或 **命令提示字元** 中，切到專案目錄後依序執行：

```powershell
cd D:\Python\Cursor1stLook
```

```powershell
git init
```

```powershell
git add .
git status
```
（確認列出的檔案正確，且**沒有** `.env`；`.env` 已在 `.gitignore` 中排除。）

```powershell
git commit -m "Initial commit: Planner/Outlook/OneNote agent with Power Automate"
```

```powershell
git branch -M main
git remote add origin https://github.com/你的帳號/你的儲存庫名稱.git
```
（把上面的網址換成你在步驟 2 複製的儲存庫網址。）

```powershell
git push -u origin main
```

若 GitHub 要求登入，請依畫面使用 **Personal Access Token** 或 **SSH key** 完成驗證。

---

## 4. 之後的日常更新

程式或文件有修改後，若要再送交到 GitHub：

```powershell
cd D:\Python\Cursor1stLook
git add .
git status
git commit -m "簡短說明這次改了什麼"
git push
```

---

## 注意事項

- **`.env` 不會被提交**（已寫入 `.gitignore`），機密資訊不會上傳到 GitHub。
- 若已把 `.env` 加入過追蹤，請先執行：  
  `git rm --cached .env`  
  再 `git commit` 與 `git push`。
