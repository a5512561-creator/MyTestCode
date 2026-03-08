# 將專案上傳到 GitHub（版本控管）

目前本機若尚未安裝 Git 或尚未初始化儲存庫，請依下列步驟操作。

---

## 若你之前已 push 過、現在 `git pull` 失敗（除錯與以本機為準上傳）

這種情況通常是：本機與 GitHub 上的歷史不一致（例如曾在別處改過、或本機重新做過很多修改），`git pull` 會報錯或要求合併。

**目標：以「目前 Cursor1stLook 專案」為準，讓 GitHub 與本機一致。**

**快速作法（確定要以本機覆蓋遠端時）：** 在專案目錄執行  
`git add .` → `git commit -m "Update: current project"`（若有未存變更）→ `git push --force origin main`

### 步驟 A：先確認狀態

在專案目錄執行（PowerShell 或 Git Bash）：

```powershell
cd D:\Python\Cursor1stLook
git status
git remote -v
git branch -a
```

- `git status`：看是否有未儲存的變更、目前分支。
- `git remote -v`：確認 `origin` 是否指到你的 GitHub 網址。
- `git branch -a`：看目前分支（例如 `main` 或 `master`）與遠端分支。

### 步驟 B：把「目前所有檔案」都納入一次 commit

若本機有改動還沒 commit，先全部加入並 commit：

```powershell
git add .
git status
git commit -m "Update: current Cursor1stLook project state"
```

（若顯示 "nothing to commit, working tree clean"，表示都已 commit，可跳過。）

### 步驟 C：以本機為準覆寫遠端（強制推送）

確定你要**用本機狀態覆蓋 GitHub 上的內容**後，執行（請把 `main` 換成你實際的分支名，多數為 `main` 或 `master`）：

```powershell
git push --force origin main
```

若你的分支叫 `master`：

```powershell
git push --force origin master
```

完成後，GitHub 上的儲存庫會與你本機目前狀態一致。

### 若不想覆蓋遠端、想先合併再推送

若 GitHub 上有別人（或你在別台電腦）的 commit 你想保留：

```powershell
git pull origin main --rebase
```
（有衝突就依提示解衝突，解完 `git rebase --continue`，最後再 `git push origin main`。）

若想直接合併（保留兩邊歷史）：

```powershell
git pull origin main
```
（有衝突就手動解衝突，再 `git add .` → `git commit` → `git push`。）

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
