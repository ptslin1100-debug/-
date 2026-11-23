# 班級分組網站（37 人）

網站：[GitHub Pages](https://ptslin1100-debug.github.io/-/)（需先啟用 Pages）

- 分組：第 1–6 組上限 6 人，第 7 組上限 7 人。
- 狀態：確定入組、考慮中、尋找組別。
- 支援退出組別、重新選組、為所在／考慮中的組別命名。
- 全班即時同步依賴 Google Apps Script / Google Sheet，**尚未完成設定前不可儲存修改**。

## 一次性後端啟用流程
1. 用班級管理者的 Google 帳號進入 https://script.google.com/home/projects/create ，開啟全新專案。
2. 將本儲存庫的 `google-apps-script.gs` **全部內容**複製至編輯器的 `Code.gs`（取代原有內容），儲存。
3. 在上方函式選單選擇 `setup`，點「執行」，允許權限。執行記錄會顯示新建立的 Google 試算表連結，裡面有「分組狀態」「個人代碼」「組別名稱」三個工作表。
4. 不要公開分享含「個人代碼」分頁的試算表；由管理者私下發放每人代碼。
5. 在 Apps Script 點選 **部署 > 新增部署 > 網頁應用程式**，執行身分選「我」，存取者選「任何人」，點部署，複製以 `/exec` 結尾的網址。
6. 在 GitHub 的 `config.js` 將 `googleScriptUrl: ''` 改為 `googleScriptUrl: '剛才複製的 /exec 網址'`，提交變更。
7. 開啟 GitHub Pages 網站，用兩台裝置測試：一位同學輸入座號與代碼，確定入組；另一台應在 5 秒內看到更新；再測試退組與組名。

**說明：** GitHub Pages 只提供網站網頁，資料由 Google Sheet 儲存。Apps Script 網頁應用程式連結需公開讀取，但所有修改會檢查個人代碼。個人代碼請勿提交到公開 GitHub。
