臺北市立建國高級中學圖書館潘威歷主任 20250326版
# 💰 退休財務幸福模擬工具

本工具是一套以 Google Sheets + Google Apps Script 建構的退休模擬與財務診斷系統，透過使用者輸入條件模擬資產變化，並自動產出報告寄送 Email。支援圖表分析、自動反推投資金額，以及精美格式化報告。

---

## 📋 使用步驟

### 1. 建立「輸入資料」工作表

建立一個 Google 試算表，並新增工作表命名為：`輸入資料`，包含以下欄位與說明：

| 欄位名稱 | 說明 |
|----------|------|
| 使用者姓名 | 將顯示在報告中 |
| 目前年齡 | 模擬起始年齡 |
| 預計退休年齡 | 模擬退休時間點（退休後停止工作收入） |
| 預期壽命 | 模擬結束年齡 |
| 現金儲蓄總額 | 起始時帳戶現金（不投資） |
| 每年儲蓄金額（不參與投資） | 年儲蓄金額，只累加至現金資產 |
| 每年投入投資金額（參與複利） | 每年投入投資（參與報酬率與複利） |
| 退休後每年持續投入投資金額 | 若退休後仍持續定期投資，填入此欄 |
| 預期年投資報酬率（%） | 投資年報酬（例：6% 請填 6） |
| 通膨率（%） | 每年生活成本上升比率（例：2.5），此參數目前沒用到，只是暫時留存 |
| 是否有退休金（Y/N） | 填 Y 表示退休後每年有退休金收入 |
| 每年退休金金額 | 若有退休金，每年金額 |
| 預期退休後每年支出 | 退休後生活支出（會依年支出成長率逐年遞增） |
| 支出年成長率（%） | 退休支出的年增率 |
| 財富自由目標（月支出） | 希望能支應的每月生活費 |
| 希望達成財富自由年齡 | 想在幾歲達成財富自由（可對照模擬） |
| 模擬間隔（每幾年計算一次） | 模擬時間粒度（預設填 1 表示每年） |
| Email 收件人（可選） | 若要寄送報告給他人，可填寫 Email |

> 📝 請將這些欄位填入 A 欄，使用者輸入值填在 B 欄。

---

### 2. 開啟 App Script，貼上本套件程式碼

- 點選「擴充功能」→「應用程式腳本」
- 新增 `.gs` 檔案並貼上完整程式碼
- 儲存後重新整理工作表

---

### 3. 點選功能選單操作

新增後會出現自訂選單 `退休模擬工具`，包含以下功能：

- 📘 模擬簡化版（以儲蓄為主）
- 🧮 自動反推年投資金額（維持正值）
- 📝 產出報告並寄送

---

## 📈 模擬功能說明

- 依據使用者輸入條件（年齡、退休年齡、儲蓄、投資、退休金等）逐年計算資產變化。
- 區分「非投資資產」與「投資資產（複利計算）」
- 若資產不足會自動動用投資收益補足
- 模擬結果顯示於 `模擬結果` 工作表並自動產生趨勢圖表

---

## 🔄 自動反推功能

- 功能：找出「每年最低投資金額」，使得模擬期間資產不會破產（總資產 ≥ 0）
- 使用二分搜尋方式測試
- 計算完成後會自動更新輸入資料中的「每年投入投資金額（參與複利）」

---

## 📝 報告產出與寄送

- 自動產生 Google Docs 報告
- 報告內容包含：
  - 使用者輸入條件
  - 財務診斷分析
  - 精簡模擬表格（年份、年齡、投資本金、投資收益、總資產）
  - 資產與投資圖表
  - Google Sheets 連結
- 匯出為 PDF 並寄送 Email 至使用者或指定地址

---

## 📌 注意事項

- 若無提供 Email 收件人，預設寄給登入者本人
- 圖表需先建立才能插入報告（由模擬完成後自動產生）
- 請確認 Google Apps Script 權限設定已允許發信與建立文件

---

## 📮 聯絡或延伸功能需求

如需擴充功能（例如封面設計、自動共享報告、更多圖表等），請聯絡作者或協作者。

---
