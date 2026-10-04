津丘 chill in POS｜Google Sheet 雲端同步設定

一、Google Sheet
1. 新增一個 Google 試算表，例如：津丘 POS 雲端資料庫
2. 點「擴充功能」→「Apps Script」
3. 把 Code.gs 全部內容貼上
4. 把 SPREADSHEET_ID = 'CHANGE_ME_SPREADSHEET_ID' 改成你的 Google Sheet ID（網址 https://docs.google.com/spreadsheets/d/【這一段】/edit）
5. 把 API_KEY = 'CHANGE_ME_JINQIU_2026' 改成你自己的英數字，例如：JQPOS2026A
6. 儲存
7. 在 Apps Script 上方函式選擇 setupSheets，按「執行」一次。
   會自動建立：Orders、OrderItems、Expenses、Menu、Categories
8. 點「部署」→「新增部署」
9. 類型選「網頁應用程式」
10. 執行身分：我
11. 誰可以存取：所有人
12. 部署並完成授權
13. 複製產生的 /exec 網址

二、POS
1. GitHub Pages 開啟後，點右上角齒輪「雲端同步設定」
2. GAS 網頁應用程式網址：貼上 /exec 網址
3. API 金鑰：填入與 Code.gs 完全相同的 API_KEY
4. 開啟「啟用雲端同步」
5. 先按「測試連線」確認成功
6. 再按「儲存」

三、資料表用途
Orders：一筆訂單一列
OrderItems：一筆訂單的每個品項一列
Expenses：每日支出
Menu：菜單品項、價格、成本、分類、活動折扣
Categories：菜單分類

四、重要說明
- 訂單、支出、菜單、分類會以 Google Sheet 為跨裝置資料來源。
- 菜單圖片目前仍保存在瀏覽器／GitHub Pages，不上傳到 Google Sheet；因此不同裝置若沒有該圖片的本機資料，圖片可能顯示預設圖。文字、價格、成本、分類仍會同步。
- POS 每 15 秒重新讀取一次雲端資料，切回瀏覽器分頁時也會同步。
- API_KEY 只是簡單的存取門檻，不是高強度安全機制；Google Sheet 本身請不要公開分享編輯權限。
