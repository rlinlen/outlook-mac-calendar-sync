# Google Calendar API 設定指南

## 步驟 1: 建立 Google Cloud 專案

1. 前往 [Google Cloud Console](https://console.cloud.google.com/)
2. 點擊「建立專案」或選擇現有專案
3. 記下專案 ID

## 步驟 2: 啟用 Google Calendar API

1. 在 Google Cloud Console 中，前往「API 和服務」→「程式庫」
2. 搜尋「Google Calendar API」
3. 點擊並啟用 API

## 步驟 3: 建立 OAuth 2.0 憑證

1. 前往「API 和服務」→「憑證」
2. 點擊「建立憑證」→「OAuth 用戶端 ID」
3. 如果是第一次，需要先設定「OAuth 同意畫面」：
   - 選擇「外部」使用者類型
   - 填寫應用程式名稱（例如：Outlook Calendar Sync）
   - 填寫使用者支援電子郵件
   - 填寫開發人員聯絡資訊
   - 在「範圍」頁面，新增 `../auth/calendar` 範圍
   - 在「測試使用者」頁面，新增你的 Google 帳號

4. 回到「憑證」頁面，建立 OAuth 用戶端 ID：
   - 應用程式類型：選擇「桌面應用程式」
   - 名稱：輸入描述性名稱（例如：Outlook Calendar Sync Client）

5. 下載 JSON 憑證檔案
6. 將檔案重新命名為 `client_secret.json` 並放在專案根目錄

## 首次執行

第一次執行時會：
1. 開啟瀏覽器進行 OAuth 認證
2. 要求授權存取 Google Calendar
3. 生成 `token.json` 檔案供後續使用

## 檔案結構

```
outlook-mac-calendar-sync/
├── requirements.txt           # pip 依賴清單
├── setup_google_calendar_api.md
├── script/                     
    ├── sync_csv_with_google_calendar.py
├── data/                     # 暫存檔案（需手動建立）
    ├── client_secret.json    # Google API 憑證（需要手動下載）
└── .venv/                    # uv 虛擬環境（自動建立）
```

## 常見問題

### 1. 憑證檔案找不到
確保 `client_secret.json` 檔案在專案根目錄，且檔名正確。

### 2. 權限錯誤
確保在 OAuth 同意畫面中新增了正確的範圍：
- `https://www.googleapis.com/auth/calendar`

### 3. 測試使用者限制
如果應用程式處於「測試」狀態，只有新增到「測試使用者」清單的帳號才能使用。

### 4. Token 過期
刪除 `token.json` 檔案並重新執行程式進行重新認證。

## 安全注意事項

1. **不要將 `client_secret.json` 提交到版本控制系統**
2. **不要分享 `token.json` 檔案**
3. **定期檢查 Google Cloud Console 中的 API 使用情況**

## API 配額限制

Google Calendar API 有以下限制：
- 每日請求數：1,000,000
- 每 100 秒請求數：10,000
- 每 100 秒每使用者請求數：250

一般個人使用不會達到這些限制。
