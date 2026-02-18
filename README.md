# Cyberbiz 訂單同步系統 (Gmail -> Google Sheets ERP)

自動從 Gmail 抓取 Cyberbiz 訂單郵件，解析訂單內容後同步至 Google Sheets 建立的 ERP 系統。

## 功能

- 自動搜尋 Gmail 中的 Cyberbiz 訂單通知郵件
- 解析郵件 HTML 內容，提取訂單編號、客戶資訊、商品明細、金額等
- 自動建立/更新 Google Sheets 表格（含表頭、格式化）
- 新訂單自動新增，已存在訂單自動更新
- 支援排程執行（預設每 10 分鐘同步一次）
- 支援單次執行模式

## 系統架構

```
Gmail (Cyberbiz 訂單郵件)
  │
  ▼
[Gmail API] 搜尋 & 讀取郵件
  │
  ▼
[Parser] 解析 HTML，提取訂單資料
  │
  ▼
[Sheets API] 寫入/更新 Google Sheets
  │
  ▼
Google Sheets (ERP 訂單管理)
```

## 前置需求

- Node.js 18+
- Google Cloud 專案（啟用 Gmail API 與 Google Sheets API）
- 一份 Google Sheets 試算表

## 設定步驟

### 1. 建立 Google Cloud 專案

1. 前往 [Google Cloud Console](https://console.cloud.google.com/)
2. 建立新專案或選擇現有專案
3. 啟用 **Gmail API** 和 **Google Sheets API**
   - 前往「API 和服務」→「資料庫」
   - 搜尋並啟用 `Gmail API`
   - 搜尋並啟用 `Google Sheets API`

### 2. 建立 OAuth2 憑證

1. 前往「API 和服務」→「憑證」
2. 點擊「建立憑證」→「OAuth 用戶端 ID」
3. 應用程式類型選擇「網路應用程式」
4. 新增授權的重新導向 URI: `http://localhost:3000/oauth2callback`
5. 記下 `Client ID` 和 `Client Secret`

### 3. 建立 Google Sheets

1. 建立一份新的 Google Sheets 試算表
2. 從 URL 複製試算表 ID（`https://docs.google.com/spreadsheets/d/`**SPREADSHEET_ID**`/edit`）
3. 系統會自動建立「訂單」工作表並填入表頭

### 4. 設定環境變數

```bash
cp .env.example .env
```

編輯 `.env` 填入：

```
GOOGLE_CLIENT_ID=你的_client_id
GOOGLE_CLIENT_SECRET=你的_client_secret
SPREADSHEET_ID=你的_spreadsheet_id
```

### 5. 安裝依賴

```bash
npm install
```

### 6. 進行 Google 授權

```bash
npm run auth
```

瀏覽器會開啟授權頁面，授權後 token 會自動儲存至 `token.json`。

## 使用方式

### 排程模式（持續運行）

```bash
npm start
```

系統會先執行一次完整同步，之後按照排程定時抓取。

### 單次執行模式

```bash
npm run fetch
```

只同步一次後結束。

## Google Sheets 欄位說明

| 欄位 | 說明 |
|------|------|
| 訂單編號 | Cyberbiz 訂單編號 |
| 訂單日期 | 下單日期 |
| 訂單狀態 | 新訂單/已付款/已出貨/已取消/退貨中/已退款 |
| 客戶姓名 | 收件人姓名 |
| 客戶Email | 客戶電子信箱 |
| 客戶電話 | 收件人電話 |
| 收件地址 | 配送地址 |
| 付款方式 | 信用卡/ATM轉帳/超商付款等 |
| 配送方式 | 宅配/超商取貨等 |
| 商品明細 | 商品名稱 x 數量 (金額) |
| 小計 | 商品小計 |
| 運費 | 運費金額 |
| 折扣 | 折扣金額 |
| 訂單總額 | 訂單總額 |
| 郵件主旨 | 原始郵件主旨 |
| 郵件日期 | 郵件寄送日期 |
| 最後更新 | 資料最後更新時間 |

## 環境變數說明

| 變數 | 必填 | 預設值 | 說明 |
|------|------|--------|------|
| GOOGLE_CLIENT_ID | 是 | - | Google OAuth2 Client ID |
| GOOGLE_CLIENT_SECRET | 是 | - | Google OAuth2 Client Secret |
| GOOGLE_REDIRECT_URI | 否 | `http://localhost:3000/oauth2callback` | OAuth2 回呼 URI |
| SPREADSHEET_ID | 是 | - | Google Sheets 試算表 ID |
| SHEET_NAME | 否 | `訂單` | 工作表名稱 |
| CYBERBIZ_SENDER | 否 | `noreply@cyberbiz.co` | Cyberbiz 寄件者 email |
| CRON_SCHEDULE | 否 | `*/10 * * * *` | 排程 cron 表達式 |
| INITIAL_FETCH_DAYS | 否 | `30` | 首次抓取天數範圍 |

## 測試

```bash
npm test
```

## 自訂解析規則

若你的 Cyberbiz 郵件格式與預設不同，可修改 `src/parser.js` 中的正規表達式：

- `extractOrderNumber()` - 訂單編號提取規則
- `extractOrderStatus()` - 狀態關鍵字對應
- `extractOrderItems()` - 商品明細解析
- 各金額提取函式的正規表達式
