# Cyberbiz 訂單同步系統 (Gmail -> Google Sheets ERP)

自動從 Gmail 抓取 Cyberbiz 訂單郵件，解析訂單內容後同步至 Google Sheets ERP 系統的「03_銷貨紀錄」工作表。

## 功能

- 自動搜尋 Gmail 中的 Cyberbiz 訂單通知郵件
- 解析郵件 HTML，提取訂單編號、客戶名稱、SKU、品名、數量、金額等
- **每個品項一列**寫入 Google Sheets（對應 ERP 銷貨紀錄結構）
- 新訂單自動新增，已存在訂單自動更新狀態（出貨、收款等）
- 不覆蓋手動填寫的欄位（成本、毛利、手續費等）
- 支援排程執行（預設每 10 分鐘）或單次執行

## 系統架構

```
Gmail (Cyberbiz 訂單郵件)
  │
  ▼
[Gmail API] 搜尋 & 讀取郵件
  │
  ▼
[Parser] 解析 HTML，提取訂單 + 品項資料
  │
  ▼ (每個品項展開為一列)
[Sheets API] 寫入/更新 03_銷貨紀錄
  │
  ▼
Google Sheets ERP (A~W 共 23 欄)
```

## 前置需求

- Node.js 18+
- Google Cloud 專案（啟用 Gmail API 與 Google Sheets API）
- 含「03_銷貨紀錄」工作表的 Google Sheets

## 設定步驟

### 1. 建立 Google Cloud 專案

1. 前往 [Google Cloud Console](https://console.cloud.google.com/)
2. 建立新專案或選擇現有專案
3. 啟用 **Gmail API** 和 **Google Sheets API**

### 2. 建立 OAuth2 憑證

1. 前往「API 和服務」→「憑證」
2. 點擊「建立憑證」→「OAuth 用戶端 ID」
3. 應用程式類型選擇「網路應用程式」
4. 新增授權的重新導向 URI: `http://localhost:3000/oauth2callback`
5. 記下 `Client ID` 和 `Client Secret`

### 3. 設定環境變數

```bash
cp .env.example .env
```

編輯 `.env` 填入：

```
GOOGLE_CLIENT_ID=你的_client_id
GOOGLE_CLIENT_SECRET=你的_client_secret
SPREADSHEET_ID=你的_spreadsheet_id
```

### 4. 安裝與授權

```bash
npm install
npm run auth    # 開啟瀏覽器完成 Google 帳號授權
```

## 使用方式

```bash
npm start       # 排程模式（持續運行，每 10 分鐘同步）
npm run fetch   # 單次執行模式
```

## ERP 欄位對應 (A ~ W)

| 欄 | 欄位名稱 | 自動填入 | 說明 |
|----|----------|---------|------|
| A | 訂單日期 | v | 從郵件解析 |
| B | 訂單編號 | v | 從郵件主旨解析 |
| C | 通路來源 | v | 固定為「Cyberbiz」 |
| D | SKU | v | 從品名中提取（如有） |
| E | 外部品名 | v | 商品名稱 |
| F | 銷貨數量 | v | 購買數量 |
| G | 售價 | v | 單價 |
| H | 銷貨金額 | v | 數量 x 單價 |
| I | 訂單狀態 | v | 新訂單/已付款/已出貨/已取消等 |
| J | 客戶代碼 | | 手動填寫 |
| K | 客戶名稱 | v | 從郵件解析收件人 |
| L | 出貨日期 | v | 從出貨通知郵件解析 |
| M | 物流單號 | v | 從出貨通知郵件解析 |
| N | 收款狀態 | v | 已收款/未收款/已退款 |
| O | 收款日期 | v | 從付款通知郵件解析 |
| P | 成本 | | 手動填寫 |
| Q | 毛利 | | 手動填寫或公式 |
| R | 毛利率 | | 手動填寫或公式 |
| S | 賣場優惠券 | v | 從郵件折扣資訊解析 |
| T | 成交手續費 | | 手動填寫 |
| U | 其他服務費 | | 手動填寫 |
| V | 金流與系統處理費 | | 手動填寫 |
| W | 備註 | v | 標記「Gmail 自動匯入」 |

## 同步邏輯

- **新訂單**：每個品項新增一列，訂單層級的資訊（日期、編號、狀態、客戶）在每列重複
- **已存在訂單**：只更新狀態相關欄位（I、L、M、N、O），不覆蓋手動填寫的欄位
- **去重機制**：以「訂單編號 + 外部品名」作為唯一鍵，避免重複新增

## 環境變數

| 變數 | 必填 | 預設值 | 說明 |
|------|------|--------|------|
| GOOGLE_CLIENT_ID | 是 | - | Google OAuth2 Client ID |
| GOOGLE_CLIENT_SECRET | 是 | - | Google OAuth2 Client Secret |
| GOOGLE_REDIRECT_URI | 否 | `http://localhost:3000/oauth2callback` | OAuth2 回呼 URI |
| SPREADSHEET_ID | 是 | - | Google Sheets 試算表 ID |
| SHEET_NAME | 否 | `03_銷貨紀錄` | 工作表名稱 |
| CYBERBIZ_SENDER | 否 | `noreply@cyberbiz.co` | Cyberbiz 寄件者 email |
| CRON_SCHEDULE | 否 | `*/10 * * * *` | 排程 cron 表達式 |
| INITIAL_FETCH_DAYS | 否 | `30` | 首次抓取天數範圍 |

## 測試

```bash
npm test
```

## 自訂解析規則

若你的 Cyberbiz 郵件格式與預設不同，可修改 `src/parser.js`：

- `extractOrderNumber()` - 訂單編號提取規則
- `extractOrderStatus()` - 狀態關鍵字對應
- `extractPaymentStatus()` - 收款狀態判斷
- `extractOrderItems()` - 商品品項與 SKU 解析
