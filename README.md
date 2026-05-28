# Outrun First Project — 訓練補給建議工具

堅心 Outrun Nutrition — 找到你的訓練補給建議

## 線上 URL

- 正式：https://tools.outrunnutrition.com/
- 舊路徑：`/nutrition-calculator.html` → 301 redirect 到根目錄（為舊 deep link 備援）

## 結構

- `index.html` — 主要工具頁（944 行，6 步驟表單）
- `nutrition-calculator.html` — 舊路徑 redirect 兼容
- `CNAME` — GitHub Pages custom domain (`tools.outrunnutrition.com`)
- `logo-white.png` — 品牌 logo

## 後端

填答資料透過 Google Apps Script Webhook 寫入 Google Sheets 後台。
- Apps Script: `https://script.google.com/macros/s/AKfycbzj.../exec`
- Sheet: `1qLSGyzhEbMHW_1SwrRvxQUUFEGpo-9F34lROuCn6wcY`

## Hosting

GitHub Pages with Cloudflare-managed CNAME。下季度評估搬遷至 Vercel / Netlify / 自家 server。
