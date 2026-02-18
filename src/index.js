import cron from 'node-cron';
import { config } from './config.js';
import { getAuthenticatedClient } from './auth.js';
import { fetchAllCyberbizEmails } from './gmail.js';
import { parseEmails } from './parser.js';
import { syncOrdersToSheet } from './sheets.js';

/**
 * 計算搜尋起始日期
 * @param {number} days - 往前推幾天
 * @returns {string} YYYY/MM/DD 格式
 */
function getAfterDate(days) {
  const date = new Date();
  date.setDate(date.getDate() - days);
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}/${m}/${d}`;
}

/**
 * 執行一次完整的同步流程
 */
async function runSync(auth, afterDate) {
  const startTime = Date.now();
  console.log(`\n${'='.repeat(50)}`);
  console.log(`[Sync] 開始同步 - ${new Date().toLocaleString('zh-TW')}`);
  console.log(`[Sync] 搜尋 ${afterDate} 之後的訂單郵件`);

  try {
    // 1. 從 Gmail 取得 Cyberbiz 訂單郵件
    const emails = await fetchAllCyberbizEmails(auth, { afterDate });

    if (emails.length === 0) {
      console.log('[Sync] 沒有找到新的訂單郵件');
      return;
    }

    // 2. 解析郵件取得訂單資訊
    const orders = parseEmails(emails);

    if (orders.length === 0) {
      console.log('[Sync] 沒有成功解析的訂單');
      return;
    }

    // 3. 同步到 Google Sheets
    const result = await syncOrdersToSheet(auth, orders);

    const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
    console.log(`[Sync] 完成 - 新增 ${result.added} 筆, 更新 ${result.updated} 筆 (耗時 ${elapsed}s)`);
  } catch (err) {
    console.error(`[Sync] 同步失敗:`, err.message);
    if (err.code === 401 || err.code === 403) {
      console.error('[Sync] Token 可能已失效，請重新執行 npm run auth 進行授權');
    }
  }
}

/**
 * 主程式
 */
async function main() {
  console.log('========================================');
  console.log(' Cyberbiz 訂單同步系統');
  console.log(' Gmail -> Google Sheets ERP');
  console.log('========================================');

  // 驗證必要設定
  if (!config.google.clientId || !config.google.clientSecret) {
    console.error('[Error] 請先設定 GOOGLE_CLIENT_ID 和 GOOGLE_CLIENT_SECRET');
    console.error('[Error] 請參考 .env.example 建立 .env 檔案');
    process.exit(1);
  }

  if (!config.sheets.spreadsheetId) {
    console.error('[Error] 請先設定 SPREADSHEET_ID');
    process.exit(1);
  }

  // 取得授權
  const auth = await getAuthenticatedClient();

  // 判斷是否為單次執行模式
  const isOnce = process.argv.includes('--once');

  // 首次同步：抓取指定天數內的訂單
  const initialAfterDate = getAfterDate(config.initialFetchDays);
  await runSync(auth, initialAfterDate);

  if (isOnce) {
    console.log('\n[System] 單次執行模式，結束。');
    return;
  }

  // 啟動排程
  console.log(`\n[Scheduler] 啟動排程: ${config.cron.schedule}`);
  console.log('[Scheduler] 按 Ctrl+C 停止\n');

  cron.schedule(config.cron.schedule, async () => {
    // 排程模式只抓取最近 2 天的郵件（避免重複處理太多）
    const afterDate = getAfterDate(2);
    await runSync(auth, afterDate);
  });
}

main().catch((err) => {
  console.error('[Fatal]', err);
  process.exit(1);
});
