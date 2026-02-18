import { google } from 'googleapis';
import { config } from './config.js';

/**
 * ERP「03_銷貨紀錄」欄位對應 (A ~ W, 共 23 欄)
 *
 * A: 訂單日期       B: 訂單編號      C: 通路來源
 * D: SKU           E: 外部品名       F: 銷貨數量
 * G: 售價           H: 銷貨金額      I: 訂單狀態
 * J: 客戶代碼       K: 客戶名稱      L: 出貨日期
 * M: 物流單號       N: 收款狀態      O: 收款日期
 * P: 成本           Q: 毛利          R: 毛利率
 * S: 賣場優惠券     T: 成交手續費    U: 其他服務費
 * V: 金流與系統處理費 W: 備註
 */

const HEADER_ROW = [
  '訂單日期',       // A
  '訂單編號',       // B
  '通路來源',       // C
  'SKU',            // D
  '外部品名',       // E
  '銷貨數量',       // F
  '售價',           // G
  '銷貨金額',       // H
  '訂單狀態',       // I
  '客戶代碼',       // J
  '客戶名稱',       // K
  '出貨日期',       // L
  '物流單號',       // M
  '收款狀態',       // N
  '收款日期',       // O
  '成本',           // P
  '毛利',           // Q
  '毛利率',         // R
  '賣場優惠券',     // S
  '成交手續費',     // T
  '其他服務費',     // U
  '金流與系統處理費', // V
  '備註',           // W
];

// 欄位範圍: A ~ W
const COL_RANGE = 'A:W';
const COL_END = 'W';

/**
 * 取得 Google Sheets API 實例
 */
function getSheetsApi(auth) {
  return google.sheets({ version: 'v4', auth });
}

/**
 * 確保工作表存在且有表頭
 */
export async function ensureSheetSetup(auth) {
  const sheets = getSheetsApi(auth);
  const { spreadsheetId, sheetName } = config.sheets;

  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
  const sheetExists = spreadsheet.data.sheets.some(
    (s) => s.properties.title === sheetName
  );

  if (!sheetExists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          { addSheet: { properties: { title: sheetName } } },
        ],
      },
    });
    console.log(`[Sheets] 已建立工作表: ${sheetName}`);
  }

  // 檢查是否有表頭
  const headerRange = `'${sheetName}'!A1:${COL_END}1`;
  const headerRes = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: headerRange,
  });

  if (!headerRes.data.values || headerRes.data.values.length === 0) {
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: headerRange,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [HEADER_ROW] },
    });
    console.log('[Sheets] 已寫入表頭');

    const sheetId = await getSheetId(sheets, spreadsheetId, sheetName);
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            repeatCell: {
              range: { sheetId, startRowIndex: 0, endRowIndex: 1 },
              cell: {
                userEnteredFormat: {
                  textFormat: { bold: true },
                  backgroundColor: { red: 0.85, green: 0.92, blue: 1.0 },
                },
              },
              fields: 'userEnteredFormat(textFormat,backgroundColor)',
            },
          },
          {
            updateSheetProperties: {
              properties: {
                sheetId,
                gridProperties: { frozenRowCount: 1 },
              },
              fields: 'gridProperties.frozenRowCount',
            },
          },
        ],
      },
    });
    console.log('[Sheets] 已格式化表頭');
  }
}

/**
 * 取得工作表 ID
 */
async function getSheetId(sheets, spreadsheetId, sheetName) {
  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
  const sheet = spreadsheet.data.sheets.find(
    (s) => s.properties.title === sheetName
  );
  return sheet?.properties?.sheetId ?? 0;
}

/**
 * 取得已存在的訂單編號集合（用於去重）
 * 讀取 B 欄（訂單編號）和 E 欄（外部品名）來建立已存在的 key
 * @returns {Promise<Set<string>>} "訂單編號|品名" 的集合
 */
export async function getExistingOrderKeys(auth) {
  const sheets = getSheetsApi(auth);
  const { spreadsheetId, sheetName } = config.sheets;

  // 讀取 B 欄（訂單編號）和 E 欄（外部品名）
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${sheetName}'!B:E`,
  });

  const existingKeys = new Set();
  if (res.data.values) {
    res.data.values.forEach((row, index) => {
      if (index === 0) return; // 跳過表頭
      const orderNumber = row[0] || '';   // B 欄 = index 0
      const itemName = row[3] || '';      // E 欄 = index 3 (B,C,D,E)
      if (orderNumber) {
        existingKeys.add(`${orderNumber}|${itemName}`);
      }
    });
  }

  return existingKeys;
}

/**
 * 取得已存在的訂單編號與其對應行號（用於狀態更新）
 * @returns {Promise<Map<string, number[]>>} 訂單編號 -> 行號陣列 (1-indexed, 含表頭)
 */
export async function getExistingOrderRows(auth) {
  const sheets = getSheetsApi(auth);
  const { spreadsheetId, sheetName } = config.sheets;

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${sheetName}'!B:B`,
  });

  const orderRows = new Map();
  if (res.data.values) {
    res.data.values.forEach((row, index) => {
      if (index === 0) return;
      const orderNumber = row[0];
      if (orderNumber) {
        const rows = orderRows.get(orderNumber) || [];
        rows.push(index + 1); // 1-indexed for Sheets API
        orderRows.set(orderNumber, rows);
      }
    });
  }

  return orderRows;
}

/**
 * 將一筆訂單展開為多列（每個品項一列）
 * @param {Object} order - 解析後的訂單物件
 * @returns {Array<Array>} Sheets rows (每列 23 個值, A ~ W)
 */
function orderToRows(order) {
  // 如果沒有品項，仍建立一列（品項欄位留空）
  if (order.items.length === 0) {
    return [buildRow(order, null)];
  }

  return order.items.map((item) => buildRow(order, item));
}

/**
 * 建立單一列資料 (A ~ W)
 */
function buildRow(order, item) {
  return [
    order.orderDate,                          // A: 訂單日期
    order.orderNumber,                        // B: 訂單編號
    order.channelSource,                      // C: 通路來源
    item?.sku || '',                          // D: SKU
    item?.name || '',                         // E: 外部品名
    item?.quantity || '',                     // F: 銷貨數量
    item?.unitPrice || '',                    // G: 售價
    item?.subtotal || '',                     // H: 銷貨金額
    order.orderStatus,                        // I: 訂單狀態
    '',                                       // J: 客戶代碼（手動填寫）
    order.customerName,                       // K: 客戶名稱
    order.shippingDate,                       // L: 出貨日期
    order.trackingNumber,                     // M: 物流單號
    order.paymentStatus,                      // N: 收款狀態
    order.paymentDate,                        // O: 收款日期
    '',                                       // P: 成本（手動填寫）
    '',                                       // Q: 毛利（手動填寫或公式）
    '',                                       // R: 毛利率（手動填寫或公式）
    order.discount || '',                     // S: 賣場優惠券
    '',                                       // T: 成交手續費（手動填寫）
    '',                                       // U: 其他服務費（手動填寫）
    '',                                       // V: 金流與系統處理費（手動填寫）
    `Gmail 自動匯入`,                         // W: 備註
  ];
}

/**
 * 將訂單資料同步至 Google Sheets ERP
 * - 新訂單品項：新增列
 * - 已存在訂單：更新狀態欄位（訂單狀態、出貨日期、物流單號、收款狀態、收款日期）
 *
 * @param {import('googleapis').Auth.OAuth2Client} auth
 * @param {Array} orders - 訂單陣列
 * @returns {Promise<{added: number, updated: number, skipped: number}>}
 */
export async function syncOrdersToSheet(auth, orders) {
  const sheets = getSheetsApi(auth);
  const { spreadsheetId, sheetName } = config.sheets;

  await ensureSheetSetup(auth);

  // 取得已存在的 key 和行號
  const [existingKeys, existingOrderRows] = await Promise.all([
    getExistingOrderKeys(auth),
    getExistingOrderRows(auth),
  ]);

  let added = 0;
  let updated = 0;
  let skipped = 0;

  const newRows = [];
  const updateRequests = [];

  for (const order of orders) {
    const existingRows = existingOrderRows.get(order.orderNumber);

    if (existingRows && existingRows.length > 0) {
      // 訂單已存在 → 更新狀態相關欄位（I, L, M, N, O）
      // 只更新有值的欄位，不覆蓋手動填寫的資料
      const statusUpdates = [];

      if (order.orderStatus) {
        statusUpdates.push({
          range: `'${sheetName}'!I${existingRows[0]}`,
          values: [[order.orderStatus]],
        });
      }
      if (order.shippingDate) {
        for (const row of existingRows) {
          statusUpdates.push({
            range: `'${sheetName}'!L${row}:M${row}`,
            values: [[order.shippingDate, order.trackingNumber]],
          });
        }
      }
      if (order.paymentStatus) {
        for (const row of existingRows) {
          statusUpdates.push({
            range: `'${sheetName}'!N${row}:O${row}`,
            values: [[order.paymentStatus, order.paymentDate]],
          });
        }
      }

      if (statusUpdates.length > 0) {
        updateRequests.push(...statusUpdates);
        updated++;
      } else {
        skipped++;
      }
    } else {
      // 新訂單 → 展開品項為多列新增
      const rows = orderToRows(order);
      for (const row of rows) {
        const key = `${order.orderNumber}|${row[4]}`; // B|E
        if (!existingKeys.has(key)) {
          newRows.push(row);
          existingKeys.add(key); // 避免同批次重複
        }
      }
      added++;
    }
  }

  // 批次更新狀態
  if (updateRequests.length > 0) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: {
        valueInputOption: 'USER_ENTERED',
        data: updateRequests,
      },
    });
  }

  // 批次新增品項列
  if (newRows.length > 0) {
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: `'${sheetName}'!${COL_RANGE}`,
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values: newRows },
    });
  }

  console.log(
    `[Sheets] 同步完成: 新增 ${added} 筆訂單 (${newRows.length} 列), 更新 ${updated} 筆, 略過 ${skipped} 筆`
  );
  return { added, updated, skipped };
}
