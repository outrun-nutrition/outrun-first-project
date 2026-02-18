import { google } from 'googleapis';
import { config } from './config.js';

const HEADER_ROW = [
  '訂單編號',
  '訂單日期',
  '訂單狀態',
  '客戶姓名',
  '客戶Email',
  '客戶電話',
  '收件地址',
  '付款方式',
  '配送方式',
  '商品明細',
  '小計',
  '運費',
  '折扣',
  '訂單總額',
  '郵件主旨',
  '郵件日期',
  '最後更新',
];

/**
 * 取得 Google Sheets API 實例
 */
function getSheetsApi(auth) {
  return google.sheets({ version: 'v4', auth });
}

/**
 * 確保工作表存在且有表頭
 * @param {import('googleapis').Auth.OAuth2Client} auth
 */
export async function ensureSheetSetup(auth) {
  const sheets = getSheetsApi(auth);
  const { spreadsheetId, sheetName } = config.sheets;

  // 檢查工作表是否存在
  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
  const sheetExists = spreadsheet.data.sheets.some(
    (s) => s.properties.title === sheetName
  );

  if (!sheetExists) {
    // 建立工作表
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: { title: sheetName },
            },
          },
        ],
      },
    });
    console.log(`[Sheets] 已建立工作表: ${sheetName}`);
  }

  // 檢查是否有表頭
  const headerRange = `${sheetName}!A1:Q1`;
  const headerRes = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: headerRange,
  });

  if (!headerRes.data.values || headerRes.data.values.length === 0) {
    // 寫入表頭
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: headerRange,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [HEADER_ROW],
      },
    });
    console.log('[Sheets] 已寫入表頭');

    // 格式化表頭（粗體 + 凍結）
    const sheetId = await getSheetId(sheets, spreadsheetId, sheetName);
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            repeatCell: {
              range: {
                sheetId,
                startRowIndex: 0,
                endRowIndex: 1,
              },
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
 * 取得所有已存在的訂單編號
 * @param {import('googleapis').Auth.OAuth2Client} auth
 * @returns {Promise<Map<string, number>>} 訂單編號 -> 行號 (0-indexed)
 */
export async function getExistingOrders(auth) {
  const sheets = getSheetsApi(auth);
  const { spreadsheetId, sheetName } = config.sheets;

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${sheetName}!A:A`,
  });

  const existingOrders = new Map();
  if (res.data.values) {
    res.data.values.forEach((row, index) => {
      if (index === 0) return; // 跳過表頭
      if (row[0]) {
        existingOrders.set(row[0], index);
      }
    });
  }

  return existingOrders;
}

/**
 * 將訂單轉為 Sheets row 資料
 */
function orderToRow(order) {
  const itemsSummary = order.items
    .map((item) => `${item.name} x${item.quantity} ($${item.price})`)
    .join('; ');

  return [
    order.orderNumber,
    order.orderDate,
    order.orderStatus,
    order.customerName,
    order.customerEmail,
    order.customerPhone,
    order.shippingAddress,
    order.paymentMethod,
    order.shippingMethod,
    itemsSummary,
    order.subtotal || '',
    order.shippingFee || '',
    order.discount || '',
    order.totalAmount || '',
    order.rawSubject,
    order.emailDate,
    new Date().toISOString(),
  ];
}

/**
 * 將訂單資料寫入/更新至 Google Sheets
 * @param {import('googleapis').Auth.OAuth2Client} auth
 * @param {Array} orders - 訂單陣列
 * @returns {Promise<{added: number, updated: number}>}
 */
export async function syncOrdersToSheet(auth, orders) {
  const sheets = getSheetsApi(auth);
  const { spreadsheetId, sheetName } = config.sheets;

  // 確保工作表已建立
  await ensureSheetSetup(auth);

  // 取得已存在的訂單
  const existingOrders = await getExistingOrders(auth);

  let added = 0;
  let updated = 0;

  const newRows = [];
  const updateRequests = [];

  for (const order of orders) {
    const row = orderToRow(order);
    const existingRowIndex = existingOrders.get(order.orderNumber);

    if (existingRowIndex !== undefined) {
      // 更新已存在的訂單（覆蓋整行）
      updateRequests.push({
        range: `${sheetName}!A${existingRowIndex + 1}:Q${existingRowIndex + 1}`,
        values: [row],
      });
      updated++;
    } else {
      // 新增訂單
      newRows.push(row);
      added++;
    }
  }

  // 批次更新已存在的訂單
  if (updateRequests.length > 0) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: {
        valueInputOption: 'USER_ENTERED',
        data: updateRequests,
      },
    });
  }

  // 批次新增訂單
  if (newRows.length > 0) {
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: `${sheetName}!A:Q`,
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: newRows,
      },
    });
  }

  console.log(`[Sheets] 同步完成: 新增 ${added} 筆, 更新 ${updated} 筆`);
  return { added, updated };
}
