/**
 * =====================================================
 * 堅心 Outrun Nutrition — 營運記帳系統 Google Apps Script
 * =====================================================
 *
 * 【整合版】寫入現有 ERP 試算表的 04_費用紀錄 工作表
 *
 * 欄位對照（A-K 為原有欄位，L-T 為新增擴充欄位）：
 * A: 費用日期 | B: 費用類別 | C: 費用項目 | D: 金額 | E: 付款方式
 * F: 對象/供應商 | G: 關聯SKU | H: 備註 | I: 建立時間 | J: 建立人員
 * K: 關聯活動 | L: 科目代碼 | M: 上層分類 | N: 稅額 | O: 未稅金額
 * P: 發票號碼 | Q: 發票類型 | R: 發票照片連結 | S: Drive檔案ID | T: 狀態
 *
 * 擴充功能：
 * - 廠商名單雲端同步（07_廠商名單）
 * - 月報自動寄送（Time Trigger + MailApp）
 * - 預算管理欄位
 *
 * 設定步驟：
 * 1. 在現有 ERP 的 Apps Script 編輯器中新增此檔案
 * 2. 部署為 Web App（存取權限：任何人）
 * 3. 將 Web App URL 填入前端 HTML 的設定頁
 */

// ==================== 設定區 ====================
const EXPENSE_CONFIG = {
  // ERP 主試算表 ID
  SPREADSHEET_ID: '1_xTeRGPmz5Y1tgDb2JRe2oKiF7XBIRLjolOULD8bLmw',

  // Google Drive 發票照片資料夾 ID
  DRIVE_FOLDER_ID: '1C5cyuKp6J3v-d97D9Il7OC0pbMhBw2fh',

  // 工作表名稱（使用現有 ERP 的命名）
  SHEET_EXPENSE: '04_費用紀錄',
  SHEET_CATEGORIES: '06_科目設定',
  SHEET_VENDORS: '07_廠商名單',
  SHEET_QUOTATIONS: '08_報價單',
  SHEET_PO: '09_採購單',
  SHEET_DOC_ITEMS: '10_單據明細',

  // 月報自動寄送設定
  REPORT_EMAIL: '',  // 預設收件人（可透過 API 設定）
};

// ==================== 欄位索引常數 ====================
// 04_費用紀錄 的欄位位置（0-based index）
const COL = {
  DATE: 0,           // A: 費用日期
  CATEGORY_NAME: 1,  // B: 費用類別（科目名稱）
  ITEM: 2,           // C: 費用項目（備註/項目描述）
  AMOUNT: 3,         // D: 金額
  PAYMENT: 4,        // E: 付款方式
  VENDOR: 5,         // F: 對象/供應商
  SKU: 6,            // G: 關聯SKU
  NOTE: 7,           // H: 備註
  CREATED_AT: 8,     // I: 建立時間
  CREATED_BY: 9,     // J: 建立人員
  CAMPAIGN: 10,      // K: 關聯活動
  CATEGORY_CODE: 11, // L: 科目代碼（新增）
  PARENT_CAT: 12,    // M: 上層分類（新增）
  TAX: 13,           // N: 稅額（新增）
  NET_AMOUNT: 14,    // O: 未稅金額（新增）
  INVOICE_NUM: 15,   // P: 發票號碼（新增）
  INVOICE_TYPE: 16,  // Q: 發票類型（新增）
  IMAGE_URL: 17,     // R: 發票照片連結（新增）
  DRIVE_ID: 18,      // S: Drive檔案ID（新增）
  STATUS: 19,        // T: 狀態（新增）
};

const TOTAL_COLS = 20; // A~T 共 20 欄

// ==================== Web App 入口 ====================

/**
 * 處理 GET 請求（查詢資料）
 */
function doGet(e) {
  try {
    const action = e.parameter.action;
    let result;

    switch (action) {
      case 'getExpenses':
        result = getExpenses(e.parameter);
        break;
      case 'getCategories':
        result = getCategories();
        break;
      case 'getMonthlySummary':
        result = getMonthlySummary(e.parameter.year, e.parameter.month);
        break;
      case 'getYearlySummary':
        result = getYearlySummary(e.parameter.year);
        break;
      case 'getVendors':
        result = getVendors();
        break;
      // === 報價單/採購單 ===
      case 'getQuotations':
        result = getQuotations(e.parameter);
        break;
      case 'getPurchaseOrders':
        result = getPurchaseOrders(e.parameter);
        break;
      case 'getDocItems':
        result = getDocItems(e.parameter.docId);
        break;
      default:
        result = { error: '未知的操作' };
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 處理 POST 請求（新增/修改/刪除資料）
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    let result;

    switch (action) {
      case 'addExpense':
        result = addExpense(data.expense);
        break;
      case 'updateExpense':
        result = updateExpense(data.id, data.expense);
        break;
      case 'deleteExpense':
        result = deleteExpense(data.id);
        break;
      case 'uploadImage':
        result = uploadImage(data.fileName, data.fileData, data.mimeType);
        break;
      case 'batchAdd':
        result = batchAddExpenses(data.expenses);
        break;
      case 'syncVendors':
        result = syncVendors(data.vendors);
        break;
      case 'setupAutoReport':
        result = setupAutoReportTrigger(data.email);
        break;
      // === 報價單/採購單 ===
      case 'addQuotation':
        result = addQuotation(data);
        break;
      case 'updateQuotationStatus':
        result = updateQuotationStatus(data.id, data.status);
        break;
      case 'addPurchaseOrder':
        result = addPurchaseOrder(data);
        break;
      case 'updatePOStatus':
        result = updatePOStatus(data.id, data.status);
        break;
      default:
        result = { error: '未知的操作' };
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==================== 費用操作 ====================

/**
 * 將 expense 物件轉換為符合 04_費用紀錄 欄位順序的陣列
 */
function expenseToRow(expense, now) {
  const row = new Array(TOTAL_COLS).fill('');
  row[COL.DATE]          = expense.date;
  row[COL.CATEGORY_NAME] = expense.categoryName;
  row[COL.ITEM]          = expense.note || expense.categoryName;
  row[COL.AMOUNT]        = expense.amount;
  row[COL.PAYMENT]       = expense.paymentMethod;
  row[COL.VENDOR]        = expense.vendor || '';
  row[COL.SKU]           = expense.relatedSku || '';
  row[COL.NOTE]          = expense.note || '';
  row[COL.CREATED_AT]    = now;
  row[COL.CREATED_BY]    = expense.recorder || 'expense-tracker';
  row[COL.CAMPAIGN]      = expense.relatedCampaign || '';
  row[COL.CATEGORY_CODE] = expense.categoryCode;
  row[COL.PARENT_CAT]    = expense.parentCategory;
  row[COL.TAX]           = expense.taxAmount || '';
  row[COL.NET_AMOUNT]    = expense.netAmount || '';
  row[COL.INVOICE_NUM]   = expense.invoiceNumber || '';
  row[COL.INVOICE_TYPE]  = expense.invoiceType || '';
  row[COL.IMAGE_URL]     = expense.imageUrl || '';
  row[COL.DRIVE_ID]      = expense.imageDriveId || '';
  row[COL.STATUS]        = '正常';
  return row;
}

/**
 * 將工作表的一列資料轉換為 expense 物件
 */
function rowToExpense(row, rowIndex) {
  return {
    id: 'ROW-' + rowIndex,  // 使用列號作為 ID
    date: row[COL.DATE],
    categoryName: row[COL.CATEGORY_NAME],
    item: row[COL.ITEM],
    amount: row[COL.AMOUNT],
    paymentMethod: row[COL.PAYMENT],
    vendor: row[COL.VENDOR],
    relatedSku: row[COL.SKU],
    note: row[COL.NOTE],
    createdAt: row[COL.CREATED_AT],
    createdBy: row[COL.CREATED_BY],
    recorder: row[COL.CREATED_BY],
    relatedCampaign: row[COL.CAMPAIGN],
    categoryCode: row[COL.CATEGORY_CODE],
    parentCategory: row[COL.PARENT_CAT],
    taxAmount: row[COL.TAX],
    netAmount: row[COL.NET_AMOUNT],
    invoiceNumber: row[COL.INVOICE_NUM],
    invoiceType: row[COL.INVOICE_TYPE],
    imageUrl: row[COL.IMAGE_URL],
    imageDriveId: row[COL.DRIVE_ID],
    status: row[COL.STATUS] || '正常',
  };
}

/**
 * 新增一筆費用
 */
function addExpense(expense) {
  const ss = SpreadsheetApp.openById(EXPENSE_CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(EXPENSE_CONFIG.SHEET_EXPENSE);
  const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');

  const row = expenseToRow(expense, now);
  sheet.appendRow(row);

  // 記錄操作日誌
  logOperation('費用登錄', expense.categoryName + ' NT$' + expense.amount, expense.parentCategory, expense.amount);

  return { id: 'ROW-' + sheet.getLastRow(), message: '新增成功' };
}

/**
 * 批次新增費用（離線同步用）
 */
function batchAddExpenses(expenses) {
  const ss = SpreadsheetApp.openById(EXPENSE_CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(EXPENSE_CONFIG.SHEET_EXPENSE);
  const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');

  const rows = expenses.map(expense => expenseToRow(expense, now));

  if (rows.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rows.length, TOTAL_COLS).setValues(rows);
  }

  return { count: rows.length, message: '批次新增成功' };
}

/**
 * 更新費用
 */
function updateExpense(id, expense) {
  const ss = SpreadsheetApp.openById(EXPENSE_CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(EXPENSE_CONFIG.SHEET_EXPENSE);
  const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');

  // 從 id 取得列號（格式：ROW-N）
  const rowNum = parseInt(id.replace('ROW-', ''));
  if (isNaN(rowNum) || rowNum < 2) return { error: '無效的 ID' };

  // 更新可修改的欄位
  sheet.getRange(rowNum, COL.DATE + 1).setValue(expense.date);
  sheet.getRange(rowNum, COL.CATEGORY_NAME + 1).setValue(expense.categoryName);
  sheet.getRange(rowNum, COL.ITEM + 1).setValue(expense.note || expense.categoryName);
  sheet.getRange(rowNum, COL.AMOUNT + 1).setValue(expense.amount);
  sheet.getRange(rowNum, COL.PAYMENT + 1).setValue(expense.paymentMethod);
  sheet.getRange(rowNum, COL.VENDOR + 1).setValue(expense.vendor || '');
  sheet.getRange(rowNum, COL.NOTE + 1).setValue(expense.note || '');
  sheet.getRange(rowNum, COL.CATEGORY_CODE + 1).setValue(expense.categoryCode);
  sheet.getRange(rowNum, COL.PARENT_CAT + 1).setValue(expense.parentCategory);
  sheet.getRange(rowNum, COL.TAX + 1).setValue(expense.taxAmount || '');
  sheet.getRange(rowNum, COL.NET_AMOUNT + 1).setValue(expense.netAmount || '');
  sheet.getRange(rowNum, COL.INVOICE_NUM + 1).setValue(expense.invoiceNumber || '');
  sheet.getRange(rowNum, COL.INVOICE_TYPE + 1).setValue(expense.invoiceType || '');

  return { id: id, message: '更新成功' };
}

/**
 * 刪除費用（軟刪除，標記為作廢）
 */
function deleteExpense(id) {
  const ss = SpreadsheetApp.openById(EXPENSE_CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(EXPENSE_CONFIG.SHEET_EXPENSE);

  const rowNum = parseInt(id.replace('ROW-', ''));
  if (isNaN(rowNum) || rowNum < 2) return { error: '無效的 ID' };

  sheet.getRange(rowNum, COL.STATUS + 1).setValue('作廢');

  return { id: id, message: '已標記作廢' };
}

/**
 * 查詢費用
 */
function getExpenses(params) {
  const ss = SpreadsheetApp.openById(EXPENSE_CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(EXPENSE_CONFIG.SHEET_EXPENSE);
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) return [];

  let expenses = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // 跳過作廢的
    if (row[COL.STATUS] === '作廢') continue;

    // 月份篩選
    if (params.year && params.month) {
      const date = new Date(row[COL.DATE]);
      if (date.getFullYear() != params.year || (date.getMonth() + 1) != params.month) continue;
    }

    // 分類篩選
    if (params.category && row[COL.PARENT_CAT] !== params.category) continue;

    expenses.push(rowToExpense(row, i + 1));
  }

  // 依日期排序（最新在前）
  expenses.sort((a, b) => new Date(b.date) - new Date(a.date));

  return expenses;
}

// ==================== 科目管理 ====================

/**
 * 取得科目列表
 */
function getCategories() {
  const ss = SpreadsheetApp.openById(EXPENSE_CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(EXPENSE_CONFIG.SHEET_CATEGORIES);
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) return [];

  return data.slice(1).map(row => ({
    code: row[0],
    name: row[1],
    parent: row[2],
    description: row[3],
  }));
}

// ==================== 月報彙總 ====================

/**
 * 取得月報彙總
 */
function getMonthlySummary(year, month) {
  const expenses = getExpenses({ year: year, month: month });

  const byParent = {};
  const byCategory = {};
  const byPayment = {};
  const byInvoiceType = {};

  let totalAmount = 0;
  let totalTax = 0;
  let count = 0;

  expenses.forEach(exp => {
    const amount = Number(exp.amount) || 0;
    const tax = Number(exp.taxAmount) || 0;
    totalAmount += amount;
    totalTax += tax;
    count++;

    // 上層分類
    const parentKey = exp.parentCategory || '未分類';
    if (!byParent[parentKey]) byParent[parentKey] = 0;
    byParent[parentKey] += amount;

    // 科目
    const catKey = (exp.categoryCode || '') + ' ' + (exp.categoryName || '未分類');
    if (!byCategory[catKey]) byCategory[catKey] = { amount: 0, count: 0 };
    byCategory[catKey].amount += amount;
    byCategory[catKey].count++;

    // 付款方式
    const payKey = exp.paymentMethod || '未記錄';
    if (!byPayment[payKey]) byPayment[payKey] = 0;
    byPayment[payKey] += amount;

    // 發票類型
    const invType = exp.invoiceType || '無';
    if (!byInvoiceType[invType]) byInvoiceType[invType] = 0;
    byInvoiceType[invType] += amount;
  });

  return {
    year: year,
    month: month,
    totalAmount: totalAmount,
    totalTax: totalTax,
    totalNet: totalAmount - totalTax,
    count: count,
    byParent: byParent,
    byCategory: byCategory,
    byPayment: byPayment,
    byInvoiceType: byInvoiceType,
    expenses: expenses,
  };
}

/**
 * 年度彙總
 */
function getYearlySummary(year) {
  const monthlySummaries = [];
  for (let m = 1; m <= 12; m++) {
    monthlySummaries.push(getMonthlySummary(year, m));
  }

  return {
    year: year,
    months: monthlySummaries,
    yearTotal: monthlySummaries.reduce((sum, m) => sum + m.totalAmount, 0),
    yearTax: monthlySummaries.reduce((sum, m) => sum + m.totalTax, 0),
    yearCount: monthlySummaries.reduce((sum, m) => sum + m.count, 0),
  };
}

// ==================== 圖片上傳 ====================

/**
 * 上傳發票照片到 Google Drive
 */
function uploadImage(fileName, fileData, mimeType) {
  const folder = DriveApp.getFolderById(EXPENSE_CONFIG.DRIVE_FOLDER_ID);

  // Base64 解碼
  const decoded = Utilities.base64Decode(fileData);
  const blob = Utilities.newBlob(decoded, mimeType, fileName);

  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    fileId: file.getId(),
    fileUrl: file.getUrl(),
    directUrl: 'https://drive.google.com/uc?id=' + file.getId(),
    fileName: fileName,
  };
}

// ==================== 廠商名單管理 ====================

/**
 * 取得廠商名單（從 07_廠商名單 工作表）
 * 欄位：A: 廠商名稱 | B: 備註 | C: 最後使用時間 | D: 建立時間
 */
function getVendors() {
  const ss = SpreadsheetApp.openById(EXPENSE_CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(EXPENSE_CONFIG.SHEET_VENDORS);

  // 如果工作表不存在，自動建立
  if (!sheet) {
    sheet = ss.insertSheet(EXPENSE_CONFIG.SHEET_VENDORS);
    sheet.appendRow(['廠商名稱', '備註', '最後使用時間', '建立時間']);
    // 設定標題列格式
    const headerRange = sheet.getRange(1, 1, 1, 4);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#3A3A3A');
    headerRange.setFontColor('#E0E0E0');
    return [];
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1).map(row => ({
    name: row[0] || '',
    note: row[1] || '',
    usedAt: row[2] || '',
    createdAt: row[3] || '',
  })).filter(v => v.name);  // 過濾空白列
}

/**
 * 同步廠商名單到 07_廠商名單 工作表
 * 策略：完整覆寫（先清除再寫入），保持前端為單一真實來源
 */
function syncVendors(vendors) {
  if (!vendors || !Array.isArray(vendors)) {
    return { error: '無效的廠商資料' };
  }

  const ss = SpreadsheetApp.openById(EXPENSE_CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(EXPENSE_CONFIG.SHEET_VENDORS);

  // 如果工作表不存在，自動建立
  if (!sheet) {
    sheet = ss.insertSheet(EXPENSE_CONFIG.SHEET_VENDORS);
  }

  // 清除現有資料（保留標題列或重建）
  sheet.clear();
  sheet.appendRow(['廠商名稱', '備註', '最後使用時間', '建立時間']);

  // 設定標題列格式
  const headerRange = sheet.getRange(1, 1, 1, 4);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#3A3A3A');
  headerRange.setFontColor('#E0E0E0');

  // 批次寫入廠商資料
  if (vendors.length > 0) {
    const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
    const rows = vendors.map(v => [
      v.name || '',
      v.note || '',
      v.usedAt || '',
      v.createdAt || now,
    ]);

    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  }

  // 記錄操作日誌
  logOperation('廠商同步', '同步 ' + vendors.length + ' 筆廠商名單', '07_廠商名單', vendors.length);

  return {
    count: vendors.length,
    message: '廠商名單同步完成',
  };
}

// ==================== 月報自動寄送 ====================

/**
 * 設定月報自動寄送 Time Trigger
 * 每月 1 號上午 8 點自動寄出上月費用報表
 */
function setupAutoReportTrigger(email) {
  if (!email) return { error: '請提供收件人 Email' };

  // 驗證 Email 格式
  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailPattern.test(email)) {
    return { error: '無效的 Email 格式' };
  }

  // 儲存 email 到 Script Properties
  const props = PropertiesService.getScriptProperties();
  props.setProperty('REPORT_EMAIL', email);

  // 移除既有的月報觸發器（避免重複）
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendMonthlyReport') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 建立新的月報觸發器：每月 1 號上午 8 點
  ScriptApp.newTrigger('sendMonthlyReport')
    .timeBased()
    .onMonthDay(1)
    .atHour(8)
    .create();

  logOperation('月報設定', '自動寄送已設定，收件人：' + email, '系統設定', '');

  return {
    email: email,
    schedule: '每月 1 號 08:00',
    message: '月報自動寄送已設定完成',
  };
}

/**
 * 每月自動寄送費用報表（由 Time Trigger 呼叫）
 */
function sendMonthlyReport() {
  const props = PropertiesService.getScriptProperties();
  const email = props.getProperty('REPORT_EMAIL');
  if (!email) return;

  const ss = SpreadsheetApp.openById(EXPENSE_CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(EXPENSE_CONFIG.SHEET_EXPENSE);
  const data = sheet.getDataRange().getValues();

  // 取得上個月的年份與月份
  const now = new Date();
  const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const year = lastMonth.getFullYear();
  const month = lastMonth.getMonth() + 1;
  const monthStr = String(month).padStart(2, '0');

  // 篩選上月費用資料
  const rows = data.slice(1).filter(row => {
    if (row[COL.STATUS] === '作廢') return false;
    const d = new Date(row[COL.DATE]);
    return d.getFullYear() === year && (d.getMonth() + 1) === month;
  });

  if (rows.length === 0) {
    // 無資料也寄一封通知
    MailApp.sendEmail({
      to: email,
      subject: '堅心 Outrun — ' + year + '/' + monthStr + ' 月費用報表（無資料）',
      body: '本月無費用紀錄。\n\n詳見 Google Sheets：\n' + ss.getUrl(),
    });
    return;
  }

  // 彙總計算
  let totalAmount = 0;
  let totalTax = 0;
  const byParent = {};
  const byPayment = {};

  rows.forEach(row => {
    const amount = Number(row[COL.AMOUNT]) || 0;
    const tax = Number(row[COL.TAX]) || 0;
    totalAmount += amount;
    totalTax += tax;

    const parent = row[COL.PARENT_CAT] || '未分類';
    byParent[parent] = (byParent[parent] || 0) + amount;

    const payment = row[COL.PAYMENT] || '未記錄';
    byPayment[payment] = (byPayment[payment] || 0) + amount;
  });

  // 建立分類明細文字
  let categoryDetail = '';
  Object.entries(byParent).sort((a, b) => b[1] - a[1]).forEach(([cat, amt]) => {
    const pct = totalAmount > 0 ? ((amt / totalAmount) * 100).toFixed(1) : '0.0';
    categoryDetail += '  ・' + cat + '：NTD ' + amt.toLocaleString() + '（' + pct + '%）\n';
  });

  // 建立付款方式明細
  let paymentDetail = '';
  Object.entries(byPayment).sort((a, b) => b[1] - a[1]).forEach(([pay, amt]) => {
    paymentDetail += '  ・' + pay + '：NTD ' + amt.toLocaleString() + '\n';
  });

  // 組合信件內容
  const subject = '堅心 Outrun — ' + year + '/' + monthStr + ' 月費用報表';
  const body = [
    '═══════════════════════════════════',
    '堅心 Outrun Nutrition — ' + year + '年' + month + '月 費用月報',
    '═══════════════════════════════════',
    '',
    '📊 本月摘要',
    '  ・費用筆數：' + rows.length + ' 筆',
    '  ・費用總額：NTD ' + totalAmount.toLocaleString(),
    '  ・稅額合計：NTD ' + totalTax.toLocaleString(),
    '  ・未稅金額：NTD ' + (totalAmount - totalTax).toLocaleString(),
    '',
    '📂 分類明細',
    categoryDetail,
    '💳 付款方式',
    paymentDetail,
    '───────────────────────────────────',
    '📎 詳細資料請見 Google Sheets：',
    ss.getUrl(),
    '',
    '此信件由「堅心記帳系統」自動產生',
    '產生時間：' + Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd HH:mm'),
  ].join('\n');

  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: body,
  });

  logOperation('月報寄送', year + '/' + monthStr + ' 月報已寄至 ' + email, '自動寄送', totalAmount);
}

/**
 * 手動觸發測試月報寄送（Debug 用）
 * 在 Apps Script 編輯器中直接執行此函數即可測試
 */
function testSendMonthlyReport() {
  sendMonthlyReport();
}

// ==================== 操作日誌 ====================

// ==================== 報價單操作 ====================

/**
 * 確保工作表存在，不存在則建立並加上標題列
 */
function ensureSheet(sheetName, headers) {
  const ss = SpreadsheetApp.openById(EXPENSE_CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#3A3A3A');
    headerRange.setFontColor('#E0E0E0');
  }
  return sheet;
}

/**
 * 新增報價單
 */
function addQuotation(data) {
  const headers = ['報價編號','日期','有效期限','客戶名稱','聯繫人','電話','Email','地址',
    '小計','折扣類型','折扣值','折扣金額','稅率','稅額','總金額','備註','狀態','建立人員','建立時間'];
  const sheet = ensureSheet(EXPENSE_CONFIG.SHEET_QUOTATIONS, headers);
  const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');

  const q = data;
  const row = [
    q.id, q.date, q.validUntil,
    q.customer?.name || '', q.customer?.contact || '', q.customer?.phone || '',
    q.customer?.email || '', q.customer?.address || '',
    q.subtotal, q.discountType, q.discountValue, q.discountAmount,
    q.taxRate, q.taxAmount, q.total,
    q.notes || '', q.status || '草稿', q.recorder || '', now,
  ];
  sheet.appendRow(row);

  // 寫入明細行
  if (q.items && q.items.length > 0) {
    addDocItems(q.id, q.items);
  }

  logOperation('報價單新增', q.id + ' ' + (q.customer?.name || '') + ' NTD ' + q.total, '08_報價單', q.total);
  return { id: q.id, message: '報價單新增成功' };
}

/**
 * 更新報價單狀態
 */
function updateQuotationStatus(id, newStatus) {
  const ss = SpreadsheetApp.openById(EXPENSE_CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(EXPENSE_CONFIG.SHEET_QUOTATIONS);
  if (!sheet) return { error: '工作表不存在' };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.getRange(i + 1, 17).setValue(newStatus); // Q欄：狀態
      logOperation('報價單狀態', id + ' → ' + newStatus, '08_報價單', '');
      return { id: id, status: newStatus, message: '狀態已更新' };
    }
  }
  return { error: '找不到報價單 ' + id };
}

/**
 * 查詢報價單列表
 */
function getQuotations(params) {
  const ss = SpreadsheetApp.openById(EXPENSE_CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(EXPENSE_CONFIG.SHEET_QUOTATIONS);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  let quotations = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (params.year && params.month) {
      const d = new Date(row[1]);
      if (d.getFullYear() != params.year || (d.getMonth() + 1) != params.month) continue;
    }
    quotations.push({
      id: row[0], date: row[1], validUntil: row[2],
      customer: { name: row[3], contact: row[4], phone: row[5], email: row[6], address: row[7] },
      subtotal: row[8], discountType: row[9], discountValue: row[10], discountAmount: row[11],
      taxRate: row[12], taxAmount: row[13], total: row[14],
      notes: row[15], status: row[16], recorder: row[17], createdAt: row[18],
    });
  }
  return quotations;
}

// ==================== 採購單操作 ====================

/**
 * 新增採購單
 */
function addPurchaseOrder(data) {
  const headers = ['採購編號','日期','供應商名稱','聯繫人','電話','Email',
    '小計','稅率','稅額','總金額','預期交貨日','付款條件',
    '備註','狀態','關聯報價單','建立人員','建立時間','確認日期','交貨日期','結款日期'];
  const sheet = ensureSheet(EXPENSE_CONFIG.SHEET_PO, headers);
  const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');

  const po = data;
  const row = [
    po.id, po.date,
    po.vendor?.name || '', po.vendor?.contact || '', po.vendor?.phone || '', po.vendor?.email || '',
    po.subtotal, po.taxRate, po.taxAmount, po.total,
    po.expectedDelivery || '', po.paymentTerms || '',
    po.notes || '', po.status || '待確認', po.relatedQuotationId || '',
    po.recorder || '', now, '', '', '',
  ];
  sheet.appendRow(row);

  if (po.items && po.items.length > 0) {
    addDocItems(po.id, po.items.map(item => ({
      productName: item.itemName, spec: item.spec,
      qty: item.qty, unit: item.unit, unitPrice: item.unitPrice, subtotal: item.subtotal,
    })));
  }

  logOperation('採購單新增', po.id + ' ' + (po.vendor?.name || '') + ' NTD ' + po.total, '09_採購單', po.total);
  return { id: po.id, message: '採購單新增成功' };
}

/**
 * 更新採購單狀態
 */
function updatePOStatus(id, newStatus) {
  const ss = SpreadsheetApp.openById(EXPENSE_CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(EXPENSE_CONFIG.SHEET_PO);
  if (!sheet) return { error: '工作表不存在' };

  const data = sheet.getDataRange().getValues();
  const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.getRange(i + 1, 14).setValue(newStatus); // N欄：狀態
      // 自動填入日期
      if (newStatus === '已確認') sheet.getRange(i + 1, 18).setValue(now);
      if (newStatus === '已交貨') sheet.getRange(i + 1, 19).setValue(now);
      if (newStatus === '已結款') sheet.getRange(i + 1, 20).setValue(now);

      logOperation('採購單狀態', id + ' → ' + newStatus, '09_採購單', '');
      return { id: id, status: newStatus, message: '狀態已更新' };
    }
  }
  return { error: '找不到採購單 ' + id };
}

/**
 * 查詢採購單列表
 */
function getPurchaseOrders(params) {
  const ss = SpreadsheetApp.openById(EXPENSE_CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(EXPENSE_CONFIG.SHEET_PO);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  let orders = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (params.year && params.month) {
      const d = new Date(row[1]);
      if (d.getFullYear() != params.year || (d.getMonth() + 1) != params.month) continue;
    }
    orders.push({
      id: row[0], date: row[1],
      vendor: { name: row[2], contact: row[3], phone: row[4], email: row[5] },
      subtotal: row[6], taxRate: row[7], taxAmount: row[8], total: row[9],
      expectedDelivery: row[10], paymentTerms: row[11],
      notes: row[12], status: row[13], relatedQuotationId: row[14],
      recorder: row[15], createdAt: row[16],
    });
  }
  return orders;
}

// ==================== 單據明細操作 ====================

/**
 * 批次寫入明細行到 10_單據明細
 */
function addDocItems(docId, items) {
  const headers = ['單據編號','行號','品名','規格','數量','單位','單價','小計','備註','建立時間'];
  const sheet = ensureSheet(EXPENSE_CONFIG.SHEET_DOC_ITEMS, headers);
  const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');

  const rows = items.map((item, idx) => [
    docId, idx + 1,
    item.productName || item.itemName || '',
    item.spec || '', item.qty || 0, item.unit || '個',
    item.unitPrice || 0, item.subtotal || 0,
    item.notes || '', now,
  ]);

  if (rows.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rows.length, 10).setValues(rows);
  }
}

/**
 * 讀取單據明細行
 */
function getDocItems(docId) {
  const ss = SpreadsheetApp.openById(EXPENSE_CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(EXPENSE_CONFIG.SHEET_DOC_ITEMS);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1)
    .filter(row => row[0] === docId)
    .map(row => ({
      docId: row[0], lineNo: row[1],
      productName: row[2], spec: row[3],
      qty: row[4], unit: row[5],
      unitPrice: row[6], subtotal: row[7],
      notes: row[8], createdAt: row[9],
    }));
}

// ==================== 操作日誌 ====================

/**
 * 記錄操作到 05_操作紀錄（與現有 ERP 整合）
 */
function logOperation(type, content, scope, amount) {
  try {
    const ss = SpreadsheetApp.openById(EXPENSE_CONFIG.SPREADSHEET_ID);
    const logSheet = ss.getSheetByName('05_操作紀錄');
    if (!logSheet) return;

    const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
    const user = Session.getActiveUser().getEmail() || 'expense-tracker';

    logSheet.appendRow([now, user, type, content, scope || '', amount || '']);
  } catch (e) {
    // 日誌寫入失敗不影響主流程
  }
}
