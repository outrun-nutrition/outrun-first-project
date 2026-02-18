import * as cheerio from 'cheerio';

/**
 * 解析 Cyberbiz 訂單郵件，提取訂單資訊
 * @param {Object} email - 從 Gmail 取得的郵件物件
 * @returns {Object|null} 訂單資訊，解析失敗回傳 null
 */
export function parseCyberbizOrderEmail(email) {
  try {
    const { subject, body, date } = email;

    // 從主旨提取訂單編號
    const orderNumber = extractOrderNumber(subject);
    if (!orderNumber) {
      console.warn(`[Parser] 無法從主旨提取訂單編號: ${subject}`);
      return null;
    }

    // 解析 HTML 內容
    const $ = cheerio.load(body);
    const textContent = $.text();

    // 提取訂單資訊
    const order = {
      emailId: email.id,
      orderNumber,
      orderDate: extractOrderDate(textContent, date),
      customerName: extractField(textContent, /收件人[：:]\s*(.+)/),
      customerEmail: extractEmail(textContent),
      customerPhone: extractField(textContent, /電話[：:]\s*([\d\-+()]+)/),
      shippingAddress: extractField(textContent, /地址[：:]\s*(.+)/),
      paymentMethod: extractField(textContent, /付款方式[：:]\s*(.+)/),
      shippingMethod: extractField(textContent, /配送方式[：:]\s*(.+)/),
      orderStatus: extractOrderStatus(subject, textContent),
      items: extractOrderItems($, textContent),
      subtotal: extractAmount(textContent, /小計[：:]\s*(?:NT\$|＄|\$)?\s*([\d,]+)/),
      shippingFee: extractAmount(textContent, /運費[：:]\s*(?:NT\$|＄|\$)?\s*([\d,]+)/),
      discount: extractAmount(textContent, /折扣[：:]\s*-?(?:NT\$|＄|\$)?\s*([\d,]+)/),
      totalAmount: extractTotalAmount(textContent),
      rawSubject: subject,
      emailDate: date,
    };

    console.log(`[Parser] 成功解析訂單: ${orderNumber}`);
    return order;
  } catch (err) {
    console.error(`[Parser] 解析失敗:`, err.message);
    return null;
  }
}

/**
 * 從主旨提取訂單編號
 */
function extractOrderNumber(subject) {
  // Cyberbiz 訂單編號常見格式: #C12345, 訂單編號 C12345, Order #12345 等
  const patterns = [
    /(?:訂單編號|訂單號碼|Order\s*#?)[：:\s]*([A-Za-z0-9\-]+)/i,
    /#([A-Za-z]?\d{4,})/,
    /\b([A-Z]{1,3}\d{6,})\b/,
  ];

  for (const pattern of patterns) {
    const match = subject.match(pattern);
    if (match) return match[1];
  }

  return null;
}

/**
 * 從內容提取訂單日期
 */
function extractOrderDate(text, fallbackDate) {
  const patterns = [
    /訂單日期[：:]\s*(\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?)/,
    /下單時間[：:]\s*(\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?)/,
    /日期[：:]\s*(\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2})/,
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) return match[1];
  }

  // fallback: 使用郵件日期
  if (fallbackDate) {
    try {
      return new Date(fallbackDate).toISOString().split('T')[0];
    } catch {
      // ignore
    }
  }

  return '';
}

/**
 * 提取通用欄位
 */
function extractField(text, pattern) {
  const match = text.match(pattern);
  return match ? match[1].trim() : '';
}

/**
 * 提取 email
 */
function extractEmail(text) {
  const match = text.match(/[\w.\-+]+@[\w.\-]+\.\w+/);
  return match ? match[0] : '';
}

/**
 * 提取訂單狀態
 */
function extractOrderStatus(subject, text) {
  const statusKeywords = {
    '已成立': '新訂單',
    '新訂單': '新訂單',
    '已付款': '已付款',
    '付款完成': '已付款',
    '已出貨': '已出貨',
    '出貨通知': '已出貨',
    '已取消': '已取消',
    '取消': '已取消',
    '退貨': '退貨中',
    '退款': '已退款',
  };

  const combined = subject + ' ' + text;
  for (const [keyword, status] of Object.entries(statusKeywords)) {
    if (combined.includes(keyword)) return status;
  }

  return '新訂單';
}

/**
 * 提取金額
 */
function extractAmount(text, pattern) {
  const match = text.match(pattern);
  if (match) {
    return parseInt(match[1].replace(/,/g, ''), 10) || 0;
  }
  return 0;
}

/**
 * 提取訂單總金額
 */
function extractTotalAmount(text) {
  const patterns = [
    /總計[：:]\s*(?:NT\$|＄|\$)?\s*([\d,]+)/,
    /合計[：:]\s*(?:NT\$|＄|\$)?\s*([\d,]+)/,
    /Total[：:]\s*(?:NT\$|＄|\$)?\s*([\d,]+)/i,
    /應付金額[：:]\s*(?:NT\$|＄|\$)?\s*([\d,]+)/,
    /訂單金額[：:]\s*(?:NT\$|＄|\$)?\s*([\d,]+)/,
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) {
      return parseInt(match[1].replace(/,/g, ''), 10) || 0;
    }
  }

  return 0;
}

/**
 * 從 HTML 表格或文字中提取訂單品項
 */
function extractOrderItems($, textContent) {
  const items = [];

  // 嘗試從 HTML 表格提取
  $('table tr').each((i, row) => {
    if (i === 0) return; // 跳過表頭

    const cells = $(row).find('td');
    if (cells.length >= 3) {
      const name = $(cells[0]).text().trim();
      const quantity = parseInt($(cells[1]).text().trim(), 10) || 0;
      const priceText = $(cells[2]).text().trim();
      const price = parseInt(priceText.replace(/[^\d]/g, ''), 10) || 0;

      if (name && quantity > 0) {
        items.push({ name, quantity, price });
      }
    }
  });

  // 若 HTML 表格解析無結果，嘗試從純文字解析
  if (items.length === 0) {
    const itemPattern = /(.+?)\s*[xX×]\s*(\d+)\s*(?:NT\$|＄|\$)?\s*([\d,]+)/g;
    let match;
    while ((match = itemPattern.exec(textContent)) !== null) {
      items.push({
        name: match[1].trim(),
        quantity: parseInt(match[2], 10),
        price: parseInt(match[3].replace(/,/g, ''), 10),
      });
    }
  }

  return items;
}

/**
 * 批量解析郵件
 * @param {Array} emails - 郵件陣列
 * @returns {Array} 成功解析的訂單陣列
 */
export function parseEmails(emails) {
  const orders = [];
  let failCount = 0;

  for (const email of emails) {
    const order = parseCyberbizOrderEmail(email);
    if (order) {
      orders.push(order);
    } else {
      failCount++;
    }
  }

  console.log(`[Parser] 共解析 ${orders.length} 筆訂單，${failCount} 封郵件解析失敗`);
  return orders;
}
