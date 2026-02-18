import * as cheerio from 'cheerio';

/**
 * 解析 Cyberbiz 訂單郵件，提取訂單資訊
 * 回傳的結構對應 ERP「03_銷貨紀錄」的欄位
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
      channelSource: 'Cyberbiz',
      customerName: extractField(textContent, /收件人[：:]\s*(.+)/) ||
                     extractField(textContent, /姓名[：:]\s*(.+)/) ||
                     extractField(textContent, /買家[：:]\s*(.+)/),
      orderStatus: extractOrderStatus(subject, textContent),
      shippingDate: extractField(textContent, /出貨日期[：:]\s*(\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2})/),
      trackingNumber: extractField(textContent, /物流單號[：:]\s*([A-Za-z0-9]+)/) ||
                      extractField(textContent, /追蹤編號[：:]\s*([A-Za-z0-9]+)/),
      paymentStatus: extractPaymentStatus(subject, textContent),
      paymentDate: extractField(textContent, /付款日期[：:]\s*(\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2})/) ||
                   extractField(textContent, /收款日期[：:]\s*(\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2})/),
      discount: extractAmount(textContent, /折扣[：:]\s*-?(?:NT\$|＄|\$)?\s*([\d,]+)/) ||
                extractAmount(textContent, /優惠[：:]\s*-?(?:NT\$|＄|\$)?\s*([\d,]+)/),
      items: extractOrderItems($, textContent),
      rawSubject: subject,
      emailDate: date,
    };

    console.log(`[Parser] 成功解析訂單: ${orderNumber} (${order.items.length} 個品項)`);
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
 * 提取收款狀態
 */
function extractPaymentStatus(subject, text) {
  const combined = subject + ' ' + text;

  if (combined.includes('已付款') || combined.includes('付款完成') || combined.includes('付款成功')) {
    return '已收款';
  }
  if (combined.includes('未付款') || combined.includes('待付款')) {
    return '未收款';
  }
  if (combined.includes('退款')) {
    return '已退款';
  }

  return '';
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
 * 從 HTML 表格或文字中提取訂單品項
 * 每個品項包含: name, sku, quantity, unitPrice, subtotal
 */
function extractOrderItems($, textContent) {
  const items = [];

  // 嘗試從 HTML 表格提取（Cyberbiz 常見格式）
  $('table tr').each((i, row) => {
    if (i === 0) return; // 跳過表頭

    const cells = $(row).find('td');
    if (cells.length >= 3) {
      const nameCell = $(cells[0]).text().trim();
      const quantity = parseInt($(cells[1]).text().trim(), 10) || 0;
      const priceText = $(cells[2]).text().trim();
      const unitPrice = parseInt(priceText.replace(/[^\d]/g, ''), 10) || 0;

      if (nameCell && quantity > 0) {
        // 嘗試從品名中提取 SKU（常見格式: "SKU: ABC123 品名" 或 "(ABC123) 品名"）
        const skuMatch = nameCell.match(/(?:SKU[：:\s]*|[（(])([A-Za-z0-9\-]+)[）)]?/i);
        const sku = skuMatch ? skuMatch[1] : '';
        const name = skuMatch ? nameCell.replace(skuMatch[0], '').trim() : nameCell;

        items.push({
          name,
          sku,
          quantity,
          unitPrice,
          subtotal: quantity * unitPrice,
        });
      }
    }
  });

  // 若 HTML 表格解析無結果，嘗試從純文字解析
  if (items.length === 0) {
    const itemPattern = /(.+?)\s*[xX×]\s*(\d+)\s*(?:NT\$|＄|\$)?\s*([\d,]+)/g;
    let match;
    while ((match = itemPattern.exec(textContent)) !== null) {
      const rawName = match[1].trim();
      const quantity = parseInt(match[2], 10);
      const unitPrice = parseInt(match[3].replace(/,/g, ''), 10);

      const skuMatch = rawName.match(/(?:SKU[：:\s]*|[（(])([A-Za-z0-9\-]+)[）)]?/i);
      const sku = skuMatch ? skuMatch[1] : '';
      const name = skuMatch ? rawName.replace(skuMatch[0], '').trim() : rawName;

      items.push({
        name,
        sku,
        quantity,
        unitPrice,
        subtotal: quantity * unitPrice,
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
