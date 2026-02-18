import { google } from 'googleapis';
import { config } from './config.js';

/**
 * 從 Gmail 搜尋 Cyberbiz 訂單郵件
 * @param {import('googleapis').Auth.OAuth2Client} auth - 已授權的 OAuth2 client
 * @param {Object} options
 * @param {string} [options.afterDate] - 搜尋此日期之後的郵件 (YYYY/MM/DD)
 * @param {string} [options.pageToken] - 分頁 token
 * @returns {Promise<{messages: Array, nextPageToken: string|null}>}
 */
export async function searchCyberbizEmails(auth, options = {}) {
  const gmail = google.gmail({ version: 'v1', auth });

  // 建立搜尋查詢
  const queryParts = [
    `from:${config.gmail.cyberbizSender}`,
    'subject:訂單',
  ];

  if (options.afterDate) {
    queryParts.push(`after:${options.afterDate}`);
  }

  const query = queryParts.join(' ');
  console.log(`[Gmail] 搜尋查詢: ${query}`);

  const res = await gmail.users.messages.list({
    userId: 'me',
    q: query,
    maxResults: 100,
    pageToken: options.pageToken || undefined,
  });

  const messages = res.data.messages || [];
  console.log(`[Gmail] 找到 ${messages.length} 封郵件`);

  return {
    messages,
    nextPageToken: res.data.nextPageToken || null,
  };
}

/**
 * 取得單封郵件的完整內容
 * @param {import('googleapis').Auth.OAuth2Client} auth
 * @param {string} messageId
 * @returns {Promise<Object>} 郵件內容
 */
export async function getEmailContent(auth, messageId) {
  const gmail = google.gmail({ version: 'v1', auth });

  const res = await gmail.users.messages.get({
    userId: 'me',
    id: messageId,
    format: 'full',
  });

  const message = res.data;
  const headers = message.payload.headers;

  const getHeader = (name) =>
    headers.find((h) => h.name.toLowerCase() === name.toLowerCase())?.value || '';

  // 解碼郵件 body
  const body = extractBody(message.payload);

  return {
    id: message.id,
    threadId: message.threadId,
    subject: getHeader('Subject'),
    from: getHeader('From'),
    date: getHeader('Date'),
    body,
    internalDate: message.internalDate,
  };
}

/**
 * 從 payload 中遞迴提取 HTML 或純文字 body
 */
function extractBody(payload) {
  // 如果 payload 有直接的 body data
  if (payload.body?.data) {
    return decodeBase64Url(payload.body.data);
  }

  // 遞迴搜尋 multipart 結構
  if (payload.parts) {
    // 優先取 HTML
    for (const part of payload.parts) {
      if (part.mimeType === 'text/html' && part.body?.data) {
        return decodeBase64Url(part.body.data);
      }
    }
    // fallback 純文字
    for (const part of payload.parts) {
      if (part.mimeType === 'text/plain' && part.body?.data) {
        return decodeBase64Url(part.body.data);
      }
    }
    // 繼續遞迴
    for (const part of payload.parts) {
      const result = extractBody(part);
      if (result) return result;
    }
  }

  return '';
}

/**
 * 解碼 Gmail 的 base64url 編碼
 */
function decodeBase64Url(encoded) {
  const base64 = encoded.replace(/-/g, '+').replace(/_/g, '/');
  return Buffer.from(base64, 'base64').toString('utf-8');
}

/**
 * 取得所有 Cyberbiz 訂單郵件（自動處理分頁）
 * @param {import('googleapis').Auth.OAuth2Client} auth
 * @param {Object} options
 * @param {string} [options.afterDate]
 * @returns {Promise<Array>} 所有郵件內容
 */
export async function fetchAllCyberbizEmails(auth, options = {}) {
  const allEmails = [];
  let pageToken = null;

  do {
    const { messages, nextPageToken } = await searchCyberbizEmails(auth, {
      ...options,
      pageToken,
    });

    // 逐封取得完整內容
    for (const msg of messages) {
      const content = await getEmailContent(auth, msg.id);
      allEmails.push(content);
    }

    pageToken = nextPageToken;
  } while (pageToken);

  console.log(`[Gmail] 共取得 ${allEmails.length} 封 Cyberbiz 訂單郵件`);
  return allEmails;
}
