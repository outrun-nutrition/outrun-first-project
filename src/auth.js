import { google } from 'googleapis';
import { createServer } from 'http';
import { URL } from 'url';
import { readFile, writeFile } from 'fs/promises';
import { config } from './config.js';

const { clientId, clientSecret, redirectUri } = config.google;

/**
 * 建立 OAuth2 客戶端
 */
export function createOAuth2Client() {
  return new google.auth.OAuth2(clientId, clientSecret, redirectUri);
}

/**
 * 載入已儲存的 token，若不存在則回傳 null
 */
export async function loadToken(oauth2Client) {
  try {
    const tokenData = await readFile(config.tokenPath, 'utf-8');
    const token = JSON.parse(tokenData);
    oauth2Client.setCredentials(token);

    // 設定 token 更新回呼，自動儲存 refreshed token
    oauth2Client.on('tokens', async (newTokens) => {
      const existing = JSON.parse(await readFile(config.tokenPath, 'utf-8'));
      const merged = { ...existing, ...newTokens };
      await writeFile(config.tokenPath, JSON.stringify(merged, null, 2));
      console.log('[Auth] Token 已自動更新並儲存');
    });

    return token;
  } catch {
    return null;
  }
}

/**
 * 透過本地 HTTP server 進行 OAuth2 授權流程
 */
async function authorizeViaLocalServer(oauth2Client) {
  const authUrl = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: config.scopes,
    prompt: 'consent',
  });

  console.log('\n========================================');
  console.log('請在瀏覽器中開啟以下連結進行授權：');
  console.log(authUrl);
  console.log('========================================\n');

  return new Promise((resolve, reject) => {
    const server = createServer(async (req, res) => {
      try {
        const url = new URL(req.url, `http://localhost:3000`);
        if (url.pathname !== '/oauth2callback') return;

        const code = url.searchParams.get('code');
        if (!code) {
          res.end('授權失敗：未取得授權碼');
          reject(new Error('未取得授權碼'));
          return;
        }

        const { tokens } = await oauth2Client.getToken(code);
        oauth2Client.setCredentials(tokens);

        await writeFile(config.tokenPath, JSON.stringify(tokens, null, 2));

        res.end('授權成功！你可以關閉這個視窗。');
        console.log('[Auth] 授權成功，Token 已儲存');

        server.close();
        resolve(tokens);
      } catch (err) {
        res.end('授權過程發生錯誤');
        reject(err);
      }
    });

    const port = new URL(redirectUri).port || 3000;
    server.listen(port, () => {
      console.log(`[Auth] 等待授權回呼中 (port ${port})...`);
    });
  });
}

/**
 * 取得已授權的 OAuth2 client
 * 若已有 token 則直接載入，否則啟動授權流程
 */
export async function getAuthenticatedClient() {
  const oauth2Client = createOAuth2Client();
  const token = await loadToken(oauth2Client);

  if (token) {
    console.log('[Auth] 已載入儲存的 Token');
    return oauth2Client;
  }

  console.log('[Auth] 未找到 Token，開始授權流程...');
  await authorizeViaLocalServer(oauth2Client);
  return oauth2Client;
}

// 若直接執行此檔案，啟動授權流程
const isDirectRun = process.argv[1]?.endsWith('auth.js');
if (isDirectRun) {
  getAuthenticatedClient()
    .then(() => {
      console.log('[Auth] 授權完成，可以開始使用系統');
      process.exit(0);
    })
    .catch((err) => {
      console.error('[Auth] 授權失敗:', err.message);
      process.exit(1);
    });
}
