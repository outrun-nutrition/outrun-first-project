import 'dotenv/config';

export const config = {
  google: {
    clientId: process.env.GOOGLE_CLIENT_ID,
    clientSecret: process.env.GOOGLE_CLIENT_SECRET,
    redirectUri: process.env.GOOGLE_REDIRECT_URI || 'http://localhost:3000/oauth2callback',
  },
  sheets: {
    spreadsheetId: process.env.SPREADSHEET_ID,
    sheetName: process.env.SHEET_NAME || '03_銷貨紀錄',
  },
  gmail: {
    cyberbizSender: process.env.CYBERBIZ_SENDER || 'noreply@cyberbiz.co',
  },
  cron: {
    schedule: process.env.CRON_SCHEDULE || '*/10 * * * *',
  },
  initialFetchDays: parseInt(process.env.INITIAL_FETCH_DAYS || '30', 10),
  tokenPath: new URL('../token.json', import.meta.url),
  scopes: [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/spreadsheets',
  ],
};
