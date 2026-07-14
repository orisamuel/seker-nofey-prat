/**
 * config.js — הגדרות המערכת (נטען בכל הדפים)
 * אחרי פריסת ה-Apps Script יש להדביק כאן את כתובת ה-Web App.
 */
const CONFIG = {
  // כתובת ה-Web App מ-Apps Script (Deploy → Manage deployments)
  SCRIPT_URL: 'PASTE_YOUR_DEPLOYED_SCRIPT_URL_HERE',

  // קישור ישיר לגיליון (לכפתור "פתח גיליון" בדשבורד)
  SHEETS_URL: 'https://docs.google.com/spreadsheets/d/PASTE_SHEET_ID/edit',

  APP_NAME: 'סקר התושבים השנתי',
  YISHUV: 'נופי פרת',

  // מצב פיתוח: כשה-SCRIPT_URL עוד לא הוגדר, הסקר עובד מקומית
  // מתוך survey-data.js ושומר תשובות ב-localStorage בלבד.
};
