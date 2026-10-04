/**
 * config.js — הגדרות המערכת (נטען בכל הדפים)
 * אחרי פריסת ה-Apps Script יש להדביק כאן את כתובת ה-Web App.
 */
const CONFIG = {
  // כתובת ה-Web App מ-Apps Script (Deploy → Manage deployments)
  SCRIPT_URL: 'https://script.google.com/macros/s/AKfycbyBz0e1WFXrhBSw_slxQ58ElikngdGcP9BgrC1dXRmCumXiQ1NkOVk3eGvSo03egJowfA/exec',

  // קישור ישיר לגיליון (לכפתור "פתח גיליון" בדשבורד)
  SHEETS_URL: 'https://docs.google.com/spreadsheets/d/1Ya4TerGGhxIcl2K5ziu6xjDssB9bAn2J2oz9Ka6xm9c/edit',

  APP_NAME: 'סקר התושבים השנתי',
  YISHUV: 'נופי פרת',

  // מצב פיתוח: כשה-SCRIPT_URL עוד לא הוגדר, הסקר עובד מקומית
  // מתוך survey-data.js ושומר תשובות ב-localStorage בלבד.
};
