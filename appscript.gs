// ============================================================
// סקר התושבים — נופי פרת · Google Apps Script Backend
// הגיליון הוא מסד הנתונים: פרקים, שאלות, תשובות, הגרלה, הגדרות, מיילים (מוצפנים)
// ============================================================
//
// פריסה (פירוט מלא ב-CLAUDE.md, עם clasp מהמחשב של אורי):
//   1. clasp create --type sheets → נוצרים גיליון + סקריפט. ה-ID של הגיליון נכנס ל-SHEET_ID.
//   2. clasp push -f && clasp deploy → כתובת ה-/exec נכנסת ל-config.js → SCRIPT_URL.
//   3. פעם אחת: לפתוח את העורך (clasp open), להריץ את setup() ולאשר הרשאות.
//      setup יוצר את הטאבים, סיסמת צוות אקראית (מודפסת בלוג ונשמרת בטאב "הגדרות")
//      וטריגר חימום.
//   4. דף הניהול באתר (admin.html) → כניסה בסיסמה → "שמירה". בשרת ריק נטען השאלון ההתחלתי מ-survey-data.js.
//
// ⚠️ אחרי כל עריכה של הקובץ הזה: ./deploy.sh "מה השתנה" (גרסה חדשה על אותה כתובת).
// ============================================================

const SHEET_ID = '1Ya4TerGGhxIcl2K5ziu6xjDssB9bAn2J2oz9Ka6xm9c';

// שמות הטאבים
const T_CHAPTERS = 'פרקים';
const T_QUESTIONS = 'שאלות';
const T_RAW = 'תשובות גולמי';
const T_FLAT = 'תוצאות';
const T_RAFFLE = 'הגרלה';
const T_SETTINGS = 'הגדרות';
const T_EMAILS = 'מיילים';

// Schema: זמן(0), קוד עונה(1), פרק(2), תשובות JSON(3) — append-only
const RAW_HEADERS = ['זמן', 'קוד עונה', 'פרק', 'תשובות (JSON)'];
// Schema: שם(0), טלפון(1) — בלי זמן ובלי קוד עונה, ממוין לפי שם (אי אפשר לקשר לתשובות)
const RAFFLE_HEADERS = ['שם', 'טלפון'];

// Schema: זמן(0), מייל מוצפן(1), קוד עונה(2). הכתובת עצמה לא נשמרת בשום מקום
const EMAIL_HEADERS = ['זמן', 'מייל (מוצפן)', 'קוד עונה'];
const MAX_NEW_EMAILS_PER_HOUR = 400; // הגנה מהצפה של הטאב

// "פרק" מיוחד בטאב הגולמי: העונה לחץ "התחלה מחדש", כל התשובות שלו לא נספרות
const DISCARD_CH = '_בוטל';

const T_NOTICE = 'קרא אותי';
const SITE = 'https://sekernofey.online/';

// הסקר הציבורי נשמר במטמון (קריאה ופענוח של הטאבים לוקחים כמה שניות). מתחלף בכל שמירה בדף הניהול
const SURVEY_CACHE = 'survey_public_v1';
const SURVEY_CACHE_SECS = 21600; // המקסימום ש-Apps Script מאפשר (6 שעות)

// ============================================================
// עזרים
// ============================================================

function getSpreadsheet() { return SpreadsheetApp.openById(SHEET_ID); }

function ensureSheet(name, headers) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length) sheet.appendRow(headers);
  }
  return sheet;
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function parseJsonSafe(s, fallback) {
  if (s === null || s === undefined || s === '') return fallback;
  try { return JSON.parse(s); } catch (e) { return fallback; }
}

function normRid(rid) {
  rid = String(rid || '').toUpperCase().trim();
  return /^[A-Z2-9]{8}$/.test(rid) ? rid : null;
}

// טקסט חופשי שמתחיל ב- = + - @ הגיליון מפרש כנוסחה. גרש בהתחלה שומר אותו כטקסט.
function safeCell(v) {
  if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
  return v;
}

// טקסט מדף הניהול: מה שנראה כמו נוסחה, מספר או תאריך הגיליון היה הופך. גרש בהתחלה שומר אותו כמו שהוא
function textCell(v) {
  if (typeof v !== 'string') return v;
  return /^[=+\-@']/.test(v) || /^[\d\s.,:\/]+$/.test(v) ? "'" + v : v;
}

// טאב חדש נוצר עם 26 עמודות, וטבלת התוצאות צריכה עמודה לכל שאלה. בלי זה הכתיבה נכשלת.
function ensureCols(sheet, n) {
  const max = sheet.getMaxColumns();
  if (max < n) sheet.insertColumnsAfter(max, n - max);
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// ── הגדרות: key → value ──────────────────────────────────
// אין סיסמת ברירת מחדל בקוד (הריפו ציבורי). בפעם הראשונה נוצרת סיסמה אקראית בטאב "הגדרות".
function settingsDefaults() {
  return [
    ['surveyOpen', 'כן', 'האם הסקר פתוח למענה (כן/לא)'],
    // מתחילה באותיות: סיסמה שכולה ספרות (או "123e45") הגיליון הופך למספר, ואז היא לא תואמת למה שמקלידים
    ['dashboardPassword', 'np' + Utilities.getUuid().replace(/-/g, '').slice(0, 8), 'סיסמת הצוות לדשבורד ולדף הניהול. נוצרה אקראית, מחליפים בדף הניהול'],
    ['publicReport', 'לא', 'האם הדוח הציבורי פעיל (כן/לא)'],
  ];
}

function getSettings() {
  let sheet = getSpreadsheet().getSheetByName(T_SETTINGS);
  if (!sheet || sheet.getLastRow() <= 1) {
    withLock(function () {
      sheet = ensureSheet(T_SETTINGS, ['מפתח', 'ערך', 'הסבר']);
      if (sheet.getLastRow() <= 1) settingsDefaults().forEach(function (r) { sheet.appendRow(r); });
    });
  }
  const data = sheet.getDataRange().getValues();
  const out = {};
  for (let i = 1; i < data.length; i++) out[String(data[i][0]).trim()] = String(data[i][1]).trim();
  return out;
}

function upsertSetting(sheet, key, value, note) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) {
      sheet.getRange(i + 1, 2).setValue(textCell(value));
      return;
    }
  }
  sheet.appendRow([key, textCell(value), note || '']);
}

function validEmail(e) {
  return e.length <= 254 && /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(e);
}

// ב-Gmail נקודות ו"+תוספת" מגיעות לאותה תיבה, אז הן לא פותחות סקר נוסף
function normEmail(email) {
  email = String(email || '').trim().toLowerCase();
  if (!validEmail(email)) return null;
  const at = email.lastIndexOf('@');
  let user = email.slice(0, at), domain = email.slice(at + 1);
  if (domain === 'gmail.com' || domain === 'googlemail.com') {
    user = user.split('+')[0].replace(/\./g, '');
    domain = 'gmail.com';
  }
  return user ? user + '@' + domain : null;
}

// המפתח הסודי של המיילים יושב בהגדרות הסקריפט, לא בגיליון: מי שרואה את הגיליון
// לא יכול לחשב את הקוד של כתובת מסוימת ולבדוק אם היא ענתה
function emailSecret() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('EMAIL_SECRET');
  if (!secret) {
    secret = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
    props.setProperty('EMAIL_SECRET', secret);
  }
  return secret;
}

function emailKey(norm) {
  return Utilities.computeHmacSha256Signature(norm, emailSecret(), Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

// ── טלפון (להגרלה) ─────────────────────────────────────
function phoneDigits(phone) {
  const d = String(phone || '').replace(/[^0-9]/g, '');
  return /^0\d{8,9}$/.test(d) ? d : null;
}

// הערכת תנאי showIf — זהה ללוגיקה בצד הלקוח (utils.js)
function evalCond(cond, answers) {
  if (!cond || !cond.q) return true;
  const val = answers[cond.q];
  if (val === undefined || val === null || val === '' ||
      (Array.isArray(val) && val.length === 0)) return true;
  const vals = cond.vals || cond.in || cond.any || cond.notIn || [];
  const neg = !!cond.neg || !!cond.notIn;
  const match = Array.isArray(val)
    ? vals.some(function (v) { return val.indexOf(v) !== -1; })
    : vals.indexOf(val) !== -1;
  return neg ? !match : match;
}

// ── סוגי שאלות: בגיליון בעברית, בקוד באנגלית ──────────────
const TYPE_TO_HE = {
  scale: 'סולם', radio: 'בחירה אחת', checkbox: 'בחירה מרובה',
  text: 'טקסט קצר', textarea: 'טקסט ארוך', number: 'מספר', rank: 'דירוג',
};

function normalizeType(t) {
  t = String(t || '').trim();
  if (TYPE_TO_HE[t]) return t; // כבר באנגלית
  if (t === 'סולם 1-10') return 'scale'; // השם הישן
  for (var k in TYPE_TO_HE) if (TYPE_TO_HE[k] === t) return k;
  return t;
}

// ── תנאי הצגה: בגיליון בתחביר פשוט, בקוד כאובייקט ─────────
//   "syn_teacher = מישהו מהיישוב"  · "about_gender != גבר"  · "about_kids = נוער (ז׳–י״ב) | יסודי (א׳–ו׳)"
function parseCondition(s) {
  s = String(s || '').trim();
  if (!s) return null;
  if (s.charAt(0) === '{') return parseJsonSafe(s, null); // תאימות ל-JSON ישן
  var neg = s.indexOf('!=') !== -1;
  var parts = s.split(neg ? '!=' : '=');
  if (parts.length < 2) return null;
  var q = parts[0].trim();
  var vals = parts.slice(1).join('=').split('|')
    .map(function (v) { return v.trim(); })
    .filter(function (v) { return v !== ''; });
  if (!q || !vals.length) return null;
  var c = { q: q, vals: vals };
  if (neg) c.neg = true;
  return c;
}

function condToString(c) {
  if (!c || !c.q) return '';
  var vals = c.vals || c.in || c.any || c.notIn || [];
  var neg = !!c.neg || !!c.notIn;
  return c.q + ' ' + (neg ? '!=' : '=') + ' ' + vals.join(' | ');
}

// ============================================================
// הפעלה ראשונה — מריצים פעם אחת מהעורך (▶ setup), זה גם שלב אישור ההרשאות
// ============================================================

function setup() {
  const settings = getSettings();
  ensureSheet(T_RAW, RAW_HEADERS);
  ensureSheet(T_RAFFLE, RAFFLE_HEADERS);
  ensureSheet(T_EMAILS, EMAIL_HEADERS);
  emailSecret();
  writeNoticeSheet();
  protectTabs();
  const hasWarm = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'keepWarm'; });
  if (!hasWarm) ScriptApp.newTrigger('keepWarm').timeBased().everyMinutes(10).create();
  Logger.log('מוכן. סיסמת הצוות (גם בטאב "הגדרות"): ' + settings.dashboardPassword);
}

// טריגר כל 10 דקות, מקצר את ההמתנה של הגולש הראשון אחרי שקט
function keepWarm() { Logger.log('warm ' + new Date().toISOString()); }

// ============================================================
// מבנה הסקר: פרקים + שאלות
// ============================================================

const CHAPTER_HEADERS = ['id', 'כותרת', 'אייקון', 'תיאור', 'תנאי הצגה', 'שער', 'פעיל', 'סדר', 'קטגוריה', 'פתיח', 'סיום', 'בלי דילוג'];
const QUESTION_HEADERS = ['פרק', 'id', 'סוג', 'שאלה', 'עזרה', 'אפשרויות ( | )', 'אחר', 'בלעדי', 'מינ', 'מקס', 'תווית מינ', 'תווית מקס', 'תנאי הצגה', 'פעיל', 'סדר', 'חובה', 'לא יודע'];

// תשובת "לא יודע/ת" בסולם: נשמרת כטקסט, ולא נכנסת לממוצעים
const DONT_KNOW = 'לא יודע/ת';

// ============================================================
// דף הניהול (admin.html): כל העריכה של השאלון, הטקסטים וההגדרות, בסיסמת הצוות.
// הטאבים "פרקים" ו"שאלות" הם רק אחסון (וגיבוי קריא). לא עורכים אותם ידנית.
// ============================================================

function checkPassword(password) {
  return !!password && String(password) === getSettings().dashboardPassword;
}

// השאלון המלא, כולל נושאים ושאלות מוסתרים. legacy = עוד לא נשמר מדף הניהול
function getAdmin(password) {
  try {
    if (!checkPassword(password)) return { success: false, message: 'סיסמה שגויה' };
    const settings = getSettings();
    return {
      success: true,
      survey: loadSurvey(true),
      rev: settings.surveyRev || '',
      legacy: !settings.surveyRev,
      settings: { surveyOpen: settings.surveyOpen !== 'לא', publicReport: settings.publicReport === 'כן' },
      answered: answeredCounts(),
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// כמה ענו על כל שאלה (מהטבלה השטוחה), כדי שבמחיקה יהיה ברור מה יוצא מהתוצאות
function answeredCounts() {
  const out = {};
  const flat = getSpreadsheet().getSheetByName(T_FLAT);
  if (!flat || flat.getLastRow() <= 1) return out;
  const data = flat.getDataRange().getValues();
  for (let j = 2; j < data[0].length; j++) {
    let n = 0;
    for (let i = 1; i < data.length; i++) if (data[i][j] !== '' && data[i][j] !== null) n++;
    if (n) out[String(data[0][j])] = n;
  }
  return out;
}

// rev: הגרסה שהדף טען. אם מישהו אחר שמר בינתיים, לא דורסים (אלא אם force)
function saveSurvey(password, surveyJson, rev, force) {
  try {
    if (!checkPassword(password)) return { success: false, message: 'סיסמה שגויה' };
    const survey = parseJsonSafe(surveyJson, null);
    const errors = validateSurvey(survey);
    if (errors.length) return { success: false, message: 'יש בשאלון דברים לתקן', errors: errors.slice(0, 30) };
    return withLock(function () {
      if (!force && (getSettings().surveyRev || '') !== String(rev || '')) {
        return { success: false, conflict: true, message: 'מישהו אחר שמר שינויים בינתיים' };
      }
      writeSurveyTabs(survey);
      const newRev = 'r' + Utilities.getUuid().replace(/-/g, '').slice(0, 12);
      upsertSetting(getSpreadsheet().getSheetByName(T_SETTINGS), 'surveyRev', newRev, 'מזהה של השמירה האחרונה בדף הניהול');
      CacheService.getScriptCache().remove(SURVEY_CACHE);
      publicSurvey(); // ממלא את המטמון בגרסה החדשה
      return { success: true, rev: newRev };
    });
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function saveSettings(password, surveyOpen, publicReport, newPassword) {
  try {
    if (!checkPassword(password)) return { success: false, message: 'סיסמה שגויה' };
    newPassword = newPassword === undefined || newPassword === null ? '' : String(newPassword).trim();
    if (newPassword && newPassword.length < 8) return { success: false, message: 'סיסמה חדשה: לפחות 8 תווים' };
    return withLock(function () {
      const sheet = getSpreadsheet().getSheetByName(T_SETTINGS);
      if (surveyOpen === true || surveyOpen === false) upsertSetting(sheet, 'surveyOpen', surveyOpen ? 'כן' : 'לא');
      if (publicReport === true || publicReport === false) upsertSetting(sheet, 'publicReport', publicReport ? 'כן' : 'לא');
      if (newPassword) upsertSetting(sheet, 'dashboardPassword', newPassword);
      return { success: true };
    });
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// בדיקות לפני שמירה. ההודעות מוצגות בדף הניהול כמו שהן
function validateSurvey(s) {
  if (!s || typeof s !== 'object' || !Array.isArray(s.chapters)) return ['מבנה שאלון לא תקין'];
  if (JSON.stringify(s).length > 300000) return ['השאלון ארוך מדי'];
  if (s.meta !== undefined && (!s.meta || typeof s.meta !== 'object' || JSON.stringify(s.meta).length > 40000)) {
    return ['הטקסטים הכלליים לא תקינים או ארוכים מדי'];
  }
  const errs = [];
  const chIds = {}, all = {};
  s.chapters.forEach(function (ch, ci) {
    const where = '"' + String((ch && ch.title) || ('נושא ' + (ci + 1))) + '"';
    if (!ch || !/^[a-z0-9_]{1,40}$/.test(String(ch.id || ''))) { errs.push(where + ': מזהה לא תקין'); return; }
    if (chIds[ch.id]) errs.push(where + ': מזהה כפול');
    chIds[ch.id] = true;
    if (!String(ch.title || '').trim()) errs.push('נושא ' + (ci + 1) + ': חסרה כותרת');
    if (!Array.isArray(ch.questions)) { errs.push(where + ': רשימת השאלות לא תקינה'); return; }
    ch.questions.forEach(function (q, qi) {
      const w = where + ', שאלה ' + (qi + 1);
      if (!q || !/^[a-z0-9_]{1,60}$/.test(String(q.id || ''))) { errs.push(w + ': מזהה לא תקין'); return; }
      if (all[q.id]) errs.push(w + ': מזהה כפול');
      all[q.id] = q;
      if (!TYPE_TO_HE[q.type]) errs.push(w + ': סוג לא מוכר');
      if (!String(q.text || '').trim()) errs.push(w + ': חסר נוסח');
      if (q.type === 'radio' || q.type === 'checkbox' || q.type === 'rank') {
        const o = Array.isArray(q.opts) ? q.opts.map(function (x) { return String(x).trim(); }) : [];
        if (o.length < 2 || o.some(function (x) { return !x; })) errs.push(w + ': צריך לפחות שתי אפשרויות, בלי אפשרות ריקה');
        if (o.some(function (x) { return x.indexOf('|') !== -1; })) errs.push(w + ': אי אפשר להשתמש בתו | באפשרות');
        if (new Set(o).size !== o.length) errs.push(w + ': אותה אפשרות מופיעה פעמיים');
        if (q.exclusive && o.indexOf(q.exclusive) === -1) errs.push(w + ': האפשרות שמבטלת את השאר לא ברשימה');
      }
      if (q.type === 'scale') {
        const lo = Number(q.min), hi = Number(q.max);
        if (!(lo >= 0 && hi > lo && hi - lo <= 10)) errs.push(w + ': טווח הסולם לא תקין');
      }
    });
  });
  // תנאי הצגה: מפנים לשאלה קיימת ולאפשרויות שלה
  s.chapters.forEach(function (ch) {
    if (!ch || !Array.isArray(ch.questions)) return;
    [ch].concat(ch.questions).forEach(function (x) {
      if (!x || !x.showIf) return;
      const name = '"' + String(x.title || x.text || '').slice(0, 40) + '"';
      const t = all[x.showIf.q];
      const vals = x.showIf.vals || x.showIf.in || x.showIf.any || x.showIf.notIn || [];
      if (!t) errs.push(name + ': תנאי ההצגה מפנה לשאלה שלא קיימת');
      else if (!vals.length || vals.some(function (v) { return (t.opts || []).indexOf(v) === -1; })) errs.push(name + ': תנאי ההצגה מפנה לתשובה שלא קיימת');
    });
  });
  return errs;
}

// השאלון לטאבים "פרקים" ו"שאלות", והטקסטים הכלליים לטאב "הגדרות"
function writeSurveyTabs(survey) {
  const chRows = [], qRows = [];
  survey.chapters.forEach(function (ch, ci) {
    chRows.push([
      ch.id, ch.title, ch.icon || '', ch.desc || '',
      condToString(ch.showIf),
      (ch.gate || ch.core) ? 'כן' : '', ch.active === false ? 'לא' : 'כן', ci + 1, ch.cat || '',
      ch.intro || '', ch.outro || '',
      ch.noSkip ? 'כן' : '',
    ].map(textCell));
    ch.questions.forEach(function (q, qi) {
      qRows.push([
        ch.id, q.id, TYPE_TO_HE[q.type] || q.type, q.text, q.help || '',
        (q.opts || []).join(' | '),
        q.other ? 'כן' : '', q.exclusive || '',
        q.min !== undefined && q.min !== '' ? Number(q.min) : '', q.max !== undefined && q.max !== '' ? Number(q.max) : '',
        q.minLabel || '', q.maxLabel || '',
        condToString(q.showIf),
        q.active === false ? 'לא' : 'כן', qi + 1,
        q.required ? 'כן' : '',
        q.dontKnow ? 'כן' : '',
      ].map(textCell));
    });
  });
  const chSheet = ensureSheet(T_CHAPTERS, CHAPTER_HEADERS);
  const qSheet = ensureSheet(T_QUESTIONS, QUESTION_HEADERS);
  chSheet.clearContents(); chSheet.appendRow(CHAPTER_HEADERS);
  qSheet.clearContents(); qSheet.appendRow(QUESTION_HEADERS);
  if (chRows.length) chSheet.getRange(2, 1, chRows.length, CHAPTER_HEADERS.length).setValues(chRows);
  if (qRows.length) qSheet.getRange(2, 1, qRows.length, QUESTION_HEADERS.length).setValues(qRows);

  upsertSetting(ensureSheet(T_SETTINGS, ['מפתח', 'ערך', 'הסבר']), 'meta', JSON.stringify(survey.meta || {}), 'הטקסטים הכלליים של הסקר (נערכים בדף הניהול)');
  ensureSheet(T_RAW, RAW_HEADERS);
  ensureSheet(T_RAFFLE, RAFFLE_HEADERS);
  ensureSheet(T_EMAILS, EMAIL_HEADERS);
  writeNoticeSheet();
  protectTabs();
  rebuildFlatHeaders();
}

// במקום מדריך עריכה: הגיליון הוא מסד הנתונים, ועורכים רק בדף הניהול
function writeNoticeSheet() {
  const ss = getSpreadsheet();
  const old = ss.getSheetByName('מדריך עריכה');
  if (old) ss.deleteSheet(old);
  const sheet = ss.getSheetByName(T_NOTICE) || ss.insertSheet(T_NOTICE);
  sheet.clearContents();
  sheet.getRange(1, 1, 3, 1).setValues([
    ['הגיליון הזה הוא מסד הנתונים של הסקר. לא עורכים בו.'],
    ['שאלות, טקסטים והגדרות: בדף הניהול, ' + SITE + 'admin.html'],
    ['תוצאות: בדשבורד, ' + SITE + 'dashboard.html'],
  ]);
  sheet.setColumnWidth(1, 700);
}

// עריכה ידנית בגיליון עוקפת את דף הניהול ואת המטמון. מי שמנסה מקבל אזהרה (לא חסימה)
function protectTabs() {
  const ss = getSpreadsheet();
  [T_CHAPTERS, T_QUESTIONS, T_SETTINGS, T_RAW, T_FLAT, T_EMAILS].forEach(function (name) {
    const sh = ss.getSheetByName(name);
    if (sh && !sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).length) {
      sh.protect().setDescription('נערך מדף הניהול באתר').setWarningOnly(true);
    }
  });
}

// קריאת מבנה הסקר מהגיליון
// includeInactive: גם נושאים ושאלות מוסתרים (active: false), לדף הניהול
function loadSurvey(includeInactive) {
  const chSheet = getSpreadsheet().getSheetByName(T_CHAPTERS);
  const qSheet = getSpreadsheet().getSheetByName(T_QUESTIONS);
  if (!chSheet || !qSheet || chSheet.getLastRow() <= 1) return null;

  const settings = getSettings();
  const meta = parseJsonSafe(settings.meta, {});

  const qData = qSheet.getDataRange().getValues();
  // עמודת "חובה" נוספה ב-10/2026. בגיליון שנזרע לפניה לא מחזירים כלום, והאתר לוקח את זה מ-survey-data.js
  const hasRequired = String(qData[0][15] || '').trim() === 'חובה';
  const byChapter = {};
  for (let i = 1; i < qData.length; i++) {
    const r = qData[i];
    if (!String(r[1]).trim()) continue;            // שורה ריקה
    const qOff = String(r[13]).trim() === 'לא';   // לא פעיל
    if (qOff && !includeInactive) continue;
    const q = { id: String(r[1]).trim(), type: normalizeType(r[2]), text: String(r[3]), order: Number(r[14]) || i };
    if (qOff) q.active = false;
    if (r[4]) q.help = String(r[4]);
    if (r[5]) q.opts = String(r[5]).split('|').map(function (s) { return s.trim(); }).filter(String);
    if (String(r[6]).trim() === 'כן') q.other = true;
    if (r[7]) q.exclusive = String(r[7]).trim();
    if (r[8] !== '') q.min = Number(r[8]);
    if (r[9] !== '') q.max = Number(r[9]);
    if (r[10]) q.minLabel = String(r[10]);
    if (r[11]) q.maxLabel = String(r[11]);
    const cond = parseCondition(String(r[12]));
    if (cond) q.showIf = cond;
    if (hasRequired) q.required = String(r[15]).trim() === 'כן';
    if (String(r[16] || '').trim() === 'כן') q.dontKnow = true;
    const chId = String(r[0]).trim();
    (byChapter[chId] = byChapter[chId] || []).push(q);
  }
  // עמודת "סדר" קובעת את המיקום בתוך הפרק (גם לשאלה שנוספה בתחתית הטאב)
  Object.keys(byChapter).forEach(function (k) {
    byChapter[k].sort(function (a, b) { return a.order - b.order; });
    byChapter[k].forEach(function (q) { delete q.order; });
  });

  const chData = chSheet.getDataRange().getValues();
  const chapters = [];
  for (let i = 1; i < chData.length; i++) {
    const r = chData[i];
    if (!String(r[0]).trim()) continue;
    const chOff = String(r[6]).trim() === 'לא'; // לא פעיל
    if (chOff && !includeInactive) continue;
    const ch = { id: String(r[0]).trim(), title: String(r[1]), questions: byChapter[String(r[0]).trim()] || [], order: Number(r[7]) || i };
    if (chOff) ch.active = false;
    if (r[2]) ch.icon = String(r[2]);
    if (r[3]) ch.desc = String(r[3]);
    const cond = parseCondition(String(r[4]));
    if (cond) ch.showIf = cond;
    if (String(r[5]).trim() === 'כן') ch.gate = true;
    if (r.length > 8 && r[8]) ch.cat = String(r[8]).trim();
    if (r.length > 9 && r[9]) ch.intro = String(r[9]);
    if (r.length > 10 && r[10]) ch.outro = String(r[10]);
    if (String(r[11] || '').trim() === 'כן') ch.noSkip = true;
    chapters.push(ch);
  }
  chapters.sort(function (a, b) { return a.order - b.order; });
  chapters.forEach(function (ch) { delete ch.order; });
  return { meta: meta, chapters: chapters };
}

// הסקר הציבורי (בלי מוסתרים), מהמטמון כשאפשר
function publicSurvey() {
  const cache = CacheService.getScriptCache();
  const hit = parseJsonSafe(cache.get(SURVEY_CACHE), null);
  if (hit) return hit;
  const survey = loadSurvey(false);
  if (survey) {
    try { cache.put(SURVEY_CACHE, JSON.stringify(survey), SURVEY_CACHE_SECS); } catch (e) { /* גדול מדי למטמון: קוראים מהגיליון */ }
  }
  return survey;
}

// legacy: השאלון עוד לא נשמר מדף הניהול, ואז האתר לוקח את הטקסטים הכלליים ואת "קצת עליך" מהקוד
function getSurvey() {
  try {
    const settings = getSettings();
    const open = settings.surveyOpen !== 'לא';
    const survey = publicSurvey();
    if (!survey) return { success: false, open: open, message: 'השאלון עוד לא נשמר. נכנסים לדף הניהול ושומרים' };
    return { success: true, survey: survey, open: open, legacy: !settings.surveyRev };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ============================================================
// תשובות
// ============================================================

// הטבלה השטוחה: עמודה לכל שאלה, שורה לכל עונה — נוחה לניתוח בגיליון
function rebuildFlatHeaders() {
  const survey = loadSurvey(true); // גם לשאלות מוסתרות יש עמודה
  if (!survey) return;
  const flat = ensureSheet(T_FLAT, ['קוד עונה', 'עדכון אחרון']);
  const headers = ['קוד עונה', 'עדכון אחרון'];
  survey.chapters.forEach(function (ch) {
    ch.questions.forEach(function (q) { headers.push(q.id); });
  });
  const existing = flat.getLastColumn() ? flat.getRange(1, 1, 1, flat.getLastColumn()).getValues()[0] : [];
  // מוסיפים רק עמודות חדשות — לא מוחקים נתונים קיימים
  const missing = headers.filter(function (h) { return existing.indexOf(h) === -1; });
  if (existing.length === 0) {
    ensureCols(flat, headers.length);
    flat.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else if (missing.length) {
    ensureCols(flat, existing.length + missing.length);
    flat.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
  }
  flat.setFrozenRows(1);
}

function flatValue(v) {
  if (Array.isArray(v)) return safeCell(v.join(' | '));
  return v === undefined || v === null ? '' : safeCell(v);
}

function findRidRow(flat, rid) {
  if (flat.getLastRow() <= 1) return -1;
  const rids = flat.getRange(2, 1, flat.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < rids.length; i++) {
    if (String(rids[i][0]) === rid) return i + 2;
  }
  return -1;
}

function submitChapter(rid, chapterId, answersJson) {
  try {
    const settings = getSettings();
    if (settings.surveyOpen === 'לא') return { success: false, message: 'הסקר סגור כרגע למענה' };

    rid = normRid(rid);
    if (!rid) return { success: false, message: 'קוד עונה לא תקין' };
    chapterId = String(chapterId || '').trim();
    if (!chapterId || chapterId.charAt(0) === '_') return { success: false, message: 'פרק לא תקין' };
    const answers = parseJsonSafe(answersJson, null);
    if (!answers || typeof answers !== 'object' || Array.isArray(answers)) return { success: false, message: 'תשובות לא תקינות' };

    withLock(function () {
      // 1. שמירה גולמית (append-only, גיבוי מלא)
      const raw = ensureSheet(T_RAW, RAW_HEADERS);
      raw.appendRow([new Date(), rid, chapterId, JSON.stringify(answers)]);

      // 2. עדכון הטבלה השטוחה — קריאה אחת וכתיבה אחת לשורה (מהיר יותר תחת עומס)
      const flat = ensureSheet(T_FLAT, ['קוד עונה', 'עדכון אחרון']);
      let headers = flat.getRange(1, 1, 1, Math.max(flat.getLastColumn(), 2)).getValues()[0];

      // עמודות חסרות? (שאלה חדשה שנוספה בגיליון)
      const missing = Object.keys(answers).filter(function (qId) { return qId.charAt(0) !== '_' && headers.indexOf(qId) === -1; });
      if (missing.length) {
        ensureCols(flat, headers.length + missing.length);
        flat.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
        headers = headers.concat(missing);
      }

      const row = findRidRow(flat, rid);
      // ערכים שנקראים בחזרה מאבדים את הגרש שהגן עליהם; safeCell מחזיר אותו לפני הכתיבה מחדש
      const vals = row === -1
        ? headers.map(function () { return ''; })
        : flat.getRange(row, 1, 1, headers.length).getValues()[0].map(safeCell);
      vals[0] = rid;
      vals[1] = new Date();
      for (const qId in answers) {
        if (qId.charAt(0) === '_') continue; // סימוני מערכת (דילוג וכו') — רק בגולמי
        const col = headers.indexOf(qId);
        if (col > 1) vals[col] = flatValue(answers[qId]);
      }
      if (row === -1) flat.appendRow(vals);
      else flat.getRange(row, 1, 1, headers.length).setValues([vals]);
    });
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// קודי עונים שלחצו "התחלה מחדש"
function discardedRids(rawData) {
  const out = {};
  for (let i = 1; i < rawData.length; i++) {
    if (String(rawData[i][2]) === DISCARD_CH) out[String(rawData[i][1])] = true;
  }
  return out;
}

// "התחלה מחדש": התשובות הקודמות נשארות בגיבוי הגולמי אבל לא נספרות, והשורה בטבלה השטוחה נמחקת
function discard(rid) {
  try {
    rid = normRid(rid);
    if (!rid) return { success: false, message: 'קוד עונה לא תקין' };
    withLock(function () {
      ensureSheet(T_RAW, RAW_HEADERS).appendRow([new Date(), rid, DISCARD_CH, '{}']);
      const flat = getSpreadsheet().getSheetByName(T_FLAT);
      if (flat) {
        const row = findRidRow(flat, rid);
        if (row > 1) flat.deleteRow(row);
      }
    });
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// כל הפרקים שהוגשו עבור עונה (שימוש פנימי: אימות ההגרלה)
function resume(rid) {
  try {
    rid = normRid(rid);
    const raw = getSpreadsheet().getSheetByName(T_RAW);
    if (!rid || !raw || raw.getLastRow() <= 1) return { success: false, message: 'לא נמצאו תשובות' };
    const data = raw.getDataRange().getValues();
    if (discardedRids(data)[rid]) return { success: false, message: 'לא נמצאו תשובות' };
    const chapters = {};
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) !== rid) continue;
      const chId = String(data[i][2]);
      const prev = chapters[chId] ? chapters[chId].answers : {};
      const cur = parseJsonSafe(String(data[i][3]), {});
      chapters[chId] = { ts: new Date(data[i][0]).toISOString(), answers: Object.assign(prev, cur) };
    }
    if (!Object.keys(chapters).length) return { success: false, message: 'לא נמצאו תשובות' };
    return { success: true, chapters: chapters };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ============================================================
// מייל = זהות: כתובת אחת, סקר אחד, והמשך מכל מכשיר. לא שולחים מיילים.
// כתובת חדשה נקשרת לקוד העונה של המכשיר. כתובת מוכרת ממכשיר אחר מחזירה את קוד העונה שלה
// ואת רשימת הנושאים שכבר נשמרו, בלי התשובות: מי שמקליד מייל של מישהו אחר לא רואה מה הוא ענה.
// ============================================================

function identify(email, rid) {
  try {
    const norm = normEmail(email);
    rid = normRid(rid);
    if (!norm) return { success: false, message: 'בדקו את כתובת המייל' };
    if (!rid) return { success: false, message: 'קוד עונה לא תקין' };

    const res = withLock(function () {
      const key = emailKey(norm);
      const sheet = ensureSheet(T_EMAILS, EMAIL_HEADERS);
      const n = sheet.getLastRow() - 1;
      const rows = n > 0 ? sheet.getRange(2, 1, n, 3).getValues() : [];
      let at = -1;
      for (let i = 0; i < rows.length; i++) if (String(rows[i][1]) === key) { at = i; break; }

      if (at === -1) {
        const cache = CacheService.getScriptCache();
        const perHour = Number(cache.get('new_emails_hour') || 0);
        if (perHour >= MAX_NEW_EMAILS_PER_HOUR) return { success: false, message: 'יש עומס רגעי. נסו שוב בעוד כמה דקות' };
        sheet.appendRow([new Date(), key, rid]);
        cache.put('new_emails_hour', String(perHour + 1), 3600);
        return { success: true, status: 'new' };
      }

      const known = String(rows[at][2]);
      if (known === rid) return { success: true, status: 'same' };

      const raw = getSpreadsheet().getSheetByName(T_RAW);
      const rawData = raw && raw.getLastRow() > 1 ? raw.getDataRange().getValues() : [[]];
      // הסקר של הכתובת בוטל ב"להתחיל מחדש": הכתובת עוברת לסקר החדש
      if (discardedRids(rawData)[known]) {
        sheet.getRange(at + 2, 3).setValue(rid);
        return { success: true, status: 'new' };
      }

      // הגולמי מסודר לפי זמן, אז השורה האחרונה של כל נושא היא המצב שלו
      const done = {};
      let about = {};
      for (let i = 1; i < rawData.length; i++) {
        if (String(rawData[i][1]) !== known) continue;
        const chId = String(rawData[i][2]);
        const ans = parseJsonSafe(String(rawData[i][3]), {});
        const real = Object.keys(ans).some(function (k) { return k.charAt(0) !== '_'; });
        done[chId] = { ts: new Date(rawData[i][0]).toISOString(), skipped: !real };
        if (chId === 'about') about = ans;
      }
      return { success: true, status: 'resume', rid: known, done: done, about: about };
    });

    // מ"קצת עליך" חוזרות רק התשובות שקובעות אילו נושאים ושאלות מוצגים (למשל ילדים בבית)
    if (res.status === 'resume') {
      const logic = profileKeysForLogic();
      res.profile = {};
      Object.keys(res.about).forEach(function (k) { if (logic[k]) res.profile[k] = res.about[k]; });
      delete res.about;
    }
    return res;
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function profileKeysForLogic() {
  const survey = publicSurvey();
  const keys = {};
  (survey ? survey.chapters : []).forEach(function (ch) {
    [ch].concat(ch.questions).forEach(function (x) {
      if (x.showIf && String(x.showIf.q).indexOf('about_') === 0) keys[x.showIf.q] = true;
    });
  });
  return keys;
}

// ============================================================
// הגרלה — האימות קורה כאן; נשמרים שם וטלפון בלבד.
// בלי קוד עונה ובלי זמן, והטאב ממוין לפי שם: אי אפשר לקשר הרשמה לתשובות לפי סדר או שעה.
// ============================================================

function enterRaffle(rid, name, phone) {
  try {
    rid = normRid(rid);
    name = String(name || '').trim().slice(0, 80);
    const digits = phoneDigits(phone);
    if (!rid || name.length < 2 || !digits) return { success: false, message: 'בדקו את השם והטלפון' };

    // אימות השלמה: כל הפרקים הרלוונטיים לפי הפרופיל הוגשו
    const res = resume(rid);
    if (!res.success) return { success: false, message: 'לא נמצאו תשובות. השלימו את הסקר קודם' };
    const submitted = res.chapters;
    const profile = (submitted.about && submitted.about.answers) || {};

    const survey = publicSurvey();
    if (survey) {
      const required = survey.chapters.filter(function (ch) { return evalCond(ch.showIf, profile); });
      const missing = required.filter(function (ch) { return !submitted[ch.id]; });
      if (missing.length) {
        return { success: false, message: 'נותרו נושאים להשלמה: ' + missing.map(function (c) { return c.title; }).join(', ') };
      }
    }

    return withLock(function () {
      const raffle = ensureSheet(T_RAFFLE, RAFFLE_HEADERS);
      // מניעת הרשמה כפולה לפי טלפון
      if (raffle.getLastRow() > 1) {
        const phones = raffle.getRange(2, 2, raffle.getLastRow() - 1, 1).getValues();
        for (let i = 0; i < phones.length; i++) {
          if (String(phones[i][0]).replace(/[^0-9]/g, '') === digits) {
            return { success: false, message: 'המספר הזה כבר רשום להגרלה 🙂' };
          }
        }
      }
      // גרש בהתחלה: הגיליון לא יהפוך את הטלפון למספר ולא יאכל את ה-0
      raffle.appendRow([safeCell(name), "'" + digits]);
      if (raffle.getLastRow() > 2) raffle.getRange(2, 1, raffle.getLastRow() - 1, RAFFLE_HEADERS.length).sort(1);
      return { success: true };
    });
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ============================================================
// תוצאות — דשבורד צוות (עם סיסמה) + דוח ציבורי (מצרפי בלבד)
// ============================================================

function getResults(password) {
  try {
    const settings = getSettings();
    if (!password || password !== settings.dashboardPassword) return { success: false, message: 'סיסמה שגויה' };
    const raw = getSpreadsheet().getSheetByName(T_RAW);
    const rows = [];
    if (raw && raw.getLastRow() > 1) {
      const data = raw.getDataRange().getValues();
      const discarded = discardedRids(data);
      for (let i = 1; i < data.length; i++) {
        const rid = String(data[i][1]);
        if (discarded[rid]) continue;
        rows.push({
          ts: new Date(data[i][0]).toISOString(),
          rid: rid,
          chapter: String(data[i][2]),
          answers: parseJsonSafe(String(data[i][3]), {}),
        });
      }
    }
    const raffle = getSpreadsheet().getSheetByName(T_RAFFLE);
    const raffleCount = raffle ? Math.max(0, raffle.getLastRow() - 1) : 0;
    const emails = getSpreadsheet().getSheetByName(T_EMAILS);
    const identified = emails ? Math.max(0, emails.getLastRow() - 1) : 0; // עברו את שלב המייל (גם מי שלא שמר כלום)
    return { success: true, rows: rows, raffleCount: raffleCount, identified: identified, survey: publicSurvey() };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// דוח ציבורי: מצרפים בלבד — בלי טקסט חופשי, בלי שמות, בלי פילוחים קטנים
function getPublicReport() {
  try {
    const settings = getSettings();
    if (settings.publicReport !== 'כן') return { success: false, message: 'הדוח הציבורי עדיין לא פורסם' };

    const survey = publicSurvey();
    const raw = getSpreadsheet().getSheetByName(T_RAW);
    if (!survey || !raw || raw.getLastRow() <= 1) return { success: false, message: 'אין עדיין נתונים' };

    // מיזוג: תשובה אחרונה לכל (עונה, שאלה)
    const data = raw.getDataRange().getValues();
    const discarded = discardedRids(data);
    const perRid = {};
    for (let i = 1; i < data.length; i++) {
      const rid = String(data[i][1]);
      if (discarded[rid]) continue;
      perRid[rid] = perRid[rid] || {};
      Object.assign(perRid[rid], parseJsonSafe(String(data[i][3]), {}));
    }
    const respondents = Object.keys(perRid).map(function (k) { return perRid[k]; });

    const EXCLUDE = { about_name: 1, about_submitted: 1 };
    const report = { totalRespondents: respondents.length, chapters: [] };

    survey.chapters.forEach(function (ch) {
      const chOut = { id: ch.id, title: ch.title, icon: ch.icon || '', questions: [] };
      ch.questions.forEach(function (q) {
        if (EXCLUDE[q.id]) return;
        if (q.type === 'text' || q.type === 'textarea') return; // טקסט חופשי לא מפורסם
        const vals = respondents.map(function (r) { return r[q.id]; })
          .filter(function (v) { return v !== undefined && v !== null && v !== ''; });
        if (vals.length < 5) return; // כלל מינימום 5 — הגנת אנונימיות
        const qOut = { id: q.id, text: q.text, type: q.type, count: vals.length };
        if (q.type === 'scale' || q.type === 'number') {
          const nums = vals.map(Number).filter(function (n) { return !isNaN(n); });
          if (vals.length > nums.length) qOut.dontKnow = vals.length - nums.length;
          if (nums.length) qOut.avg = Math.round((nums.reduce(function (a, b) { return a + b; }, 0) / nums.length) * 10) / 10;
          if (q.type === 'scale') {
            qOut.min = q.min || 1; qOut.max = q.max || 10;
            qOut.minLabel = q.minLabel || ''; qOut.maxLabel = q.maxLabel || '';
            qOut.hist = {};
            nums.forEach(function (n) { qOut.hist[n] = (qOut.hist[n] || 0) + 1; });
          }
        } else if (q.type === 'radio' || q.type === 'checkbox') {
          qOut.counts = {};
          vals.forEach(function (v) {
            (Array.isArray(v) ? v : [v]).forEach(function (o) {
              const key = String(o).indexOf('אחר: ') === 0 ? 'אחר' : String(o);
              qOut.counts[key] = (qOut.counts[key] || 0) + 1;
            });
          });
        } else if (q.type === 'rank') {
          // ממוצע מיקום (1 = הכי חשוב)
          qOut.avgPos = {};
          const sums = {}, ns = {};
          vals.forEach(function (arr) {
            if (!Array.isArray(arr)) return;
            arr.forEach(function (opt, idx) {
              sums[opt] = (sums[opt] || 0) + idx + 1;
              ns[opt] = (ns[opt] || 0) + 1;
            });
          });
          Object.keys(sums).forEach(function (opt) {
            qOut.avgPos[opt] = Math.round((sums[opt] / ns[opt]) * 10) / 10;
          });
        }
        chOut.questions.push(qOut);
      });
      if (chOut.questions.length) report.chapters.push(chOut);
    });

    return { success: true, report: report, meta: survey.meta };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ============================================================
// ראוטר — endpoint יחיד
// GET לקריאות ציבוריות · POST (text/plain) לכתיבות ולכל מה שיש בו מידע אישי או סיסמה
// ============================================================

function doGet(e) { return route(e, null); }

function doPost(e) {
  let body = null;
  if (e && e.postData && e.postData.contents) body = parseJsonSafe(e.postData.contents, null);
  return route(e, body);
}

function route(e, body) {
  try {
    const p = Object.assign({}, (e && e.parameter) || {}, body || {});
    const action = p.action;

    switch (action) {
      case 'ping':
        return jsonResponse({ success: true, version: 'v4' });

      case 'getSurvey':
        return jsonResponse(getSurvey());

      case 'submitChapter':
        return jsonResponse(submitChapter(p.rid, p.chapter, typeof p.answers === 'string' ? p.answers : JSON.stringify(p.answers || {})));

      case 'discard':
        return jsonResponse(discard(p.rid));

      case 'identify':
        return jsonResponse(identify(p.email, p.rid));

      case 'enterRaffle':
        return jsonResponse(enterRaffle(p.rid, p.name, p.phone));

      case 'getResults':
        return jsonResponse(getResults(p.password));

      case 'getPublicReport':
        return jsonResponse(getPublicReport());

      case 'getAdmin':
        return jsonResponse(getAdmin(p.password));

      case 'saveSurvey':
        return jsonResponse(saveSurvey(p.password, typeof p.survey === 'string' ? p.survey : JSON.stringify(p.survey || null), p.rev, p.force === true));

      case 'saveSettings':
        return jsonResponse(saveSettings(p.password, p.surveyOpen, p.publicReport, p.newPassword));

      default:
        return jsonResponse({ success: false, message: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonResponse({ success: false, message: err.toString() });
  }
}
