// ============================================================
// סקר התושבים — נופי פרת · Google Apps Script Backend
// הגיליון הוא מסד הנתונים: פרקים, שאלות, תשובות, הגרלה, הגדרות
// ============================================================
//
// פריסה (פירוט מלא ב-CLAUDE.md, עם clasp מהמחשב של אורי):
//   1. clasp create --type sheets → נוצרים גיליון + סקריפט. ה-ID של הגיליון נכנס ל-SHEET_ID.
//   2. clasp push -f && clasp deploy → כתובת ה-/exec נכנסת ל-config.js → SCRIPT_URL.
//   3. פעם אחת: לפתוח את העורך (clasp open), להריץ את setup() ולאשר הרשאות.
//      setup יוצר את הטאבים, סיסמת צוות אקראית (מודפסת בלוג ונשמרת בטאב "הגדרות")
//      וטריגר חימום.
//   4. setup.html באתר → "סנכרון שאלות לגיליון" (בפעם הראשונה בלי סיסמה).
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
const T_LINKS = 'קודי המשך';
const T_SETTINGS = 'הגדרות';

// Schema: זמן(0), קוד עונה(1), פרק(2), תשובות JSON(3) — append-only
const RAW_HEADERS = ['זמן', 'קוד עונה', 'פרק', 'תשובות (JSON)'];
// Schema: שם(0), טלפון(1) — בלי זמן ובלי קוד עונה, ממוין לפי שם (אי אפשר לקשר לתשובות)
const RAFFLE_HEADERS = ['שם', 'טלפון'];
// Schema: עדכון(0), מפתח טלפון(1), קוד מוצפן(2), קוד עונה(3)
const LINK_HEADERS = ['עדכון', 'מפתח טלפון (מוצפן)', 'קוד (מוצפן)', 'קוד עונה'];

// "פרק" מיוחד בטאב הגולמי: העונה לחץ "התחלה מחדש", כל התשובות שלו לא נספרות
const DISCARD_CH = '_בוטל';
const MAX_RESUME_TRIES = 5; // ניסיונות שגויים לטלפון, לשעה

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
    ['dashboardPassword', Utilities.getUuid().replace(/-/g, '').slice(0, 10), 'סיסמת הצוות לדשבורד ולסנכרון השאלות. נוצרה אקראית, אפשר להחליף'],
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
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value, note || '']);
}

// ── טלפון להמשך ממכשיר אחר: חתימת HMAC עם מפתח סודי ────────
// המפתח נשמר ב-Script Properties (לא בגיליון ולא בקוד), כך שמי שרואה את הגיליון
// לא יכול לשחזר מספר מהחתימה.
function phonePepper() {
  const props = PropertiesService.getScriptProperties();
  let p = props.getProperty('PHONE_PEPPER');
  if (!p) {
    p = withLock(function () {
      let cur = props.getProperty('PHONE_PEPPER');
      if (!cur) { cur = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('PHONE_PEPPER', cur); }
      return cur;
    });
  }
  return p;
}

function hmacHex(value) {
  const bytes = Utilities.computeHmacSha256Signature(value, phonePepper());
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function phoneDigits(phone) {
  const d = String(phone || '').replace(/[^0-9]/g, '');
  return /^0\d{8,9}$/.test(d) ? d : null;
}

function phoneKey(phone) {
  const d = phoneDigits(phone);
  return d ? hmacHex('phone:' + d) : null;
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
  phonePepper();
  ensureSheet(T_RAW, RAW_HEADERS);
  ensureSheet(T_RAFFLE, RAFFLE_HEADERS);
  ensureSheet(T_LINKS, LINK_HEADERS);
  const hasWarm = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'keepWarm'; });
  if (!hasWarm) ScriptApp.newTrigger('keepWarm').timeBased().everyMinutes(10).create();
  Logger.log('מוכן. סיסמת הצוות (גם בטאב "הגדרות"): ' + settings.dashboardPassword);
}

// טריגר כל 10 דקות, מקצר את ההמתנה של הגולש הראשון אחרי שקט
function keepWarm() { Logger.log('warm ' + new Date().toISOString()); }

// ============================================================
// מבנה הסקר: פרקים + שאלות
// ============================================================

const CHAPTER_HEADERS = ['id', 'כותרת', 'אייקון', 'תיאור', 'תנאי הצגה', 'שער', 'פעיל', 'סדר', 'קטגוריה', 'פתיח', 'סיום'];
const QUESTION_HEADERS = ['פרק', 'id', 'סוג', 'שאלה', 'עזרה', 'אפשרויות ( | )', 'אחר', 'בלעדי', 'מינ', 'מקס', 'תווית מינ', 'תווית מקס', 'תנאי הצגה', 'פעיל', 'סדר'];

// זריעת המבנה מהלקוח (setup.html שולח את survey-data.js המלא).
// בפעם הראשונה (טאב השאלות ריק) אין צורך בסיסמה; אחר כך חובה, כי הזריעה דורסת עריכות בגיליון.
function seedSurvey(surveyJson, password) {
  const qExisting = getSpreadsheet().getSheetByName(T_QUESTIONS);
  const firstTime = !qExisting || qExisting.getLastRow() <= 1;
  const settings = getSettings();
  if (!firstTime && password !== settings.dashboardPassword) {
    return { success: false, message: 'סיסמה שגויה (הסיסמה נמצאת בטאב "הגדרות" בגיליון)' };
  }
  const survey = parseJsonSafe(surveyJson, null);
  if (!survey || !survey.chapters) return { success: false, message: 'מבנה סקר לא תקין' };

  const chSheet = ensureSheet(T_CHAPTERS, CHAPTER_HEADERS);
  const qSheet = ensureSheet(T_QUESTIONS, QUESTION_HEADERS);
  chSheet.clearContents(); chSheet.appendRow(CHAPTER_HEADERS);
  qSheet.clearContents(); qSheet.appendRow(QUESTION_HEADERS);

  const chRows = [], qRows = [];
  survey.chapters.forEach(function (ch, ci) {
    chRows.push([
      ch.id, ch.title, ch.icon || '', ch.desc || '',
      condToString(ch.showIf),
      (ch.gate || ch.core) ? 'כן' : '', 'כן', ci + 1, ch.cat || '',
      ch.intro || '', ch.outro || '',
    ]);
    (ch.questions || []).forEach(function (q, qi) {
      qRows.push([
        ch.id, q.id, TYPE_TO_HE[q.type] || q.type, q.text, q.help || '',
        (q.opts || []).join(' | '),
        q.other ? 'כן' : '', q.exclusive || '',
        q.min !== undefined ? q.min : '', q.max !== undefined ? q.max : '',
        q.minLabel || '', q.maxLabel || '',
        condToString(q.showIf),
        'כן', qi + 1,
      ]);
    });
  });
  if (chRows.length) chSheet.getRange(2, 1, chRows.length, CHAPTER_HEADERS.length).setValues(chRows);
  if (qRows.length) qSheet.getRange(2, 1, qRows.length, QUESTION_HEADERS.length).setValues(qRows);

  // שמירת המטא (כותרות, פתיח, הגרלה) בהגדרות
  const sSheet = ensureSheet(T_SETTINGS, ['מפתח', 'ערך', 'הסבר']);
  upsertSetting(sSheet, 'meta', JSON.stringify(survey.meta || {}), 'מטא של הסקר (כותרת, פתיח, הגרלה) — JSON');

  // הכנת שאר הטאבים + כותרות הטבלה השטוחה
  ensureSheet(T_RAW, RAW_HEADERS);
  ensureSheet(T_RAFFLE, RAFFLE_HEADERS);
  ensureSheet(T_LINKS, LINK_HEADERS);
  writeGuideSheet();
  rebuildFlatHeaders();

  return { success: true, message: 'נטענו ' + chRows.length + ' פרקים ו-' + qRows.length + ' שאלות' };
}

// טאב "מדריך עריכה" — הוראות לצוות, נכתב מחדש בכל סנכרון
function writeGuideSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('מדריך עריכה');
  if (!sheet) sheet = ss.insertSheet('מדריך עריכה');
  sheet.clearContents();
  const rows = [
    ['📝 איך עורכים את הסקר? (השינויים מופיעים באתר בטעינה הבאה, בלי פריסה מחדש)'],
    [''],
    ['עריכת נוסח', 'פשוט עורכים את התא בעמודה "שאלה" בטאב "שאלות". אותו דבר לגבי אפשרויות, עזרה ותוויות.'],
    ['הוספת שאלה', 'מוסיפים שורה בטאב "שאלות": פרק קיים, id חדש באנגלית (למשל post_x1), סוג ונוסח. עמודת "סדר" קובעת את המיקום בפרק.'],
    ['השבתת שאלה/פרק', 'עמודת "פעיל" = לא. השאלה נעלמת מהאתר, והתשובות שכבר נאספו נשמרות.'],
    ['⚠️ חשוב', 'לא לשנות id של שאלה קיימת. התוצאות נשמרות לפי ה-id.'],
    [''],
    ['סוגי שאלות', 'סולם · בחירה אחת · בחירה מרובה · טקסט קצר · טקסט ארוך · מספר · דירוג'],
    ['אפשרויות', 'מפרידים בקו אנכי | . עמודת "אחר" = כן מוסיפה אפשרות "אחר" עם שדה חופשי.'],
    ['סולם', 'עמודות מינ/מקס קובעות את הטווח (למשל 1 ו-10, או 1 ו-7) + תווית מינ/תווית מקס לטקסט בקצוות.'],
    ['בחירה מרובה', 'מקס = מספר הבחירות המרבי (למשל 3 בשאלת "בחר 3 נושאים"). ריק = בלי הגבלה.'],
    ['בלעדי', 'בבחירה מרובה: אפשרות שמבטלת את כל השאר (למשל "אין ילדים בבית").'],
    [''],
    ['תנאי הצגה', 'מציג שאלה/פרק רק לפי תשובה קודמת. תחביר: id = ערך  (או כמה ערכים עם | )'],
    ['דוגמה 1', 'syn_teacher = מישהו מהיישוב   ← מוצג רק למי שבחר באפשרות הזו'],
    ['דוגמה 2', 'about_gender != גבר   ← מוצג לכולם חוץ ממי שענה "גבר"'],
    ['דוגמה 3', 'about_kids = נוער (ז׳–י״ב) | יסודי (א׳–ו׳)  ← מוצג אם סומן אחד מאלה'],
    [''],
    ['פתיח / סיום (טאב "פרקים")', 'טקסט שמוצג בראש הפרק / בסופו, למשל המילים של ועדת התרבות. ריק = בלי.'],
    ['קטגוריות (עמודה בטאב "פרקים")', 'gov=ניהול · infra=תשתיות · space=מרחב ציבורי · edu=חינוך · comm=קהילה · cult=תרבות'],
    [''],
    ['הגדרות (טאב "הגדרות")', 'surveyOpen=לא סוגר את הסקר · publicReport=כן מפרסם את הדוח הציבורי · dashboardPassword: סיסמת הצוות'],
    ['תוצאות', 'טאב "תוצאות": שורה לכל עונה, עמודה לכל שאלה. טאב "תשובות גולמי": גיבוי מלא, לא לערוך.'],
  ];
  sheet.getRange(1, 1, rows.length, 2).setValues(rows.map(function (r) { return [r[0] || '', r[1] || '']; }));
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 700);
}

// קריאת מבנה הסקר מהגיליון
function loadSurvey() {
  const chSheet = getSpreadsheet().getSheetByName(T_CHAPTERS);
  const qSheet = getSpreadsheet().getSheetByName(T_QUESTIONS);
  if (!chSheet || !qSheet || chSheet.getLastRow() <= 1) return null;

  const settings = getSettings();
  const meta = parseJsonSafe(settings.meta, {});

  const qData = qSheet.getDataRange().getValues();
  const byChapter = {};
  for (let i = 1; i < qData.length; i++) {
    const r = qData[i];
    if (!String(r[1]).trim()) continue;            // שורה ריקה
    if (String(r[13]).trim() === 'לא') continue;   // לא פעיל
    const q = { id: String(r[1]).trim(), type: normalizeType(r[2]), text: String(r[3]), order: Number(r[14]) || i };
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
    if (String(r[6]).trim() === 'לא') continue; // לא פעיל
    const ch = { id: String(r[0]).trim(), title: String(r[1]), questions: byChapter[String(r[0]).trim()] || [], order: Number(r[7]) || i };
    if (r[2]) ch.icon = String(r[2]);
    if (r[3]) ch.desc = String(r[3]);
    const cond = parseCondition(String(r[4]));
    if (cond) ch.showIf = cond;
    if (String(r[5]).trim() === 'כן') ch.gate = true;
    if (r.length > 8 && r[8]) ch.cat = String(r[8]).trim();
    if (r.length > 9 && r[9]) ch.intro = String(r[9]);
    if (r.length > 10 && r[10]) ch.outro = String(r[10]);
    chapters.push(ch);
  }
  chapters.sort(function (a, b) { return a.order - b.order; });
  chapters.forEach(function (ch) { delete ch.order; });
  return { meta: meta, chapters: chapters };
}

function getSurvey() {
  try {
    const settings = getSettings();
    const open = settings.surveyOpen !== 'לא';
    const survey = loadSurvey();
    if (!survey) return { success: false, open: open, message: 'הסקר טרם נטען לגיליון. הריצו סנכרון מ-setup.html' };
    return { success: true, survey: survey, open: open };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ============================================================
// תשובות
// ============================================================

// הטבלה השטוחה: עמודה לכל שאלה, שורה לכל עונה — נוחה לניתוח בגיליון
function rebuildFlatHeaders() {
  const survey = loadSurvey();
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

// כל הפרקים שהוגשו עבור עונה (שימוש פנימי: המשך ממכשיר אחר, אימות הגרלה)
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
// המשך ממכשיר אחר: טלפון + קוד בן 4 ספרות
// נשמרות חתימות מוצפנות בלבד. בלי הקוד אי אפשר לטעון תשובות של אחר, גם אם יודעים את המספר שלו.
// ============================================================

function linkResume(phone, pin, rid) {
  try {
    const key = phoneKey(phone);
    rid = normRid(rid);
    pin = String(pin || '').trim();
    if (!key || !/^\d{4}$/.test(pin) || !rid) return { success: false, message: 'בדקו את הטלפון והקוד' };
    const pinHash = hmacHex('pin:' + key + ':' + pin);
    withLock(function () {
      const sheet = ensureSheet(T_LINKS, LINK_HEADERS);
      // upsert לפי טלפון: תמיד מצביע על העונה האחרון שקישר אותו
      if (sheet.getLastRow() > 1) {
        const keys = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
        for (let i = 0; i < keys.length; i++) {
          if (String(keys[i][0]) === key) {
            sheet.getRange(i + 2, 1, 1, 4).setValues([[new Date(), key, pinHash, rid]]);
            return;
          }
        }
      }
      sheet.appendRow([new Date(), key, pinHash, rid]);
    });
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function resumeByPhone(phone, pin) {
  try {
    const key = phoneKey(phone);
    pin = String(pin || '').trim();
    if (!key || !/^\d{4}$/.test(pin)) return { success: false, message: 'בדקו את הטלפון והקוד' };

    const cache = CacheService.getScriptCache();
    const triesKey = 'tries_' + key.slice(0, 40);
    const tries = Number(cache.get(triesKey) || 0);
    if (tries >= MAX_RESUME_TRIES) return { success: false, message: 'יותר מדי ניסיונות. נסו שוב בעוד שעה' };
    const fail = function () {
      cache.put(triesKey, String(tries + 1), 3600);
      return { success: false, message: 'לא מצאנו. ודאו שהטלפון והקוד זהים למה שהזנתם ב"אמשיך אחר כך"' };
    };

    const sheet = getSpreadsheet().getSheetByName(T_LINKS);
    if (!sheet || sheet.getLastRow() <= 1) return fail();
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    const pinHash = hmacHex('pin:' + key + ':' + pin);
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][1]) !== key) continue;
      if (String(data[i][2]) !== pinHash) return fail();
      cache.remove(triesKey);
      const rid = String(data[i][3]);
      const res = resume(rid);
      if (!res.success) return res;
      res.rid = rid;
      return res;
    }
    return fail();
  } catch (e) {
    return { success: false, message: e.toString() };
  }
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

    const survey = loadSurvey();
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
    return { success: true, rows: rows, raffleCount: raffleCount, survey: loadSurvey() };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// דוח ציבורי: מצרפים בלבד — בלי טקסט חופשי, בלי שמות, בלי פילוחים קטנים
function getPublicReport() {
  try {
    const settings = getSettings();
    if (settings.publicReport !== 'כן') return { success: false, message: 'הדוח הציבורי עדיין לא פורסם' };

    const survey = loadSurvey();
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
          qOut.avg = Math.round((nums.reduce(function (a, b) { return a + b; }, 0) / nums.length) * 10) / 10;
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
        return jsonResponse({ success: true, version: 'v2' });

      case 'getSurvey':
        return jsonResponse(getSurvey());

      case 'submitChapter':
        return jsonResponse(submitChapter(p.rid, p.chapter, typeof p.answers === 'string' ? p.answers : JSON.stringify(p.answers || {})));

      case 'discard':
        return jsonResponse(discard(p.rid));

      case 'linkResume':
        return jsonResponse(linkResume(p.phone, p.pin, p.rid));

      case 'resumeByPhone':
        return jsonResponse(resumeByPhone(p.phone, p.pin));

      case 'enterRaffle':
        return jsonResponse(enterRaffle(p.rid, p.name, p.phone));

      case 'getResults':
        return jsonResponse(getResults(p.password));

      case 'getPublicReport':
        return jsonResponse(getPublicReport());

      case 'seedSurvey':
        return jsonResponse(seedSurvey(typeof p.survey === 'string' ? p.survey : JSON.stringify(p.survey || {}), p.password));

      default:
        return jsonResponse({ success: false, message: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonResponse({ success: false, message: err.toString() });
  }
}
