// ============================================================
// סקר התושבים — נופי פרת · Google Apps Script Backend
// הגיליון הוא מסד הנתונים: פרקים, שאלות, תשובות, הגרלה, הגדרות
// ============================================================
//
// התקנה:
//   1. צרו גיליון Google Sheets חדש והעתיקו את ה-ID שלו
//      (המחרוזת הארוכה בין /d/ ל-/edit בכתובת הגיליון).
//   2. הדביקו את ה-ID למטה ב-SHEET_ID.
//   3. Deploy → New deployment → Web app:
//        Execute as: Me · Who has access: Anyone
//   4. העתיקו את כתובת ה-Web app אל config.js → SCRIPT_URL.
//   5. פתחו את setup.html באתר ולחצו "סנכרון שאלות לגיליון"
//      (זה ממלא את טאבי הפרקים והשאלות מקובץ survey-data.js).
//
// ⚠️ חשוב: אחרי כל עריכה של הקובץ הזה חובה לפרסם מחדש:
//    Deploy → Manage deployments → ✏ → Version: New version → Deploy
// ============================================================

const SHEET_ID = 'PASTE_YOUR_SHEET_ID_HERE';

// שמות הטאבים
const T_CHAPTERS = 'פרקים';
const T_QUESTIONS = 'שאלות';
const T_RAW = 'תשובות גולמי';
const T_FLAT = 'תוצאות';
const T_RAFFLE = 'הגרלה';
const T_LINKS = 'קודי המשך';
const T_SETTINGS = 'הגדרות';

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

// הגדרות: key → value
const SETTINGS_DEFAULTS = [
  ['surveyOpen', 'כן', 'האם הסקר פתוח למענה (כן/לא)'],
  ['dashboardPassword', 'nofim2026', 'סיסמת הדשבורד — החליפו אותה!'],
  ['publicReport', 'לא', 'האם הדוח הציבורי פעיל (כן/לא)'],
];

function getSettings() {
  const sheet = ensureSheet(T_SETTINGS, ['מפתח', 'ערך', 'הסבר']);
  if (sheet.getLastRow() <= 1) {
    SETTINGS_DEFAULTS.forEach(function (r) { sheet.appendRow(r); });
  }
  const data = sheet.getDataRange().getValues();
  const out = {};
  for (let i = 1; i < data.length; i++) out[String(data[i][0]).trim()] = String(data[i][1]).trim();
  return out;
}

// הערכת תנאי showIf — זהה ללוגיקה בצד הלקוח
function evalCond(cond, answers) {
  if (!cond || !cond.q) return true;
  const val = answers[cond.q];
  if (val === undefined || val === null || val === '' ||
      (Array.isArray(val) && val.length === 0)) return true;
  if (cond.in) return cond.in.indexOf(val) !== -1;
  if (cond.notIn) return cond.notIn.indexOf(val) === -1;
  if (cond.any) return Array.isArray(val) && cond.any.some(function (o) { return val.indexOf(o) !== -1; });
  return true;
}

// ============================================================
// מבנה הסקר: פרקים + שאלות
// ============================================================

const CHAPTER_HEADERS = ['id', 'כותרת', 'אייקון', 'תיאור', 'תנאי הצגה (JSON)', 'שער', 'פעיל', 'סדר', 'קטגוריה'];
const QUESTION_HEADERS = ['פרק', 'id', 'סוג', 'שאלה', 'עזרה', 'אפשרויות ( | )', 'אחר', 'בלעדי', 'מינ', 'מקס', 'תווית מינ', 'תווית מקס', 'תנאי הצגה (JSON)', 'פעיל', 'סדר'];

// זריעת המבנה מהלקוח (setup.html שולח את survey-data.js המלא)
function seedSurvey(surveyJson, password) {
  const settings = getSettings();
  if (password !== settings.dashboardPassword) {
    return { success: false, message: 'סיסמה שגויה' };
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
      ch.showIf ? JSON.stringify(ch.showIf) : '',
      (ch.gate || ch.core) ? 'כן' : '', 'כן', ci + 1, ch.cat || '',
    ]);
    (ch.questions || []).forEach(function (q, qi) {
      qRows.push([
        ch.id, q.id, q.type, q.text, q.help || '',
        (q.opts || []).join(' | '),
        q.other ? 'כן' : '', q.exclusive || '',
        q.min !== undefined ? q.min : '', q.max !== undefined ? q.max : '',
        q.minLabel || '', q.maxLabel || '',
        q.showIf ? JSON.stringify(q.showIf) : '',
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
  ensureSheet(T_RAW, ['זמן', 'קוד עונה', 'פרק', 'תשובות (JSON)']);
  ensureSheet(T_RAFFLE, ['זמן', 'שם', 'טלפון']);
  ensureSheet(T_LINKS, ['זמן', 'hash', 'קוד עונה']);
  rebuildFlatHeaders();

  return { success: true, message: 'נטענו ' + chRows.length + ' פרקים ו-' + qRows.length + ' שאלות' };
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
    if (String(r[13]).trim() === 'לא') continue; // לא פעיל
    const q = { id: String(r[1]).trim(), type: String(r[2]).trim(), text: String(r[3]) };
    if (r[4]) q.help = String(r[4]);
    if (r[5]) q.opts = String(r[5]).split('|').map(function (s) { return s.trim(); }).filter(String);
    if (String(r[6]).trim() === 'כן') q.other = true;
    if (r[7]) q.exclusive = String(r[7]).trim();
    if (r[8] !== '') q.min = Number(r[8]);
    if (r[9] !== '') q.max = Number(r[9]);
    if (r[10]) q.minLabel = String(r[10]);
    if (r[11]) q.maxLabel = String(r[11]);
    const cond = parseJsonSafe(String(r[12]), null);
    if (cond) q.showIf = cond;
    const chId = String(r[0]).trim();
    (byChapter[chId] = byChapter[chId] || []).push(q);
  }

  const chData = chSheet.getDataRange().getValues();
  const chapters = [];
  for (let i = 1; i < chData.length; i++) {
    const r = chData[i];
    if (String(r[6]).trim() === 'לא') continue; // לא פעיל
    const ch = { id: String(r[0]).trim(), title: String(r[1]), questions: byChapter[String(r[0]).trim()] || [] };
    if (r[2]) ch.icon = String(r[2]);
    if (r[3]) ch.desc = String(r[3]);
    const cond = parseJsonSafe(String(r[4]), null);
    if (cond) ch.showIf = cond;
    if (String(r[5]).trim() === 'כן') ch.gate = true;
    if (r.length > 8 && r[8]) ch.cat = String(r[8]).trim();
    chapters.push(ch);
  }
  return { meta: meta, chapters: chapters };
}

function getSurvey() {
  try {
    const survey = loadSurvey();
    if (!survey) return { success: false, message: 'הסקר טרם נטען לגיליון — הריצו סנכרון מ-setup.html' };
    const settings = getSettings();
    return { success: true, survey: survey, open: settings.surveyOpen !== 'לא' };
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
    flat.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else if (missing.length) {
    flat.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
  }
}

function flatValue(v) {
  if (Array.isArray(v)) return v.join(' | ');
  return v === undefined || v === null ? '' : v;
}

function submitChapter(rid, chapterId, answersJson) {
  try {
    const settings = getSettings();
    if (settings.surveyOpen === 'לא') return { success: false, message: 'הסקר סגור כרגע למענה' };

    rid = String(rid || '').toUpperCase().trim();
    if (!/^[A-Z2-9]{8}$/.test(rid)) return { success: false, message: 'קוד עונה לא תקין' };
    const answers = parseJsonSafe(answersJson, null);
    if (!answers || typeof answers !== 'object') return { success: false, message: 'תשובות לא תקינות' };

    // נעילה — מונע דריסה בכתיבות מקבילות
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      // 1. שמירה גולמית (append-only, גיבוי מלא)
      const raw = ensureSheet(T_RAW, ['זמן', 'קוד עונה', 'פרק', 'תשובות (JSON)']);
      raw.appendRow([new Date(), rid, chapterId, JSON.stringify(answers)]);

      // 2. עדכון הטבלה השטוחה
      const flat = ensureSheet(T_FLAT, ['קוד עונה', 'עדכון אחרון']);
      let headers = flat.getRange(1, 1, 1, Math.max(flat.getLastColumn(), 2)).getValues()[0];

      // עמודות חסרות? (שאלה חדשה שנוספה בגיליון)
      const missing = Object.keys(answers).filter(function (qId) { return headers.indexOf(qId) === -1; });
      if (missing.length) {
        flat.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
        headers = headers.concat(missing);
      }

      // איתור/יצירת שורת העונה
      let row = -1;
      if (flat.getLastRow() > 1) {
        const rids = flat.getRange(2, 1, flat.getLastRow() - 1, 1).getValues();
        for (let i = 0; i < rids.length; i++) {
          if (String(rids[i][0]) === rid) { row = i + 2; break; }
        }
      }
      if (row === -1) {
        flat.appendRow([rid, new Date()]);
        row = flat.getLastRow();
      } else {
        flat.getRange(row, 2).setValue(new Date());
      }

      for (const qId in answers) {
        const col = headers.indexOf(qId) + 1;
        if (col > 0) flat.getRange(row, col).setValue(flatValue(answers[qId]));
      }
    } finally {
      lock.releaseLock();
    }
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// המשך ממכשיר אחר: מחזיר את כל הפרקים שהוגשו עבור הקוד
function resume(rid) {
  try {
    rid = String(rid || '').toUpperCase().trim();
    const raw = getSpreadsheet().getSheetByName(T_RAW);
    if (!raw || raw.getLastRow() <= 1) return { success: false, message: 'לא נמצאו תשובות' };
    const data = raw.getDataRange().getValues();
    const chapters = {};
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) !== rid) continue;
      const chId = String(data[i][2]);
      const prev = chapters[chId] ? chapters[chId].answers : {};
      const cur = parseJsonSafe(String(data[i][3]), {});
      chapters[chId] = { ts: new Date(data[i][0]).toISOString(), answers: Object.assign(prev, cur) };
    }
    if (!Object.keys(chapters).length) return { success: false, message: 'לא נמצאו תשובות לקוד הזה' };
    return { success: true, chapters: chapters };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ============================================================
// המשך ממכשיר אחר לפי טלפון — נשמר hash בלבד, לא המספר עצמו
// ============================================================

function linkResume(hash, rid) {
  try {
    hash = String(hash || '').trim();
    rid = String(rid || '').toUpperCase().trim();
    if (hash.length < 8 || !/^[A-Z2-9]{8}$/.test(rid)) return { success: false, message: 'נתונים לא תקינים' };
    const sheet = ensureSheet(T_LINKS, ['זמן', 'hash', 'קוד עונה']);
    // upsert לפי hash — טלפון אחד מצביע תמיד על העונה האחרון שקישר אותו
    if (sheet.getLastRow() > 1) {
      const hashes = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < hashes.length; i++) {
        if (String(hashes[i][0]) === hash) {
          sheet.getRange(i + 2, 1).setValue(new Date());
          sheet.getRange(i + 2, 3).setValue(rid);
          return { success: true };
        }
      }
    }
    sheet.appendRow([new Date(), hash, rid]);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function resumeByHash(hash) {
  try {
    hash = String(hash || '').trim();
    const sheet = getSpreadsheet().getSheetByName(T_LINKS);
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'לא נמצא — ודאו שלחצתם "אמשיך אחר־כך" במכשיר הקודם' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === hash) {
        const rid = String(data[i][2]);
        const res = resume(rid);
        if (!res.success) return res;
        res.rid = rid;
        return res;
      }
    }
    return { success: false, message: 'לא נמצא — ודאו שלחצתם "אמשיך אחר־כך" במכשיר הקודם' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ============================================================
// הגרלה — האימות קורה כאן; נשמרים שם וטלפון בלבד (בלי קוד עונה)
// ============================================================

function enterRaffle(rid, name, phone) {
  try {
    rid = String(rid || '').toUpperCase().trim();
    name = String(name || '').trim();
    phone = String(phone || '').trim();
    if (name.length < 2 || phone.length < 9) return { success: false, message: 'חסרים פרטים' };

    // אימות השלמה: כל הפרקים הרלוונטיים לפי הפרופיל הוגשו
    const res = resume(rid);
    if (!res.success) return { success: false, message: 'לא נמצאו תשובות — השלימו את הסקר קודם' };
    const submitted = res.chapters;
    const profile = (submitted.about && submitted.about.answers) || {};

    const survey = loadSurvey();
    if (survey) {
      const required = survey.chapters.filter(function (ch) { return evalCond(ch.showIf, profile); });
      const missing = required.filter(function (ch) { return !submitted[ch.id]; });
      if (missing.length) {
        return { success: false, message: 'נותרו פרקים להשלמה: ' + missing.map(function (c) { return c.title; }).join(', ') };
      }
    }

    // מניעת הרשמה כפולה לפי טלפון
    const raffle = ensureSheet(T_RAFFLE, ['זמן', 'שם', 'טלפון']);
    const clean = phone.replace(/[^0-9]/g, '');
    if (raffle.getLastRow() > 1) {
      const phones = raffle.getRange(2, 3, raffle.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < phones.length; i++) {
        if (String(phones[i][0]).replace(/[^0-9]/g, '') === clean) {
          return { success: false, message: 'המספר הזה כבר רשום להגרלה 🙂' };
        }
      }
    }
    raffle.appendRow([new Date(), name, phone]);
    return { success: true };
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
    if (password !== settings.dashboardPassword) return { success: false, message: 'סיסמה שגויה' };
    const raw = getSpreadsheet().getSheetByName(T_RAW);
    const rows = [];
    if (raw && raw.getLastRow() > 1) {
      const data = raw.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        rows.push({
          ts: new Date(data[i][0]).toISOString(),
          rid: String(data[i][1]),
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
    if (settings.publicReport === 'לא') return { success: false, message: 'הדוח הציבורי עדיין לא פורסם' };

    const survey = loadSurvey();
    const raw = getSpreadsheet().getSheetByName(T_RAW);
    if (!survey || !raw || raw.getLastRow() <= 1) return { success: false, message: 'אין עדיין נתונים' };

    // מיזוג: תשובה אחרונה לכל (עונה, שאלה)
    const data = raw.getDataRange().getValues();
    const perRid = {};
    for (let i = 1; i < data.length; i++) {
      const rid = String(data[i][1]);
      perRid[rid] = perRid[rid] || {};
      Object.assign(perRid[rid], parseJsonSafe(String(data[i][3]), {}));
    }
    const respondents = Object.keys(perRid).map(function (k) { return perRid[k]; });

    const EXCLUDE = { about_name: 1, about_submitted: 1, comm_volunteer_details: 1 };
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
// תחזוקה
// ============================================================

// מומלץ: טריגר כל 10 דקות למניעת cold start
// Apps Script → Triggers → Add Trigger → keepWarm → Time-driven → Every 10 min
function keepWarm() { Logger.log('warm ' + new Date().toISOString()); }

// ============================================================
// ראוטר — endpoint יחיד
// GET לקריאות · POST (text/plain) לכתיבות ארוכות
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
        return jsonResponse({ success: true, version: 'v1' });

      case 'getSurvey':
        return jsonResponse(getSurvey());

      case 'submitChapter':
        return jsonResponse(submitChapter(p.rid, p.chapter, typeof p.answers === 'string' ? p.answers : JSON.stringify(p.answers || {})));

      case 'resume':
        return jsonResponse(resume(p.rid));

      case 'linkResume':
        return jsonResponse(linkResume(p.hash, p.rid));

      case 'resumeByHash':
        return jsonResponse(resumeByHash(p.hash));

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
