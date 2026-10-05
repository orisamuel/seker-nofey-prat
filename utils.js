/**
 * utils.js — עזרים משותפים לכל הדפים
 */

// ── API ──────────────────────────────────────────────────

// שגיאה שהשרת ניסח (למשל "הסקר סגור כרגע למענה"). מציגים אותה למשתמש כמו שהיא,
// להבדיל מתקלת רשת שעליה מציגים הודעה כללית.
class ServerSaid extends Error {}

function isLocalMode() {
  return typeof CONFIG === 'undefined' || !CONFIG.SCRIPT_URL || CONFIG.SCRIPT_URL.includes('PASTE_');
}

// חימום מוקדם של Apps Script (cold start ~3-5 שניות)
function warmupServer() {
  if (isLocalMode()) return;
  fetch(CONFIG.SCRIPT_URL + '?action=ping').catch(() => {});
}

// Apps Script עונה בהפניה לכתובת זמנית של גוגל, ובערך פעם בעשר זה נכשל בצד של גוגל: מגיע דף שגיאה
// (שהדפדפן מציג כשגיאת רשת), או שהבקשה מגיעה לשרת בלי הפרמטרים ("Unknown action: undefined").
// אז מנסים שוב, פעם אחת. כל הפעולות בטוחות לניסיון חוזר (בהגרלה: ראו showRaffleModal ב-index.html).
async function withRetry(send) {
  for (let attempt = 1; ; attempt++) {
    try {
      const data = await send();
      const flaky = data.raw !== undefined || /^Unknown action: undefined/.test(data.message || '');
      if (!flaky) { if (attempt > 1) data.retried = true; return data; }
      if (attempt >= 2) return { success: false, message: 'לא הצליח, נסו שוב', retried: true };
    } catch (e) {
      if (attempt >= 2) throw e;
    }
    await new Promise(r => setTimeout(r, 700));
  }
}

async function readJson(res) {
  if (!res.ok) throw new Error('Server error: ' + res.status);
  const text = await res.text();
  try { return JSON.parse(text); } catch { return { success: false, raw: text }; }
}

// קריאות ציבוריות — GET (CORS פשוט, בלי preflight)
async function apiCall(action, params = {}) {
  if (isLocalMode()) throw new Error('SCRIPT_URL not configured');
  const url = CONFIG.SCRIPT_URL + '?' + new URLSearchParams({ action, ...params });
  return withRetry(async () => readJson(await fetch(url, { redirect: 'follow' })));
}

// כתיבות, וכל קריאה שיש בה מידע אישי או סיסמה (מייל, תשובות, טלפון, הגרלה, דשבורד) — POST עם גוף text/plain,
// כדי שהמידע לא ייכנס לכתובת ולא יישמר בלוגים ובהיסטוריה.
// זו "בקשה פשוטה" מבחינת CORS (בלי preflight), ו-Apps Script קורא אותה מ-e.postData.
async function apiPost(action, payload = {}) {
  if (isLocalMode()) throw new Error('SCRIPT_URL not configured');
  return withRetry(async () => readJson(await fetch(CONFIG.SCRIPT_URL, {
    method: 'POST',
    redirect: 'follow',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action, ...payload }),
  })));
}

// ── טוסטים ───────────────────────────────────────────────

function showToast(message, type = 'info', duration = 3500) {
  let container = document.getElementById('toastContainer');
  if (!container) {
    container = document.createElement('div');
    container.id = 'toastContainer';
    container.className = 'toast-container';
    document.body.appendChild(container);
  }
  const toast = document.createElement('div');
  toast.className = 'toast toast-' + type;
  const icons = { success: '✓', error: '✕', warning: '⚠', info: 'ℹ' };
  const icon = document.createElement('span');
  icon.className = 'toast-icon';
  icon.textContent = icons[type] || 'ℹ';
  const msg = document.createElement('span');
  msg.className = 'toast-msg';
  msg.textContent = message;
  toast.append(icon, msg);
  container.appendChild(toast);
  requestAnimationFrame(() => toast.classList.add('toast-show'));
  setTimeout(() => {
    toast.classList.remove('toast-show');
    toast.classList.add('toast-hide');
    setTimeout(() => toast.remove(), 350);
  }, duration);
}

// ── המתנה ארוכה ──────────────────────────────────────────
// השרת של גוגל עונה לפעמים רק אחרי עשרות שניות. במקום ספינר שקט: שלבים מתחלפים שמתארים מה באמת
// קורה, ופס שמתקדם לאט ולא מגיע לסוף עד שהתשובה מגיעה. ככה המתנה ארוכה מרגישה סבירה.
// opts.slow: משפט למקרה שכבר עברו כל השלבים ועדיין מחכים. opts.inline: גרסה קטנה, בתוך חלון.
const SLOW_MSG = 'השרת של גוגל לוקח היום את הזמן שלו. עוד רגע…';

function loadingSteps(steps, opts = {}) {
  let text = el('p', { class: 'ld-text' }, steps[0]);
  const fill = el('i');
  const node = el('div', { class: 'ld' + (opts.inline ? ' ld-inline' : ''), role: 'status', 'aria-live': 'polite' },
    el('div', { class: 'spinner' }), text, el('div', { class: 'ld-bar', 'aria-hidden': 'true' }, fill));
  const start = Date.now();
  let ticks = 0;
  const timer = setInterval(() => {
    if (++ticks > 2 && !node.isConnected) { clearInterval(timer); return; } // המסך כבר התחלף
    const secs = (Date.now() - start) / 1000;
    fill.style.width = Math.max(6, Math.round(92 * (1 - Math.exp(-secs / 9)))) + '%';
    const step = Math.floor(secs / 2.4);
    const msg = step < steps.length ? steps[step] : (opts.slow && secs > steps.length * 2.4 + 4 ? opts.slow : null);
    if (msg && msg !== text.textContent) {
      const t = el('p', { class: 'ld-text' }, msg);
      text.replaceWith(t);
      text = t;
    }
  }, 400);
  return node;
}

// ── עזרי DOM ─────────────────────────────────────────────

function el(tag, attrs = {}, ...children) {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === 'class') node.className = v;
    else if (k === 'text') node.textContent = v;
    else if (k === 'html') node.innerHTML = v;
    else if (k.startsWith('on') && typeof v === 'function') node.addEventListener(k.slice(2), v);
    else if (v !== null && v !== undefined && v !== false) node.setAttribute(k, v === true ? '' : v);
  }
  for (const c of children) {
    if (c === null || c === undefined) continue;
    node.append(c.nodeType ? c : document.createTextNode(c));
  }
  return node;
}

// ── מזהה אנונימי + קוד המשך ─────────────────────────────

const RID_KEY = 'np_survey_rid';

function makeRid() {
  // 8 תווים קריאים — בלי תווים דו-משמעיים (0/O, 1/I/L)
  const chars = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
  let out = '';
  const rnd = new Uint32Array(8);
  crypto.getRandomValues(rnd);
  for (let i = 0; i < 8; i++) out += chars[rnd[i] % chars.length];
  return out;
}

function getRid() {
  let rid = localStorage.getItem(RID_KEY);
  if (!rid) {
    rid = makeRid();
    localStorage.setItem(RID_KEY, rid);
  }
  return rid;
}

function setRid(rid) {
  localStorage.setItem(RID_KEY, rid.toUpperCase().trim());
}

// ── הסקר מהשרת מול survey-data.js ─────────────────────────
// מה שנערך בדף הניהול גובר. שרת שעוד לא נשמר מדף הניהול (legacy) נזרע מהקוד לפני שהיה דף ניהול,
// והקוד עדכני ממנו: עד השמירה הראשונה שם, לוקחים את הכול מהקוד.
function mergeSurvey(server, bundled, legacy) {
  if (legacy || !server || !Array.isArray(server.chapters)) return bundled;
  const sm = server.meta || {}, bm = bundled.meta;
  const meta = Object.assign({}, bm, sm, {
    version: bm.version, // איפוס טיוטות נקבע בקוד
    raffle: Object.assign({}, bm.raffle, sm.raffle || {}),
    // תחום: הסדר, האייקון והצבע מהקוד, השם מדף הניהול
    categories: bm.categories.map(c => Object.assign({}, c, { title: ((sm.categories || []).find(x => x.id === c.id) || c).title })),
  });
  return { meta, chapters: server.chapters };
}

// תשובת "לא יודע/ת" בסולם (dontKnow). טקסט ולא מספר, אז לא נכנסת לממוצעים
const DONT_KNOW = 'לא יודע/ת';

// ── "משהו שהיית רוצה להוסיף?" בסוף כל נושא ────────────────
// נוסף לכל נושא בקוד ולא בגיליון, כדי שיופיע תמיד בלי לזרוע מחדש. לא בפרופיל, ולא בפאב
// שכבר מסתיים בפידבק פתוח. שאלה בגיליון עם אותו id (למשל post_more) מחליפה את זו האוטומטית.
const MORE_SKIP = ['about', 'pub'];

function withMore(ch) {
  const qs = ch.questions || [];
  if (MORE_SKIP.includes(ch.id) || qs.some(q => q.id === ch.id + '_more')) return qs;
  return qs.concat({ id: ch.id + '_more', type: 'textarea', text: 'משהו שהיית רוצה להוסיף?' });
}

// ── הערכת תנאי showIf ────────────────────────────────────
// פורמט מאוחד: {q, vals:[..], neg?} — מוצג אם התשובה נמצאת ב-vals
// (בבחירה מרובה: אם יש חפיפה כלשהי). neg הופך את התנאי.
// תומך גם בפורמט הישן: {in / notIn / any}.
// אם אין תשובה לשאלת התנאי — מציגים (ברירת מחדל פתוחה).

function evalCond(cond, answers) {
  if (!cond || !cond.q) return true;
  const val = answers[cond.q];
  if (val === undefined || val === null || val === '' ||
      (Array.isArray(val) && val.length === 0)) return true;
  const vals = cond.vals || cond.in || cond.any || cond.notIn || [];
  const neg = !!cond.neg || !!cond.notIn;
  const match = Array.isArray(val)
    ? vals.some(v => val.includes(v))
    : vals.includes(val);
  return neg ? !match : match;
}
