/**
 * utils.js — עזרים משותפים לכל הדפים
 */

// ── API ──────────────────────────────────────────────────

function isLocalMode() {
  return typeof CONFIG === 'undefined' || !CONFIG.SCRIPT_URL || CONFIG.SCRIPT_URL.includes('PASTE_');
}

// חימום מוקדם של Apps Script (cold start ~3-5 שניות)
function warmupServer() {
  if (isLocalMode()) return;
  fetch(CONFIG.SCRIPT_URL + '?action=ping').catch(() => {});
}

// הדרך היחידה לקרוא לשרת — תמיד GET (CORS פשוט, בלי preflight)
async function apiCall(action, params = {}) {
  if (isLocalMode()) throw new Error('SCRIPT_URL not configured');
  const url = CONFIG.SCRIPT_URL + '?' + new URLSearchParams({ action, ...params });
  const res = await fetch(url, { redirect: 'follow' });
  if (!res.ok) throw new Error('Server error: ' + res.status);
  const text = await res.text();
  try { return JSON.parse(text); } catch { return { success: false, message: 'Bad response', raw: text }; }
}

// כתיבות עם מטען גדול (תשובות פרק, זריעת שאלות) — POST עם גוף text/plain.
// זו "בקשה פשוטה" מבחינת CORS (בלי preflight), ו-Apps Script קורא אותה מ-e.postData.
async function apiPost(action, payload = {}) {
  if (isLocalMode()) throw new Error('SCRIPT_URL not configured');
  const res = await fetch(CONFIG.SCRIPT_URL, {
    method: 'POST',
    redirect: 'follow',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action, ...payload }),
  });
  if (!res.ok) throw new Error('Server error: ' + res.status);
  const text = await res.text();
  try { return JSON.parse(text); } catch { return { success: false, message: 'Bad response', raw: text }; }
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

// ── הערכת תנאי showIf ────────────────────────────────────
// תנאי: {q, in:[..]} | {q, notIn:[..]} | {q, any:[..]}
// אם אין תשובה לשאלת התנאי — מציגים (ברירת מחדל פתוחה).

function evalCond(cond, answers) {
  if (!cond) return true;
  const val = answers[cond.q];
  if (val === undefined || val === null || val === '' ||
      (Array.isArray(val) && val.length === 0)) return true;
  if (cond.in)    return cond.in.includes(val);
  if (cond.notIn) return !cond.notIn.includes(val);
  if (cond.any)   return Array.isArray(val) && cond.any.some(o => val.includes(o));
  return true;
}
