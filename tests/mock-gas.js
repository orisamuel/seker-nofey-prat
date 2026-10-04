// סימולציה של Google Sheets ו-Apps Script, שמריצה את appscript.gs האמיתי מחוץ לגוגל.
// משמשת את tests/backend.test.js ואת tests/e2e-server.js.
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');
// ── מוק של הגיליון ─────────────────────────────────────────
// מחקה את ההמרות של Sheets: גרש בהתחלה = טקסט, מחרוזת מספרית = מספר (ה-0 נאכל), '=' = נוסחה
function toCell(v) {
  if (v instanceof Date || typeof v === 'number' || typeof v === 'boolean') return v;
  if (v === null || v === undefined) return '';
  v = String(v);
  if (v.startsWith("'")) return v.slice(1);
  if (v.startsWith('=')) return '#FORMULA(' + v + ')';
  if (/^-?\d+(\.\d+)?$/.test(v)) return Number(v);
  return v;
}

class Sheet {
  constructor(name) { this.name = name; this.rows = []; this.maxCols = 26; this.frozen = 0; }
  getName() { return this.name; }
  getLastRow() {
    for (let i = this.rows.length - 1; i >= 0; i--) if (this.rows[i].some(x => x !== '' && x !== undefined)) return i + 1;
    return 0;
  }
  getLastColumn() {
    let m = 0;
    for (const r of this.rows) for (let j = r.length - 1; j >= 0; j--) if (r[j] !== '' && r[j] !== undefined) { m = Math.max(m, j + 1); break; }
    return m;
  }
  getMaxColumns() { return this.maxCols; }
  insertColumnsAfter(after, n) { this.maxCols += n; }
  appendRow(vals) {
    if (vals.length > this.maxCols) throw new Error('appendRow beyond max columns (' + vals.length + ' > ' + this.maxCols + ')');
    const at = this.getLastRow();
    this.rows.length = at;
    this.rows.push(vals.map(toCell));
  }
  getRange(r, c, nr = 1, nc = 1) {
    if (r < 1 || c < 1 || nr < 1 || nc < 1) throw new Error('bad range ' + [r, c, nr, nc]);
    if (c + nc - 1 > this.maxCols) throw new Error('The coordinates of the range are outside the dimensions of the sheet. (cols ' + (c + nc - 1) + ' > ' + this.maxCols + ')');
    return new Range(this, r, c, nr, nc);
  }
  getDataRange() { return new Range(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
  clearContents() { this.rows = []; }
  deleteRow(r) { this.rows.splice(r - 1, 1); }
  setFrozenRows(n) { this.frozen = n; }
  setColumnWidth() {}
}

class Range {
  constructor(sheet, r, c, nr, nc) { Object.assign(this, { sheet, r, c, nr, nc }); }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const row = this.sheet.rows[this.r - 1 + i] || [];
      const vals = [];
      for (let j = 0; j < this.nc; j++) { const v = row[this.c - 1 + j]; vals.push(v === undefined ? '' : v); }
      out.push(vals);
    }
    return out;
  }
  setValues(values) {
    if (values.length !== this.nr || values.some(v => v.length !== this.nc)) throw new Error('setValues dimension mismatch');
    for (let i = 0; i < this.nr; i++) {
      const ri = this.r - 1 + i;
      while (this.sheet.rows.length <= ri) this.sheet.rows.push([]);
      for (let j = 0; j < this.nc; j++) this.sheet.rows[ri][this.c - 1 + j] = toCell(values[i][j]);
    }
    return this;
  }
  setValue(v) { return this.setValues([[v]]); }
  sort(col) {
    const rows = this.getValues();
    const k = col - this.c;
    rows.sort((a, b) => String(a[k]).localeCompare(String(b[k]), 'he'));
    // כתיבה ישירה, בלי המרות (הערכים כבר בתאים)
    for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sheet.rows[this.r - 1 + i][this.c - 1 + j] = rows[i][j];
    return this;
  }
}

class Spreadsheet {
  constructor() { this.sheets = {}; }
  getSheetByName(n) { return this.sheets[n] || null; }
  insertSheet(n) { if (this.sheets[n]) throw new Error('exists ' + n); return (this.sheets[n] = new Sheet(n)); }
}

function createGas(ROOT){
const SS = new Spreadsheet();
const props = {};
const cache = {};
const triggers = [];
const logs = [];

const ctx = {
  SpreadsheetApp: { openById: () => SS },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  Utilities: {
    getUuid: () => crypto.randomUUID(),
    computeHmacSha256Signature: (value, key) => Array.from(crypto.createHmac('sha256', key).update(value, 'utf8').digest()).map(b => (b > 127 ? b - 256 : b)),
  },
  CacheService: { getScriptCache: () => ({ get: k => (k in cache ? cache[k] : null), put: (k, v) => { cache[k] = v; }, remove: k => { delete cache[k]; } }) },
  PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = v; } }) },
  ScriptApp: {
    getProjectTriggers: () => triggers,
    newTrigger: (fn) => ({ timeBased: () => ({ everyMinutes: () => ({ create: () => triggers.push({ getHandlerFunction: () => fn }) }) }) }),
  },
  ContentService: { MimeType: { JSON: 'json' }, createTextOutput: (s) => ({ s, setMimeType() { return this; }, getContent() { return this.s; } }) },
  Logger: { log: (m) => logs.push(m) },
  console,
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(ROOT, 'appscript.gs'), 'utf8'), ctx);

return { ctx, SS, props, cache, triggers, logs };
}
module.exports = { createGas };
