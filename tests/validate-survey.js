// בדיקת שלמות לשאלון: מזהים, תנאים, סוגים, מקפים. הרצה: node tests/validate-survey.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.resolve(process.argv[2] || path.join(__dirname, '..'));
const ctx = { window: {} };
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(ROOT, 'survey-data.js'), 'utf8'), ctx);
const S = ctx.window.SURVEY_DATA;

const errors = [];
const warn = [];
const catIds = new Set(S.meta.categories.map(c => c.id));
const chIds = new Set();
const qById = {};
let qCount = 0;

for (const ch of S.chapters) {
  if (chIds.has(ch.id)) errors.push('duplicate chapter id ' + ch.id);
  chIds.add(ch.id);
  if (ch.id !== 'about' && !catIds.has(ch.cat)) errors.push('bad cat in ' + ch.id + ': ' + ch.cat);
  for (const q of ch.questions) {
    qCount++;
    if (qById[q.id]) errors.push('duplicate question id ' + q.id);
    qById[q.id] = { q, ch };
    if (!/^[a-z0-9_]+$/.test(q.id)) errors.push('id not snake_case: ' + q.id);
    if (!['scale', 'radio', 'checkbox', 'text', 'textarea', 'rank', 'number'].includes(q.type)) errors.push('bad type ' + q.id);
    if ((q.type === 'radio' || q.type === 'checkbox') && (!q.opts || q.opts.length < 2)) errors.push('choice without opts ' + q.id);
    if (q.type === 'checkbox' && q.max && q.max >= (q.opts.length + (q.other ? 1 : 0))) warn.push('max >= options ' + q.id);
    if (q.type === 'scale' && !(q.min >= 0 && q.max > q.min)) errors.push('bad scale range ' + q.id);
    if (q.opts && new Set(q.opts).size !== q.opts.length) errors.push('duplicate option in ' + q.id);
    if (q.opts && q.opts.includes('אחר') && q.other) errors.push('both other:true and "אחר" option in ' + q.id);
    if (q.exclusive && !q.opts.includes(q.exclusive)) errors.push('exclusive not in opts ' + q.id);
    if ('required' in q && typeof q.required !== 'boolean') errors.push('required must be true/false ' + q.id);
  }
}
// הפרטים הדמוגרפיים חובה, השם רשות
const about = S.chapters.find(c => c.id === 'about');
if (!about || about.questions.filter(q => q.required).length < 4) errors.push('about: main demographics should be required');
if (about && about.questions.some(q => q.id === 'about_name' && q.required)) errors.push('about_name must stay optional');

function checkCond(where, cond) {
  if (!cond) return;
  const target = qById[cond.q];
  if (!target) { errors.push(where + ': showIf refers to missing question ' + cond.q); return; }
  const vals = cond.vals || cond.in || cond.any || cond.notIn || [];
  for (const v of vals) {
    if (!(target.q.opts || []).includes(v)) errors.push(where + ': showIf value "' + v + '" not an option of ' + cond.q);
  }
}
for (const ch of S.chapters) {
  checkCond('chapter ' + ch.id, ch.showIf);
  for (const q of ch.questions) checkCond('question ' + q.id, q.showIf);
}

// מקף ארוך / מקף עברי בטקסט שהתושבים רואים
const dashRe = /[—־]/;
const scanText = (where, s) => { if (typeof s === 'string' && dashRe.test(s)) errors.push('dash in ' + where + ': ' + s); };
const walk = (obj, where) => {
  if (typeof obj === 'string') return scanText(where, obj);
  if (Array.isArray(obj)) return obj.forEach((x, i) => walk(x, where + '[' + i + ']'));
  if (obj && typeof obj === 'object') for (const k of Object.keys(obj)) walk(obj[k], where + '.' + k);
};
walk(S, 'SURVEY');

const perCat = {};
for (const ch of S.chapters) if (ch.cat) perCat[ch.cat] = (perCat[ch.cat] || 0) + 1;

console.log('chapters:', S.chapters.length, '(incl. about)', '| questions:', qCount, '| per category:', JSON.stringify(perCat));
console.log('types:', JSON.stringify(Object.values(qById).reduce((a, { q }) => (a[q.type] = (a[q.type] || 0) + 1, a), {})));
if (warn.length) console.log('WARN:\n  ' + warn.join('\n  '));
if (errors.length) { console.log('ERRORS:\n  ' + errors.join('\n  ')); process.exit(1); }
console.log('OK: no errors');
