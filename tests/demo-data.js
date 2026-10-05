// נתוני דמה לשרת המדומה: עונים מגוונים עם פרופיל, נטישה באמצע, "לא יודע/ת" ותשובות פתוחות,
// מפוזרים על פני כמה ימים ושעות. רק לבדיקה של הדשבורד (tests/e2e-server.js, /__demo).
const fs = require('fs');
const path = require('path');
const vm = require('vm');

function loadSurvey(root) {
  const ctx = { window: {} };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(root, 'survey-data.js'), 'utf8'), ctx);
  return JSON.parse(JSON.stringify(ctx.window.SURVEY_DATA));
}

// מחולל אקראי עם זרע, כדי שכל הרצה תיתן אותם נתונים
function rng(seed) {
  let s = seed >>> 0;
  return () => ((s = (s * 1664525 + 1013904223) >>> 0) / 4294967296);
}

const TEXTS = {
  budget: ['יותר תקציב לנוער ולתרבות', 'לשקף את התקציב לתושבים בצורה ברורה', 'פחות על אירועים, יותר על תשתיות'],
  service: ['שעות קבלה גם בערב', 'לענות לטלפונים מהר יותר', 'אפשרות לפנות בוואטסאפ'],
  security: ['תאורה בכניסה ליישוב', 'השער נפתח לאט בבוקר', 'יותר סיורים בלילה'],
  parks: ['הצללה בגינה המרכזית', 'מתקנים לגיל הרך', 'ניקיון בגינה אחרי שבת'],
  garden: ['הגינון ליד הגן מוזנח', 'יותר עצים בצד המזרחי', 'השקיה בלילה ולא בבוקר'],
  clean: ['פחים נוספים ליד הוואדי', 'ניקיון אחרי אירועים', 'שלטים נגד השלכת פסולת'],
  post: ['חלוקה גם בשישי', 'הודעה כשמגיעה חבילה'],
  community: ['יותר מפגשים לותיקים', 'קבוצת הליכה בערב', 'מקום להתנדב בתרבות'],
  culture: ['ערבי מוזיקה בחוץ', 'טיול משפחות פעם בשנה', 'שבת קהילה עם כל השכונות'],
  pub: ['עוד ערבי מוזיקה חיה', 'שעות פתיחה מוקדמות יותר', 'אוכל טוב יותר'],
  _more: ['תודה על הסקר', 'חשוב לחזור לתושבים עם התוצאות', 'יותר תאורה ברחובות', 'חניה ליד בית הכנסת', 'כלבים משוחררים בגינה'],
};

function fillDemo(g, root, n) {
  const survey = loadSurvey(root);
  const R = rng(42);
  const pick = (opts, w) => { let x = R() * w.reduce((a, b) => a + b, 0); for (let i = 0; i < opts.length; i++) { x -= w[i]; if (x <= 0) return opts[i]; } return opts[opts.length - 1]; };
  const chance = (p) => R() < p;
  const rid = () => Array.from({ length: 8 }, () => 'ABCDEFGHJKMNPQRSTUVWXYZ23456789'[Math.floor(R() * 31)]).join('');
  const raw = g.SS.getSheetByName('תשובות גולמי');
  const emails = g.SS.getSheetByName('מיילים');
  const raffle = g.SS.getSheetByName('הגרלה');
  const about = survey.chapters.find(c => c.id === 'about');
  const opt = (id) => about.questions.find(q => q.id === id).opts;
  const cats = survey.meta.categories.map(c => c.id);
  const topics = cats.flatMap(c => survey.chapters.filter(ch => ch.cat === c));

  for (let i = 0; i < n; i++) {
    const id = rid();
    const p = {
      about_age: pick(opt('about_age'), [8, 30, 32, 20, 10]),
      about_seniority: pick(opt('about_seniority'), [25, 30, 20, 15, 10]),
      about_housing: pick(opt('about_housing'), [55, 20, 15, 10]),
      about_submitted: 'כן',
    };
    const kids = opt('about_kids');
    const k = [];
    if (chance(0.45)) k.push(kids[0]);
    if (chance(0.5)) k.push(kids[1]);
    if (chance(0.35)) k.push(kids[2]);
    p.about_kids = k.length ? k : [kids[3]];
    if (chance(0.9)) p.about_gender = pick(opt('about_gender'), [52, 45, 3]);
    if (chance(0.85)) p.about_aguda = pick(opt('about_aguda'), [50, 35, 15]);
    if (chance(0.2)) p.about_name = 'עונה ' + (i + 1);

    // יום והשעה של ההתחלה: רוב בערב, קצת בבוקר
    const day = Math.floor(R() * 9);
    const hour = pick([7, 8, 12, 13, 17, 19, 20, 21, 22, 23], [5, 6, 4, 3, 5, 9, 14, 16, 10, 4]);
    let t = new Date(); t.setDate(t.getDate() - day); t.setHours(hour, Math.floor(R() * 60), 0, 0);
    const stamp = () => { t = new Date(t.getTime() + (1 + Math.floor(R() * 3)) * 60000); return new Date(t); };

    const newcomer = p.about_seniority === opt('about_seniority')[0];
    const renter = p.about_housing === opt('about_housing')[1];
    const ctxAns = Object.assign({}, p);
    const relevant = topics.filter(ch => g.ctx.evalCond(ch.showIf, ctxAns));
    // כמה נושאים יעבור: רובם הכול, חלק עוצרים באמצע
    const stamina = chance(0.62) ? relevant.length : Math.floor(R() * relevant.length);
    const aboutAfterFirst = chance(0.5);

    emails.appendRow([new Date(t), 'demo' + id, id]);
    if (!aboutAfterFirst) raw.appendRow([stamp(), id, 'about', JSON.stringify(p)]);
    relevant.slice(0, stamina).forEach((ch, idx) => {
      if (!ch.noSkip && chance(0.08)) { raw.appendRow([stamp(), id, ch.id, JSON.stringify({ _skipped: 'כן' })]); return; }
      const a = {};
      ch.questions.forEach(q => {
        if (!g.ctx.evalCond(q.showIf, Object.assign({}, ctxAns, a))) return;
        if (q.type === 'scale') {
          const lo = q.min || 1, hi = q.max || 10;
          if (q.dontKnow && chance(newcomer ? 0.3 : 0.1)) { a[q.id] = 'לא יודע/ת'; return; }
          let m = lo + (hi - lo) * 0.62;
          if (renter && (ch.id === 'security' || ch.id === 'service')) m -= 1.6;
          if (newcomer && (ch.id === 'community' || ch.id === 'klita')) m -= (hi - lo) * 0.22;
          if (ch.id === 'clean' || ch.id === 'post') m -= 1.2;
          if (ch.id === 'zoo' || ch.id === 'lib') m += 0.8;
          a[q.id] = Math.max(lo, Math.min(hi, Math.round(m + (R() - 0.5) * (hi - lo) * 0.5)));
        } else if (q.type === 'radio') {
          a[q.id] = pick(q.opts, q.opts.map((_, j) => (j === 0 ? 3 : 2 - j * 0.3 > 0.4 ? 2 - j * 0.3 : 0.4)));
        } else if (q.type === 'checkbox') {
          let sel = q.opts.filter(() => chance(q.max ? 0.55 : 0.35));
          if (q.exclusive) sel = sel.filter(o => o !== q.exclusive);
          if (q.max) sel = sel.slice(0, q.max);
          if (!sel.length) sel = [q.opts[0]];
          if (q.other && chance(0.1)) sel.push('אחר: ' + pick(['טיולים', 'הרצאות', 'ספורט'], [1, 1, 1]));
          a[q.id] = sel;
        } else if ((q.type === 'text' || q.type === 'textarea') && chance(0.35)) {
          const pool = TEXTS[ch.id] || TEXTS._more;
          a[q.id] = pick(pool, pool.map(() => 1));
        }
      });
      if (chance(0.25)) a[ch.id + '_more'] = pick(TEXTS._more, TEXTS._more.map(() => 1));
      raw.appendRow([stamp(), id, ch.id, JSON.stringify(a)]);
      if (aboutAfterFirst && idx === 0 && chance(0.9)) raw.appendRow([stamp(), id, 'about', JSON.stringify(p)]);
    });
    if (stamina === relevant.length && chance(0.7)) raffle.appendRow(['עונה ' + (i + 1), "'05" + String(10000000 + i)]);
  }
  // נכנסו עם מייל ולא שמרו כלום
  for (let i = 0; i < Math.round(n * 0.15); i++) emails.appendRow([new Date(), 'demo-x' + i, rid()]);
}

module.exports = { fillDemo, loadSurvey };
