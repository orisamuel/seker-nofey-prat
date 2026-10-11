// בדיקת כל זרימות השרת של appscript.gs מול גיליון מדומה. הרצה: node tests/backend.test.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');

const ROOT = path.resolve(process.argv[2] || path.join(__dirname, '..'));

const { createGas } = require('./mock-gas.js');
const { ctx, SS, props, cache, triggers, logs } = createGas(ROOT);

const sctx = { window: {} };
vm.createContext(sctx);
vm.runInContext(fs.readFileSync(path.join(ROOT, 'survey-data.js'), 'utf8'), sctx);
const SURVEY = JSON.parse(JSON.stringify(sctx.window.SURVEY_DATA));

// קריאה דרך הראוטר, כמו שהדפדפן קורא
const post = (body) => JSON.parse(ctx.doPost({ parameter: {}, postData: { contents: JSON.stringify(body) } }).getContent());
const get = (params) => JSON.parse(ctx.doGet({ parameter: params }).getContent());

let pass = 0, fail = 0;
function check(name, cond, extra) {
  if (cond) { pass++; } else { fail++; console.log('✗ ' + name + (extra !== undefined ? '  →  ' + JSON.stringify(extra).slice(0, 300) : '')); }
}

// ── 1. setup ──
ctx.setup();
check('setup: notice tab', !!SS.getSheetByName('קרא אותי'));
const pw = () => SS.getSheetByName('הגדרות').rows.find(r => r[0] === 'dashboardPassword')[1];
check('setup: random password created', typeof pw() === 'string' && pw().length === 10 && pw() !== 'nofim2026', pw());
check('setup: keepWarm trigger', triggers.length === 1);
ctx.setup();
check('setup twice: still one trigger', triggers.length === 1);
check('setup: password logged', logs.some(l => String(l).includes(pw())));
check('setup: no phone secret needed anymore', !props.PHONE_PEPPER);
check('setup: email secret created in script properties', typeof props.EMAIL_SECRET === 'string' && props.EMAIL_SECRET.length === 64);
check('setup: emails tab', !!SS.getSheetByName('מיילים'));

// ── 2. שמירה ראשונה מדף הניהול ──
check('ping', get({ action: 'ping' }).success === true);
check('getSurvey before the first save fails gracefully', get({ action: 'getSurvey' }).success === false);
let r = post({ action: 'seedSurvey', survey: SURVEY, password: '' });
check('old seed action is gone (no write without a password)', !r.success && /Unknown action/.test(r.message), r);
r = post({ action: 'saveSurvey', survey: SURVEY, password: '' });
check('save without password fails and writes nothing', !r.success && !SS.getSheetByName('שאלות'), r);
r = post({ action: 'saveSurvey', survey: SURVEY, password: 'wrong' });
check('save with wrong password fails', !r.success, r);
r = post({ action: 'saveSurvey', survey: SURVEY, password: pw(), rev: '' });
check('first save with the password', r.success && /^r[0-9a-f]{12}$/.test(r.rev), r);
let rev = r.rev;
r = post({ action: 'saveSurvey', survey: SURVEY, password: pw(), rev: '' });
check('save with a stale rev is a conflict', !r.success && r.conflict === true, r);
r = post({ action: 'saveSurvey', survey: SURVEY, password: pw(), rev });
check('save with the current rev', r.success && r.rev !== rev, r);
rev = r.rev;

// ── 3. round-trip: הגיליון מחזיר את אותו שאלון ──
const gs = get({ action: 'getSurvey' });
check('getSurvey ok + open', gs.success && gs.open === true, gs.message);
check('saved from admin = not legacy', gs.legacy === false, gs.legacy);
const back = gs.survey;
check('same chapter order', JSON.stringify(back.chapters.map(c => c.id)) === JSON.stringify(SURVEY.chapters.map(c => c.id)));
for (const ch of SURVEY.chapters) {
  const b = back.chapters.find(x => x.id === ch.id);
  for (const k of ['title', 'icon', 'desc', 'cat', 'intro', 'outro', 'noSkip']) check('chapter ' + ch.id + '.' + k, (b[k] || '') === (ch[k] || ''), [b[k], ch[k]]);
  check('chapter ' + ch.id + ' showIf', JSON.stringify(normCond(b.showIf)) === JSON.stringify(normCond(ch.showIf)), [b.showIf, ch.showIf]);
  check('chapter ' + ch.id + ' question order', JSON.stringify(b.questions.map(q => q.id)) === JSON.stringify(ch.questions.map(q => q.id)));
  for (const q of ch.questions) {
    const bq = b.questions.find(x => x.id === q.id);
    for (const k of ['type', 'text', 'help', 'other', 'exclusive', 'min', 'max', 'minLabel', 'maxLabel', 'dontKnow']) {
      check('q ' + q.id + '.' + k, (bq[k] === undefined ? '' : bq[k]) === (q[k] === undefined ? '' : q[k]), [bq[k], q[k]]);
    }
    check('q ' + q.id + '.opts', JSON.stringify(bq.opts || []) === JSON.stringify(q.opts || []));
    check('q ' + q.id + '.required', typeof bq.required === 'boolean' && bq.required === !!q.required, [bq.required, q.required]);
    check('q ' + q.id + '.showIf', JSON.stringify(normCond(bq.showIf)) === JSON.stringify(normCond(q.showIf)), [bq.showIf, q.showIf]);
  }
}
function normCond(c) {
  if (!c) return null;
  return { q: c.q, vals: c.vals || c.in || c.any || c.notIn, neg: !!(c.neg || c.notIn) };
}
check('meta stored', back.meta && back.meta.version === SURVEY.meta.version);
check('results tab is not built on save (only by the trigger or the dashboard)', !SS.getSheetByName('תוצאות'));

// ── 4. הגשות ──
const A = 'ABCDEFGH', B = 'BCDEFGHJ';
r = post({ action: 'submitChapter', rid: A, chapter: 'about', answers: { about_name: '=HYPERLINK("x")', about_kids: ['נוער (ז׳–י״ב)'], about_submitted: 'כן' } });
check('submit about', r.success, r);
r = post({ action: 'submitChapter', rid: A, chapter: 'community', answers: { comm_belong: 6, comm_budget3: ['נוער', 'מבוגרים', 'קליטת פנים'] } });
check('submit community', r.success, r);
r = post({ action: 'submitChapter', rid: A, chapter: 'budget', answers: { _skipped: 'כן' } });
check('skip budget', r.success, r);
r = post({ action: 'submitChapter', rid: A, chapter: 'community', answers: { comm_belong: 7 } });
check('second save of a topic', r.success, r);
r = post({ action: 'submitChapter', rid: 'bad', chapter: 'x', answers: {} });
check('bad rid rejected', !r.success);
r = post({ action: 'submitChapter', rid: A, chapter: '_בוטל', answers: {} });
check('system chapter name rejected', !r.success);
r = post({ action: 'submitChapter', rid: A, chapter: 'community', answers: { brand_new_q: 'x' } });
check('unknown question saved', r.success, r);
post({ action: 'submitChapter', rid: A, chapter: 'post', answers: { post_hours: 'לא יודע/ת', post_best_hours: '17-19' } });
post({ action: 'submitChapter', rid: A, chapter: 'lib', answers: { lib_services: '10' } });
check('submits do not touch the results tab', !SS.getSheetByName('תוצאות'));

// ── 4ב. טאב "תוצאות": כל התשובות בטבלה אחת, עם נוסח השאלות ──
r = post({ action: 'refreshResults', password: 'nope' });
check('refresh results needs the password', !r.success && !SS.getSheetByName('תוצאות'), r);
r = post({ action: 'refreshResults', password: pw() });
check('refresh results builds the tab', r.success && r.respondents === 1, r);
const flat = SS.getSheetByName('תוצאות');
const top = flat.rows[0], headers = flat.rows[1];
const rowOf = (rid) => flat.rows.slice(2).find(x => x[2] === rid);
const colOf = (text, group) => {
  // העמודה של שאלה לפי הנוסח, בתוך קבוצת הנושא שלה (כמה נושאים שואלים "משהו שהיית רוצה להוסיף?")
  let g = '';
  for (let j = 0; j < headers.length; j++) { if (top[j]) g = top[j]; if (headers[j] === text && (!group || g === group)) return j; }
  return -1;
};
const rowA = rowOf(A);
check('results: fixed columns', JSON.stringify(headers.slice(0, 7)) === JSON.stringify(['#', 'שם', 'קוד עונה', 'התחיל', 'עדכון אחרון', 'נושאים שעבר', 'סימן כלא רלוונטי']), headers.slice(0, 7));
check('results: update time in the corner', /^עונים · עודכן \d\d\/\d\d\/\d{4} \d\d:\d\d$/.test(top[0]), top[0]);
check('results: question text in the header, not ids', headers.includes('גיל') && headers.includes('עד כמה אתה מרגיש תחושת שייכות לקהילה?') && !headers.includes('about_age') && !headers.includes('comm_belong'), headers.slice(0, 14));
const titlesInOrder = top.filter(Boolean).slice(1);
const expectOrder = ['קצת עליך'].concat(SURVEY.meta.categories.flatMap(c => SURVEY.chapters.filter(ch => ch.cat === c.id).map(ch => ch.title)));
check('results: topics in the order of the survey, about first', JSON.stringify(titlesInOrder) === JSON.stringify(expectOrder), titlesInOrder);
check('results: one column per question, plus "more" per topic', headers.length === 7 + SURVEY.chapters.reduce((a, c) => a + c.questions.filter(q => q.id !== 'about_name').length + (['about', 'pub'].includes(c.id) ? 0 : 1), 0) + 1, headers.length);
check('results: "more" question in each topic', headers.filter(h => h === 'משהו שהיית רוצה להוסיף?').length === SURVEY.chapters.filter(c => !['about', 'pub'].includes(c.id)).length);
check('results: system keys hidden', !headers.includes('about_submitted') && !headers.includes('_skipped') && !headers.includes('שם מלא'));
check('results: one row per respondent', flat.rows.length === 3 && rowA[0] === 1, flat.rows.length);
check('results: name in its own column, formula neutralized', rowA[1] === '=HYPERLINK("x")', rowA[1]);
const isDate = (x) => Object.prototype.toString.call(x) === '[object Date]';
check('results: start and last update are dates', isDate(rowA[3]) && isDate(rowA[4]) && rowA[4] >= rowA[3]);
check('results: skipped topic named', rowA[6] === 'תקציב היישוב', rowA[6]);
check('results: progress against relevant topics', /^\d+ מתוך \d+$/.test(rowA[5]), rowA[5]);
check('results: multi choice readable', rowA[colOf('בחר את 3 הנושאים שלדעתך נדרשים להקצאת כספים מהתקציב')] === 'נוער · מבוגרים · קליטת פנים', rowA[colOf('בחר את 3 הנושאים שלדעתך נדרשים להקצאת כספים מהתקציב')]);
check('results: latest answer wins, scale stays a number', rowA[colOf('עד כמה אתה מרגיש תחושת שייכות לקהילה?')] === 7);
check('results: kids', rowA[colOf('ילדים בבית')] === 'נוער (ז׳–י״ב)');
check('results: dont-know kept as text', rowA[colOf('עד כמה אתה מרוצה משעות קבלת הקהל וחלוקת החבילות בדואר?')] === 'לא יודע/ת');
check('results: number-looking text stays text', rowA[colOf('לאילו שירותים משמשת אתכם הספרייה?')] === '10' && rowA[colOf('באילו שעות יהיה לכם נוח יותר לאסוף חבילות?')] === '17-19', [rowA[colOf('לאילו שירותים משמשת אתכם הספרייה?')]]);
check('results: unknown question kept, at the end of its topic', colOf('brand_new_q', 'קהילה') > colOf('עד כמה אתה מרגיש תחושת שייכות לקהילה?') && rowA[colOf('brand_new_q')] === 'x');
check('results: frozen headers, right to left, formatted', flat.frozen === 2 && flat.frozenCols === 2 && flat.rtl === true && Object.keys(flat.formats).some(k => k.startsWith('backgrounds')));
r = post({ action: 'refreshResults', password: pw() });
check('results: no rebuild when nothing changed', r.success && r.fresh === true, r);
ctx.keepWarm();
check('trigger: nothing to rebuild either', flat.rows.length === 3);
post({ action: 'submitChapter', rid: 'NEWRESP2', chapter: 'about', answers: { about_age: '26–35' } });
ctx.keepWarm();
check('trigger rebuilds after a new answer', flat.rows.length === 4 && rowOf('NEWRESP2')[0] === 2, flat.rows.length);

// ── 5. מייל = זהות (בלי שליחה) ──
// פעולות ההמשך הישנות (טלפון + קוד, קישור במייל) לא חשופות, ואין פעולה שמחזירה תשובות לפי קוד עונה
for (const action of ['linkResume', 'resumeByPhone', 'resumeByHash', 'resume', 'sendResumeLink', 'resumeByToken']) {
  r = post({ action, rid: A, phone: '0541234567', pin: '1234', email: 'dana@example.co.il', token: 'a'.repeat(64) });
  check(action + ' not exposed', !r.success && /Unknown action/.test(r.message), r);
}
r = post({ action: 'identify', email: 'not-an-email', rid: A });
check('bad email rejected', !r.success && r.message.includes('מייל'), r);
r = post({ action: 'identify', email: 'dana@example.co.il', rid: 'bad' });
check('bad rid rejected on identify', !r.success, r);
r = post({ action: 'identify', email: 'Dana@Example.co.il ', rid: A });
check('new email binds to the device rid', r.success && r.status === 'new', r);
const emails = SS.getSheetByName('מיילים');
const allCells = () => JSON.stringify(Object.values(SS.sheets).map(sh => sh.rows));
check('emails tab: one row with rid A', emails.rows.length === 2 && emails.rows[1][2] === A, emails.rows);
check('email not stored anywhere in the sheet', !allCells().toLowerCase().includes('dana'));
check('email key is keyed, not a plain sha256', /^[a-f0-9]{64}$/.test(emails.rows[1][1]) && emails.rows[1][1] !== crypto.createHash('sha256').update('dana@example.co.il').digest('hex'), emails.rows[1][1]);
check('secret not in the sheet', !allCells().includes(props.EMAIL_SECRET));
r = post({ action: 'identify', email: 'dana@example.co.il', rid: A });
check('same email, same device', r.success && r.status === 'same', r);
// מכשיר אחר עם אותו מייל: ממשיכים את הסקר של A
r = post({ action: 'identify', email: ' DANA@example.co.il', rid: 'NEWDEV22' });
check('same email from another device resumes A', r.success && r.status === 'resume' && r.rid === A, r);
check('resume: saved topics, skipped marked', r.done && r.done.about && r.done.community && r.done.community.skipped === false && r.done.budget.skipped === true, r.done);
check('resume: profile has only what drives the logic', JSON.stringify(r.profile) === JSON.stringify({ about_kids: ['נוער (ז׳–י״ב)'] }), r.profile);
check('resume: no topic answers, no name', !/comm_belong|comm_budget3|brand_new_q|about_name|HYPERLINK|מבוגרים/.test(JSON.stringify(r)), r);
check('resume adds no rows', emails.rows.length === 2);
// Gmail: נקודות, "+תוספת" ו-googlemail הם אותה תיבה
r = post({ action: 'identify', email: 'yossi.cohen+seker@gmail.com', rid: 'GMAILAA2' });
check('gmail new', r.success && r.status === 'new', r);
r = post({ action: 'identify', email: 'YossiCohen@googlemail.com', rid: 'GMAILBB3' });
check('gmail variants = same mailbox', r.success && r.status === 'resume' && r.rid === 'GMAILAA2', r);
r = post({ action: 'identify', email: 'yossi.cohen+x@walla.co.il', rid: 'WALLAAA2' });
check('other providers keep dots and plus', r.success && r.status === 'new', r);
r = post({ action: 'identify', email: 'yossicohen+x@walla.co.il', rid: 'WALLABB2' });
check('other providers: different address = different survey', r.success && r.status === 'new', r);
// הגנה מהצפה: כתובות חדשות נעצרות אחרי 400 בשעה, מוכרות ממשיכות לעבוד
cache.new_emails_hour = '400';
r = post({ action: 'identify', email: 'flood@example.com', rid: 'FLOODAA2' });
check('new emails capped per hour', !r.success && r.message.includes('עומס'), r);
r = post({ action: 'identify', email: 'dana@example.co.il', rid: A });
check('known email still works under the cap', r.success && r.status === 'same', r);
delete cache.new_emails_hour;
check('no old link/phone tabs', !SS.getSheetByName('קישורי המשך') && !SS.getSheetByName('קודי המשך'));

// ── 6. הגרלה ──
r = post({ action: 'enterRaffle', rid: A, name: 'ישראל ישראלי', phone: '0541234567' });
check('raffle blocked before completion', !r.success && r.message.includes('נותרו'), r);
// משלימים את כל הנושאים שרלוונטיים לפרופיל של A (נוער בבית)
const profile = { about_kids: ['נוער (ז׳–י״ב)'] };
const required = SURVEY.chapters.filter(ch => ctx.evalCond(ch.showIf, profile));
check('profile with teens: scouts visible, toddlers hidden', required.some(c => c.id === 'scouts') && !required.some(c => c.id === 'gilrach') && !required.some(c => c.id === 'yesodi'));
for (const ch of required) if (!['about', 'community', 'budget'].includes(ch.id)) post({ action: 'submitChapter', rid: A, chapter: ch.id, answers: { _skipped: 'כן' } });
r = post({ action: 'enterRaffle', rid: A, name: 'ישראל ישראלי', phone: '0541234567' });
check('raffle ok after completion', r.success, r);
r = post({ action: 'enterRaffle', rid: A, name: 'מישהו אחר', phone: '054-123-4567' });
check('duplicate phone rejected', !r.success && r.message.includes('כבר רשום'), r);
// עונה נוסף שמשלים הכול ונרשם, כדי לבדוק מיון ושמירת 0
const C = 'CDEFGHJK';
for (const ch of SURVEY.chapters) if (ctx.evalCond(ch.showIf, {})) post({ action: 'submitChapter', rid: C, chapter: ch.id, answers: { _skipped: 'כן' } });
r = post({ action: 'enterRaffle', rid: C, name: 'אבי אבירם', phone: '0529876543' });
check('second raffle entry', r.success, r);
const raffle = SS.getSheetByName('הגרלה');
check('raffle has no rid and no time', raffle.rows[0].length === 2 && !JSON.stringify(raffle.rows).includes(A) && !JSON.stringify(raffle.rows).includes(C));
check('raffle sorted by name', raffle.rows[1][0] === 'אבי אבירם' && raffle.rows[2][0] === 'ישראל ישראלי', raffle.rows);
check('raffle phone keeps leading zero', raffle.rows.slice(1).every(x => String(x[1]).startsWith('05')), raffle.rows);

// ── 7. התחלה מחדש ──
post({ action: 'submitChapter', rid: B, chapter: 'community', answers: { comm_belong: 1 } });
ctx.keepWarm();
check('B in results', !!rowOf(B));
r = post({ action: 'identify', email: 'b@example.com', rid: B });
check('B identified', r.success && r.status === 'new', r);
r = post({ action: 'discard', rid: B });
check('discard ok', r.success, r);
ctx.keepWarm();
check('B removed from results after discard', !rowOf(B) && !!rowOf(A));
// אחרי "להתחיל מחדש" אותו מייל עובר לקוד העונה החדש, והסקר שבוטל לא חוזר
const B2 = 'BBBBBBB2';
r = post({ action: 'identify', email: 'b@example.com', rid: B2 });
check('email of a discarded survey moves to the new one', r.success && r.status === 'new' && !r.rid, r);
check('emails row now points to the new rid', emails.rows.some(x => x[2] === B2) && !emails.rows.some(x => x[2] === B), emails.rows);
r = post({ action: 'identify', email: 'b@example.com', rid: 'OTHERDV2' });
check('another device then resumes the new survey', r.success && r.status === 'resume' && r.rid === B2 && Object.keys(r.done).length === 0, r);

// ── 8. דשבורד ──
r = post({ action: 'getResults', password: 'nope' });
check('results wrong password', !r.success);
r = post({ action: 'getResults', password: '' });
check('results empty password', !r.success);
r = post({ action: 'getResults', password: pw() });
check('results ok', r.success && r.raffleCount === 2, r.message);
check('results: how many passed the email step', r.identified === SS.getSheetByName('מיילים').rows.length - 1 && r.identified > 0, r.identified);
check('results exclude discarded B', !r.rows.some(x => x.rid === B));
check('results include A', r.rows.some(x => x.rid === A));

// ── 9. דוח ציבורי ──
r = get({ action: 'getPublicReport' });
check('public report closed by default', !r.success);
const sset = SS.getSheetByName('הגדרות');
sset.rows.find(x => x[0] === 'publicReport')[1] = 'כן';
for (let i = 0; i < 5; i++) post({ action: 'submitChapter', rid: 'R' + 'ABCDE'[i] + 'GHJKMN', chapter: 'community', answers: { comm_belong: 3 + i, comm_focus: 'טקסט חופשי ' + i } });
r = get({ action: 'getPublicReport' });
check('public report open', r.success, r.message);
const commRep = r.report.chapters.find(c => c.id === 'community');
const belong = commRep && commRep.questions.find(q => q.id === 'comm_belong');
check('report: 1-7 scale carries its range', belong && belong.min === 1 && belong.max === 7, belong);
check('report: no free text', !commRep.questions.some(q => q.id === 'comm_focus'));
check('report: <5 answers hidden', !r.report.chapters.some(c => c.questions.some(q => q.id === 'comm_budget3')));
check('report: excludes discarded B', belong && !Object.keys(belong.hist).includes('1'), belong && belong.hist);
// "לא יודע/ת" בסולם: נספר בנפרד ולא נכנס לממוצע
['RDKAAAA2', 'RDKBBBB2', 'RDKCCCC2', 'RDKDDDD2', 'RDKEEEE2'].forEach((x, i) => post({ action: 'submitChapter', rid: x, chapter: 'budget', answers: { budget_priorities: i < 3 ? 4 + i : 'לא יודע/ת' } }));
const bp = get({ action: 'getPublicReport' }).report.chapters.find(c => c.id === 'budget').questions.find(q => q.id === 'budget_priorities');
check('report: dont-know counted apart from the average', bp && bp.count === 5 && bp.dontKnow === 2 && bp.avg === 5 && !Object.keys(bp.hist).some(k => isNaN(Number(k))), bp);

// ── 10. סגירת הסקר ──
sset.rows.find(x => x[0] === 'surveyOpen')[1] = 'לא';
r = post({ action: 'submitChapter', rid: A, chapter: 'community', answers: { comm_belong: 2 } });
check('closed survey rejects submits with a clear message', !r.success && r.message === 'הסקר סגור כרגע למענה', r);
check('getSurvey reports closed', get({ action: 'getSurvey' }).open === false);

// ── 11. עריכה ידנית בגיליון: שאלה שנוספה בתחתית עם "סדר" באמצע, ושאלה מושבתת ──
const qs = SS.getSheetByName('שאלות');
qs.rows.push(['post', 'post_new', 'טקסט קצר', 'שאלה חדשה', '', '', '', '', '', '', '', '', '', 'כן', 1.5]);
qs.rows.find(x => x[1] === 'post_best_hours')[13] = 'לא';
const post2 = ctx.loadSurvey(false).chapters.find(c => c.id === 'post');
check('manual row sorted by its order column', JSON.stringify(post2.questions.map(q => q.id)) === JSON.stringify(['post_hours', 'post_new']), post2.questions.map(q => q.id));
const post2all = ctx.loadSurvey(true).chapters.find(c => c.id === 'post');
check('inactive question kept for the admin page', post2all.questions.find(q => q.id === 'post_best_hours').active === false);
check('manual sheet edits do not reach the public survey (cached)', get({ action: 'getSurvey' }).survey.chapters.find(c => c.id === 'post').questions.some(q => q.id === 'post_best_hours'));

// ── 12. גיליון שנזרע לפני עמודת "חובה": השרת לא מחזיר required, והאתר לוקח מ-survey-data.js ──
const qTab = SS.getSheetByName('שאלות');
const savedRows = qTab.rows.map(row => row.slice());
qTab.rows.forEach(row => { row.length = Math.min(row.length, 15); });
const legacy = ctx.loadSurvey(false);
check('legacy sheet: required not sent', legacy.chapters.every(c => c.questions.every(q => q.required === undefined)));
qTab.rows = savedRows;
check('with the column: about demographics required', ctx.loadSurvey(false).chapters.find(c => c.id === 'about').questions.filter(q => q.required).length === SURVEY.chapters[0].questions.filter(q => q.required).length);
// שרת שעוד לא נשמר מדף הניהול מסומן legacy, והאתר לוקח אז טקסטים ו"קצת עליך" מהקוד
const revRow = SS.getSheetByName('הגדרות').rows.find(x => x[0] === 'surveyRev');
const keepRev = revRow[1]; revRow[1] = '';
check('no rev = legacy', get({ action: 'getSurvey' }).legacy === true);
revRow[1] = keepRev;

// ── 13. דף הניהול ──
sset.rows.find(x => x[0] === 'surveyOpen')[1] = 'כן';
r = post({ action: 'getAdmin', password: 'nope' });
check('admin: wrong password', !r.success && !r.survey, r);
r = post({ action: 'getAdmin', password: pw() });
check('admin: loads survey, rev, settings', r.success && r.rev === rev && r.legacy === false && r.settings.surveyOpen === true && r.settings.publicReport === true, r.message || [r.rev, rev]);
check('admin: answered counts from results', r.answered.comm_belong >= 5 && !r.answered.about_name === false, r.answered);
// עריכה: טקסטים שהגיליון היה הופך, השבתה, אפשרות ושאלה חדשות
const ed = JSON.parse(JSON.stringify(r.survey));
const edPost = ed.chapters.find(c => c.id === 'post');
edPost.questions[0].help = '=SUM(1)';
edPost.questions[0].minLabel = '-ממש לא';
edPost.questions[0].maxLabel = '1/10';
edPost.questions.push({ id: 'post_x1', type: 'radio', text: '007', opts: ['+כן', "'לא", '10'], required: true });
ed.chapters.find(c => c.id === 'zoo').active = false;
ed.chapters.find(c => c.id === 'lib').questions[0].active = false;
ed.meta = Object.assign({}, ed.meta, { tagline: 'כותרת חדשה' });
r = post({ action: 'saveSurvey', survey: ed, password: pw(), rev });
check('admin: save edits', r.success, r);
rev = r.rev;
const pub = get({ action: 'getSurvey' }).survey;
const pq = pub.chapters.find(c => c.id === 'post').questions;
check('tricky texts survive the sheet', pq[0].help === '=SUM(1)' && pq[0].minLabel === '-ממש לא' && pq[0].maxLabel === '1/10', pq[0]);
const nq = pq.find(q => q.id === 'post_x1');
check('new question with tricky options', nq && nq.text === '007' && JSON.stringify(nq.opts) === JSON.stringify(['+כן', "'לא", '10']) && nq.required === true, nq);
check('inactive chapter hidden from the public survey', !pub.chapters.some(c => c.id === 'zoo'));
check('inactive question hidden from the public survey', !pub.chapters.find(c => c.id === 'lib').questions.some(q => q.id === 'lib_services'));
check('meta saved and served', pub.meta.tagline === 'כותרת חדשה');
const adm = post({ action: 'getAdmin', password: pw() }).survey;
check('admin still sees hidden items', adm.chapters.find(c => c.id === 'zoo').active === false && adm.chapters.find(c => c.id === 'lib').questions[0].active === false);
ctx.keepWarm();
check('results rebuilt after an admin save: new question has a column', SS.getSheetByName('תוצאות').rows[1].includes('007'), SS.getSheetByName('תוצאות').rows[1]);
check('results: hidden topic without answers left out', !SS.getSheetByName('תוצאות').rows[0].includes('פינת חי'));
check('results: hidden question with answers kept', SS.getSheetByName('תוצאות').rows[1].includes('לאילו שירותים משמשת אתכם הספרייה?'));
// אימות: שום דבר לא נכתב כשיש שגיאה
const before = JSON.stringify(SS.getSheetByName('שאלות').rows);
const bad = (fn) => { const x = JSON.parse(JSON.stringify(adm)); fn(x); return post({ action: 'saveSurvey', survey: x, password: pw(), rev }); };
const errOf = (fn) => { const res = bad(fn); return res.success ? null : (res.errors || [res.message]).join(' / '); };
check('validate: duplicate question id', /כפול/.test(errOf(x => { x.chapters[1].questions[1].id = x.chapters[1].questions[0].id; })));
check('validate: pipe in an option', /\|/.test(errOf(x => { x.chapters[0].questions[1].opts[0] = 'א | ב'; })));
check('validate: empty option', /אפשרויות/.test(errOf(x => { x.chapters[0].questions[1].opts.push(' '); })));
check('validate: one option only', /אפשרויות/.test(errOf(x => { x.chapters[0].questions[1].opts = ['רק אחת']; })));
check('validate: bad scale', /סולם/.test(errOf(x => { const q = x.chapters[1].questions[0]; q.min = 5; q.max = 3; })));
check('validate: condition to a missing question', /שלא קיימת/.test(errOf(x => { x.chapters.find(c => c.id === 'scouts').showIf = { q: 'nope', vals: ['x'] }; })));
check('validate: condition to a removed option', /שלא קיימת/.test(errOf(x => { x.chapters.find(c => c.id === 'about').questions.find(q => q.id === 'about_kids').opts.splice(1, 1); })));
check('validate: missing text', /נוסח/.test(errOf(x => { x.chapters[1].questions[0].text = ' '; })));
check('validate: bad id', /מזהה/.test(errOf(x => { x.chapters[1].questions[0].id = 'Bad Id'; })));
check('validate: not a survey', /מבנה/.test(errOf(x => { x.chapters = 'x'; })));
check('nothing written on validation errors', JSON.stringify(SS.getSheetByName('שאלות').rows) === before);
r = post({ action: 'saveSurvey', survey: adm, password: pw(), rev: 'rold' });
check('conflict on a stale rev', r.conflict === true);
r = post({ action: 'saveSurvey', survey: adm, password: pw(), rev: 'rold', force: true });
check('force overwrites', r.success, r);
rev = r.rev;
// הגדרות
r = post({ action: 'saveSettings', password: pw(), surveyOpen: false, publicReport: false });
check('settings: close survey + report', r.success && get({ action: 'getSurvey' }).open === false && !get({ action: 'getPublicReport' }).success, r);
r = post({ action: 'saveSettings', password: pw(), surveyOpen: true });
check('settings: reopen', r.success && get({ action: 'getSurvey' }).open === true);
r = post({ action: 'saveSettings', password: pw(), newPassword: 'short' });
check('settings: short password rejected', !r.success && /8/.test(r.message), r);
const oldPw = pw();
r = post({ action: 'saveSettings', password: oldPw, newPassword: '00123456' });
check('settings: numeric password kept as text', r.success && pw() === '00123456' && post({ action: 'getAdmin', password: '00123456' }).success && !post({ action: 'getAdmin', password: oldPw }).success, [r, pw()]);
// הגיליון: הסבר במקום מדריך עריכה, ואזהרה על הטאבים
check('notice tab instead of the editing guide', !!SS.getSheetByName('קרא אותי') && !SS.getSheetByName('מדריך עריכה') && SS.getSheetByName('קרא אותי').rows[1][0].includes('admin.html'));
check('tabs protected with a warning, once', ['פרקים', 'שאלות', 'הגדרות', 'תוצאות'].every(n => SS.getSheetByName(n).protections.length === 1 && SS.getSheetByName(n).protections[0].warningOnly === true));

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
