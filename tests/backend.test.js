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
const pw = () => SS.getSheetByName('הגדרות').rows.find(r => r[0] === 'dashboardPassword')[1];
check('setup: random password created', typeof pw() === 'string' && pw().length === 10 && pw() !== 'nofim2026', pw());
check('setup: keepWarm trigger', triggers.length === 1);
ctx.setup();
check('setup twice: still one trigger', triggers.length === 1);
check('setup: password logged', logs.some(l => String(l).includes(pw())));
check('setup: no phone secret needed anymore', !props.PHONE_PEPPER);
check('setup: email secret created in script properties', typeof props.EMAIL_SECRET === 'string' && props.EMAIL_SECRET.length === 64);
check('setup: emails tab', !!SS.getSheetByName('מיילים'));

// ── 2. seed ──
check('ping', get({ action: 'ping' }).success === true);
check('getSurvey before seed fails gracefully', get({ action: 'getSurvey' }).success === false);
let r = post({ action: 'seedSurvey', survey: SURVEY, password: '' });
check('seed first time without password', r.success, r);
r = post({ action: 'seedSurvey', survey: SURVEY, password: 'wrong' });
check('seed again with wrong password fails', !r.success, r);
r = post({ action: 'seedSurvey', survey: SURVEY, password: pw() });
check('seed again with right password', r.success, r);

// ── 3. round-trip: הגיליון מחזיר את אותו שאלון ──
const gs = get({ action: 'getSurvey' });
check('getSurvey ok + open', gs.success && gs.open === true, gs.message);
const back = gs.survey;
check('same chapter order', JSON.stringify(back.chapters.map(c => c.id)) === JSON.stringify(SURVEY.chapters.map(c => c.id)));
for (const ch of SURVEY.chapters) {
  const b = back.chapters.find(x => x.id === ch.id);
  for (const k of ['title', 'icon', 'desc', 'cat', 'intro', 'outro']) check('chapter ' + ch.id + '.' + k, (b[k] || '') === (ch[k] || ''), [b[k], ch[k]]);
  check('chapter ' + ch.id + ' showIf', JSON.stringify(normCond(b.showIf)) === JSON.stringify(normCond(ch.showIf)), [b.showIf, ch.showIf]);
  check('chapter ' + ch.id + ' question order', JSON.stringify(b.questions.map(q => q.id)) === JSON.stringify(ch.questions.map(q => q.id)));
  for (const q of ch.questions) {
    const bq = b.questions.find(x => x.id === q.id);
    for (const k of ['type', 'text', 'help', 'other', 'exclusive', 'min', 'max', 'minLabel', 'maxLabel']) {
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
const flat = SS.getSheetByName('תוצאות');
check('flat headers: one column per question', flat.getLastColumn() === 2 + SURVEY.chapters.reduce((a, c) => a + c.questions.length, 0), flat.getLastColumn());

// ── 4. הגשות ──
const A = 'ABCDEFGH', B = 'BCDEFGHJ';
r = post({ action: 'submitChapter', rid: A, chapter: 'about', answers: { about_name: '=HYPERLINK("x")', about_kids: ['נוער (ז׳–י״ב)'], about_submitted: 'כן' } });
check('submit about', r.success, r);
r = post({ action: 'submitChapter', rid: A, chapter: 'community', answers: { comm_belong: 6, comm_budget3: ['נוער', 'מבוגרים', 'קליטת פנים'] } });
check('submit community', r.success, r);
r = post({ action: 'submitChapter', rid: A, chapter: 'budget', answers: { _skipped: 'כן' } });
check('skip budget', r.success, r);
const headers = flat.rows[0];
const rowA = flat.rows.find(x => x[0] === A);
check('flat row for A', !!rowA);
check('formula injection neutralized', rowA[headers.indexOf('about_name')] === '=HYPERLINK("x")', rowA[headers.indexOf('about_name')]);
check('multi joined', rowA[headers.indexOf('comm_budget3')] === 'נוער | מבוגרים | קליטת פנים');
check('scale numeric', rowA[headers.indexOf('comm_belong')] === 6);
check('_skipped not in flat', !headers.includes('_skipped'));
r = post({ action: 'submitChapter', rid: A, chapter: 'community', answers: { comm_belong: 7 } });
check('update keeps one row per rid', flat.rows.filter(x => x[0] === A).length === 1 && flat.rows.find(x => x[0] === A)[headers.indexOf('comm_belong')] === 7);
check('update keeps other answers', flat.rows.find(x => x[0] === A)[headers.indexOf('comm_budget3')] === 'נוער | מבוגרים | קליטת פנים');
r = post({ action: 'submitChapter', rid: 'bad', chapter: 'x', answers: {} });
check('bad rid rejected', !r.success);
r = post({ action: 'submitChapter', rid: A, chapter: '_בוטל', answers: {} });
check('system chapter name rejected', !r.success);
r = post({ action: 'submitChapter', rid: A, chapter: 'community', answers: { brand_new_q: 'x' } });
check('new question column added on the fly', r.success && flat.rows[0].includes('brand_new_q'), r);

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
check('B in flat', flat.rows.some(x => x[0] === B));
r = post({ action: 'identify', email: 'b@example.com', rid: B });
check('B identified', r.success && r.status === 'new', r);
r = post({ action: 'discard', rid: B });
check('discard ok', r.success, r);
check('B removed from flat', !flat.rows.some(x => x[0] === B));
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

// ── 10. סגירת הסקר ──
sset.rows.find(x => x[0] === 'surveyOpen')[1] = 'לא';
r = post({ action: 'submitChapter', rid: A, chapter: 'community', answers: { comm_belong: 2 } });
check('closed survey rejects submits with a clear message', !r.success && r.message === 'הסקר סגור כרגע למענה', r);
check('getSurvey reports closed', get({ action: 'getSurvey' }).open === false);

// ── 11. עריכה ידנית בגיליון: שאלה שנוספה בתחתית עם "סדר" באמצע, ושאלה מושבתת ──
const qs = SS.getSheetByName('שאלות');
qs.rows.push(['post', 'post_new', 'טקסט קצר', 'שאלה חדשה', '', '', '', '', '', '', '', '', '', 'כן', 1.5]);
qs.rows.find(x => x[1] === 'post_best_hours')[13] = 'לא';
const post2 = get({ action: 'getSurvey' }).survey.chapters.find(c => c.id === 'post');
check('manual row sorted by its order column', JSON.stringify(post2.questions.map(q => q.id)) === JSON.stringify(['post_hours', 'post_new']), post2.questions.map(q => q.id));

// ── 12. גיליון שנזרע לפני עמודת "חובה": השרת לא מחזיר required, והאתר לוקח מ-survey-data.js ──
const qTab = SS.getSheetByName('שאלות');
const savedRows = qTab.rows.map(row => row.slice());
qTab.rows.forEach(row => { row.length = Math.min(row.length, 15); });
const legacy = get({ action: 'getSurvey' }).survey;
check('legacy sheet: required not sent', legacy.chapters.every(c => c.questions.every(q => q.required === undefined)));
qTab.rows = savedRows;
check('with the column: about demographics required', get({ action: 'getSurvey' }).survey.chapters.find(c => c.id === 'about').questions.filter(q => q.required).length === 6);

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
