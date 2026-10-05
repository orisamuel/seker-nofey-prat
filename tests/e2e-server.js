// שרת בדיקה מקומי: מגיש את האתר, ומריץ את appscript.gs האמיתי על גיליון מדומה בכתובת /exec.
// config.js מוחלף בגרסה שמצביעה ל-/exec, כך שהאתר עובד במצב שרת מלא.
const http = require('http');
const fs = require('fs');
const path = require('path');
const { createGas } = require('./mock-gas.js');
const { fillDemo, loadSurvey } = require('./demo-data.js');

const ROOT = path.resolve(process.argv[2] || path.join(__dirname, '..'));
const PORT = Number(process.argv[3] || 5178);
const g = createGas(ROOT);
g.ctx.setup();
let delay = 350, failNext = 0;

const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.png': 'image/png', '.svg': 'image/svg+xml', '.json': 'application/json' };

http.createServer((req, res) => {
  const url = new URL(req.url, 'http://localhost');
  const cors = { 'Access-Control-Allow-Origin': '*' };

  if (url.pathname === '/exec') {
    let body = '';
    req.on('data', c => (body += c));
    req.on('end', () => {
      // תקלה מדומה (כמו דף השגיאה של גוגל): /__fail?n=3 מפיל את 3 הבקשות הבאות
      if (failNext > 0) { failNext--; setTimeout(() => { res.writeHead(502); res.end('<html>error</html>'); }, delay); return; }
      const parameter = Object.fromEntries(url.searchParams);
      const out = req.method === 'POST'
        ? g.ctx.doPost({ parameter, postData: { contents: body } })
        : g.ctx.doGet({ parameter });
      // השהיה כמו Apps Script אמיתי. /__slow?ms=8000 מדמה את התגובות האיטיות שלו
      setTimeout(() => { res.writeHead(200, { ...cors, 'Content-Type': 'application/json' }); res.end(out.getContent()); }, delay);
    });
    return;
  }

  if (url.pathname === '/__slow') { delay = Number(url.searchParams.get('ms')) || 350; res.writeHead(200, cors); res.end(String(delay)); return; }
  if (url.pathname === '/__fail') { failNext = Number(url.searchParams.get('n')) || 1; res.writeHead(200, cors); res.end(String(failNext)); return; }

  // נתוני דמה לבדיקת הדשבורד: /__demo?n=80 (שומר את השאלון מהקוד ומוסיף עונים)
  if (url.pathname === '/__demo') {
    const n = Math.min(500, Number(url.searchParams.get('n')) || 80);
    const pw = g.SS.getSheetByName('הגדרות').rows.find(r => r[0] === 'dashboardPassword')[1];
    g.ctx.saveSurvey(pw, JSON.stringify(loadSurvey(ROOT)), '', true);
    fillDemo(g, ROOT, n);
    res.writeHead(200, { ...cors, 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({ ok: true, n }));
    return;
  }

  // מבט פנימי על הגיליון המדומה, לבדיקה
  if (url.pathname === '/__sheet') {
    const name = url.searchParams.get('name');
    const sh = name ? g.SS.getSheetByName(name) : null;
    res.writeHead(200, { ...cors, 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify(sh ? sh.rows : Object.keys(g.SS.sheets)));
    return;
  }

  if (url.pathname === '/config.js') {
    res.writeHead(200, { 'Content-Type': TYPES['.js'] });
    res.end("const CONFIG = { SCRIPT_URL: 'http://localhost:" + PORT + "/exec', SHEETS_URL: '#', APP_NAME: 'סקר התושבים השנתי', YISHUV: 'נופי פרת' };");
    return;
  }

  let p = decodeURIComponent(url.pathname);
  if (p === '/') p = '/index.html';
  const file = path.join(ROOT, p);
  if (!file.startsWith(path.resolve(ROOT)) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) {
    res.writeHead(404); res.end('not found'); return;
  }
  res.writeHead(200, { 'Content-Type': TYPES[path.extname(file)] || 'application/octet-stream', 'Cache-Control': 'no-store' });
  fs.createReadStream(file).pipe(res);
}).listen(PORT, () => console.log('e2e server on http://localhost:' + PORT + '  password=' + g.SS.getSheetByName('הגדרות').rows.find(r => r[0] === 'dashboardPassword')[1]));
