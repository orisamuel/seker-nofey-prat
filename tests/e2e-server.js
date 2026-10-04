// שרת בדיקה מקומי: מגיש את האתר, ומריץ את appscript.gs האמיתי על גיליון מדומה בכתובת /exec.
// config.js מוחלף בגרסה שמצביעה ל-/exec, כך שהאתר עובד במצב שרת מלא.
const http = require('http');
const fs = require('fs');
const path = require('path');
const { createGas } = require('./mock-gas.js');

const ROOT = process.argv[2] || path.join(__dirname, '..');
const PORT = Number(process.argv[3] || 5178);
const g = createGas(ROOT);
g.ctx.setup();

const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.png': 'image/png', '.svg': 'image/svg+xml', '.json': 'application/json' };

http.createServer((req, res) => {
  const url = new URL(req.url, 'http://localhost');
  const cors = { 'Access-Control-Allow-Origin': '*' };

  if (url.pathname === '/exec') {
    let body = '';
    req.on('data', c => (body += c));
    req.on('end', () => {
      const parameter = Object.fromEntries(url.searchParams);
      const out = req.method === 'POST'
        ? g.ctx.doPost({ parameter, postData: { contents: body } })
        : g.ctx.doGet({ parameter });
      // השהיה קטנה, כמו Apps Script אמיתי
      setTimeout(() => { res.writeHead(200, { ...cors, 'Content-Type': 'application/json' }); res.end(out.getContent()); }, 350);
    });
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
