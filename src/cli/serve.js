'use strict';
/** Web nhe de xem / tim kiem du lieu. Chay: qlnd.cmd serve  ->  http://localhost:5173 */
const fs = require('fs');
const http = require('http');
const path = require('path');
const { DEFAULTS, locate } = require('../lib/sources');
const { Grid } = require('../lib/grid');

const WEB = path.join(__dirname, '..', 'web');
const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
};

function readJson(f) {
  return JSON.parse(fs.readFileSync(f, 'utf8'));
}

function json(res, code, obj) {
  const body = JSON.stringify(obj);
  res.writeHead(code, {
    'content-type': MIME['.json'],
    'content-length': Buffer.byteLength(body),
    'cache-control': 'no-store',
  });
  res.end(body);
}

function docJson(req) {
  return new Promise((resolve, reject) => {
    let raw = '';
    req.on('data', (c) => {
      raw += c;
      if (raw.length > 1e6) reject(new Error('du lieu qua lon'));
    });
    req.on('end', () => {
      try {
        resolve(raw ? JSON.parse(raw) : {});
      } catch (e) {
        reject(new Error('JSON khong hop le'));
      }
    });
    req.on('error', reject);
  });
}

async function main() {
  const args = process.argv.slice(2);
  const opt = (n, d) => {
    const i = args.indexOf(`--${n}`);
    return i >= 0 && args[i + 1] ? args[i + 1] : d;
  };
  const srcDir = path.resolve(opt('src', DEFAULTS.srcDir));
  const outDir = path.resolve(opt('out', DEFAULTS.outDir));
  const port = parseInt(opt('port', '5173'), 10);
  const FIX_FROM = parseInt(opt('tu', '1'), 10);
  const FIX_TO = parseInt(opt('den', '148'), 10);
  const to = parseInt(opt('to', '1'), 10);

  const doc = readJson(path.join(outDir, `to${to}-ho.json`));
  const issues = readJson(path.join(outDir, `to${to}-loi.json`));

  // tra du lieu gon cho web (bo phan dat dai/ghi chu thua)
  const houses = doc.ho.map((h) => ({
    no: h.no,
    sttHo: h.sttHo,
    khung: h.khung,
    chuHo: h.chuHo,
    rowMin: h.rowMin,
    rowMax: h.rowMax,
    soNguoi: h.soNguoi,
    canhBao: h.canhBao,
    goiY: h.goiY,
    nguoi: h.thanhVien.map((m) => ({
      sttNguoi: m.sttNguoi,
      row: m.row,
      hoTen: m.hoTen,
      cccd: m.cccd,
      ngaySinh: m.ngaySinh,
      quanHe: m.vaiTro,
      dienThoai: m.dienThoai,
      nguonHo: m.nguonHo,
      canhBao: m.canhBao,
    })),
  }));
  const payload = {
    ap: doc.ap,
    to: doc.to,
    tongHo: doc.tongHo,
    tongNguoi: doc.tongNguoi,
    thongKe: doc.thongKe,
    quyTac: doc.quyTac,
    canhBao: issues.filter((i) => (i.muc || 'canh-bao') === 'canh-bao'),
    ghiChu: issues.filter((i) => i.muc === 'ghi-chu'),
    ho: houses,
  };

  // --- fix route: mo file Excel goc de xem / sua tung dong ---
  const loc = locate(srcDir, { to });
  let grid = null;
  let gridErr = null;
  if (loc.toFile) {
    try {
      grid = await Grid.open(loc.toFile.path, 0);
    } catch (e) {
      gridErr = e.message;
    }
  } else {
    gridErr = `khong tim thay file To 1 trong ${srcDir}`;
  }

  // tieu de cot: dung 1 lan cho ca hai nhanh API
  const cotTieuDe = () => grid.headerRows(2).find((h) => Object.keys(h).length) || {};

  const srv = http.createServer((req, res) => {
    // bat moi loi de server khong bi sap khi handler nem exception
    xuLy(req, res).catch((e) => {
      console.error('LOI xu ly yeu cau:', e.message);
      if (!res.headersSent) json(res, 500, { loi: e.message });
      else res.end();
    });
  });

  async function xuLy(req, res) {
    const url = new URL(req.url, `http://${req.headers.host}`);

    if (url.pathname === '/api/du-lieu') {
      const body = JSON.stringify(payload);
      res.writeHead(200, {
        'content-type': MIME['.json'],
        'content-length': Buffer.byteLength(body),
        'cache-control': 'no-store',
      });
      return res.end(body);
    }

    if (url.pathname === '/api/fix-route') {
      const from = parseInt(url.searchParams.get('from'), 10) || FIX_FROM;
      const to = parseInt(url.searchParams.get('to'), 10) || FIX_TO;
      if (!grid) return json(res, 200, { loi: gridErr });
      const s = grid.thongKe(from, to);
      return json(res, 200, {
        file: path.basename(grid.srcFile),
        sheet: grid.ws.name,
        from,
        to,
        daSua: grid.log.length,
        ...s,
        tieuDe: cotTieuDe(),
      });
    }

    if (url.pathname === '/api/fix-route/fix' && req.method === 'POST') {
      if (!grid) return json(res, 400, { loi: gridErr });
      const b = await docJson(req);
      const from = parseInt(b.from, 10) || FIX_FROM;
      const to = parseInt(b.to, 10) || FIX_TO;
      grid.fixRange(from, to);
      const s = grid.thongKe(from, to);
      return json(res, 200, { daSua: grid.log.length, ...s });
    }

    if (url.pathname === '/api/fix-route/set' && req.method === 'POST') {
      if (!grid) return json(res, 400, { loi: gridErr });
      const b = await docJson(req);
      const row = parseInt(b.row, 10);
      if (!row || !b.col) return json(res, 400, { loi: 'thieu row hoac col' });
      const r = grid.setCell(row, b.col, b.value);
      return json(res, 200, { doi: r, kiemTra: grid.check(row) });
    }

    if (url.pathname === '/api/fix-route/save' && req.method === 'POST') {
      if (!grid) return json(res, 400, { loi: gridErr });
      fs.mkdirSync(outDir, { recursive: true });
      const b = await docJson(req);
      const from = parseInt(b.from, 10) || FIX_FROM;
      const to = parseInt(b.to, 10) || FIX_TO;
      const xlsx = path.join(outDir, `Danh sách Tổ 1 - đã sửa dòng ${from}-${to}.xlsx`);
      const logFile = path.join(outDir, `fixroute-${from}-${to}.json`);
      await grid.save(xlsx);
      fs.writeFileSync(logFile, JSON.stringify(grid.log, null, 2), 'utf8');
      return json(res, 200, { file: xlsx, nhatKy: logFile, soDoi: grid.log.length, goc: grid.srcFile });
    }

    if (url.pathname === '/api/fix-route/doi-chieu') {
      if (!grid) return json(res, 200, { loi: gridErr });
      return json(res, 200, { nhatKy: grid.log });
    }

    const name = url.pathname === '/' ? 'index.html' : url.pathname.slice(1);
    const file = path.join(WEB, name);
    if (!file.startsWith(WEB) || !fs.existsSync(file)) {
      res.writeHead(404, { 'content-type': 'text/plain; charset=utf-8' });
      return res.end('Khong tim thay');
    }
    res.writeHead(200, { 'content-type': MIME[path.extname(file)] || 'text/plain; charset=utf-8' });
    res.end(fs.readFileSync(file));
  }

  srv.listen(port, () => {
    console.log(`Dang chay tai  http://localhost:${port}`);
    console.log(`  ${doc.tongHo} ho / ${doc.tongNguoi} nguoi`);
    if (grid) {
      const s = grid.thongKe(FIX_FROM, FIX_TO);
      console.log(`  fix route dong ${FIX_FROM}..${FIX_TO}: ${s.ok} ok, ${s.nghi} nen xem, ${s.sai} co loi`);
    } else {
      console.log(`  [canh bao] fix route khong dung: ${gridErr}`);
    }
    console.log('  Bam Ctrl+C de dung');
  });
}

main().catch((e) => {
  console.error('LOI:', e.message);
  process.exit(1);
});
