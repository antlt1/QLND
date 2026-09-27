'use strict';
/** Xem sheet trong file output da gom, quanh cac ho bat ky */
const E = require('exceljs');
const { DEFAULTS } = require('../src/lib/sources');
const path = require('path');

const FILE = process.argv[2] || path.join(DEFAULTS.outDir, 'Danh sách Tổ 1 - đã gom.xlsx');
const SHEET = process.argv[3] || 'Danh sách người';
const HO = (process.argv[4] || '27,28,29,30').split(',');
const SO = parseInt(process.argv[5] || '10', 10);

(async () => {
  const wb = new E.Workbook();
  await wb.xlsx.readFile(FILE);
  console.log(`File: ${FILE}`);
  console.log(`Sheet: ${wb.worksheets.map((w) => w.name).join(' | ')}\n`);
  const ws = wb.getWorksheet(SHEET);
  if (!ws) return console.log('khong co sheet nay');
  console.log(`--- ${SHEET} (${ws.rowCount} dong) ---`);
  for (let r = 1; r <= 3; r++) {
    console.log(`  tieu de r${r}: ${ws.getRow(r).values.slice(1, 11).map((x) => (x == null ? '' : String(x))).join(' | ')}`);
  }
  const idx = [];
  ws.eachRow((row, n) => {
    const v = row.values.slice(1, 11).map((x) => (x == null ? '' : String(x)));
    if (HO.includes((v[0] || '').trim())) idx.push(n);
  });
  console.log(`\ndong bat dau ho ${HO.join('/')}: ${idx.join(', ')}\n`);
  if (idx.length) {
    for (let r = idx[0] - 1; r <= idx[idx.length - 1] + SO; r++) {
      if (r < 1 || r > ws.rowCount) continue;
      const v = ws.getRow(r).values.slice(1, 11).map((x) => (x == null ? '' : String(x).slice(0, 24)));
      console.log(`  r${String(r).padStart(4)}: ${v.join(' | ')}`);
    }
  }
})();
