'use strict';
// In nguyen cot 1..10 cua dong 1..148 de thay dung cau truc can fix
const X = require('../src/lib/xlsx');
const { locate } = require('../src/lib/sources');

const H = (s, n) => String(s ?? '').slice(0, n).padEnd(n);

(async () => {
  const f = locate();
  const a = await X.readSheet(f.to1, 0);
  const byRow = new Map(a.rows.map((r) => [r.row, r]));

  console.log('dong |  A  | STTng | Ho ten                 | E ngay sinh | F cccd        | G ngay cap   | H tuoi | I quan he      | J ten bo');
  console.log('-'.repeat(120));
  for (let row = 1; row <= 148; row++) {
    const r = byRow.get(row);
    const g = (i) => (r ? X.txt(r.c[i]) : '');
    console.log(
      `${H(row, 4)}| ${H(JSON.stringify(r ? X.txt(r.c[1]) : null), 3)}| ${H(g(2), 6)}| ${H(g(3), 22)}| ${H(g(5), 12)}| ${H(g(6), 13)}| ${H(g(7), 12)}| ${H(g(8), 6)}| ${H(g(9), 14)}| ${H(g(10), 10)}`
    );
  }
})();
