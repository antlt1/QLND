'use strict';
/** In tieu de + mau dong, tat ca cot, cho tung to de so sanh schema */
const X = require('../src/lib/xlsx');
const { locate } = require('../src/lib/sources');

const tos = process.argv.slice(2).length ? process.argv.slice(2).map(Number) : [1, 2, 4];

(async () => {
  for (const to of tos) {
    const loc = locate(undefined, { to });
    if (!loc.toFile) {
      console.log(`\n########## TO ${to}: KHONG CO FILE ##########`);
      continue;
    }
    const a = await X.readSheet(loc.toFile.path, 0);
    console.log(`\n########## TO ${to}: ${loc.toFile.ten} (${a.rows.length} dong) ##########`);

    const b = a.rows[0].c; // dong tieu de 1
    console.log('TIE DE 1: ' + Object.keys(b).sort((x, y) => x - y).map((k) => `[${k}]${b[k]}`).join(' '));

    // cot 8,9,10... co gi?
    const cnt = new Map();
    for (const r of a.rows) for (const [k, v] of Object.entries(r.c)) if (v != null && String(v).trim() !== '') cnt.set(k, (cnt.get(k) || 0) + 1);
    console.log('SO DONG CO GIA TRI THEO COT: ' + [...cnt.entries()].sort((x, y) => Number(x[0]) - Number(y[0])).map(([k, n]) => `${k}=${n}`).join(' '));

    for (const r of a.rows.slice(1, 4)) {
      const keys = Object.keys(r.c).sort((x, y) => x - y);
      console.log(`  r${r.row}: ` + keys.map((k) => `${k}=${r.c[k]}`).join(' | ').slice(0, 260));
    }
  }
})().catch((e) => {
  console.error('LOI:', e.message);
  process.exit(1);
});
