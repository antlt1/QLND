'use strict';
/** Tim mot nguoi trong file To 1 va file tra, in ra moi dong cham toi */
const X = require('../src/lib/xlsx');
const { locate } = require('../src/lib/sources');

const CCCD = process.argv[2] || '093090011197';
const TEN = process.argv[3] || null;

(async () => {
  const f = locate();
  const a = await X.readSheet(f.to1, 0);
  const b = f.tra ? await X.readSheet(f.tra, 0) : { rows: [] };

  console.log(`CCCD can tim: ${CCCD}${TEN ? ' / ' + TEN : ''}\n`);

  console.log('=== TRONG FILE TO 1 ===');
  const hit = [];
  a.rows.forEach((r, i) => {
    const s = JSON.stringify(r);
    if (s.includes(CCCD) || (TEN && s.includes(TEN))) hit.push(i + 1);
  });
  console.log(`  dong cham: ${hit.join(', ') || 'khong co'}`);
  for (const i of hit) console.log(`  r${i}: ${JSON.stringify(a.rows[i - 1])}`);

  console.log('\n=== TRONG FILE TRA ===');
  const hit2 = [];
  b.rows.forEach((r, i) => {
    const s = JSON.stringify(r);
    if (s.includes(CCCD) || (TEN && s.includes(TEN))) hit2.push(i + 1);
  });
  console.log(`  dong cham: ${hit2.join(', ') || 'khong co'}`);
  for (const i of hit2.slice(0, 10)) console.log(`  r${i}: ${JSON.stringify(b.rows[i - 1])}`);

  // xem context toi file To 1
  if (hit.length) {
    const c = hit[0];
    console.log(`\n=== FILE TO 1: dong ${c - 4} den ${c + 4} ===`);
    for (let r = c - 4; r <= c + 4; r++) {
      if (r < 1) continue;
      console.log(`  r${String(r).padStart(4)}: ${JSON.stringify(a.rows[r - 1])}`);
    }
  }
})();
