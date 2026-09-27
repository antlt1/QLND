'use strict';
/** In mot doan dong cua file tra, de hieu cach phan biet ho */
const X = require('../src/lib/xlsx');
const { locate } = require('../src/lib/sources');

const tu = +(process.argv[2] || 1560);
const den = +(process.argv[3] || 1600);

(async () => {
  const f = locate();
  const b = await X.readSheet(f.tra, 0);
  console.log(`File tra: ${f.tra}`);
  console.log(`Tong dong doc duoc: ${b.rows.length}\n`);
  console.log('tieu de:', JSON.stringify(b.rows[0]));
  console.log('');
  for (let r = tu; r <= den; r++) {
    const o = b.rows[r - 1];
    if (!o) continue;
    console.log(`  r${String(r).padStart(4)}: ${JSON.stringify(o.c)}`);
  }
})();
