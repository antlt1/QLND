'use strict';
// In cot A tho cua tung dong de xac nhan nhom "." (dong 50-85)
const X = require('../src/lib/xlsx');
const { locate } = require('../src/lib/sources');

(async () => {
  const f = locate();
  const a = await X.readSheet(f.to1, 0);
  console.log('dong | colA (raw)      | quanHe      | hoTen');
  for (const r of a.rows) {
    if (r.row < 50 || r.row > 85) continue;
    const c1 = r.c[1];
    console.log(
      `  ${String(r.row).padStart(3)} | ${JSON.stringify(c1 === undefined ? null : c1).padEnd(16)}` +
        ` | ${JSON.stringify(X.txt(r.c[9])).padEnd(12)} | ${X.txt(r.c[3])}`
    );
  }
})();
