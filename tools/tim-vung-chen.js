'use strict';
/** Tim vung dong bi "chon dong" cua mot to: nhung nguoi sttHo == 0 trong file nguon */
const X = require('../src/lib/xlsx');
const to1 = require('../src/lib/to1');
const { locate } = require('../src/lib/sources');

const to = parseInt(process.argv[2] || '4', 10);
const loc = locate(undefined, { to });

(async () => {
  const a = await X.readSheet(loc.toFile.path, 0);
  const b = await X.readSheet(loc.traChon.path, 0);
  const d = to1.build({ to1Rows: a.rows, lookupRows: b.rows, ap: loc.ap, to, recoverRange: { from: -1, to: -1 } });

  const khong = d.people.filter((p) => p.sttHo === 0);
  console.log(`TO ${to}: ${d.people.length} nguoi, ${khong.length} nguoi sttHo = 0`);

  // gom theo doan lien tiec
  const rows = khong.map((p) => p.row).sort((x, y) => x - y);
  const doan = [];
  for (const r of rows) {
    const cuoi = doan[doan.length - 1];
    if (cuoi && r === cuoi[cuoi.length - 1] + 1) cuoi.push(r);
    else doan.push([r]);
  }
  console.log('\nDOAN DONG LIEN TIEP CHUA XAC DINH DUOC HO:');
  for (const dd of doan) console.log(`  dong ${dd[0]}-${dd[dd.length - 1]}  (${dd.length} nguoi)`);

  // cot A xung quanh vung do
  const all = a.rows.filter((r) => r.row >= 3);
  const gan = all.filter((r) => r.row >= rows[0] - 6 && r.row <= rows[rows.length - 1] + 4);
  console.log('\nCOT A / ten / quan he quanh vung do:');
  for (const r of gan) {
    const C = to1.COT_THEO_TO[to] || to1.COT_THEO_TO[1];
    console.log(`  r${String(r.row).padStart(4)} A="${r.c[1] ?? ''}" | ${X.txt(r.c[C.ten]) || '(trong)'} | qh="${X.txt(r.c[C.quanHe])}"`);
  }

  // so dong co CCCD -> co the khoi phuc bang CCCD
  const coCccd = khong.filter((p) => p.cccd).length;
  console.log(`\nTrong ${khong.length} nguoi chua xac dinh ho: ${coCccd} co CCCD (doi chieu duoc voi DS tra)`);
  console.log('=> neu bat khoi phuc vung nay, se gan lai STT ho bang CCCD.');
})().catch((e) => {
  console.error('LOI:', e.message);
  process.exit(1);
});
