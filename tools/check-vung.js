'use strict';
// Kiem tra rieng vung bi chon dong 149-325: ai khoi phuc, ai giu nguyen
const d = require('../output/to1-ho.json');
const byRow = new Map();
for (const h of d.ho) for (const m of h.thanhVien) byRow.set(m.row, { h, m });

const R = { from: 149, to: 325 };
const rows = [];
for (let r = R.from; r <= R.to; r++) if (byRow.has(r)) rows.push(r);
console.log(`Vung ${R.from}-${R.to}: ${rows.length} nguoi\n`);

const ho = new Map();
for (const r of rows) {
  const { h, m } = byRow.get(r);
  ho.set(h.no, (ho.get(h.no) || 0) + 1);
}
console.log('Phan bo theo ho sau khi khoi phuc:');
for (const [no, n] of [...ho].sort((a, b) => b[1] - a[1])) {
  const h = d.ho.find((x) => x.no === no);
  console.log(`  ho ${String(h.sttHo || '?').padStart(3)} (no ${no}) : ${n} nguoi   ${h.chuHo}`);
}

console.log('\n4 nguoi KHONG khop duoc (giu nguyen cot A):');
for (const r of [186, 190, 228, 255]) {
  const { m } = byRow.get(r);
  console.log(`  r${r}  ${m.hoTen}  cccd=${m.cccd}  -> ho ${byRow.get(r).h.sttHo || '?'} (${m.nguonHo})`);
}

console.log('\nr160 (ho tra 201 bi xung dot):');
{
  const { h, m } = byRow.get(160);
  console.log(`  ${m.hoTen}  cccd=${m.cccd}  hoTra=${m.hoTra ?? '?'}  -> ho ${h.sttHo}  [${m.nguonHo}]`);
  console.log(`  canh bao: ${(m.canhBao || []).join(', ') || '(khong co)'}`);
}
