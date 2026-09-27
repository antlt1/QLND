'use strict';
// Kiem tra nhom ho: cot A so vs rac, va phan bo STT
const d = require('../output/to1-ho.json');
const isNum = (s) => /^\d+$/.test(String(s).trim());
const n = d.ho.filter((h) => isNum(h.cotARaw));
const r = d.ho.filter((h) => !isNum(h.cotARaw));

console.log(`Tong ${d.ho.length} nhom: ${n.length} nhom co STT, ${r.length} nhom rac\n`);

console.log('--- nhom so (cot A = so) ---');
for (const h of n) {
  const bad = h.canhBao && h.canhBao.length ? ' !!' + h.canhBao.join(',') : '';
  console.log(
    `  A=${String(h.cotARaw).padStart(4)}  soNguoi=${String(h.soNguoi).padStart(3)}` +
      `  rows=${h.rowMin}-${h.rowMax}  ${h.chuHo}${bad}`
  );
}

console.log('\n--- nhom rac (cot A khong phai so) ---');
for (const h of r) {
  console.log(`  A="${h.cotARaw}"  soNguoi=${h.soNguoi}  rows=${h.rowMin}-${h.rowMax}  ${h.chuHo}`);
  console.log(`      goiY: ${h.goiY || '(khong co)'}`);
  console.log(`      cb: ${(h.canhBao || []).join(', ') || '(khong co)'}`);
  for (const m of h.thanhVien) console.log(`        - ${m.sttNguoi}. ${m.hoTen}  [${m.nguonHo}] ${(m.canhBao || []).join(',')}`);
}

// STT bi thieu / trung
const ids = n.map((h) => Number(h.cotARaw)).sort((a, b) => a - b);
const seen = new Set();
const trung = [];
for (const i of ids) {
  if (seen.has(i)) trung.push(i);
  seen.add(i);
}
const thieu = [];
for (let i = 1; i <= Math.max(...ids); i++) if (!seen.has(i)) thieu.push(i);
console.log(`\nSTT 1..${Math.max(...ids)}: trung=[${trung.join(',')}] thieu=[${thieu.join(',')}]`);

const tong = d.ho.reduce((s, h) => s + h.soNguoi, 0);
console.log(`Tong nguoi trong cac nhom: ${tong} (khai bao ${d.tongNguoi})`);
