'use strict';
// Xem raw cac dong dau ho trong nguon de hieu STT 36 / 105 bi xu ly sai o dau
const X = require('../src/lib/xlsx');
const { locate } = require('../src/lib/sources');

(async () => {
  const f = locate();
  const a = await X.readSheet(f.to1, 0);
  const byRow = new Map(a.rows.map((r) => [r.row, r]));
  const { people } = require('../src/lib/to1').parseTo1(a.rows);

  console.log('=== Tat ca dong dau ho (col A khong rong) trong khoang 60-200 ===');
  for (const p of people.filter((p) => p.isHeadRow && p.row >= 60 && p.row <= 200)) {
    console.log(
      `  r${String(p.row).padStart(3)}  A=${JSON.stringify(p.cotARaw).padEnd(8)} sttHo=${String(p.sttHo).padStart(3)}` +
        `  qh=${JSON.stringify(p.quanHe).padEnd(12)} ten=${p.hoTen}`
    );
  }

  console.log('\n=== Cac dong co col A rong nhung la "dau ho" (quanHe=chu ho) ===');
  for (const p of people.filter((p) => p.isHeadRow && p.cotARaw === '')) {
    const r = byRow.get(p.row);
    console.log(`  r${p.row} sttHo=${p.sttHo} ten=${p.hoTen}`);
    console.log('       cot 1..10:', Array.from({ length: 10 }, (_, i) => `${i + 1}=${JSON.stringify(X.txt(r.c[i + 1]))}`).join(' '));
  }

  console.log('\n=== Moi xuat hien gia tri 105 trong col A ===');
  for (const p of people.filter((p) => p.cotARaw === '105')) {
    console.log(`  r${p.row}  ten=${p.hoTen}  quanHe=${p.quanHe}  cccd=${p.cccd}`);
  }

  console.log('\n=== Moi xuat hien 36 trong col A ===');
  for (const p of people.filter((p) => p.cotARaw === '36')) {
    console.log(`  r${p.row}  ten=${p.hoTen}  quanHe=${p.quanHe}  cccd=${p.cccd}`);
  }
})();
