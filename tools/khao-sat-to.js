'use strict';
/** Khao sat cau truc 1 file danh sach to: so dong, cot, vung nghi nham */
const X = require('../src/lib/xlsx');
const { locate } = require('../src/lib/sources');

const to = parseInt(process.argv[2] || '2', 10);
const loc = locate(undefined, { to });
if (!loc.toFile) {
  console.error('Khong tim thay to ' + to);
  process.exit(1);
}

(async () => {
  const a = await X.readSheet(loc.toFile.path, 0);
  console.log(`TO ${to}: ${loc.toFile.ten}`);
  console.log(`sheet "${a.sheet}", ${a.rows.length} dong co du lieu`);

  const show = [1, 2, 3, 4, 5, 6, 7];
  console.log('\n--- 7 dong dau ---');
  for (const r of a.rows.slice(0, 7)) {
    console.log(`  r${String(r.row).padStart(4)} | ` + show.map((c) => `${c}:${r.c[c] ?? ''}`).join(' | '));
  }

  // cot A: phan bo gia tri
  const cotA = a.rows.filter((r) => r.c[1] != null && String(r.c[1]).trim() !== '');
  const so = new Map();
  for (const r of cotA) {
    const v = String(r.c[1]).trim();
    so.set(v, (so.get(v) || 0) + 1);
  }
  const nhan = [...so.entries()].sort((x, y) => Number(x[0]) - Number(y[0]));
  console.log(`\n--- cot A: ${cotA.length} dong co gia tri ---`);
  console.log('  gia tri lap lai nhieu lan (>1):');
  for (const [v, n] of nhan) if (n > 1) console.log(`    A="${v}" x ${n}`);
  console.log('  khong phai so:', nhan.filter(([v]) => !/^\d+$/.test(v)).map(([v, n]) => `"${v}" x${n}`).join(', ') || '(khong co)');

  // vung co STT A = "." hoac rong -> nghi nham chon dòng
  const dot = a.rows.filter((r) => /^\s*[.x]\s*$/i.test(String(r.c[1] ?? '')));
  console.log(`\n--- dong co A = "." hoac "x": ${dot.length} ---`);
  if (dot.length) console.log(`  tu dong ${dot[0].row} den dong ${dot[dot.length - 1].row}`);

  // kiem tra vung chon dong: A khong lap trong mot doan dai
  const trong = a.rows.filter((r) => r.c[1] == null || String(r.c[1]).trim() === '');
  if (trong.length) console.log(`\n--- dong A rong: ${trong.length}, tu dong ${trong[0].row} den ${trong[trong.length - 1].row} ---`);

  // cot nao co du lieu, theo tung doan 50 dong
  console.log('\n--- so dong co du lieu / cot nao chuyen nhieu ---');
  const buoc = 50;
  for (let s = 1; s <= a.rows[a.rows.length - 1].row; s += buoc) {
    const nhom = a.rows.filter((r) => r.row >= s && r.row < s + buoc);
    if (!nhom.length) continue;
    const cnt = new Map();
    for (const r of nhom) for (const [k, v] of Object.entries(r.c)) if (v != null && String(v).trim() !== '') cnt.set(k, (cnt.get(k) || 0) + 1);
    const cols = [...cnt.entries()].sort((x, y) => y[1] - x[1]).slice(0, 5).map(([k, n]) => `${k}:${n}`).join(' ');
    console.log(`  r${s}-${s + buoc - 1}: ${nhom.length} dong | ${cols}`);
  }
})().catch((e) => {
  console.error('LOI:', e.message);
  process.exit(1);
});
