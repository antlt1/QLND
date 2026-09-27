'use strict';
// Kiem tra bang do ho: dong dau ho nao trong To 1 co hoTra = X
const X = require('../src/lib/xlsx');
const to1 = require('../src/lib/to1');
const { locate } = require('../src/lib/sources');

(async () => {
  const f = locate();
  const a = await X.readSheet(f.to1, 0);
  const b = await X.readSheet(f.tra, 0);
  const { people } = to1.parseTo1(a.rows);
  const idx = to1.indexLookup(to1.parseLookup(b.rows));
  to1.matchLookup(people, idx);

  for (const target of [201, 113]) {
    console.log(`\n=== hoTra ${target}: dong dau ho To 1 tuong ung ===`);
    const heads = people.filter((p) => p.isHeadRow && p.hoTra === target);
    if (!heads.length) console.log('  (khong co dong dau ho nao)');
    for (const p of heads) {
      console.log(`  r${p.row}  A=${p.cotARaw}  ${p.hoTen}  cccd=${p.cccd}  matchBy=${p.matchBy}`);
    }
    console.log(`  => so ung vien: ${heads.length}`);
  }

  // Dem xung dot that bang cach nay
  const cand = new Map();
  for (const p of people) {
    if (!p.isHeadRow || !p.hoTra || !X.intOf(p.cotARaw)) continue;
    if (!cand.has(p.hoTra)) cand.set(p.hoTra, new Set());
    cand.get(p.hoTra).add(p.cotARaw);
  }
  const conf = [...cand].filter(([, s]) => s.size > 1);
  console.log(`\n=== hoTra bi xung dot that: ${conf.length} ===`);
  for (const [ht, s] of conf) console.log(`  hoTra ${ht} -> STT ho To 1: ${[...s].join(' / ')}`);
})();
