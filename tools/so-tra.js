'use strict';
/**
 * SO SANH ho To 1 voi ho tuong ung trong file tra.
 *
 * Voi moi ho To 1, lay CCCD chu ho -> tim trong file tra -> lay tat ca nguoi
 * cung hoTra, roi doi chieu CCCD. Ket qua:
 *   - THIEU  : co trong file tra nhung khong co trong To 1
 *   - THUA   : co trong To 1 nhung khong co trong file tra
 *   - KHOIP  : khop
 */
const X = require('../src/lib/xlsx');
const to1 = require('../src/lib/to1');
const { locate } = require('../src/lib/sources');

(async () => {
  const f = locate();
  const a = await X.readSheet(f.to1, 0);
  const b = await X.readSheet(f.tra, 0);
  const d = to1.build({ to1Rows: a.rows, lookupRows: b.rows, ap: 'Trường Bình', to: 1 });

  // nhom nguoi trong file tra theo hoTra
  const traNhom = new Map();
  for (const p of d.people) {
    if (!p.hoTra) continue;
    if (!traNhom.has(p.hoTra)) traNhom.set(p.hoTra, []);
    traNhom.get(p.hoTra).push(p);
  }
  // nhom nguoi trong To 1 theo ho (dung CCCD de so)
  const dsNguoi = d.people;
  const theoCccd = new Map(dsNguoi.filter((p) => p.cccd).map((p) => [p.cccd, p]));

  let tongThieu = 0,
    tongThua = 0,
    tongKhop = 0,
    tongKhongCo = 0;
  const dsThieu = [];
  const dsThua = [];

  for (const h of d.households) {
    const chu = h.members.find((m) => m.vaiTro === 'chuHo') || h.members[0];
    if (!chu) continue;
    const nguoiTra = chu.hoTra ? traNhom.get(chu.hoTra) : null;
    if (!nguoiTra) {
      tongKhongCo++;
      continue;
    }
    const cccdHo = new Set(h.members.map((m) => m.cccd).filter(Boolean));

    // thieu: co trong tra, khong co trong To 1
    for (const t of nguoiTra) {
      if (!t.cccd) continue;
      if (!cccdHo.has(t.cccd)) {
        tongThieu++;
        dsThieu.push({ sttHo: h.sttHo, chuHo: chu.hoTen, hoTra: chu.hoTra, nguoi: t.hoTen, cccd: t.cccd, quanHe: t.quanHe, sttTra: t.stt });
      }
    }
    // thua: co trong To 1, khong co trong tra
    const cccdTra = new Set(nguoiTra.map((t) => t.cccd).filter(Boolean));
    for (const m of h.members) {
      if (!m.cccd) continue;
      if (!cccdTra.has(m.cccd)) {
        tongThua++;
        dsThua.push({ sttHo: h.sttHo, hoTra: chu.hoTra, nguoi: m.hoTen, cccd: m.cccd, row: m.row });
      }
    }
    tongKhop++;
  }

  console.log(`Tong ho To 1: ${d.households.length}`);
  console.log(`  co ho tra tuong ung : ${tongKhop}`);
  console.log(`  khong tim thay ho tra: ${tongKhongCo}`);
  console.log(`  nguoi THIEU (co trong tra, thieu o To 1): ${tongThieu}`);
  console.log(`  nguoi THUA  (co o To 1, khong co trong tra): ${tongThua}\n`);

  if (dsThieu.length) {
    console.log('=== THIEU ===');
    for (const x of dsThieu) console.log(`  ho ${String(x.sttHo).padStart(3)} (tra ${x.hoTra})  ${x.cccd}  ${x.quanHe.padEnd(10)} ${x.nguoi}`);
  }
  if (dsThua.length) {
    console.log('\n=== THUA ===');
    for (const x of dsThua) console.log(`  ho ${String(x.sttHo).padStart(3)} (tra ${x.hoTra})  row ${x.row}  ${x.cccd}  ${x.nguoi}`);
  }
})();
