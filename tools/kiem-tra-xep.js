'use strict';
/** Kiem tra file da xep: moi ho phai la 1 khoi lien, 1 chu ho, cot A khop */
const path = require('path');
const E = require('exceljs');
const { DEFAULTS } = require('../src/lib/sources');

const FILE = process.argv[2] || path.join(DEFAULTS.outDir, 'Danh sách Tổ 1 - đã xếp.xlsx');
const s = (v) => (v == null ? '' : String(v).trim());

(async () => {
  const wb = new E.Workbook();
  await wb.xlsx.readFile(FILE);
  const ws = wb.worksheets[0];
  const tong = ws.actualRowCount;
  console.log(`File: ${FILE}`);
  console.log(`Sheet "${ws.name}" - ${tong} dong\n`);

  // quet tung dong: cot A (STT ho) / cot B (STT nguoi) / cot C (ten) / cot I (quan he)
  const loi = [];
  const khoi = new Map(); // sttHo -> {dau, het, nChuHo, soNguoi}
  let chuaXacDinh = null; // nhom dot A = "." xep cuoi file
  let cur = null;
  let sttNguoiTruoc = 0;

  for (let r = 3; r <= tong; r++) {
    const row = ws.getRow(r);
    const a = s(row.getCell(1).value);
    const b = s(row.getCell(2).value);
    const ten = s(row.getCell(3).value);
    const cccd = s(row.getCell(6).value);
    // quan he nam o cot 9 (khoi 1) hoac cot 8 (khoi 2 rut gon)
    const c8 = s(row.getCell(8).value);
    const c9 = s(row.getCell(9).value);
    const qh = /^(chủ hộ|thành viên)$/.test(c9) ? c9 : /^(chủ hộ|thành viên)$/.test(c8) ? c8 : '';

    if (a !== '') {
      // dong chu ho moi -> bat dau khoi moi
      if (cur) cur.het = r - 1;
      if (a === '.') {
        // nhom "chua xac dinh ho" xep cuoi file, gom vao mot cho (khong can chu ho)
        if (!chuaXacDinh) chuaXacDinh = { dau: r, het: r, soNguoi: 0 };
        chuaXacDinh.het = r;
        chuaXacDinh.soNguoi++;
        cur = null;
        const nb = Number(b);
        if (Number.isInteger(nb) && nb === sttNguoiTruoc + 1) sttNguoiTruoc = nb;
        else loi.push(`r${r}: STT người "${b}" không nối tiếp (${sttNguoiTruoc})`);
        continue;
      }
      if (khoi.has(a)) loi.push(`r${r}: STT hộ ${a} xuất hiện lần 2 (lần trước ở khoi dòng ${khoi.get(a).dau})`);
      cur = { dau: r, het: r, nChuHo: 0, soNguoi: 0, sttHo: a, ten, cccd };
      khoi.set(a, cur);
    }
    if (!cur) {
      loi.push(`r${r}: có người nhưng chưa vào hộ nào (cot A trống)`);
      continue;
    }
    if (!ten && !cccd) continue; // dong rong (vd dong chi con STT ho)
    cur.soNguoi++;
    if (qh === 'chủ hộ' || (a !== '' && !cur.nChuHo && cur.soNguoi === 1)) {
      cur.nChuHo++;
      if (qh !== 'chủ hộ') loi.push(`r${r}: chủ hộ ${ten} thiếu ghi "chủ hộ" (đang là "${qh || 'trống'}")`);
    } else if (qh !== 'thành viên') {
      loi.push(`r${r}: thành viên ${ten} có quan hệ "${qh || 'trống'}"`);
    }
    // STT nguoi phai chay lien tuc tu 1
    const nb = Number(b);
    if (!Number.isInteger(nb)) loi.push(`r${r}: STT người "${b}" không phải số`);
    else if (nb !== sttNguoiTruoc + 1) loi.push(`r${r}: STT người nhảy ${sttNguoiTruoc} -> ${nb}`);
    else sttNguoiTruoc = nb;
  }
  if (cur) cur.het = tong;

  let nguoi = 0;
  for (const k of khoi.values()) nguoi += k.soNguoi;
  // dong chi con STT ho (khong ten, khong CCCD) khong phai ho that
  const rong = [...khoi.values()].filter((k) => k.soNguoi === 0);
  const hoThat = [...khoi.values()].filter((k) => k.soNguoi > 0);
  const hoNhieuChuHo = hoThat.filter((k) => k.nChuHo !== 1);

  console.log('=== KET QUA ===');
  console.log(`  so ho         : ${hoThat.length} (cong ${rong.length} dong chi con STT ho)`);
  console.log(`  so nguoi      : ${nguoi}${chuaXacDinh ? ` + ${chuaXacDinh.soNguoi} chua xac dinh ho` : ''}`);
  console.log(`  STT nguoi cuoi: ${sttNguoiTruoc}`);
  console.log(`  ho khong co dung 1 chu ho: ${hoNhieuChuHo.length}${hoNhieuChuHo.length ? ' -> ' + hoNhieuChuHo.map((k) => `hộ ${k.sttHo} (${k.nChuHo})`).join(', ') : ''}`);
  console.log(`  loi           : ${loi.length}`);
  loi.slice(0, 25).forEach((l) => console.log('    - ' + l));
  if (loi.length > 25) console.log(`    ... va ${loi.length - 25} loi nua`);

  console.log('\n=== VAI HO DA YEU (28-78) ===');
  [28, 29, 36, 40, 55, 78].forEach((n) => {
    const k = khoi.get(String(n));
    if (!k) return console.log(`  ho ${n}: KHONG CO`);
    const ds = [];
    for (let r = k.dau; r <= k.het; r++) {
      const row = ws.getRow(r);
      ds.push(`      ${r === k.dau ? 'CHU HO ' : '         '}${s(row.getCell(3).value)} [${s(row.getCell(6).value)}]`);
    }
    console.log(`  ho ${n}: ${k.soNguoi} nguoi, dong ${k.dau}-${k.het}`);
    ds.forEach((x) => console.log(x));
  });

  console.log(loi.length === 0 ? '\nPASS' : '\nCO LOI');
})();
