'use strict';
/**
 * XEP LAI file To 1: moi ho gom dung mot khoi - chu ho roi den thanh vien.
 *
 * File goc co van de:
 *   dong 3..97    ho 1..27  - chu ho + thanh vien ngay duoi     (da dung)
 *   dong 98..148  ho 28..78 - CHI CO chu ho, cot B trong         (sai)
 *   dong 149..325           - thanh vien cua ho 28..78 vang rai rac
 *
 * Cach lam: dung CCCD doi chieu voi file "Ho ap Truong Binh su dung de tra"
 * de biet moi nguoi thuoc ho nao, roi sap xep lai theo thu tu ho.
 * Khong xoa dong nao, khong bo nguoi nao; chi thay doi THU TU va dien cot A/B.
 */
const path = require('path');
const ExcelJS = require('exceljs');
const X = require('./xlsx');
const to1 = require('./to1');
const { freezeFormulas } = require('./grid');
const { locate, DEFAULTS } = require('./sources');

const COT_STT_HO = 1; // cot A
const COT_STT_NGUOI = 2; // cot B
const COT_TEN = 3; // cot C

/** Copy nguyen mot dong (gia tri + dinh dang) tu dong nguon sang dong dich */
function copyRow(wsNguon, rNguon, wsDich, rDich) {
  const src = wsNguon.getRow(rNguon);
  const dst = wsDich.getRow(rDich);
  for (let c = 1; c <= Math.max(wsNguon.columnCount, 26); c++) {
    const s = src.getCell(c);
    const t = dst.getCell(c);
    if (s.value !== null && s.value !== undefined) {
      t.value = s.value;
      if (s.numFmt) t.numFmt = s.numFmt;
    }
    if (s.style && s.style !== t.style) t.style = { ...s.style };
  }
  dst.height = src.height;
  return dst;
}

/** Che do rong cot cua dong nguon sang dong dich */
function copyWidths(wsNguon, wsDich) {
  for (let c = 1; c <= Math.max(wsNguon.columnCount, 26); c++) {
    const w = wsNguon.getColumn(c).width;
    if (w) wsDich.getColumn(c).width = w;
  }
}

/**
 * @param opts.srcFile  file danh sach To N
 * @param opts.traFile  file tra cuu
 * @param opts.ap, opts.to
 * @param opts.recoverRange  vung duoc phep khoi phuc STT ho. {from:-1,to:-1} = tat.
 * @param opts.headerRows    so dong tieu de o dau sheet (mac dinh 2)
 * @param opts.outFile       file xlsx dau ra
 */
async function xep({ srcFile, traFile, ap, to, recoverRange, headerRows: headerRowsOpt, outFile }) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(srcFile);
  const ws = wb.worksheets[0];
  if (!ws) throw new Error('File khong co sheet nao');

  // ExcelJS khong ghi duoc workbook khi chep o cong thuc, nen bien truoc thanh
  // gia tri tinh: cot B (tuoi) va cot H (nam sinh -> tuoi) se ghi nhu so thuan
  const dongCongThuc = freezeFormulas(ws);

  const a = await X.readSheet(srcFile, 0);
  const b = traFile ? await X.readSheet(traFile, 0) : { rows: [] };
  const d = to1.build({ to1Rows: a.rows, lookupRows: b.rows, ap, to, recoverRange });

  // a.rows la danh sach loc, khong index theo so dong -> dung bang tra cuu
  const theoDong = new Map(a.rows.map((r) => [r.row, r]));
  const gocDong = (r) => theoDong.get(r) || { row: r, c: {} };

  // ExcelJS bao cao actualRowCount = 601 nhung van co du lieu o dong 604
  // (chi co cot A = "141"), nen phai quet theo rowCount roi bo dong trong.
  const tongDongGoc = ws.rowCount;
  const headerRows = headerRowsOpt || 2; // 2 dong tieu de o dau sheet

  // --- thu tu xep: theo thu tu ho da chot, nhung day nhom CHUA XAC DINH HO
  // (key M...) xuong cuoi file de khong cat ngang danh sach ho 79..140 ---
  const thuTu = [];
  const rowDaDung = new Set();
  d.households.forEach((h) => {
    const members = [...h.members].sort((x, y) => (x.row === y.row ? x.sttNguoi - y.sttNguoi : x.row - y.row));
    thuTu.push({ hh: h, members });
    members.forEach((m) => rowDaDung.add(m.row));
  });
  const chuaXacDinh = thuTu.filter((x) => /^M\d+$/.test(x.hh.key));
  thuTu.splice(0, thuTu.length, ...thuTu.filter((x) => !/^M\d+$/.test(x.hh.key)), ...chuaXacDinh);

  // dong co du lieu nhung khong gan duoc ho nao -> xep cuoi file
  const chuaXep = [];
  for (let r = headerRows + 1; r <= tongDongGoc; r++) {
    if (rowDaDung.has(r)) continue;
    // giu moi dong co bat ky o nao khac rong (ke ca dong chi co STT ho o cot A)
    const o = gocDong(r).c;
    const coDuLieu = Object.keys(o).some(
      (k) => o[k] !== null && o[k] !== undefined && String(o[k]).trim() !== ''
    );
    if (!coDuLieu) continue;
    chuaXep.push(r);
    rowDaDung.add(r);
  }

  // --- ghi sheet moi ---
  const out = new ExcelJS.Workbook();
  const wsOut = out.addWorksheet(ws.name, { views: [{ state: 'frozen', ySplit: headerRows }] });
  copyWidths(ws, wsOut);

  // 2 dong tieu de giu nguyen
  for (let r = 1; r <= headerRows; r++) copyRow(ws, r, wsOut, r);

  const nhatKy = [];
  let rDich = headerRows;
  let sttNguoi = 0;

  const ghi = (rowNguon, hh, vaiTro, nguonHo) => {
    rDich++;
    copyRow(ws, rowNguon, wsOut, rDich);
    sttNguoi++;
    const dong = wsOut.getRow(rDich);
    if (vaiTro === 'chuHo') {
      dong.getCell(COT_STT_HO).value = hh.sttHo || null;
    } else if (vaiTro === 'thanhVien') {
      dong.getCell(COT_STT_HO).value = null; // chi dong chu ho moi ghi STT ho
    } else if (vaiTro === 'chuaXacDinh') {
      dong.getCell(COT_STT_HO).value = '.'; // quy uoc cua chinh file goc
    }
    // dong chua gan ho: giu nguyen cot A cua file goc
    // cot B: STT nguoi chay tu 1
    dong.getCell(COT_STT_NGUOI).value = sttNguoi;
    nhatKy.push({
      sttNguoi,
      sttHo: hh.sttHo || 0,
      vaiTro,
      hoTen: X.txt(gocDong(rowNguon).c[COT_TEN]),
      cccd: X.cccdOf(gocDong(rowNguon).c[6]),
      dongCu: rowNguon,
      dongMoi: rDich,
      trangThai: rowNguon === rDich ? 'giu nguyen' : 'di chuyen',
      nguon: nguonHo || '',
    });
    return rDich;
  };

  for (const { hh, members } of thuTu) {
    const chuaXacDinh = /^M\d+$/.test(hh.key);
    members.forEach((m, i) =>
      ghi(m.row, hh, chuaXacDinh ? 'chuaXacDinh' : i === 0 ? 'chuHo' : 'thanhVien', m.nguonHo)
    );
  }
  const dongChuaXep = chuaXep.map((r) => ghi(r, { sttHo: 0 }, 'chuaXep', ''));
  const tongDongDaGhi = rDich;

  // --- nhat ky xep ---
  const wsLog = out.addWorksheet('Nhat ký xếp');
  wsLog.columns = [
    { header: 'STT người', key: 'sttNguoi', width: 10 },
    { header: 'STT hộ', key: 'sttHo', width: 9 },
    { header: 'Vai trò', key: 'vaiTro', width: 11 },
    { header: 'Họ và tên', key: 'hoTen', width: 26 },
    { header: 'CCCD', key: 'cccd', width: 14 },
    { header: 'Dòng cũ', key: 'dongCu', width: 9 },
    { header: 'Dòng mới', key: 'dongMoi', width: 10 },
    { header: 'Trạng thái', key: 'trangThai', width: 12 },
    { header: 'Nguồn', key: 'nguon', width: 22 },
  ];
  wsLog.addRows(nhatKy);
  wsLog.getRow(1).font = { bold: true };

  // --- tom tat ---
  const wsSum = out.addWorksheet('Tóm tắt');
  const giu = nhatKy.filter((x) => x.trangThai === 'giu nguyen').length;
  const chuyen = nhatKy.filter((x) => x.trangThai === 'di chuyen').length;
  const nguoiChuaXacDinh = nhatKy.filter((x) => x.vaiTro === 'chuaXacDinh').length;
  wsSum.columns = [
    { header: 'Mục', key: 'k', width: 34 },
    { header: 'Giá trị', key: 'v', width: 16 },
  ];
  wsSum.addRows([
    { k: 'File nguồn', v: path.basename(srcFile) },
    { k: 'Sheet', v: ws.name },
    { k: 'Số hộ', v: d.households.length },
    { k: 'Số người', v: d.people.length },
    { k: 'Dòng tiêu đề giữ nguyên', v: headerRows },
    { k: 'Tổng dòng file gốc', v: tongDongGoc },
    { k: 'Tổng dòng file mới', v: tongDongDaGhi },
    { k: 'Dòng giữ nguyên vị trí', v: giu },
    { k: 'Dòng di chuyển', v: chuyen },
    { k: 'Dòng chưa gán hộ (xếp cuối)', v: dongChuaXep.length },
    {
      k: 'Người chưa xác định được hộ (cot A = ".")',
      v: nguoiChuaXacDinh,
    },
  ]);
  wsSum.getRow(1).font = { bold: true };

  await X.writeXlsx(out, outFile, 'file da xep');
  return { outFile, nhatKy, tongDongGoc, tongDongDaGhi, giu, chuyen, chuaXep: dongChuaXep, tongHo: d.households.length, tongNguoi: d.people.length, dongCongThuc, nguoiChuaXacDinh };
}

module.exports = { xep, copyRow, copyWidths, COT_STT_HO, COT_STT_NGUOI, COT_TEN };

if (require.main === module) {
  // chay truc tiep: node src/lib/arrange.js --to 2
  require('../cli/xep');
}
