'use strict';
/**
 * Lop doc Excel dung chung.
 * - ExcelJS xu ly .xlsx binh thuong
 * - File .xls / .xlsx "hoi" (chu ky OLE2 D0CF11E0) thi phai dung so lay mau tu zip
 */
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const CHU_HO = 'chủ hộ';
const THANH_VIEN = 'thành viên';
const KHONG_CO = 'không có';
const RELATIONSHIPS = [CHU_HO, THANH_VIEN];

/**
 * Chuan hoa 1 ExcelJS cell thanh gia tri tho.
 * QUAN TRONG: Date phai duoc giu nguyen la Date - dung dung cell.text (no se
 * format theo numFmt "mm-dd-yy" va lam sai moi ngay).
 */
function s(cell) {
  if (cell === null || cell === undefined) return '';
  const v = cell instanceof Object && 'value' in cell ? cell.value : cell;
  return raw(v);
}

/** Gia tri tho -> chuoi da trim. Date giu nguyen (de isoDoc xu ly) */
function raw(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v; // giu nguyen Date
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map((r) => r.text).join('').trim();
    if (v.text !== undefined) return String(v.text).trim();
    if (v.result !== undefined && v.result !== null) return raw(v.result);
    if (v.error !== undefined) return '';
    if (v.hyperlink !== undefined) return String(v.text || v.hyperlink).trim();
    return '';
  }
  return String(v).trim();
}

/** Chuan hoa thanh chuoi (ve dien dat thuong cho moi noi goi string) */
function str(v) {
  const r = raw(v);
  return r instanceof Date ? r : r;
}

/** Chuoi so nguyen, hoac null */
function intOf(cell) {
  const t = raw(cell instanceof Object && 'value' in cell ? cell.value : cell);
  if (t instanceof Date || !/^-?\d+$/.test(t)) return null;
  return parseInt(t, 10);
}

/** Chuoi 12 so CCCD hop le, hoac '' */
function cccdOf(cell) {
  const t = raw(cell instanceof Object && 'value' in cell ? cell.value : cell);
  if (t instanceof Date) return '';
  const d = String(t).replace(/\s+/g, '');
  return /^\d{12}$/.test(d) ? d : '';
}

/** Ngay thang -> 'yyyy-MM-dd'. Nhan Date, chuoi dd/MM/yyyy, hoac chuoi ISO. */
function isoDate(cell) {
  const d = toDate(cell);
  return d ? d.toISOString().slice(0, 10) : '';
}

/** Cell -> Date | null */
function toDate(cell) {
  const v = cell instanceof Object && 'value' in cell ? cell.value : cell;
  if (v instanceof Date && !isNaN(v)) return v;
  const t = raw(v);
  if (t instanceof Date) return isNaN(t) ? null : t;
  if (typeof t !== 'string' || !t) return null;
  let m = t.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
  if (m) return new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
  m = t.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2})$/);
  if (m) {
    let y = +m[3];
    y += y < 30 ? 2000 : 1900;
    return new Date(Date.UTC(y, +m[2] - 1, +m[1]));
  }
  m = t.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
  return null;
}

/** Cell -> chuoi (Date -> 'yyyy-MM-dd', con lai -> chuoi da trim) */
function txt(cell) {
  const v = cell instanceof Object && 'value' in cell ? cell.value : cell;
  const d = toDate(v);
  if (d) return d.toISOString().slice(0, 10);
  const r = raw(v);
  return r instanceof Date ? r.toISOString().slice(0, 10) : r;
}

/** Chuan hoa ten de so sanh: bo khoang trang, lowercase, bo dau tieng Viet */
function normName(name) {
  return txt(name)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Doc mot sheet thanh mang cac hang.
 * @returns {Promise<{rows: Array<object>, sheet: string}>}
 *   moi hang: { row: <so dong Excel>, c: { <cot 1-based>: <chuoi> } }
 */
async function readSheet(file, sheetName) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(file);
  const ws = typeof sheetName === 'number' ? wb.worksheets[sheetName] : wb.getWorksheet(sheetName);
  if (!ws) {
    throw new Error(
      `Khong tim thay sheet "${sheetName}" trong ${path.basename(file)}. ` +
        `Cac sheet co: ${wb.worksheets.map((w) => w.name).join(', ')}`
    );
  }
  const rows = [];
  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    const c = {};
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      // luon doc cell.value (giu nguyen Date), khong dung cell.text
      const v = cell.value;
      if (v === null || v === undefined || v === '') return;
      const r = raw(v);
      if (r === '' || r === null) return;
      c[colNumber] = r; // Date giu nguyen
    });
    if (Object.keys(c).length > 0) rows.push({ row: rowNumber, c });
  });
  return { rows, sheet: ws.name };
}

/** Ten cac sheet trong file (de bao loi ro rang) */
async function listSheets(file) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(file);
  return wb.worksheets.map((w) => w.name);
}

/** Kiem tra file co phai .xlsx that khong */
function isRealXlsx(file) {
  const fd = fs.openSync(file, 'r');
  try {
    const buf = Buffer.alloc(4);
    fs.readSync(fd, buf, 0, 4, 0);
    return buf.toString('hex') === '504b0304';
  } finally {
    fs.closeSync(fd);
  }
}

/** Kiem tra file co phai .xls cu (OLE2) khong */
function isOle2(file) {
  const fd = fs.openSync(file, 'r');
  try {
    const buf = Buffer.alloc(8);
    fs.readSync(fd, buf, 0, 8, 0);
    return buf.toString('hex').startsWith('d0cf11e0');
  } finally {
    fs.closeSync(fd);
  }
}

/** Tim file theo tu khoa trong thu muc */
/**
 * File .xlsx that la bi bo qua: file khoa cua Excel (~$...), file tam, file an.
 * Excel dang mo file se tao "~$Ten.xlsx" rat nho - doc vao se loi
 * "Can't find end of central directory".
 */
function isJunkName(name) {
  const n = String(name);
  return n.startsWith('~$') || n.startsWith('.~lock.') || n.startsWith('~') || n.startsWith('.');
}

function findFile(dir, predicate) {
  if (!fs.existsSync(dir)) return null;
  const f = fs
    .readdirSync(dir, { withFileTypes: true })
    .filter((e) => e.isFile() && e.name.toLowerCase().endsWith('.xlsx') && !isJunkName(e.name))
    .map((e) => path.join(dir, e.name))
    .find((p) => predicate(p, path.basename(p), fs.statSync(p).size));
  return f || null;
}

/** Tien to goc de tim so dong (khong phu thuoc dau tieng Viet) */
function deaccent(str) {
  return str
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd');
}

/**
 * Ghi workbook ra file, bao loi ro rang khi file dang bi mo (thuong la Excel).
 * @param wb     workbook cua exceljs
 * @param dest   duong dan file dich
 * @param nhan   ten hien thi trong loi (vd "file da xep")
 */
async function writeXlsx(wb, dest, nhan = 'file') {
  try {
    await wb.xlsx.writeFile(dest);
  } catch (err) {
    const code = err && err.code;
    if (code === 'EBUSY' || code === 'EPERM' || code === 'EACCES') {
      throw new Error(
        `Khong ghi duoc ${nhan}: "${path.basename(dest)}" dang duoc mo (Excel hay chuong trinh khac).\n` +
          `       Dong file do lai roi chay lai. File cu cua ban khong bi anh huong.`
      );
    }
    throw err;
  }
}

module.exports = {
  CHU_HO,  THANH_VIEN,
  KHONG_CO,
  RELATIONSHIPS,
  s,
  txt,
  raw,
  toDate,
  intOf,
  cccdOf,
  isoDate,
  normName,
  readSheet,
  listSheets,
  isRealXlsx,
  isOle2,
  findFile,
  isJunkName,
  writeXlsx,
  deaccent,
};
