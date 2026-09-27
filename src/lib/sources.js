'use strict';
/** Tim cac file nguon trong source/fileexcel theo dac tinh, khong hard-code ten co dau */
const fs = require('fs');
const path = require('path');
const X = require('./xlsx');

const DEFAULTS = {
  srcDir: path.join(__dirname, '..', '..', 'source', 'fileexcel'),
  outDir: path.join(__dirname, '..', '..', 'output'),
};

/**
 * File tra cuu dung cho tung to.
 *  - to 1 va to 4: danh sach "Truong Binh"
 *  - to 2         : danh sach "Truong Binh A"
 * Cac to khong co quy tac san -> hoi nguoi dung chon.
 */
const TRA_THEO_TO = { 1: 'tra', 4: 'tra', 2: 'traA' };

/** Vung duoc phep khoi phuc STT ho. Chi to 1 da xac dinh; to khac de trong de KHONG doan. */
const RECOVER_THEO_TO = { 1: { from: 149, to: 325 } };

/** So dong tieu de o dau sheet (Tổ 1 xác định từ file, tổ khác dùng mặc định) */
const HEADER_THEO_TO = { 1: 2 };

/** Cac file .xlsx trong thu muc (bo qua file khoa ~$ cua Excel) */
function listXlsx(srcDir = DEFAULTS.srcDir) {
  if (!fs.existsSync(srcDir)) return [];
  return fs
    .readdirSync(srcDir, { withFileTypes: true })
    .filter((e) => e.isFile() && e.name.toLowerCase().endsWith('.xlsx') && !X.isJunkName(e.name))
    .map((e) => {
      const p = path.join(srcDir, e.name);
      const size = fs.statSync(p).size;
      return { ten: e.name, path: p, size, doc: X.deaccent(e.name), ole2: X.isOle2(p), xlsx: X.isRealXlsx(p) };
    });
}

/**
 * Danh sach cac to tim thay trong thu muc.
 * Nhan dang: ten chua "to <so>", va KHONG phai danh sach khac cung co "to" nhu
 * "DS nguoi chua tham gia BHYT To 1-2-3" (do la danh sach nguoi, khong phai ho).
 * Neu nhieu file cho cung mot to, uu tien: doc duoc > khong phai "- Copy" > dung lon nhat.
 * @returns [{ to, ten, path, size, doc, docDuoc, banSao, khac }]
 */
function listTos(srcDir = DEFAULTS.srcDir) {
  const theoTo = new Map();
  for (const f of listXlsx(srcDir)) {
    const m = f.doc.match(/to\s*(\d+)/);
    if (!m) continue;
    if (/tra|tat\s*ca|chua\s*tham\s*gia|khong\s*tham\s*gia/.test(f.doc)) continue;
    const to = parseInt(m[1], 10);
    const banSao = /copy|ban\s*sao/.test(f.doc);
    const item = {
      to,
      ten: f.ten,
      path: f.path,
      size: f.size,
      ole2: f.ole2,
      banSao,
      // thu doc duoc khong: phai la zip/xlsx, va khong phai OLE2
      docDuoc: f.xlsx && !f.ole2,
    };
    if (!theoTo.has(to)) theoTo.set(to, []);
    theoTo.get(to).push(item);
  }

  const xep = (x, y) =>
    (y.docDuoc ? 1 : 0) - (x.docDuoc ? 1 : 0) ||
    (x.banSao ? 1 : 0) - (y.banSao ? 1 : 0) ||
    y.size - x.size;

  const ra = [];
  for (const [to, list] of theoTo) {
    list.sort(xep);
    const chinh = list[0];
    ra.push({ ...chinh, khac: list.slice(1) });
  }
  ra.sort((a, b) => a.to - b.to);
  return ra;
}

/** Cac file danh sach tra cuu (ten chua "tra") */
function listTra(srcDir = DEFAULTS.srcDir) {
  const ra = listXlsx(srcDir).filter((f) => f.doc.includes('tra'));
  ra.sort((a, b) => a.size - b.size); // nho nhat truoc
  return ra;
}

/**
 * Tim file cho mot to.
 * @param srcDir
 * @param opts.to      so to (mac dinh 1)
 * @param opts.tra     khoa file tra ('tra' | 'traA') de ghi de quy tac
 * @returns { to, toFile, tra, traA, traKhoa, ap, recoverRange, headerRows, tos, tras }
 */
function locate(srcDir = DEFAULTS.srcDir, opts = {}) {
  const to = parseInt(opts.to, 10) || 1;
  const tos = listTos(srcDir);
  const tras = listTra(srcDir);
  const tra = tras.find((f) => !f.doc.includes(' a ')) || tras[0] || null; // khong co " A" -> ban goc
  const traA = tras.find((f) => f.doc.includes(' a ')) || (tras.length > 1 ? tras[tras.length - 1] : null);

  const traKhoa = opts.tra || TRA_THEO_TO[to] || 'tra';
  const traChon = traKhoa === 'traA' ? traA : tra;

  const toFile = tos.find((t) => t.to === to) || null;

  return {
    to,
    tos,
    tras,
    toFile,
    toDocDuoc: !!(toFile && toFile.docDuoc),
    tra,
    traA,
    traKhoa,
    traChon,
    traDocDuoc: !!(traChon && traChon.xlsx && !traChon.ole2),
    // ten ap lay tu file tra da chon: "truong binh" hoac "truong binh a"
    ap: traChon ? apOf(traChon.doc) : 'Trường Bình',
    recoverRange: RECOVER_THEO_TO[to] || { from: -1, to: -1 },
    coKhoiPhuc: !!RECOVER_THEO_TO[to],
    headerRows: HEADER_THEO_TO[to] || 2,
  };
}

/** "ho ap truong binh a su dung de tra" -> "Truong Binh A" */
function apOf(doc) {
  if (/truong\s*binh\s*a/.test(doc)) return 'Trường Bình A';
  if (/truong\s*binh/.test(doc)) return 'Trường Bình';
  return 'Trường Bình';
}

/** Bao cao nhung file .xlsx ma may doc khong duoc (thuong la .xls duoi duoi .xlsx) */
async function scanUnreadable(srcDir = DEFAULTS.srcDir) {
  const out = [];
  for (const f of listXlsx(srcDir)) {
    if (!f.xlsx) {
      out.push({
        file: f.ten,
        lyDo: f.ole2 ? 'File .xls cu (OLE2) duoi duoi .xlsx' : 'Khong phai zip/xlsx',
      });
      continue;
    }
    try {
      await X.listSheets(f.path);
    } catch (err) {
      out.push({ file: f.ten, lyDo: `Doc loi: ${err.message}` });
    }
  }
  return out;
}

module.exports = {
  locate,
  listTos,
  listTra,
  listXlsx,
  scanUnreadable,
  apOf,
  TRA_THEO_TO,
  RECOVER_THEO_TO,
  DEFAULTS,
};
