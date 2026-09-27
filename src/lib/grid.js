'use strict';
/**
 * Thu vien xem va sua Excel theo tung dong ("fix route").
 *
 * Nguyen tac: KHONG xoa dong, KHONG doan. Chi chuan hoa nhung gi tri co the
 * chuan hoa chac chan (cat khoang trang, sua ngay thang, bo dau cach trong CCCD)
 * va ghi lai toan bo thay doi de doi chieu.
 *
 * Cach dung:
 *   const g = await Grid.open(file, 0);
 *   g.route(1, 148).forEach(step => ...);
 *   g.fix(5);
 *   await g.save(dest);
 */
const ExcelJS = require('exceljs');
const X = require('./xlsx');

const A = 'A';
/** Chuoi cot -> so: A=1, AA=27, ... */
function colToNum(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}
/** So -> chuoi cot: 1=A, 27=AA */
function numToCol(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Gia tri cell -> chuoi de hien thi (Date -> dd/MM/yyyy) */
function show(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) {
    const p = (x) => String(x).padStart(2, '0');
    return `${p(v.getUTCDate())}/${p(v.getUTCMonth() + 1)}/${v.getUTCFullYear()}`;
  }
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map((r) => r.text).join('');
    if (v.text !== undefined && typeof v.text !== 'object') return String(v.text);
    if (v.result !== undefined && v.result !== null) return show(v.result);
    if (v.error !== undefined) return `#${v.error}`;
    // cong thuc chung (sharedFormula) chua co ket qua -> Excel can tinh lai
    if (v.sharedFormula || v.formula) return '';
    return '';
  }
  return String(v);
}

/** Co phai cong thuc chua co ket qua khong */
function isUnresolvedFormula(v) {
  if (!v || typeof v !== 'object') return false;
  if (v instanceof Date) return false;
  return (v.sharedFormula !== undefined || v.formula !== undefined) && v.result === undefined;
}

/**
 * So thu tu ngay cua Excel (khong co dinh dang ngay).
 * 1..100000 ~ nam 1900..2173. 46013 = 22/12/2025.
 */
function serialToDate(n) {
  if (typeof n !== 'number' || !Number.isFinite(n)) return null;
  if (n < 1 || n > 100000) return null;
  return new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000);
}

/** Chuoi -> Date (dd/MM/yyyy, dd/MM/yy) hoac null */
function toDate(s) {
  const t = String(s).trim();
  // dd/MM/yyyy
  let m = /^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})$/.exec(t);
  if (m) return new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
  // dd/MM/yy  -> 79 = 1979, 08 = 2008
  m = /^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2})$/.exec(t);
  if (m) {
    const y = +m[3] < 30 ? 2000 + +m[3] : 1900 + +m[3];
    return new Date(Date.UTC(y, +m[2] - 1, +m[1]));
  }
  return null;
}

const REL = ['chủ hộ', 'thành viên'];

/**
 * Bien moi o cong thuc cua worksheet thanh gia tri tinh.
 *  - co ket qua da luu -> giu ket qua do
 *  - khong co ket qua  -> de trong va ghi nho de canh bao
 * Ly do: de sua/chép duoc o nay ma khong lam hong shared formula, va de Excel
 * khong tu tinh lai ghi de len gia tri da sua tay.
 * @returns [{o, ketQua}] danh sach o da doi
 */
function freezeFormulas(ws) {
  const ds = [];
  for (let r = 1; r <= ws.rowCount; r++) {
    for (let c = 1; c <= ws.columnCount; c++) {
      const cell = ws.getRow(r).getCell(c);
      const v = cell.value;
      if (!v || typeof v !== 'object' || v instanceof Date) continue;
      if (v.formula === undefined && v.sharedFormula === undefined) continue;
      if (v.result !== undefined && v.result !== null) {
        ds.push({ o: cell.address, ketQua: show(v.result) });
        cell.value = v.result;
      } else {
        ds.push({ o: cell.address, ketQua: null });
        cell.value = null;
      }
    }
  }
  return ds;
}

class Grid {
  /**
   * @param headerRows so dong tieu de o dau sheet (mac dinh 2)
   * @param freeze     true = bien cong thuc thanh gia tri tinh (mac dinh).
   *                  ExcelJS khong ghi duoc workbook khi sua o la master cua
   *                  shared formula, nen phai co dinh lai truoc.
   */
  constructor(wb, ws, srcFile, headerRows = 2, freeze = true) {
    this.wb = wb;
    this.ws = ws;
    this.srcFile = srcFile;
    this.headerRowCount = headerRows;
    this.log = []; // [{ row, col, giaTriCu, giaTriMoi, lyDo }]
    this.dongCongThuc = []; // cac o da bien cong thuc thanh gia tri
    if (freeze) this.dongCongThuc = freezeFormulas(this.ws);
  }

  static async open(file, sheetIndex = 0, headerRows = 2, freeze = true) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(file);
    const ws = wb.worksheets[sheetIndex] || wb.worksheets[0];
    if (!ws) throw new Error('File khong co sheet nao');
    return new Grid(wb, ws, file, headerRows, freeze);
  }

  get maxRow() {
    return this.ws.actualRowCount ? this.ws.rowCount : 0;
  }
  get maxCol() {
    return this.ws.actualColumnCount ? this.ws.columnCount : 0;
  }

  /** So cot cua tung dong (dung de doc header) */
  headerRows(n = 2) {
    const out = [];
    for (let r = 1; r <= n; r++) {
      const o = {};
      for (let c = 1; c <= this.maxCol; c++) {
        const t = show(this.ws.getRow(r).getCell(c).value);
        if (t) o[numToCol(c)] = t;
      }
      out.push(o);
    }
    return out;
  }

  /** Doc 1 dong duoi dang { row, cells: {A:..., B:...}, goc: {...} } */
  read(row) {
    const r = this.ws.getRow(row);
    const cells = {};
    for (let c = 1; c <= this.maxCol; c++) {
      const L = numToCol(c);
      cells[L] = show(r.getCell(c).value);
    }
    return { row, cells };
  }

  /**
   * Kiem tra 1 dong. Tra ve danh sach loi:
   *   muc  'sai'     - du lieu sai, can sua
   *   muc  'nghi'    - nen sua, nhung de quyet dinh
   *   muc  'ok'      - khong can lam gi
   */
  check(row, ctx = {}) {
    const d = this.read(row);
    const c = d.cells;
    // dong tieu de -> khong kiem tra
    if (row <= this.headerRowCount) {
      return { row, cells: c, loi: [], sai: 0, nghi: 0, laTieuDe: true };
    }
    const v = [];
    const add = (cot, muc, loi, goiY) => v.push({ cot, muc, loi, goiY: goiY || '' });

    const a = (c.A || '').trim();
    if (a !== '' && !/^\d+$/.test(a)) add('A', 'sai', `Cột A = "${a}" không phải số`, 'để trống hoặc nhập số hộ');
    if (a !== '' && parseInt(a, 10) > 141) add('A', 'nghi', `STT hộ ${a} vượt quá 141`);

    if (!(c.C || '').trim()) add('C', 'sai', 'Thiếu họ và tên');
    if ((c.C || '').trim() !== (c.C || '')) add('C', 'sai', 'Họ tên có khoảng trắng thừa', 'bỏ khoảng trắng đầu/cuối');

    if ((c.B || '').trim() !== '' && !/^\d+$/.test((c.B || '').trim()))
      add('B', 'sai', `STT người = "${c.B}" không phải số`);

    const e = (c.E || '').trim();
    if (e !== '' && serialToDate(+e)) {
      add('E', 'nghi', `Ngày sinh ghi dạng số thứ tự Excel (${e})`, `tương đương ${show(serialToDate(+e))} — nên định dạng lại thành ngày`);
    } else if (e !== '' && !toDate(e) && !/^\d{1,2}$/.test(e))
      add('E', 'sai', `Ngày sinh = "${e}" không đọc được`, 'định dạng dd/MM/yyyy');

    const f = (c.F || '').trim();
    if (f !== '' && !/^\d{12}$/.test(f.replace(/\s+/g, '')))
      add('F', 'sai', `CCCD = "${f}" không phải 12 số`);
    if (f.replace(/\s+/g, '').length !== f.length && /^\d{12}$/.test(f.replace(/\s+/g, '')))
      add('F', 'sai', 'CCCD có khoảng trắng thừa', 'bỏ khoảng trắng');

    const g = (c.G || '').trim();
    if (g !== '' && serialToDate(+g)) {
      add('G', 'nghi', `Ngày cấp ghi dạng số thứ tự Excel (${g})`, `tương đương ${show(serialToDate(+g))} — nên định dạng lại thành ngày`);
    } else if (g !== '' && !toDate(g) && !/^\d{1,2}$/.test(g))
      add('G', 'sai', `Ngày cấp = "${g}" không đọc được`, 'định dạng dd/MM/yyyy');

    const h = (c.H || '').trim();
    if (h === '' && e !== '' && toDate(e)) add('H', 'nghi', 'Thiếu tuổi', 'điền số tuổi');
    else if (h !== '' && !/^\d{1,3}$/.test(h)) add('H', 'sai', `Tuổi = "${h}" không phải số`);

    const i = (c.I || '').trim();
    if (i !== '' && !REL.includes(i))
      add('I', 'sai', `Quan hệ = "${i}"`, `chỉ nhận "${REL[0]}" hoặc "${REL[1]}"`);

    // chu ho phai la dong dau tien cua ho
    if (i === REL[0] && ctx.truocDo) {
      const pa = (ctx.truocDo.cells.A || '').trim();
      if (pa === a && a !== '') add('I', 'sai', 'Hai dòng liên tiếp cùng ghi "chủ hộ"', 'chỉ dòng đầu là chủ hộ');
    }
    if (i === REL[0] && a === '') add('I', 'nghi', 'Dòng ghi "chủ hộ" nhưng cột A trống', 'nhập STT hộ vào cột A');

    // tuoi co khop voi nam sinh khong
    if (h && toDate(e)) {
      const nam = toDate(e).getUTCFullYear();
      const namNay = new Date().getUTCFullYear();
      const tuoiTinh = namNay - nam;
      if (Math.abs(tuoiTinh - parseInt(h, 10)) >= 2)
        add('H', 'nghi', `Tuổi ghi ${h}, tính từ năm sinh ra ${tuoiTinh}`);
    }

    return { row, cells: c, loi: v, sai: v.filter((x) => x.muc === 'sai').length, nghi: v.filter((x) => x.muc === 'nghi').length };
  }

  /**
   * Sua 1 dong: chi chuan hoa gi tri chac chan.
   * Tra ve danh sach thay doi. Khong sua khi co loi "sai" chua xu ly.
   */
  fix(row) {
    if (row <= this.headerRowCount) return { row, doi: [], boQua: 0, lyDo: 'dong tieu de' };
    const kq = this.check(row);
    if (kq.sai > 0) return { row, doi: [], boQua: kq.sai, lyDo: 'còn lỗi sai, chưa sửa' };
    const r = this.ws.getRow(row);
    const c = kq.cells;
    const doi = [];

    const ghi = (col, giaMoi, lyDo, numFmt) => {
      const cell = r.getCell(colToNum(col));
      const cu = show(cell.value);
      const moi = show(giaMoi); // so sanh theo dang hien thi, khong so Date voi chuoi
      if (cu === moi) return;
      doi.push({ row, col, giaTriCu: cu, giaTriMoi: moi, lyDo });
      cell.value = giaMoi;
      if (numFmt) cell.numFmt = numFmt;
      this.log.push({ row, col, giaTriCu: cu, giaTriMoi: moi, lyDo });
    };

    ghi('A', (c.A || '').trim(), 'bỏ khoảng trắng');
    ghi('B', (c.B || '').trim(), 'bỏ khoảng trắng');
    ghi('C', (c.C || '').trim(), 'bỏ khoảng trắng');
    ghi('I', (c.I || '').trim(), 'bỏ khoảng trắng');

    const f = (c.F || '').trim().replace(/\s+/g, '');
    if (f) ghi('F', f, 'bỏ khoảng trắng trong CCCD');

    for (const col of ['E', 'G']) {
      const t = (c[col] || '').trim();
      if (!t) continue;
      // so thu tu ngay cua Excel chua co dinh dang -> chuyen thanh ngay that
      const serial = serialToDate(+t);
      if (serial) {
        ghi(col, serial, `chuyển số thứ tự ngày ${t} thành ${show(serial)}`, 'dd/mm/yyyy');
        continue;
      }
      const d = toDate(t);
      if (d) ghi(col, d, 'chuyển ngày về dạng thật', 'dd/mm/yyyy');
    }

    return { row, doi, boQua: 0 };
  }

  /**
   * Duyet tung dong tu `from` den `to`. KHONG sua gi - chi kiem tra.
   * Muon sua dung fixRange(from, to).
   */
  route(from, to) {
    return this.scan(from, to);
  }

  /** Chi kiem tra, khong sua */
  scan(from, to) {
    const buoc = [];
    for (let r = from; r <= to; r++) {
      buoc.push(this.check(r, { truocDo: r > from ? buoc[r - 1 - from] : null }));
    }
    return buoc;
  }

  /** Sua het mot doan */
  fixRange(from, to) {
    const ketQua = [];
    for (let r = from; r <= to; r++) ketQua.push(this.fix(r));
    return ketQua;
  }

  /** Cac cot chua ngay -> gan dinh dang khi sua */
  laCotNgay(col) {
    return ['E', 'G'].includes(String(col).toUpperCase());
  }

  /**
   * Sua tay mot o (tu web). Ghi nhat ky nhu cac sua tu dong.
   * Gia tri rong -> xoa o. Ngay -> luu dang Date + dd/mm/yyyy.
   */
  setCell(row, col, value) {
    const cell = this.ws.getRow(row).getCell(colToNum(col));
    const cu = show(cell.value);
    const s = value === null || value === undefined ? '' : String(value).trim();

    let giaMoi = s;
    if (s !== '' && this.laCotNgay(col)) {
      const d = toDate(s) || serialToDate(+s);
      if (d) {
        giaMoi = d;
        cell.numFmt = 'dd/mm/yyyy';
      }
    } else if (s !== '' && /^\d+$/.test(s) && ['A', 'B', 'F', 'H'].includes(String(col).toUpperCase())) {
      giaMoi = s.length <= 11 ? parseInt(s, 10) : s; // CCCD 12 so giu dang chuoi
    }

    if (show(giaMoi) === cu) return { row, col, giaTriCu: cu, giaTriMoi: show(giaMoi), doi: false };

    cell.value = s === '' ? null : giaMoi;
    const rec = { row, col, giaTriCu: cu, giaTriMoi: s, lyDo: 'sua tay', thuCong: true };
    this.log = this.log.filter((l) => !(l.row === row && l.col === col));
    this.log.push(rec);
    return { ...rec, doi: true };
  }

  /** Thong ke ngan cho mot doan */
  thongKe(from, to) {
    const buoc = this.scan(from, to);
    return {
      tong: buoc.length,
      soTieuDe: buoc.filter((b) => b.laTieuDe).length,
      ok: buoc.filter((b) => !b.laTieuDe && !b.sai && !b.nghi).length,
      nghi: buoc.filter((b) => b.nghi).length,
      sai: buoc.filter((b) => b.sai).length,
      buoc,
    };
  }

  async save(dest) {
    await X.writeXlsx(this.wb, dest, 'file Excel');
    return dest;
  }
}

module.exports = { Grid, show, toDate, colToNum, numToCol, serialToDate, isUnresolvedFormula, freezeFormulas, REL, A };
