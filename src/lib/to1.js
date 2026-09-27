'use strict';
/**
 * Nghiep vu: doc "Danh sach To 1" + "Danh sach Trau Binh", gom thanh ho,
 * khoi phuc STT ho cho dong bi chon dong bang cach doi chieu CCCD.
 *
 * Nguyen tac: KHONG XOA DONG NAO. Moi thay doi deu ghi lai canh bao + giu nguyen gia tri goc.
 */
const X = require('./xlsx');
const { CHU_HO, RELATIONSHIPS } = X;

const BLOCK_BOUNDARY = 328; // duoi dong nay la khoi du lieu thu 2 (schema rut gon)

/**
 * Pham vi khoi phuc STT ho bang CCCD.
 * Nguoi dung xac nhan: chi khoi phuc vung bi chon dong o giua hộ 20 và hộ 78
 * (dong 149..325). Cac dong thanh vien ben ngoai giu nguyen cot A trong.
 * Dat null de khoi phuc toan bo.
 */
const RECOVER_RANGE = { from: 149, to: 325 };

/** Dong nao duoc phep khoi phuc STT ho */
function canRecover(row, range) {
  if (!range) return true;
  return row >= range.from && row <= range.to;
}

/** Trang thai 3 khoi du lieu nguoi dung da xac nhan */
const KHUNG = {
  DA_FIX: '1_DA_FIX',
  CHUA_FIX: '2_CHUA_FIX_THANH_VIEN_BI_CHON',
  CHUA_CAP: '3_CHUA_CAP_NHAT',
  RAC: 'STT_HO_RAC',
  NGOAI: 'NGOAI_PHU',
};

function khungOf(sttHo) {
  if (sttHo == null || sttHo <= 0) return KHUNG.RAC;
  if (sttHo <= 27) return KHUNG.DA_FIX;
  if (sttHo <= 78) return KHUNG.CHUA_FIX;
  if (sttHo <= 141) return KHUNG.CHUA_CAP;
  return KHUNG.NGOAI;
}

/**
 * Phan loai thong bao:
 *   CANH_BAO = loi trong du lieu, can nguoi doi chieu / sua
 *   GHI_CHU  = trang thai xu ly, khong phai loi (vd "chua co STT ho trong DS tra")
 */
const MUC = { CANH_BAO: 'canh-bao', GHI_CHU: 'ghi-chu' };
const GHI_CHU_CODES = new Set([
  'KHONG_KHOP_DS_TRA',
  'HO_TRA_CHUA_CO_MAPPING',
  'HO_TRA_XUNG_DOT',
]);
const mucOf = (code) => (GHI_CHU_CODES.has(code) ? MUC.GHI_CHU : MUC.CANH_BAO);

/**
 * @param {Array<{row:number,c:object}>} rows  hang cua sheet
 * @returns {{people: Array<object>, issues: Array<object>}}
 */
function parseTo1(rows) {
  const people = [];
  const issues = [];
  const flag = (p, code, msg) =>
    issues.push({ row: p.row, ten: p.hoTen, code, muc: mucOf(code), moTa: msg });

  let curA = '';
  let curHoFwd = 0;

  for (const r of rows) {
    if (r.row < 3) continue;
    const hoTen = X.txt(r.c[3]);
    if (!hoTen) continue;

    const aRaw = X.txt(r.c[1]);
    if (aRaw !== '') {
      curA = aRaw;
      const n = X.intOf(r.c[1]);
      if (n != null) curHoFwd = n;
    }
    const isHeadRow = aRaw !== '';

    // --- quan he: vi tri khac nhau giua 2 khoi, dung fallback de tim ---
    let quanHe = '';
    let quanHeCol = '';
    for (const [col, label] of [
      [9, 'I'],
      [8, 'H'],
      [7, 'G'],
    ]) {
      if (RELATIONSHIPS.includes(X.txt(r.c[col]))) {
        quanHe = X.txt(r.c[col]);
        quanHeCol = label;
        break;
      }
    }

    // --- ngay sinh / CCCD: dong bi lech cot khi CCCD nham o cot E ---
    let nsRaw = X.txt(r.c[5]);
    const cccdRaw = X.txt(r.c[6]);
    let rowShifted = false;
    if (/^\d{12}$/.test(nsRaw) && cccdRaw === '') rowShifted = true;

    let ngaySinh = rowShifted ? '' : X.isoDate(r.c[5]);
    const cccd = rowShifted ? nsRaw : X.cccdOf(r.c[6]);
    const ngayCap = rowShifted ? '' : X.isoDate(r.c[7]);

    // --- tuoi: chi khoi 1 co ---
    let tuoi = null;
    if (r.row <= BLOCK_BOUNDARY) {
      const t = X.intOf(r.c[8]);
      if (t != null) tuoi = t;
      else if (X.txt(r.c[8]) !== '') flag({ row: r.row, hoTen }, 'TUOI_KHONG_PHAI_SO', `Gia tri "${X.txt(r.c[8])}"`);
    }

    // --- ma bao hiem ---
    let maBH = X.txt(r.c[4]);
    const khongBH = maBH === X.KHONG_CO;
    if (khongBH) maBH = '';

    const p = {
      row: r.row,
      block: r.row <= BLOCK_BOUNDARY ? 1 : 2,
      rowShifted,
      cotARaw: aRaw,
      hoGoc: curHoFwd,
      isHeadRow,
      hoTen,
      sttNguoiGoc: X.txt(r.c[2]),
      quanHe,
      quanHeCol,
      maBH,
      khongBH,
      ngaySinh,
      ngaySinhRaw: X.txt(r.c[5]),
      cccd,
      ngayCap,
      tuoi,
      dienThoai: X.txt(r.c[10]),
      duHoc: X.txt(r.c[11]),
      layChong: X.txt(r.c[12]),
      xkld: X.txt(r.c[13]),
      ngheNghiep: X.txt(r.c[14]),
      ghiChu: X.txt(r.c[25]),
      datDai: {
        lua: { dienTich: X.txt(r.c[15]) },
        vuon: { cay: X.txt(r.c[16]), soLuong: X.txt(r.c[17]), dienTich: X.txt(r.c[18]) },
        rauMau: { tenRau: X.txt(r.c[19]), dienTich: X.txt(r.c[20]) },
        ao: { tenNuoi: X.txt(r.c[21]), soLuong: X.txt(r.c[22]) },
        chanNuoi: { giaXuc: X.txt(r.c[23]), giaCam: X.txt(r.c[24]) },
      },
      canhBao: [],
      hoTra: null,
      matchBy: '',
      sttHo: 0,
      sttNguoi: 0,
      vaiTro: '',
      nguonHo: '',
    };
    if (rowShifted) p.canhBao.push('DONG_BI_LECH_COT_DA_KHOI_PHUCCCD');

    // canh bao
    if (!quanHe) flag(p, 'THIEU_QUAN_HE', "Khong co o 'chu ho' / 'thanh vien'");
    if (rowShifted) {
      flag(p, 'DONG_BI_LECH_COT', `CCCD "${nsRaw}" nam o cot E, quan he o cot G (thay vi F va I)`);
    } else if (!ngaySinh) {
      if (nsRaw === '') flag(p, 'THIEU_NGAY_SINH', 'Cot E trong');
      else flag(p, 'NGAY_SINH_KHONG_DOC_DUOC', `Gia tri "${nsRaw}" khong phai ngay sinh`);
    }
    if (!rowShifted && X.txt(r.c[7]) !== '' && !ngayCap) {
      flag(p, 'NGAY_CAP_KHONG_DOC_DUOC', `Gia tri "${X.txt(r.c[7])}"`);
    }
    if (!cccd) {
      if (cccdRaw !== '') flag(p, 'CCCD_KHONG_DUNG_DOI', `Gia tri "${cccdRaw}" khong phai 12 so`);
      else flag(p, 'THIEU_CCCD', 'Khong co so dinh danh');
    }
    if (!maBH) flag(p, 'THIEU_MA_BH', 'Khong co ma bao hiem');

    people.push(p);
  }
  return { people, issues };
}

/** Doc danh sach tra cuu (khong co cot so ho, chi co "chu ho" de tach nhom) */
function parseLookup(rows) {
  const list = [];
  let hhNo = 0;
  for (const r of rows) {
    if (r.row < 2) continue;
    const hoTen = X.txt(r.c[2]);
    if (!hoTen) continue;
    if (X.txt(r.c[7]) === CHU_HO) hhNo++;
    const sdd = X.txt(r.c[3]);
    list.push({
      stt: X.intOf(r.c[1]) || 0,
      hoTen,
      soDinhDanh: sdd === X.KHONG_CO ? '' : sdd,
      ngaySinh: X.isoDate(r.c[4]),
      cccd: X.cccdOf(r.c[5]),
      ngayCap: X.txt(r.c[6]),
      quanHe: X.txt(r.c[7]),
      hoTraStt: hhNo,
    });
  }
  return list;
}

/** Tao bang tra cuu nhanh theo CCCD va theo ten da chuan hoa */
function indexLookup(list) {
  const byCccd = new Map();
  const byName = new Map();
  for (const p of list) {
    if (p.cccd && !byCccd.has(p.cccd)) byCccd.set(p.cccd, p);
    const nk = X.normName(p.hoTen);
    if (nk && !byName.has(nk)) byName.set(nk, p);
  }
  return { byCccd, byName };
}

/** Gan hoTra cho tung nguoi To 1 bang CCCD, fallback ten */
function matchLookup(people, idx) {
  let c = 0,
    t = 0,
    k = 0;
  for (const p of people) {
    let hit = null;
    if (p.cccd && idx.byCccd.has(p.cccd)) {
      hit = idx.byCccd.get(p.cccd);
      p.matchBy = 'CCCD';
      c++;
    } else {
      const nk = X.normName(p.hoTen);
      if (nk && idx.byName.has(nk)) {
        hit = idx.byName.get(nk);
        p.matchBy = 'ten';
        t++;
      }
    }
    if (hit) p.hoTra = hit.hoTraStt;
    else k++;
  }
  return { cccd: c, ten: t, khongKhop: k };
}

/** Danh dau CCCD bi trung giua cac dong */
function markDuplicateCccd(people, issues) {
  const groups = new Map();
  for (const p of people) {
    if (!p.cccd) continue;
    if (!groups.has(p.cccd)) groups.set(p.cccd, []);
    groups.get(p.cccd).push(p);
  }
  const dupRows = new Map();
  for (const [cccd, g] of groups) {
    if (g.length < 2) continue;
    const rows = g.map((m) => m.row).join(', ');
    for (const m of g) {
      dupRows.set(m.row, true);
      m.canhBao.push('CCCD_TRUNG');
      issues.push({
        row: m.row,
        ten: m.hoTen,
        code: 'CCCD_TRUNG',
        muc: mucOf('CCCD_TRUNG'),
        moTa: `CCCD ${cccd} xuat hien o dong ${rows}`,
      });
    }
  }
  return dupRows;
}

/**
 * Bang do hoTra (so ho trong DS tra) -> STT ho cua To 1.
 * Chi dung cac dong CHU HO trong To 1 (dong co STT ho o cot A) lam chuan.
 * Neu mot hoTra chi khop mot STT ho -> an toan. Neu nhieu -> bo dong trung CCCD roi thu lai,
 * con xung dot thi bo qua (khong doan).
 */
function buildHouseMap(people, dupRows) {
  const cand = new Map();
  for (const p of people) {
    if (!p.isHeadRow || !p.hoTra || p.hoGoc <= 0) continue;
    if (!X.intOf(p.cotARaw)) continue; // cot A rac (vd '.') -> khong dua vao chuan
    if (!cand.has(p.hoTra)) cand.set(p.hoTra, new Map());
    const m = cand.get(p.hoTra);
    m.set(p.hoGoc, (m.get(p.hoGoc) || 0) + 1);
  }

  const map = new Map();
  const ambiguous = new Map();
  const log = [];
  for (const [hoTra, counts] of cand) {
    const keys = [...counts.keys()];
    if (keys.length === 1) {
      map.set(hoTra, keys[0]);
      continue;
    }
    // thu loai bo dong bi trung CCCD khoi phia nguon
    const clean = keys.filter((k) =>
      people.some(
        (p) => p.isHeadRow && p.hoTra === hoTra && p.hoGoc === k && !dupRows.has(p.row)
      )
    );
    if (clean.length === 1) {
      map.set(hoTra, clean[0]);
      log.push({ hoTra, chon: clean[0], lyDo: 'bo dong trung CCCD', tru: keys });
    } else {
      ambiguous.set(hoTra, keys);
      log.push({ hoTra, chon: null, lyDo: 'xung dot', tru: keys });
    }
  }
  return { map, ambiguous, log };
}

/** Gán lai STT ho cho dong thanh vien dua tren bang do */
function recoverHouseholds(people, houseMap, ambiguous, issues, range) {
  let khoiPhuc = 0,
    giuNguyen = 0;
  for (const p of people) {
    if (p.isHeadRow) {
      p.sttHo = X.intOf(p.cotARaw) || 0;
      p.nguonHo = 'cot-A';
      continue;
    }
    if (p.hoTra && houseMap.has(p.hoTra) && canRecover(p.row, range)) {
      p.sttHo = houseMap.get(p.hoTra);
      p.nguonHo = 'khoi-phuc-CCCD';
      p.canhBao.push('DA_KHOI_PHU_HO_TU_DS_TRAU_BINH');
      khoiPhuc++;
    } else if (canRecover(p.row, range)) {
      // Trong vung 149..325 cac dong thanh vien nam tan manh, khong the suy ra
      // ho tu vi tri. File tra la nguon chuan: khong tra duoc thi de trong,
      // KHONG doan vao chu hộ dang mo (truoc do "cot-A(forward)" lam ho 78
      // nhan nham 10 nguoi thuoc ho khac).
      p.sttHo = 0;
      p.nguonHo = 'chua-xac-dinh-ho';
      p.canhBao.push('CHUA_XAC_DINH_HO_TU_DS_TRA');
      issues.push({
        row: p.row,
        ten: p.hoTen,
        code: 'CHUA_XAC_DINH_HO_TU_DS_TRA',
        muc: mucOf('CHUA_XAC_DINH_HO_TU_DS_TRA'),
        moTa: p.hoTra
          ? `Co CCCD ${p.cccd} khop hoTra ${p.hoTra} trong DS tra, nhung ho do khong co chu hộ tuong ung trong Tổ 1`
          : 'Khong tim thay trong DS tra; cung khong xac dinh duoc ho nao tu vi tri',
      });
      giuNguyen++;
    } else {
      p.sttHo = p.hoGoc;
      p.nguonHo = 'cot-A(forward)';
      p.canhBao.push('CHUA_KHOI_PHU_HO');
      if (!p.hoTra) {
        issues.push({
          row: p.row,
          ten: p.hoTen,
          code: 'KHONG_KHOP_DS_TRA',
          muc: mucOf('KHONG_KHOP_DS_TRA'),
          moTa: 'Khong tim thay trong danh sach Trau Binh theo CCCD hay ten',
        });
      } else if (ambiguous.has(p.hoTra)) {
        issues.push({
          row: p.row,
          ten: p.hoTen,
          code: 'HO_TRA_XUNG_DOT',
          muc: mucOf('HO_TRA_XUNG_DOT'),
          moTa: `Ho #${p.hoTra} trong DS tra khop voi nhieu STT ho To 1 (${ambiguous.get(p.hoTra).join('/')})`,
        });
      } else {
        issues.push({
          row: p.row,
          ten: p.hoTen,
          code: 'HO_TRA_CHUA_CO_MAPPING',
          muc: mucOf('HO_TRA_CHUA_CO_MAPPING'),
          moTa: `Ho #${p.hoTra} trong DS tra chua co nguoi To 1 nao de xac dinh STT ho`,
        });
      }
      giuNguyen++;
    }
  }
  return { khoiPhuc, giuNguyen };
}

/**
 * Gom nguoi thanh ho.
 *
 * QUAN TRONG: chi dong DAU HO (cot A co gia tri) moi duoc phep tao nhom.
 * Neu de dong thanh vien tao nhom truoc, no se "chim" STT ho cua chinh no
 * (vi du dong 74 la thanh vien nhung lai duoc gan ve ho 36 -> nhom 36
 * mat so 36). Vi vay phai tach 2 luot:
 *   Luot 1: moi dong dau ho tao 1 nhom rieng, giu nguyen cot A goc.
 *   Luot 2: dong thanh vien gan vao nhom co STT ho khop.
 */
function buildHouseholds(people, issues) {
  const byStt = new Map(); // sttHo -> [nhom] (co the nhieu, vi du ho 105 trung)
  const order = [];

  const newGroup = (head, key) => ({
    key,
    sttHo: /^\d+$/.test(head.cotARaw) ? parseInt(head.cotARaw, 10) : 0,
    cotARaw: head.cotARaw,
    members: [],
    chuHo: '',
    chuHoRow: 0,
    rowMin: head.row,
    rowMax: head.row,
    canhBao: [],
    goiY: '',
  });

  const attach = (g, p) => {
    g.members.push(p);
    if (p.row < g.rowMin) g.rowMin = p.row;
    if (p.row > g.rowMax) g.rowMax = p.row;
    if (p.quanHe === CHU_HO && !g.chuHo) {
      g.chuHo = p.hoTen;
      g.chuHoRow = p.row;
    }
  };

  let racOpen = null;
  const orphan = [];

  for (const p of people) {
    // --- dong dau ho: luon tao nhom rieng ---
    if (p.isHeadRow) {
      if (/^\d+$/.test(p.cotARaw)) {
        const g = newGroup(p, `H${p.row}`);
        order.push(g);
        if (!byStt.has(g.sttHo)) byStt.set(g.sttHo, []);
        byStt.get(g.sttHo).push(g);
        attach(g, p);
        racOpen = null;
      } else {
        if (!racOpen) {
          racOpen = newGroup(p, `A${p.row}`);
          order.push(racOpen);
        } else {
          racOpen.cotARaw = p.cotARaw;
        }
        attach(racOpen, p);
      }
      continue;
    }

    // --- dong thanh vien: con trong khoi "." dang mo -> giu vao nhom do ---
    if (racOpen) {
      attach(racOpen, p);
      continue;
    }
    const cands = byStt.get(p.sttHo);
    if (cands && cands.length === 1) {
      attach(cands[0], p);
    } else if (cands && cands.length > 1) {
      const c = cands.reduce((best, g) =>
        Math.abs(g.rowMin - p.row) < Math.abs(best.rowMin - p.row) ? g : best
      );
      attach(c, p);
      p.canhBao.push('STT_HO_TRUNG_TRONG_DANH_SACH');
    } else {
      p.canhBao.push('CHUA_TIM_THAY_HO_CHUA_DUNG_STT');
      orphan.push(p);
    }
  }

  // nhom mo co: khong xac dinh duoc ho -> tach rieng, canh bao
  for (const p of orphan) {
    const g = {
      key: `M${p.row}`,
      sttHo: 0,
      cotARaw: '',
      members: [p],
      chuHo: p.quanHe === CHU_HO ? p.hoTen : '',
      chuHoRow: p.quanHe === CHU_HO ? p.row : 0,
      rowMin: p.row,
      rowMax: p.row,
      canhBao: [],
      goiY: p.hoTra
        ? `Xem lai ho so ${p.cccd} trong DS trau Binh: khop ho #${p.hoTra} nhung ho do khong co chu ho tuong ung trong Tổ 1`
        : 'Xem lai ho so nay o nguon: khong tim thay trong DS trau Binh theo CCCD hay ten',
    };
    order.push(g);
  }

  // sap xep theo dong chu ho
  order.sort((a, b) => (a.chuHoRow || a.rowMin) - (b.chuHoRow || b.rowMin));

  // gan STT nguoi lien tuc + vai tro + canh bao
  let stt = 0;
  order.forEach((hh, i) => {
    hh.no = i + 1;
    hh.id = `HO-${hh.rowMin}`;
    hh.khung = khungOf(hh.sttHo);
    hh.soNguoi = hh.members.length;
    hh.members.sort((a, b) => a.row - b.row);
    hh.members.forEach((m, j) => {
      stt++;
      m.sttNguoi = stt;
      m.sttHo = hh.sttHo;
      m.vaiTro = j === 0 && m.quanHe === CHU_HO ? 'chuHo' : 'thanhVien';
      if (j > 0 && m.quanHe === CHU_HO) m.canhBao.push('DONG_DAU_KHONG_PHAI_CHU_HO');
    });

    // canh bao cap ho
    const w = new Set();
    if (hh.cotARaw && !/^\d+$/.test(hh.cotARaw)) w.add('STT_HO_RAC');
    if (hh.soNguoi === 1) w.add('HO_CHI_CO_1_NGUOI');
    if (hh.khung === KHUNG.CHUA_FIX) w.add('CAN_THEM_THANH_VIEN_TU_DS_TRAU_BINH');
    if (hh.khung === KHUNG.CHUA_CAP) w.add('CHUA_CAP_NHAT_TU_DS_TRAU_BINH');
    if (!hh.chuHo) w.add('THIEU_CHU_HO');
    const nFix = hh.members.filter((m) => m.nguonHo === 'khoi-phuc-CCCD').length;
    if (nFix > 0) w.add(`DA_KHOI_PHU_${nFix}_THANH_VIEN`);
    hh.members.forEach((m) => m.canhBao.forEach((c) => w.add(`NGUOI:${c}`)));
    hh.canhBao = [...w];

    // nhieu dong danh "chu ho" trong cung nhom
    const heads = hh.members.filter((m) => m.quanHe === CHU_HO);
    if (heads.length > 1) {
      issues.push({
        row: heads[1].row,
        ten: heads[1].hoTen,
        code: 'TRONG_HO_HAI_CHU_HO',
        muc: mucOf('TRONG_HO_HAI_CHU_HO'),
        moTa: `Nhom co ${heads.length} dong danh "chu ho" (dong ${heads.map((h) => h.row).join(', ')})`,
      });
    }

    // goi y khi cot A rac
    if (hh.cotARaw && !/^\d+$/.test(hh.cotARaw)) {
      for (let k = i - 1; k >= 0; k--) {
        if (/^\d+$/.test(order[k].cotARaw)) {
          hh.goiY = `Cot A dong ${hh.rowMin}-${hh.rowMax} = "${hh.cotARaw}"; theo vi tri thi thuoc ho ${order[k].sttHo} (bat dau dong ${order[k].rowMin}). Can kiem tra lai.`;
          break;
        }
      }
    }
  });

  return order;
}

/** Chay tron: duong dan file -> du lieu da gom nhom */
function build({ to1Rows, lookupRows, ap, to, recoverRange }) {
  const { people, issues } = parseTo1(to1Rows);
  const lookup = parseLookup(lookupRows);
  const idx = indexLookup(lookup);
  const match = matchLookup(people, idx);
  const dupRows = markDuplicateCccd(people, issues);
  const { map, ambiguous, log } = buildHouseMap(people, dupRows);
  const range = recoverRange === undefined ? RECOVER_RANGE : recoverRange;
  const rec = recoverHouseholds(people, map, ambiguous, issues, range);
  const households = buildHouseholds(people, issues);

  issues.sort((a, b) => a.row - b.row || a.code.localeCompare(b.code));

  const byKhung = {};
  for (const h of households) {
    byKhung[h.khung] = byKhung[h.khung] || { soHo: 0, soNguoi: 0 };
    byKhung[h.khung].soHo++;
    byKhung[h.khung].soNguoi += h.soNguoi;
  }

  return {
    ap,
    to,
    tongNguoi: people.length,
    tongHo: households.length,
    match,
    houseMapSize: map.size,
    houseMapAmbiguous: ambiguous.size,
    houseMapLog: log,
    recover: rec,
    recoverRange: range,
    byKhung,
    households,
    people,
    lookup,
    issues,
  };
}

module.exports = {
  build,
  parseTo1,
  parseLookup,
  indexLookup,
  matchLookup,
  buildHouseholds,
  khungOf,
  canRecover,
  KHUNG,
  MUC,
  BLOCK_BOUNDARY,
  RECOVER_RANGE,
};
