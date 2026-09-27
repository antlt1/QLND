'use strict';
/** CLI: xuat output/*.json -> file .xlsx sach, da gom ho, STT nguoi lien tuc, co cot canh bao */
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const X = require('../lib/xlsx');
const { resolve } = require('../lib/chon');

const HEADER_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
const HEADER_FONT = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
const THIN = { style: 'thin', color: { argb: 'FFBFBFBF' } };
const BORDER = { top: THIN, left: THIN, bottom: THIN, right: THIN };

const FILL = {
  DA_FIX: 'FFD9EAD3',
  CHUA_FIX: 'FFFFF2CC',
  CHUA_CAP: 'FFE4D9F3',
  RAC: 'FFF4CCCC',
  NONE: null,
};

function styleHeader(ws, ncol) {
  const r = ws.getRow(1);
  r.height = 30;
  for (let i = 1; i <= ncol; i++) {
    const c = r.getCell(i);
    c.fill = HEADER_FILL;
    c.font = HEADER_FONT;
    c.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    c.border = BORDER;
  }
  ws.views = [{ state: 'frozen', ySplit: 1 }];
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: ncol } };
}

function addSheet(wb, name, cols, rows, widths) {
  const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 1 }] });
  ws.columns = cols.map((c) => ({ header: c.h, key: c.k, width: widths?.[c.k] || c.w || 14 }));
  for (const r of rows) ws.addRow(r);
  styleHeader(ws, cols.length);
  return ws;
}

function joinList(a) {
  return Array.isArray(a) && a.length ? a.join(' | ') : '';
}

async function main() {
  const args = process.argv.slice(2);
  const cfg = resolve(args);
  const { outDir, to } = cfg;
  const doc = JSON.parse(fs.readFileSync(path.join(outDir, `to${to}-ho.json`), 'utf8'));
  const issues = JSON.parse(fs.readFileSync(path.join(outDir, `to${to}-loi.json`), 'utf8'));
  const lookup = JSON.parse(fs.readFileSync(path.join(outDir, 'tra-cuoc.json'), 'utf8'));

  const wb = new ExcelJS.Workbook();
  wb.creator = 'QLND tool';
  wb.created = new Date();

  // ---------- 1. Bang tong hop ----------
  const sum = wb.addWorksheet('Tổng hợp');
  sum.columns = [{ width: 34 }, { width: 12 }, { width: 12 }, { width: 60 }];
  sum.addRow(['ẤP', doc.ap, '', '']);
  sum.addRow(['Tổ', doc.to, '', '']);
  sum.addRow(['Tổng số hộ', doc.tongHo, '', '']);
  sum.addRow(['Tổng số người', doc.tongNguoi, '', '']);
  sum.addRow(['Cảnh báo (cần đối chiếu)', issues.filter((i) => (i.muc || 'canh-bao') === 'canh-bao').length, '', 'lỗi trong dữ liệu nguồn']);
  sum.addRow(['Ghi chú (trạng thái xử lý)', issues.filter((i) => i.muc === 'ghi-chu').length, '', 'không phải lỗi']);
  sum.addRow([]);
  sum.addRow(['Khối dữ liệu', 'Số hộ', 'Số người', 'Ý nghĩa']);
  const ds = `Danh sách ${doc.ap || 'tra'}`;
  const YNGHIA = {
    '1_DA_FIX': 'Đã sửa xong. Mỗi hộ có 1 dòng "chủ hộ" ở đầu.',
    '2_CHUA_FIX_THANH_VIEN_BI_CHON': `Thành viên bị chồn dòng. Cần tra ${ds} để bổ sung.`,
    '3_CHUA_CAP_NHAT': `Chưa cập nhật từ ${ds}.`,
    STT_HO_RAC: 'Cột A gõ dấu "." — cần kiểm tra lại.',
    NGOAI_PHU: 'Ngoài phạm vi STT 1–141.',
  };
  for (const [k, v] of Object.entries(doc.thongKe?.byKhung || {})) {
    sum.addRow([k, v.soHo, v.soNguoi, YNGHIA[k] || '']);
  }
  sum.addRow([]);
  const rr = doc.quyTac?.khoiPhuc || '';
  const rrDong = doc.recoverRange && doc.recoverRange.from >= 0
    ? ` (dòng ${doc.recoverRange.from}–${doc.recoverRange.to})`
    : ' (tắt)';
  sum.addRow([`Khôi phục STT hộ${rrDong}`, '', '', '']);
  sum.addRow(['  đã khôi phục', doc.thongKe?.recover?.khoiPhuc ?? '', '', `đối chiếu CCCD với ${ds}`]);
  sum.addRow(['  giữ nguyên', doc.thongKe?.recover?.giuNguyen ?? '', '', 'ngoài phạm vi, hoặc chưa xác định được hộ']);
  sum.addRow(['Bản đồ hộ', doc.thongKe?.houseMapSize ?? '', '', `số hộ tra → STT hộ Tổ ${to} xác định được`]);
  void rr;
  for (let i = 1; i <= 4; i++) sum.getRow(i).font = { bold: true, size: 12 };
  for (let i = 1; i <= 6; i++) {
    sum.getCell(i, 1).font = { bold: true };
    sum.getCell(i, 2).font = { bold: true, size: 12 };
  }
  const hr = sum.getRow(7);
  for (let i = 1; i <= 4; i++) {
    hr.getCell(i).fill = HEADER_FILL;
    hr.getCell(i).font = HEADER_FONT;
  }
  const khungRows = Object.keys(doc.thongKe?.byKhung || {}).length;
  for (let i = 0; i < khungRows; i++) {
    const r = sum.getRow(8 + i);
    const f = FILL[r.getCell(1).value];
    if (f) for (let c = 1; c <= 4; c++) r.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: f } };
  }

  // ---------- 2. Danh sach ho (moi ho 1 dong) ----------
  const hh = doc.ho;
  addSheet(
    wb,
    'Danh sách hộ',
    [
      { h: 'TT hộ', k: 'no' },
      { h: 'STT hộ', k: 'sttHo' },
      { h: 'Giá trị cột A gốc', k: 'cotARaw', w: 16 },
      { h: 'Khối', k: 'khung', w: 32 },
      { h: 'Chủ hộ', k: 'chuHo', w: 26 },
      { h: 'Số người', k: 'soNguoi' },
      { h: 'Dòng Excel đầu', k: 'rowMin' },
      { h: 'Dòng Excel cuối', k: 'rowMax' },
      { h: 'Danh sách thành viên', k: 'ds', w: 80 },
      { h: 'Cảnh báo', k: 'cb', w: 44 },
      { h: 'Gợi ý', k: 'goiY', w: 60 },
    ],
    hh.map((h) => ({
      no: h.no,
      sttHo: h.sttHo || '',
      cotARaw: h.cotARaw,
      khung: h.khung,
      chuHo: h.chuHo,
      soNguoi: h.soNguoi,
      rowMin: h.rowMin,
      rowMax: h.rowMax,
      ds: h.thanhVien
        .map((m) => `${m.sttNguoi}. ${m.hoTen}${m.ngaySinh ? ` (${m.ngaySinh.slice(0, 4)})` : ''}`)
        .join('\n'),
      cb: joinList(h.canhBao),
      goiY: h.goiY,
    })),
    { ds: 80, cb: 44, goiY: 60, khung: 32, chuHo: 26 }
  );
  const wsHh = wb.getWorksheet('Danh sách hộ');
  for (let i = 0; i < hh.length; i++) {
    const f = FILL[hh[i].khung];
    if (f) {
      for (let c = 1; c <= 11; c++) {
        wsHh.getRow(i + 2).getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: f } };
      }
    }
    wsHh.getRow(i + 2).alignment = { vertical: 'top' };
  }

  // ---------- 3. Danh sach nguoi (1 dong 1 nguoi) ----------
  const people = [];
  for (const h of doc.ho) {
    for (const m of h.thanhVien) people.push({ h, m });
  }
  const wsN = addSheet(
    wb,
    'Danh sách người',
    [
      { h: 'STT người', k: 'sttNguoi' },
      { h: 'STT hộ', k: 'sttHo' },
      { h: 'Khối', k: 'khung', w: 30 },
      { h: 'Vai trò', k: 'vaiTro' },
      { h: 'Họ và tên', k: 'hoTen', w: 26 },
      { h: 'Năm sinh', k: 'namSinh' },
      { h: 'Tuổi', k: 'tuoi' },
      { h: 'CCCD', k: 'cccd', w: 16 },
      { h: 'Mã BH', k: 'maBH', w: 14 },
      { h: 'Ngày cấp', k: 'ngayCap', w: 12 },
      { h: 'Quan hệ (gốc)', k: 'quanHe', w: 13 },
      { h: 'Nguồn STT hộ', k: 'nguonHo', w: 17 },
      { h: 'Hộ trong DS tra', k: 'hoTra' },
      { h: 'Điện thoại', k: 'dienThoai', w: 14 },
      { h: 'Học vấn', k: 'duHoc' },
      { h: 'Lấy chồng', k: 'layChong' },
      { h: 'XKLD', k: 'xkld', w: 20 },
      { h: 'Nghề nghiệp', k: 'ngheNghiep', w: 20 },
      { h: 'Lúa (m²)', k: 'lua' },
      { h: 'Vườn', k: 'vuon', w: 20 },
      { h: 'Rau màu', k: 'rau', w: 18 },
      { h: 'Ao', k: 'ao', w: 18 },
      { h: 'Chăn nuôi', k: 'chanNuoi', w: 18 },
      { h: 'Ghi chú', k: 'ghiChu', w: 26 },
      { h: 'Dòng Excel', k: 'row' },
      { h: 'Cảnh báo', k: 'cb', w: 40 },
    ],
    people.map(({ h, m }) => {
      const dd = m.datDai || {};
      const vuon = dd.vuon || {};
      const rau = dd.rauMau || {};
      const ao = dd.ao || {};
      const cn = dd.chanNuoi || {};
      return {
        sttNguoi: m.sttNguoi,
        sttHo: h.sttHo || '',
        khung: h.khung,
        vaiTro: m.vaiTro,
        hoTen: m.hoTen,
        namSinh: m.ngaySinh ? m.ngaySinh.slice(0, 4) : '',
        tuoi: m.tuoi ?? '',
        cccd: m.cccd,
        maBH: m.maBH,
        ngayCap: m.ngayCap,
        quanHe: m.quanHe,
        nguonHo: m.nguonHo,
        hoTra: m.hoTra ?? '',
        dienThoai: m.dienThoai,
        duHoc: m.duHoc,
        layChong: m.layChong,
        xkld: m.xkld,
        ngheNghiep: m.ngheNghiep,
        lua: dd.lua?.dienTich || '',
        vuon: [vuon.cay, vuon.soLuong && `${vuon.soLuong} cây`, vuon.dienTich && `${vuon.dienTich}m²`]
          .filter(Boolean)
          .join(' '),
        rau: [rau.tenRau, rau.dienTich && `${rau.dienTich}m²`].filter(Boolean).join(' '),
        ao: [ao.tenNuoi, ao.soLuong && `${ao.soLuong}`].filter(Boolean).join(' '),
        chanNuoi: [cn.giaXuc && `gia xúc: ${cn.giaXuc}`, cn.giaCam && `gia cầm: ${cn.giaCam}`]
          .filter(Boolean)
          .join('; '),
        ghiChu: m.ghiChu,
        row: m.row,
        cb: joinList(m.canhBao),
      };
    }),
    { hoTen: 26, cccd: 16, khung: 30, cb: 40, vuon: 20, rau: 18, ao: 18, chanNuoi: 18, ghiChu: 26, nguonHo: 17 }
  );
  for (let i = 0; i < people.length; i++) {
    const f = FILL[people[i].h.khung];
    if (f) for (let c = 1; c <= 26; c++) {
      wsN.getRow(i + 2).getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: f } };
    }
  }

  // ---------- 4. Canh bao + ghi chu ----------
  const cb = issues.filter((i) => (i.muc || 'canh-bao') === 'canh-bao');
  const gc = issues.filter((i) => (i.muc || 'canh-bao') === 'ghi-chu');
  addSheet(
    wb,
    'Cảnh báo',
    [
      { h: 'Dòng Excel', k: 'row' },
      { h: 'Họ tên', k: 'ten', w: 26 },
      { h: 'Mã', k: 'code', w: 30 },
      { h: 'Mô tả', k: 'moTa', w: 90 },
    ],
    cb
  );
  const wsC = wb.getWorksheet('Cảnh báo');
  for (let i = 0; i < cb.length; i++) {
    wsC.getRow(i + 2).getCell(3).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: cb[i].code === 'CCCD_TRUNG' ? 'FFF4CCCC' : 'FFFFF2CC' },
    };
  }

  addSheet(
    wb,
    'Ghi chú',
    [
      { h: 'Dòng Excel', k: 'row' },
      { h: 'Họ tên', k: 'ten', w: 26 },
      { h: 'Mã', k: 'code', w: 30 },
      { h: 'Mô tả', k: 'moTa', w: 90 },
    ],
    gc
  );

  // ---------- 5. Danh sach tra cuu ----------
  addSheet(
    wb,
    `DS ${doc.ap || 'tra'} (tra)`,
    [
      { h: 'TT', k: 'stt' },
      { h: 'Họ tên', k: 'hoTen', w: 26 },
      { h: 'Quan hệ', k: 'quanHe', w: 13 },
      { h: 'Số hộ', k: 'hoTraStt' },
      { h: 'Năm sinh', k: 'namSinh' },
      { h: 'CCCD', k: 'cccd', w: 16 },
      { h: 'Số định danh', k: 'soDinhDanh', w: 14 },
    ],
    lookup.map((p) => ({
      stt: p.stt,
      hoTen: p.hoTen,
      quanHe: p.quanHe,
      hoTraStt: p.hoTraStt,
      namSinh: p.ngaySinh ? p.ngaySinh.slice(0, 4) : '',
      cccd: p.cccd,
      soDinhDanh: p.soDinhDanh,
    }))
  );

  // ---------- 6. Ghi chu quy tac ----------
  const wsN2 = wb.addWorksheet('Quy tắc');
  wsN2.columns = [{ width: 26 }, { width: 100 }];
  wsN2.addRow(['Mục', 'Nội dung']);
  const rules = [
    ['Cột A', doc.quyTac?.cotA || ''],
    ['STT người', 'Đánh lại liên tục 1..' + doc.tongNguoi + ', theo thứ tự hộ rồi tới dòng trong hộ.'],
    ['Khôi phục', doc.quyTac?.khoiPhuc || ''],
    ['Cột A rác', 'Dấu "." hoặc chữ trong cột A: giữ nguyên, tách thành nhóm riêng, ghi cảnh báo + gợi ý. Không tự gộp.'],
    ['CCCD trùng', 'Giữ toàn bộ các dòng, chỉ ghi cảnh báo. Không tự khử trùng.'],
    ['Xoá dòng', 'Không xoá dòng nào. Mọi sai lệch đều được ghi lại.'],
    ['Màu nền', 'Xanh = đã sửa · Vàng = chưa sửa · Tím = chưa cập nhật · Đỏ = lỗi cột A'],
  ];
  for (const [a, b] of rules) wsN2.addRow([a, b]);
  const r2 = wsN2.getRow(1);
  for (let i = 1; i <= 2; i++) {
    r2.getCell(i).fill = HEADER_FILL;
    r2.getCell(i).font = HEADER_FONT;
  }

  const dest = path.join(outDir, `Danh sách Tổ ${doc.to} - đã gom.xlsx`);
  await X.writeXlsx(wb, dest, 'file Excel');
  console.log(`Da tao: ${dest}`);
  console.log(`  ${wb.worksheets.map((w) => `${w.name} (${w.rowCount - 1})`).join(', ')}`);
}

main().catch((e) => {
  console.error('LOI:', e.message);
  process.exit(1);
});
