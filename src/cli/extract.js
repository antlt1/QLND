'use strict';
/** CLI: doc Excel -> JSON trong output/
 *  qlnd.cmd extract            dung lua chon da luu (xem `qlnd.cmd chon`)
 *  qlnd.cmd extract --to 2     chay cho To 2
 *  qlnd.cmd extract --to 2 --tra traA
 */
const fs = require('fs');
const path = require('path');
const X = require('../lib/xlsx');
const to1 = require('../lib/to1');
const { scanUnreadable } = require('../lib/sources');
const { resolve, assertChay, inNguon } = require('../lib/chon');

async function main() {
  const cfg = resolve(process.argv.slice(2));
  const { outDir, to, ap } = cfg;
  inNguon(cfg);
  assertChay(cfg);
  const { toFile, traChon, recoverRange, coKhoiPhuc } = cfg.loc;

  const unread = await scanUnreadable(cfg.srcDir);
  if (unread.length) {
    console.log('\nBo qua (may doc khong duoc):');
    for (const u of unread) console.log(`  - ${u.file}: ${u.lyDo}`);
  }

  console.log('\nDang doc...');
  const a = await X.readSheet(toFile.path, 0);
  const b = await X.readSheet(traChon.path, 0);
  console.log(`  "${a.sheet}": ${a.rows.length} dong co du lieu`);
  console.log(`  "${b.sheet}": ${b.rows.length} dong co du lieu`);

  const d = to1.build({
    to1Rows: a.rows,
    lookupRows: b.rows,
    ap,
    to,
    recoverRange,
  });

  fs.mkdirSync(outDir, { recursive: true });
  const doc = {
    ap: d.ap,
    to: d.to,
    nguon: { toFile: toFile.ten, traFile: traChon.ten },
    xuatAt: new Date().toISOString(),
    tongNguoi: d.tongNguoi,
    tongHo: d.tongHo,
    ghiChu:
      'Cot A = STT ho (chi ghi o dong chu ho). STT nguoi duoc danh lai lien tuc 1..N. Khong xoa dong nao.',
    quyTac: {
      cotA: 'STT hộ, chỉ ghi ở dòng chủ hộ',
      khoiPhuc: coKhoiPhuc
        ? `Chỉ dòng ${recoverRange.from}–${recoverRange.to} của tổ ${to} được gán lại STT hộ bằng cách đối chiếu CCCD/tên với danh sách ${ap}. Ngoài phạm vi: giữ nguyên.`
        : `Tổ ${to}: CHƯA xác định vùng chồn dòng nên KHÔNG khôi phục STT hộ — giữ nguyên vị trí cũ, không đoán.`,
      rac: 'Dấu "." hoặc chữ trong cột A: giữ nguyên, tách thành nhóm riêng kèm gợi ý, không tự gộp.',
    },
    thongKe: {
      match: d.match,
      recover: d.recover,
      recoverRange: d.recoverRange,
      houseMapSize: d.houseMapSize,
      houseMapAmbiguous: d.houseMapAmbiguous,
      houseMapLog: d.houseMapLog,
      byKhung: d.byKhung,
      soCanhBao: d.issues.filter((i) => (i.muc || 'canh-bao') === 'canh-bao').length,
      soGhiChu: d.issues.filter((i) => i.muc === 'ghi-chu').length,
    },
    ho: d.households.map((h) => ({
      id: h.id,
      no: h.no,
      sttHo: h.sttHo,
      cotARaw: h.cotARaw,
      khung: h.khung,
      chuHo: h.chuHo,
      rowMin: h.rowMin,
      rowMax: h.rowMax,
      soNguoi: h.soNguoi,
      canhBao: h.canhBao,
      goiY: h.goiY,
      thanhVien: h.members.map((m) => ({
        sttNguoi: m.sttNguoi,
        row: m.row,
        block: m.block,
        hoTen: m.hoTen,
        quanHe: m.quanHe,
        vaiTro: m.vaiTro,
        maBH: m.maBH,
        khongBH: m.khongBH,
        ngaySinh: m.ngaySinh,
        cccd: m.cccd,
        ngayCap: m.ngayCap,
        tuoi: m.tuoi,
        dienThoai: m.dienThoai,
        duHoc: m.duHoc,
        layChong: m.layChong,
        xkld: m.xkld,
        ngheNghiep: m.ngheNghiep,
        ghiChu: m.ghiChu,
        cotARaw: m.cotARaw,
        hoGoc: m.hoGoc,
        hoTra: m.hoTra,
        matchBy: m.matchBy,
        nguonHo: m.nguonHo,
        canhBao: m.canhBao,
        datDai: m.datDai,
      })),
    })),
  };

  const write = (name, data) => {
    const p = path.join(outDir, name);
    fs.writeFileSync(p, JSON.stringify(data, null, 2), 'utf8');
    console.log(`  ghi ${name} (${(fs.statSync(p).size / 1024).toFixed(0)} KB)`);
  };

  console.log('\nKet qua:');
  write(`to${to}-ho.json`, doc);
  write(`to${to}-nguoi.json`, d.people.map((m) => ({
    sttNguoi: m.sttNguoi, sttHo: m.sttHo, row: m.row, hoTen: m.hoTen,
    quanHe: m.quanHe, vaiTro: m.vaiTro, maBH: m.maBH, ngaySinh: m.ngaySinh,
    cccd: m.cccd, ngayCap: m.ngayCap, tuoi: m.tuoi, dienThoai: m.dienThoai,
    ngheNghiep: m.ngheNghiep, nguonHo: m.nguonHo, canhBao: m.canhBao, datDai: m.datDai,
  })));
  write(`to${to}-loi.json`, d.issues);
  write(`tra-cuoc.json`, d.lookup);

  const nCb = doc.thongKe.soCanhBao;
  const nGc = doc.thongKe.soGhiChu;
  console.log(`\n  ${d.tongHo} ho / ${d.tongNguoi} nguoi`);
  console.log(`  ${nCb} canh bao, ${nGc} ghi chu`);
  console.log(`  khop CCCD ${d.match.cccd}, theo ten ${d.match.ten}, khong khop ${d.match.khongKhop}`);
  const rr = d.recoverRange ? ` (dong ${d.recoverRange.from}-${d.recoverRange.to})` : ' (toan bo)';
  console.log(`  khoi phuc STT ho${rr}: ${d.recover.khoiPhuc} dong, giu nguyen ${d.recover.giuNguyen} dong`);
  console.log(`  bang do ho: ${d.houseMapSize} hop le, ${d.houseMapAmbiguous} xung dot`);
  for (const [k, v] of Object.entries(d.byKhung)) {
    console.log(`    ${k.padEnd(34)} ${String(v.soHo).padStart(4)} ho  ${String(v.soNguoi).padStart(4)} nguoi`);
  }
}

main().catch((e) => {
  console.error('\nLOI:', e.message);
  process.exit(1);
});
