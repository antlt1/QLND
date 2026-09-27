'use strict';
/**
 * xep - sap xep lai danh sach theo tung ho: chu ho roi den thanh vien.
 *
 *   qlnd.cmd xep
 *   qlnd.cmd xep --to 2
 *   qlnd.cmd xep --to 4 --out D:\ketqua
 *
 * Xuat: <out>\Danh sách Tổ <n> - đã xếp.xlsx  + sheet "Nhat ký xếp" + "Tóm tắt"
 */
const path = require('path');
const { xep } = require('../lib/arrange');
const { resolve, assertChay, inNguon } = require('../lib/chon');

(async () => {
  const cfg = resolve(process.argv.slice(2));
  const { outDir, to, ap } = cfg;
  inNguon(cfg);
  assertChay(cfg);
  const { toFile, traChon, recoverRange } = cfg.loc;

  const outFile = path.join(outDir, `Danh sách Tổ ${to} - đã xếp.xlsx`);
  const r = await xep({
    srcFile: toFile.path,
    traFile: traChon.path,
    ap,
    to,
    recoverRange,
    outFile,
  });

  console.log('=== XEP LAI DANH SACH ===');
  console.log('  nguon  : ' + toFile.ten);
  console.log('  tra    : ' + traChon.ten);
  console.log('  ket qua: ' + r.outFile);
  console.log('');
  console.log('  so ho            : ' + r.tongHo);
  console.log('  so nguoi         : ' + r.tongNguoi);
  console.log('  dong giu nguyen  : ' + r.giu);
  console.log('  dong di chuyen   : ' + r.chuyen);
  console.log('  dong chua gan ho : ' + r.chuaXep.length + (r.chuaXep.length ? ' -> xep cuoi, o dong ' + r.chuaXep.join(', ') + ' cua file moi' : ''));
  console.log('  nguoi chua xac dinh ho: ' + r.nguoiChuaXacDinh + ' (cot A = "." , xep cuoi file, can nguoi duyet)');
  console.log('  o cong thuc -> gia tri: ' + r.dongCongThuc.length);
  console.log('');
  console.log('  Moi hộ deu co 1 dong "chu ho" roi den cac dong "thành vien".');
  console.log('  Cot A = STT hộ (chi dong chu hộ), cot B = STT nguoi chay tu 1.');
  console.log('  File goc khong bi dong vao.');
})().catch((e) => {
  console.error('[LOI] ' + e.message);
  process.exit(1);
});

