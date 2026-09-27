'use strict';
/** In chi tiet khoi du lieu theo STT ho, kem nguon (dong file goc) */
const fs = require('fs');
const path = require('path');
const { DEFAULTS } = require('../src/lib/sources');

const doc = JSON.parse(fs.readFileSync(path.join(DEFAULTS.outDir, 'to1-ho.json'), 'utf8'));
const args = process.argv.slice(2).map(Number);
const tu = args[0] || 25;
const den = args[1] || 32;

for (const h of doc.ho) {
  if (h.sttHo < tu || h.sttHo > den) continue;
  console.log(`\n=== HO ${h.sttHo}  (${h.khung})  ${h.soNguoi} nguoi  dong ${h.rowMin}..${h.rowMax} ===`);
  if (h.chuHo) {
    console.log(`  chu ho: ${h.chuHo.hoTen}  cccd=${h.chuHo.cccd}  nguon=${h.chuHo.nguonHo}  row=${h.chuHo.row}`);
  } else {
    console.log('  chu ho: (khong co)');
  }
  for (const m of h.thanhVien) {
    console.log(
      `   - ${String(m.sttNguoi).padStart(3)} ${String(m.hoTen).padEnd(24)} cccd=${String(m.cccd).padEnd(14)} ${String(m.vaiTro).padEnd(10)} row=${String(m.row).padStart(4)} nguon=${m.nguonHo}`
    );
  }
  if (h.canhBao && h.canhBao.length) console.log('  canh bao:', h.canhBao.join('; '));
  if (h.goiY) console.log('  goi y   :', h.goiY);
}
