'use strict';
/** In ra cac to va danh sach tra theo tung to, de kiem tra logic chon */
const s = require('../src/lib/sources');

console.log('=== CAC TO TIM THAY ===');
for (const t of s.listTos()) {
  const trang = t.docDuoc ? 'DOC DUOC      ' : 'DOC KHONG DUOC';
  const bo = t.khac.length ? '   (bo qua: ' + t.khac.map((k) => k.ten).join(', ') + ')' : '';
  console.log(`  To ${t.to}  ${trang}  ${t.ten}${bo}`);
}

console.log('\n=== FILE TRA ===');
for (const t of s.listTra()) {
  console.log(`  [${/ a /.test(t.doc) ? 'traA' : 'tra '}] ${t.ten}`);
}

console.log('\n=== locate() TUNG TO ===');
for (const n of [1, 2, 3, 4, 5]) {
  const l = s.locate(undefined, { to: n });
  const file = l.toFile ? l.toFile.ten.padEnd(32) : 'KHONG TIM THAY'.padEnd(32);
  const tra = l.traChon ? l.traChon.ten.padEnd(44) : 'KHONG TIM THAY';
  const kp = l.coKhoiPhuc ? `dòng ${l.recoverRange.from}-${l.recoverRange.to}` : 'TAT (khong doan)';
  console.log(`  To ${n}: ${file}${tra}ap=${l.ap.padEnd(14)}khoiPhuc=${kp}`);
}
