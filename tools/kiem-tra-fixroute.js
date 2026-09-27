'use strict';
/**
 * Kiem tra vong tron cua fix route:
 *   sua tay -> kiem tra lai -> luu file moi -> file GOC khong doi
 * Chay: node tools/kiem-tra-fixroute.js
 */
const fs = require('fs');
const path = require('path');
const http = require('http');
const { Grid } = require('../src/lib/grid');
const { locate, DEFAULTS } = require('../src/lib/sources');

const PORT = 5311;
const GOC = path.join(DEFAULTS.srcDir, 'Danh sách Tổ 1.xlsx');
let sai = 0;

function ok(cond, msg) {
  console.log(`  ${cond ? 'DONG' : 'SAI '}  ${msg}`);
  if (!cond) sai++;
}

function api(url, body) {
  return new Promise((resolve, reject) => {
    const u = new URL(url, 'http://x');
    const data = body ? JSON.stringify(body) : null;
    const req = http.request(
      { hostname: '127.0.0.1', port: PORT, path: u.pathname + u.search, method: data ? 'POST' : 'GET',
        headers: data ? { 'content-type': 'application/json', 'content-length': Buffer.byteLength(data) } : {} },
      (res) => { let b = ''; res.on('data', (c) => (b += c)); res.on('end', () => { try { resolve(JSON.parse(b)); } catch (e) { reject(e); } }); }
    );
    req.on('error', reject);
    if (data) req.write(data);
    req.end();
  });
}

(async () => {
  const hashGoc = require('crypto').createHash('sha256').update(fs.readFileSync(GOC)).digest('hex');
  const kichThuocBanDau = fs.statSync(GOC).size;
  console.log(`File goc: ${GOC}`);
  console.log(`  sha256 bat dau: ${hashGoc.slice(0, 16)}...  (${kichThuocBanDau} bytes)\n`);

  const { spawn } = require('child_process');
  const srv = spawn(process.execPath, [path.join(__dirname, '..', 'src', 'cli', 'serve.js'), '--port', String(PORT)], {
    stdio: 'ignore',
  });

  let du = null;
  for (let i = 0; i < 40; i++) {
    await new Promise((r) => setTimeout(r, 700));
    try { du = await api(`http://127.0.0.1:${PORT}/api/fix-route`); break; } catch {}
  }
  if (!du) { console.log('  SAI  khong ket noi duoc server'); srv.kill(); process.exit(1); }

  try {
    console.log('1. Nap du lieu dong 1..148');
    ok(du.tong === 148, `nhan duoc ${du.tong} dong (mong doi 148)`);
    ok(Object.keys(du.tieuDe || {}).length > 10, `co ${Object.keys(du.tieuDe || {}).length} cot tieu de`);
    ok(du.sai + du.nghi + du.ok + du.soTieuDe === 148, 'tong so dong khop thong ke');
    ok(du.buoc[0].row === 1 && du.buoc[0].laTieuDe, 'dong 1 la tieu de');
    ok(du.buoc[1].row === 2 && du.buoc[1].laTieuDe, 'dong 2 la tieu de');
    ok(du.buoc[2].row === 3, 'dong 3 la dong du lieu dau tien');
    ok(du.buoc[147].row === 148, `dong cuoi la ${du.buoc[147].row} (mong doi 148)`);
    const r3 = du.buoc[2].cells;
    ok(r3.C === 'Trần Công Lập', `C3 = "${r3.C}"`);
    ok(r3.F === '093079008311', `F3 = "${r3.F}"`);
    ok(r3.E === '01/01/1979', `E3 = "${r3.E}"`);

    console.log('\n2. Chi doc - khong duoc doi gi');
    ok(du.daSua === 0, `chua co thay doi nao (daSua=${du.daSua})`);

    console.log('\n3. Sua tay mot o (dung so, khong phai CCCD)');
    const truoc = du.buoc[2].cells.H;
    const r = await api(`http://127.0.0.1:${PORT}/api/fix-route/set`, { row: 3, col: 'H', value: '99' });
    ok(r.doi.doi === true, 'ghi nhan thay doi');
    ok(r.doi.giaTriCu === truoc, `gia tri cu giu nguyen: ${truoc} -> ${r.doi.giaTriCu}`);
    ok(r.kiemTra.cells.H === '99', `doc lai H3 = "${r.kiemTra.cells.H}"`);
    ok(r.kiemTra.nghi >= 1, 'tuoi 99 duoc canh bao "nen xem" so voi nam sinh');

    console.log('\n4. Sua CCCD phai duoc chuoi 12 so');
    const rs = await api(`http://127.0.0.1:${PORT}/api/fix-route/set`, { row: 3, col: 'F', value: '093079008311' });
    ok(rs.kiemTra.cells.F === '093079008311', `F3 = "${rs.kiemTra.cells.F}"`);
    ok(!rs.kiemTra.loi.some((l) => l.cot === 'F'), 'khong con loi o cot F');

    console.log('\n5. Tra ve gia tri cu');
    const rt = await api(`http://127.0.0.1:${PORT}/api/fix-route/set`, { row: 3, col: 'H', value: truoc });
    ok(rt.kiemTra.cells.H === truoc, `H3 da khoi phuc "${truoc}"`);

    console.log('\n6. Sua ngan (ngay o dang so thu tu Excel)');
    const d87 = (await api(`http://127.0.0.1:${PORT}/api/fix-route`)).buoc.find((b) => b.row === 87);
    ok(d87.cells.G === '46013', `G87 truoc khi sua = "${d87.cells.G}"`);
    await api(`http://127.0.0.1:${PORT}/api/fix-route/fix`, { from: 1, to: 148 });
    const s87 = (await api(`http://127.0.0.1:${PORT}/api/fix-route`)).buoc.find((b) => b.row === 87);
    ok(s87.cells.G === '22/12/2025', `G87 sau khi sua = "${s87.cells.G}"`);

    console.log('\n7. Khong duoc xoa dong nao');
    const d = await api(`http://127.0.0.1:${PORT}/api/fix-route`);
    ok(d.buoc.length === 148, `van con ${d.buoc.length} dong`);
    ok(d.buoc.every((b, i) => b.row === i + 1), 'dong lien tuc tu 1 den 148');

    console.log('\n8. Luu file moi');
    const sv = await api(`http://127.0.0.1:${PORT}/api/fix-route/save`, { from: 1, to: 148 });
    if (sv.loi) console.log(`  [server bao loi] ${sv.loi}`);
    ok(!sv.loi, `luu thanh cong: ${sv.file}`);
    ok(fs.existsSync(sv.file), 'file da tao tren dia');
    ok(fs.existsSync(sv.nhatKy), 'nhat ky da tao');
    const log = JSON.parse(fs.readFileSync(sv.nhatKy, 'utf8'));
    ok(Array.isArray(log), `nhat ky la mang, ${log.length} muc`);
    ok(log.every((l) => 'row' in l && 'col' in l && 'giaTriCu' in l), 'muc nhat ky co du cot doi chieu');

    console.log('\n9. Doc lai file da sua');
    const g2 = await Grid.open(sv.file, 0);
    ok(g2.ws.getRow(87).getCell(7).value instanceof Date, 'G87 la doi tuong Date that');
    ok(g2.read(87).cells.G === '22/12/2025', `G87 doc lai = "${g2.read(87).cells.G}"`);
    ok(g2.read(3).cells.C === 'Trần Công Lập', 'ten nguoi dong 3 giu nguyen');
    ok(g2.read(3).cells.F === '093079008311', 'CCCD dong 3 giu nguyen');
    ok(g2.maxRow >= 325, `file van con ${g2.maxRow} dong (khong bi cat)`);
    const t3 = g2.thongKe(1, 148);
    ok(t3.sai === 5, `van con ${t3.sai} loi canh bao (4 dau "." + dong 67 lech cot)`);

    console.log('\n10. File GOC khong bi thay doi');
    const hashSau = require('crypto').createHash('sha256').update(fs.readFileSync(GOC)).digest('hex');
    ok(hashSau === hashGoc, 'sha256 giong het truoc khi sua');
    ok(fs.statSync(GOC).size === kichThuocBanDau, `kich thuoc giong (${kichThuocBanDau} bytes)`);
  } catch (e) {
    console.log(`  SAI  loi khong mong doi: ${e.message}`);
    sai++;
  } finally {
    srv.kill();
  }

  console.log(`\n${sai === 0 ? 'PASS' : 'FAIL'} - ${sai} loi`);
  process.exit(sai === 0 ? 0 : 1);
})();
