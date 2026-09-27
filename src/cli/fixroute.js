'use strict';
/**
 * FIX ROUTE - duyet tung dong mot, tu dong `from` den `to` (mac dinh 1..148).
 *
 *   qlnd.cmd fixroute              chi kiem tra, khong sua
 *   qlnd.cmd fixroute --sua        sua nhung gi chan chan
 *   qlnd.cmd fixroute --from 1 --to 148 --sua
 */
const fs = require('fs');
const path = require('path');
const { Grid } = require('../lib/grid');
const { locate, DEFAULTS } = require('../lib/sources');

const args = process.argv.slice(2);
const has = (n) => args.includes(`--${n}`);
const opt = (n, d) => {
  const i = args.indexOf(`--${n}`);
  return i >= 0 && args[i + 1] ? args[i + 1] : d;
};

const FROM = parseInt(opt('from', '1'), 10);
const TO = parseInt(opt('to', '148'), 10);
const SUA = has('sua');

function pad(s, n) {
  s = String(s ?? '');
  return s.length > n ? s.slice(0, n - 1) + '~' : s.padEnd(n);
}

async function main() {
  const srcDir = path.resolve(opt('src', DEFAULTS.srcDir));
  const outDir = path.resolve(opt('out', DEFAULTS.outDir));
  const files = locate(srcDir);
  if (!files.to1) throw new Error(`Khong tim thay file To 1 trong ${srcDir}`);

  console.log(`FIX ROUTE  ${path.basename(files.to1)}`);

  const g = await Grid.open(files.to1, 0);
  console.log(`  sheet: "${g.ws.name}"  dong ${FROM}..${TO}  che do: ${SUA ? 'SUA' : 'CHI KIEM TRA'}`);
  console.log('  tieu de:');
  const hdr = g.headerRows(2).find((h) => Object.keys(h).length) || {};
  for (const [c, t] of Object.entries(hdr)) console.log(`    ${c}: ${t}`);
  console.log('');

  let buoc;
  if (SUA) {
    buoc = g.fixRange(FROM, TO);
    console.log(`  da sua ${g.log.length} o:\n`);
    for (const l of g.log) {
      console.log(`    dong ${String(l.row).padStart(3)}  ${l.col}  "${l.giaTriCu}" -> "${l.giaTriMoi}"   (${l.lyDo})`);
    }
    if (!g.log.length) console.log('    (khong co gi can sua)');
    console.log('');
    buoc = g.scan(FROM, TO);
  } else {
    buoc = g.scan(FROM, TO);
  }

  // in tung dong
  let sai = 0,
    nghi = 0,
    on = 0;
  console.log(`  ${'dong'.padEnd(5)}${'A'.padEnd(5)}${'nguoi'.padEnd(7)}${'ho ten'.padEnd(24)}${'ngay sinh'.padEnd(12)}${'cccd'.padEnd(14)}${'tuoi'.padEnd(5)}${'quan he'.padEnd(12)}trang thai`);
  console.log('  ' + '-'.repeat(104));
  for (const b of buoc) {
    const c = b.cells;
    const st = b.laTieuDe ? 'tieu de' : b.sai ? `${b.sai} LOI` : b.nghi ? `${b.nghi} nen xem` : 'ok';
    if (b.laTieuDe) { on++; }
    else if (b.sai) sai++;
    else if (b.nghi) nghi++;
    else on++;
    console.log(
      `  ${pad(b.row, 5)}${pad(JSON.stringify(c.A || ''), 5)}${pad(c.B, 7)}${pad(c.C, 24)}${pad(c.E, 12)}${pad(c.F, 14)}${pad(c.H, 5)}${pad(c.I, 12)}${st}`
    );
  }
  console.log('  ' + '-'.repeat(104));
  console.log(`  ${buoc.length} dong:  ${on} ok   ${nghi} nen xem   ${sai} co loi\n`);

  // chi tiet loi
  const coLoi = buoc.filter((b) => b.loi.length);
  if (coLoi.length) {
    console.log('  CHI TIET');
    for (const b of coLoi) {
      for (const l of b.loi) {
        console.log(`    dong ${String(b.row).padStart(3)} [${l.cot}] ${l.muc === 'sai' ? 'LOI ' : 'NEN '}${l.loi}${l.goiY ? '  -> ' + l.goiY : ''}`);
      }
    }
    console.log('');
  }

  if (SUA) {
    fs.mkdirSync(outDir, { recursive: true });
    const dest = path.join(outDir, `Danh sách Tổ 1 - đã sửa dòng ${FROM}-${TO}.xlsx`);
    await g.save(dest);
    console.log(`  Da luu: ${dest}`);

    const logFile = path.join(outDir, `fixroute-${FROM}-${TO}.json`);
    fs.writeFileSync(logFile, JSON.stringify(g.log, null, 2), 'utf8');
    console.log(`  Nhat ky: ${logFile}  (${g.log.length} thay doi)`);
    console.log('  File goc KHONG bi thay doi.');
  }
}

main().catch((e) => {
  console.error('LOI:', e.message);
  process.exit(1);
});
