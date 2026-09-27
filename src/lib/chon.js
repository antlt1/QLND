'use strict';
/**
 * Chon to + chon danh sach tra, va nho lai de cac lenh sau dung chung.
 *
 * Thu tu uu tien:
 *   1. co so truyen tren dong lenh   --to 2 --tra traA
 *   2. lua chon da luu               output\.chon.json   (xem `qlnd.cmd chon`)
 *   3. quy tac mac dinh             to 1,4 -> Truong Binh | to 2 -> Truong Binh A
 */
const fs = require('fs');
const path = require('path');
const readline = require('readline');
const { locate, DEFAULTS, TRA_THEO_TO } = require('./sources');

const TEN_CHON = '.chon.json';

/** tach --ten giatri tu argv (hoac --ten=value) */
function opts(args) {
  const o = {};
  const rest = [];
  for (let i = 0; i < args.length; i++) {
    const a = args[i];
    if (a.startsWith('--')) {
      const eq = a.indexOf('=');
      if (eq > 0) o[a.slice(2, eq)] = a.slice(eq + 1);
      else if (i + 1 < args.length && !args[i + 1].startsWith('--')) o[a.slice(2)] = args[++i];
      else o[a.slice(2)] = true;
    } else rest.push(a);
  }
  return { o, rest };
}

const duongDanhChon = (outDir) => path.join(outDir, TEN_CHON);

function docChon(outDir = DEFAULTS.outDir) {
  try {
    return JSON.parse(fs.readFileSync(duongDanhChon(outDir), 'utf8'));
  } catch {
    return null;
  }
}

function luuChon(chon, outDir = DEFAULTS.outDir) {
  fs.mkdirSync(outDir, { recursive: true });
  fs.writeFileSync(duongDanhChon(outDir), JSON.stringify(chon, null, 2), 'utf8');
  return duongDanhChon(outDir);
}

function xoaChon(outDir = DEFAULTS.outDir) {
  try {
    fs.unlinkSync(duongDanhChon(outDir));
    return true;
  } catch {
    return false;
  }
}

/**
 * Giai quyet cau hinh chay lenh.
 * @returns { to, ap, loc, nhan, ... } loc = ket qua locate() cho dung to
 */
function resolve(argv = process.argv.slice(2)) {
  const { o } = opts(argv);
  const srcDir = path.resolve(o.src || DEFAULTS.srcDir);
  const outDir = path.resolve(o.out || DEFAULTS.outDir);
  const daChon = docChon(outDir);
  const to = parseInt(o.to || (daChon && daChon.to) || 1, 10);
  const traKhoa = o.tra || (daChon && daChon.tra) || TRA_THEO_TO[to] || 'tra';
  const loc = locate(srcDir, { to, tra: traKhoa });
  return {
    o,
    srcDir,
    outDir,
    to,
    traKhoa,
    loc,
    ap: o.ap || (daChon && daChon.ap) || loc.ap,
    daChon,
    nhan: `Tổ ${to}`,
  };
}

/** Kiem tra du dieu kien chay, nem loi ro rang neu thieu */
function assertChay(cfg) {
  const { loc, to } = cfg;
  if (!loc.toFile) throw new Error(`Khong tim thay file danh sach To ${to} trong ${cfg.srcDir}`);
  if (!loc.toDocDuoc)
    throw new Error(
      `"${loc.toFile.ten}" doc khong duoc (${loc.toFile.ole2 ? 'la file .xls cu OLE2' : 'khong phai .xlsx'}).\n` +
        `       Tool chi doc duoc .xlsx. Hay mo file do trong Excel roi "Save As" thanh .xlsx roi thu lai.`
    );
  if (!loc.traChon) throw new Error('Khong tim thay danh sach tra cuu trong ' + cfg.srcDir);
  if (!loc.traDocDuoc)
    throw new Error(`Danh sach tra cuu "${loc.traChon.ten}" doc khong duoc (OLE2 / khong phai .xlsx).`);
}

/** In thong tin nguon theo tung to, kem canh bao khi khong khoi phuc duoc */
function inNguon(cfg) {
  const { loc, to, ap, traKhoa } = cfg;
  console.log('Nguon:');
  console.log(`  To ${to}   :`, loc.toFile ? loc.toFile.ten : '(khong tim thay)');
  console.log(`  Tra cuu:`, loc.traChon ? loc.traChon.ten : '(khong tim thay)');
  console.log(`  -> ap "${ap}" (khoa ${traKhoa})`);
  const khac = loc.tras.filter((t) => t.path !== (loc.traChon && loc.traChon.path));
  if (khac.length) console.log('  (file tra khac, dung --tra de doi):', khac.map((k) => k.ten).join(' | '));
  if (!loc.coKhoiPhuc)
    console.log(
      `  ! To ${to}: CHUA xac dinh vung khoi phuc STT ho -> tat khoi phuc, giu nguyen vi tri cu (khong doan).`
    );
}

/** Hoi nguoi dung 1 cau hoi, tra ve gia tri da chon */
function hoi(cauHoi) {
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question(cauHoi, (a) => {
      rl.close();
      resolve(String(a).trim());
    });
  });
}

/** Menu chon tay: chon to, chon file tra, luu lai */
async function menu(argv = process.argv.slice(2)) {
  const { o } = opts(argv);
  const srcDir = path.resolve(o.src || DEFAULTS.srcDir);
  const outDir = path.resolve(o.out || DEFAULTS.outDir);
  const loc0 = locate(srcDir, { to: 1 });

  if (!loc0.tos.length) {
    console.log('Khong tim thay file danh sach To nao trong', srcDir);
    console.log('Ten file can chua "To <so>", vi du: "Danh sách Tổ 2.xlsx".');
    return null;
  }

  console.log('\n=== CHON DANH SACH ===\n');
  console.log('Cac to tim thay trong thu muc nguon:\n');
  const docDuoc = new Set(loc0.tos.filter((t) => t.docDuoc).map((t) => t.to));
  loc0.tos.forEach((t, i) => {
    const ghi = docDuoc.has(t.to) ? '' : '   <-- doc khong duoc, la .xls cu (OLE2)';
    console.log(`  ${i + 1}) To ${t.to}   ${t.ten}${ghi}`);
  });

  let to;
  for (;;) {
    const a = await hoi(`\nChon to (1-${loc0.tos.length}, Enter = ${loc0.tos[0].to}): `);
    if (!a) {
      to = loc0.tos[0].to;
      break;
    }
    const i = parseInt(a, 10) - 1;
    if (i >= 0 && i < loc0.tos.length) {
      to = loc0.tos[i].to;
      break;
    }
    console.log('  So khong hop le.');
  }

  const loc = locate(srcDir, { to });
  console.log(`\nTo ${to}:`);
  console.log(`  file danh sach : ${loc.toFile ? loc.toFile.ten : '(khong tim thay)'}`);
  if (!loc.toDocDuoc)
    console.log('  ! File nay doc khong duoc (OLE2). Hay "Save As" thanh .xlsx trong Excel roi chon lai.');

  console.log('\nDanh sach tra cuu dung cho to nay:\n');
  const quyTac = TRA_THEO_TO[to];
  loc.tras.forEach((t, i) => {
    const khoa = / a /.test(t.doc) ? 'traA' : 'tra';
    const laMacDinh = khoa === quyTac;
    console.log(`  ${i + 1}) [${khoa}] ${t.ten}${laMacDinh ? '   <-- mac dinh cho to nay' : ''}`);
  });
  if (quyTac) console.log(`\n  Quy tac: To ${to} dung danh sach "${loc.ap}".`);

  let traKhoa = quyTac || 'tra';
  for (;;) {
    const a = await hoi(`\nChon danh sach tra (1-${loc.tras.length}, Enter = mac dinh): `);
    if (!a) break;
    const i = parseInt(a, 10) - 1;
    if (i >= 0 && i < loc.tras.length) {
      traKhoa = / a /.test(loc.tras[i].doc) ? 'traA' : 'tra';
      break;
    }
    console.log('  So khong hop le.');
  }

  const chon = { to, tra: traKhoa, ap: locate(srcDir, { to, tra: traKhoa }).ap, chonLuc: new Date().toISOString() };
  const p = luuChon(chon, outDir);
  const tenTra = locate(srcDir, { to, tra: traKhoa }).traChon;
  console.log('\nDa luu lua chon:', p);
  console.log(`  To ${chon.to} + ${tenTra ? tenTra.ten : chon.tra} -> ap "${chon.ap}"`);
  console.log('  Cac lenh sau (build, xep, fixroute) se dung lua chon nay.');
  return chon;
}

module.exports = {
  opts,
  resolve,
  assertChay,
  inNguon,
  menu,
  hoi,
  docChon,
  luuChon,
  xoaChon,
  TEN_CHON,
};
