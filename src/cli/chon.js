'use strict';
/**
 * chon - chon danh sach can xu ly, nho lai cho cac lenh sau.
 *
 *   qlnd.cmd chon              chon tay tu ban phim
 *   qlnd.cmd chon --to 2       chon tinh, khong hoi
 *   qlnd.cmd chon --xoa       xoa lua chon da luu
 *   qlnd.cmd chon --chay       chon xong chay tiep build + xep
 */
const path = require('path');
const { opts, menu, resolve, assertChay, inNguon, docChon, xoaChon } = require('../lib/chon');

(async () => {
  const argv = process.argv.slice(2);
  const { o } = opts(argv);

  if (o.xoa) {
    console.log(xoaChon() ? 'Da xoa lua chon.' : 'Chua co lua chon nao de xoa.');
    return;
  }

  if (o.to) {
    // chi dinh ro tren dong lenh -> khong hoi, chi kiem tra lai
    const cfg = resolve(['--to', String(o.to), ...(o.tra ? ['--tra', String(o.tra)] : [])]);
    inNguon(cfg);
    assertChay(cfg);
    console.log('\nOK');
    return;
  }

  const chon = await menu(argv);
  if (!chon) return;

  if (o.chay) {
    console.log('\n=== CHAY TIEP ===\n');
    const cfg = resolve([]);
    inNguon(cfg);
    assertChay(cfg);
    console.log('');
    const { execFileSync } = require('child_process');
    const root = path.join(__dirname, '..', '..');
    const node = process.execPath;
    for (const [ten, script] of [
      ['build', 'build.js'],
      ['xep', 'xep.js'],
    ]) {
      console.log(`--- ${ten} ---`);
      execFileSync(node, [path.join(root, 'src', 'cli', script)], { stdio: 'inherit' });
    }
    return;
  }

  const daChon = docChon();
  if (daChon) {
    console.log(`\nLua chon hien tai: To ${daChon.to}, ap "${daChon.ap}"`);
  }
})().catch((e) => {
  console.error('\n[LOI] ' + e.message);
  process.exit(1);
});
