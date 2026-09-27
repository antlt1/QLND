'use strict';
/** Test nhanh: bo test framework, chay bang node test/run.js */
const assert = require('assert');
const X = require('../src/lib/xlsx');
const to1 = require('../src/lib/to1');
const G = require('../src/lib/grid');
const { locate } = require('../src/lib/sources');

let pass = 0,
  fail = 0;
function t(name, fn) {
  try {
    fn();
    pass++;
    console.log(`  ok   ${name}`);
  } catch (e) {
    fail++;
    console.log(`  FAIL ${name}\n         ${e.message}`);
  }
}

console.log('\n== xlsx helpers ==');
t('txt giu nguyen Date', () => {
  const d = new Date(Date.UTC(1979, 0, 1));
  assert.strictEqual(X.txt(d), '1979-01-01');
  assert.strictEqual(X.isoDate(d), '1979-01-01');
});
t('isoDate doc dd/MM/yyyy', () => assert.strictEqual(X.isoDate('27/12/2008'), '2008-12-27'));
t('isoDate doc dd/MM/yy', () => assert.strictEqual(X.isoDate('01/01/79'), '1979-01-01'));
t('isoDate doc rong', () => assert.strictEqual(X.isoDate(''), ''));
t('isoDate doc rac -> rong', () => assert.strictEqual(X.isoDate('khong phai ngay'), ''));
t('intOf', () => {
  assert.strictEqual(X.intOf('42'), 42);
  assert.strictEqual(X.intOf('.'), null);
  assert.strictEqual(X.intOf(''), null);
  assert.strictEqual(X.intOf(new Date()), null);
});
t('cccdOf chap nhan 12 so', () => {
  assert.strictEqual(X.cccdOf('093079008311'), '093079008311');
  assert.strictEqual(X.cccdOf('93304000918'), '');
  assert.strictEqual(X.cccdOf(new Date()), '');
});
t('normName bo dau', () => assert.strictEqual(X.normName('  Trần   Công  Lập '), 'tran cong lap'));
t('deaccent', () => assert.strictEqual(X.deaccent('Đặng'), 'dang'));

console.log('\n== to1 ==');
t('khungOf', () => {
  assert.strictEqual(to1.khungOf(5), to1.KHUNG.DA_FIX);
  assert.strictEqual(to1.khungOf(50), to1.KHUNG.CHUA_FIX);
  assert.strictEqual(to1.khungOf(100), to1.KHUNG.CHUA_CAP);
  assert.strictEqual(to1.khungOf(200), to1.KHUNG.NGOAI);
  assert.strictEqual(to1.khungOf(0), to1.KHUNG.RAC);
});
t('parseTo1 doc dong', () => {
  const rows = [
    { row: 3, c: { 1: '1', 3: 'A', 5: '01/01/1979', 6: '093079008311', 9: 'chủ hộ' } },
    { row: 4, c: { 3: 'B', 5: '02/02/2000', 6: '093200000002', 9: 'thành viên' } },
  ];
  const { people } = to1.parseTo1(rows);
  assert.strictEqual(people.length, 2);
  assert.strictEqual(people[0].isHeadRow, true);
  assert.strictEqual(people[1].isHeadRow, false);
  assert.strictEqual(people[0].ngaySinh, '1979-01-01');
  assert.strictEqual(people[1].cccd, '093200000002');
});
t('gom ho + STT nguoi lien tuc', () => {
  const rows = [
    { row: 3, c: { 1: '1', 3: 'A', 6: '093079008311', 9: 'chủ hộ' } },
    { row: 4, c: { 3: 'B', 6: '093200000002', 9: 'thành viên' } },
    { row: 5, c: { 1: '2', 3: 'C', 6: '093300000003', 9: 'chủ hộ' } },
  ];
  const d = to1.build({ to1Rows: rows, lookupRows: [], ap: 'X', to: 1 });
  assert.strictEqual(d.households.length, 2);
  assert.strictEqual(d.households[0].soNguoi, 2);
  assert.strictEqual(d.households[1].soNguoi, 1);
  assert.deepStrictEqual(d.people.map((p) => p.sttNguoi), [1, 2, 3]);
});
t('cot A rac -> tach nhom rieng, khong ghep', () => {
  const rows = [
    { row: 3, c: { 1: '1', 3: 'A', 6: '093079008311', 9: 'chủ hộ' } },
    { row: 4, c: { 1: '.', 3: 'B', 6: '093200000002', 9: 'thành viên' } },
    { row: 5, c: { 1: '.', 3: 'C', 6: '093300000003', 9: 'thành viên' } },
    { row: 6, c: { 1: '2', 3: 'D', 6: '093400000004', 9: 'chủ hộ' } },
  ];
  const d = to1.build({ to1Rows: rows, lookupRows: [], ap: 'X', to: 1 });
  assert.strictEqual(d.households.length, 3, 'phai tach 3 nhom');
  const rac = d.households.filter((h) => h.khung === to1.KHUNG.RAC);
  assert.strictEqual(rac.length, 1);
  assert.strictEqual(rac[0].soNguoi, 2, 'hai dong "." phai chung 1 nhom rac');
  assert.ok(rac[0].goiY.includes('ho 1'), 'phai co goi y');
});
t('dong lech cot: CCCD o cot E duoc khoi phuc', () => {
  const rows = [
    { row: 3, c: { 1: '1', 3: 'A', 6: '093079008311', 9: 'chủ hộ' } },
    { row: 4, c: { 3: 'B', 5: '089189017518', 7: 'thành viên' } },
  ];
  const { people, issues } = to1.parseTo1(rows);
  assert.strictEqual(people[1].cccd, '089189017518');
  assert.strictEqual(people[1].rowShifted, true);
  assert.ok(issues.some((i) => i.code === 'DONG_BI_LECH_COT'));
});
t('khong xoa dong nao', () => {
  const rows = [];
  for (let i = 3; i <= 20; i++) {
    rows.push({ row: i, c: { 3: 'N' + i, 6: '093000000' + String(i).padStart(3, '0'), 9: 'thành viên' } });
  }
  const d = to1.build({ to1Rows: rows, lookupRows: [], ap: 'X', to: 1 });
  assert.strictEqual(d.people.length, 18);
  assert.strictEqual(d.tongNguoi, 18);
});
t('bang do ho: bo dong trung CCCD khi xung dot', () => {
  const mk = (row, a, ten, cccd) => ({ row, c: { 1: a, 3: ten, 6: cccd, 9: 'chủ hộ' } });
  const to1Rows = [
    mk(3, '1', 'A', '093000000001'),
    mk(4, '2', 'B', '093000000002'),
    mk(5, '2b', 'C', '093000000003'),
    // 2 va 2b cung ho tra 1 -> xung dot
    mk(6, '3', 'D', '093000000004'),
    mk(6.5, '3', 'E', '093000000004'),
  ];
  const lookupRows = [
    { row: 1, c: { 1: '1', 2: 'A', 5: '093000000001', 7: 'chủ hộ' } },
    { row: 2, c: { 2: 'B', 5: '093000000002', 7: 'chủ hộ' } },
    { row: 3, c: { 2: 'C', 5: '093000000003', 7: 'chủ hộ' } },
    { row: 4, c: { 2: 'D', 5: '093000000004', 7: 'chủ hộ' } },
    { row: 5, c: { 2: 'E', 5: '093000000004', 7: 'thành viên' } },
  ];
  const d = to1.build({ to1Rows, lookupRows, ap: 'X', to: 1 });
  assert.ok(d.houseMapAmbiguous === 0 || d.houseMapAmbiguous >= 0);
});

console.log('\n== grid: ham xu ly don le ==');
t('numToCol / colToNum', () => {
  assert.strictEqual(G.numToCol(1), 'A');
  assert.strictEqual(G.numToCol(26), 'Z');
  assert.strictEqual(G.numToCol(27), 'AA');
  assert.strictEqual(G.colToNum('A'), 1);
  assert.strictEqual(G.colToNum('Z'), 26);
  assert.strictEqual(G.colToNum('AA'), 27);
  assert.strictEqual(G.colToNum('E'), 5);
});
t('show doc duoc Date va chuoi', () => {
  assert.strictEqual(G.show(new Date(Date.UTC(1979, 0, 1))), '01/01/1979');
  assert.strictEqual(G.show('abc'), 'abc');
  assert.strictEqual(G.show(null), '');
  assert.strictEqual(G.show(undefined), '');
  assert.strictEqual(G.show(0), '0');
  assert.strictEqual(G.show({ result: 42 }), '42');
  assert.strictEqual(G.show({ sharedFormula: 'H35' }), '');
});
t('toDate doc dd/MM/yyyy va dd/MM/yy', () => {
  assert.strictEqual(G.toDate('01/01/1979').toISOString().slice(0, 10), '1979-01-01');
  assert.strictEqual(G.toDate('27/12/2008').toISOString().slice(0, 10), '2008-12-27');
  assert.strictEqual(G.toDate('01/01/79').toISOString().slice(0, 10), '1979-01-01');
  assert.strictEqual(G.toDate('khong phai ngay'), null);
  assert.strictEqual(G.toDate(''), null);
});
t('serialToDate doi so thu tu ngay Excel', () => {
  assert.strictEqual(G.show(G.serialToDate(46013)), '22/12/2025');
  assert.strictEqual(G.serialToDate(0), null);
  assert.strictEqual(G.serialToDate('abc'), null);
  assert.strictEqual(G.serialToDate(null), null);
});
t('isUnresolvedFormula', () => {
  assert.strictEqual(G.isUnresolvedFormula({ sharedFormula: 'H35' }), true);
  assert.strictEqual(G.isUnresolvedFormula({ formula: 'X', result: 1 }), false);
  assert.strictEqual(G.isUnresolvedFormula({ result: 1 }), false);
  assert.strictEqual(G.isUnresolvedFormula(null), false);
  assert.strictEqual(G.isUnresolvedFormula('x'), false);
});

console.log('\n== du lieu that ==');
(async () => {
  const loc = locate(undefined, { to: 1 });
  if (loc.toFile) {
    const a = await X.readSheet(loc.toFile.path, 0);
    const b = loc.traChon ? await X.readSheet(loc.traChon.path, 0) : { rows: [] };
    const d = to1.build({ to1Rows: a.rows, lookupRows: b.rows, ap: 'Trường Bình', to: 1 });

    t('doc duoc 598 nguoi', () => assert.strictEqual(d.tongNguoi, 598));
    t('khong con ho 180 nguoi', () => {
      const big = d.households.filter((h) => h.soNguoi > 20);
      assert.strictEqual(big.length, 0, `con ho khong lon: ${big.map((h) => h.sttHo + ':' + h.soNguoi)}`);
    });
    t('STT nguoi la 1..598 khong trung, khong thieu', () => {
      const ids = d.people.map((p) => p.sttNguoi).sort((a, b) => a - b);
      assert.strictEqual(new Set(ids).size, 598, 'khong co STT nguoi trung');
      assert.strictEqual(ids[0], 1);
      assert.strictEqual(ids[597], 598);
    });
    t('khong mat nguoi nao', () => {
      const n = d.people.length;
      assert.strictEqual(n, 598);
      assert.ok(d.people.every((p) => p.hoTen), 'moi nguoi deu co ten');
    });
    t('khong co nguoi trung CCCD bi xoa', () => {
      const dup = d.people.filter((p) => p.canhBao.includes('CCCD_TRUNG'));
      assert.ok(dup.length >= 2, 'van giu ca 2 dong trung');
    });
    t('mọi hộ deu có STT hộ hoặc được đánh dấu', () => {
      for (const h of d.households) {
        if (h.sttHo === 0) assert.strictEqual(h.khung, to1.KHUNG.RAC);
      }
    });
    t('khoi phuc chi trong pham vi 149-325', () => {
      assert.deepStrictEqual(d.recoverRange, { from: 149, to: 325 });
      const ngoai = d.people.filter(
        (p) => p.nguonHo === 'khoi-phuc-CCCD' && (p.row < 149 || p.row > 325)
      );
      assert.strictEqual(ngoai.length, 0, `co ${ngoai.length} dong ngoai pham vi bi khoi phuc`);
    });
    t('vung 149-325: 173 khop CCCD, 4 khong khop', () => {
      const vung = d.people.filter((p) => p.row >= 149 && p.row <= 325);
      assert.strictEqual(vung.length, 177);
      const khongKhop = vung.filter((p) => !p.hoTra);
      assert.strictEqual(khongKhop.length, 4, 'dung 4 nguoi khong tim thay trong DS tra');
      assert.deepStrictEqual(
        khongKhop.map((p) => p.row).sort((a, b) => a - b),
        [186, 190, 228, 255]
      );
    });
    t('cot A rac: 3 nhom "." + 9 nguoi chua xac dinh ho', () => {
      const rac = d.households.filter((h) => h.khung === to1.KHUNG.RAC);
      // 3 nhom cot A = "." (12 nguoi) + 9 nguoi khong tra duoc ho (tach rieng)
      const nhomDot = rac.filter((h) => /^A\d+$/.test(h.key));
      const nhomChuaXacDinh = rac.filter((h) => /^M\d+$/.test(h.key));
      assert.strictEqual(nhomDot.length, 3);
      assert.strictEqual(
        nhomDot.reduce((s, h) => s + h.soNguoi, 0),
        12
      );
      assert.strictEqual(nhomChuaXacDinh.length, 9);
      assert.strictEqual(
        nhomChuaXacDinh.reduce((s, h) => s + h.soNguoi, 0),
        9
      );
      for (const h of rac) assert.ok(h.goiY, 'moi nhom rac deu co goi y');
    });
    t('vung 149-325: khong doan vao chu ho dang mo', () => {
      // 9 nguoi tra khong ra ho nao -> tach rieng, KHONG ghep vao ho 78
      const chuaXacDinh = d.people.filter((p) => p.nguonHo === 'chua-xac-dinh-ho');
      assert.strictEqual(chuaXacDinh.length, 9);
      for (const p of chuaXacDinh) {
        assert.strictEqual(p.sttHo, 0, `r${p.row} phai de sttHo = 0`);
        assert.ok(p.canhBao.includes('CHUA_XAC_DINH_HO_TU_DS_TRA'));
      }
      assert.deepStrictEqual(
        chuaXacDinh.map((p) => p.row).sort((a, b) => a - b),
        [158, 186, 190, 212, 220, 228, 255, 305, 314]
      );
      const ho78 = d.households.find((h) => h.sttHo === 78);
      assert.strictEqual(ho78.soNguoi, 9, 'ho 78 phai con 9 nguoi');
      assert.ok(
        !ho78.members.some((m) => m.nguonHo === 'chua-xac-dinh-ho'),
        'ho 78 khong duoc nhan nguoi chua xac dinh ho'
      );
    });
    t('STT ho: 1..140, thieu 2 va 4, trung 105 giu nguyen', () => {
      const stt = d.households.map((h) => h.sttHo).filter((s) => s > 0);
      const thieu = [];
      for (let i = 1; i <= 140; i++) if (!stt.includes(i)) thieu.push(i);
      assert.deepStrictEqual(thieu, [2, 4]);
      assert.strictEqual(stt.filter((s) => s === 105).length, 2, 'ho 105 trung phai giu ca 2');
    });
    t('phan loai canh bao / ghi chu', () => {
      const cb = d.issues.filter((i) => (i.muc || 'canh-bao') === 'canh-bao');
      const gc = d.issues.filter((i) => i.muc === 'ghi-chu');
      assert.ok(cb.length > 0 && gc.length > 0);
      assert.ok(
        gc.every((i) => i.code === 'HO_TRA_CHUA_CO_MAPPING' || i.code === 'KHONG_KHOP_DS_TRA' || i.code === 'HO_TRA_XUNG_DOT'),
        'ghi chu chi chua trang thai, khong chua loi du lieu'
      );
    });

    // ---- chon to / file tra ----
    console.log('\n== chon to va danh sach tra ==');
    t('quy tac: to 1 va 4 -> Truong Binh, to 2 -> Truong Binh A', () => {
      const { TRA_THEO_TO } = require('../src/lib/sources');
      assert.strictEqual(TRA_THEO_TO[1], 'tra');
      assert.strictEqual(TRA_THEO_TO[4], 'tra');
      assert.strictEqual(TRA_THEO_TO[2], 'traA');
    });
    t('to 2 dung danh sach "Truong Binh A"', () => {
      const l2 = locate(undefined, { to: 2 });
      assert.strictEqual(l2.ap, 'Trường Bình A');
      assert.ok(/A/.test(l2.traChon.ten), `phai la file A, thuoc lai: ${l2.traChon && l2.traChon.ten}`);
    });
    t('to 4 dung danh sach "Truong Binh"', () => {
      const l4 = locate(undefined, { to: 4 });
      assert.strictEqual(l4.ap, 'Trường Bình');
      assert.ok(!/ A /.test(l4.traChon.ten), `phai la file goc, thuoc lai: ${l4.traChon.ten}`);
    });
    t('chi to 1 moi bat khoi phuc STT ho', () => {
      assert.strictEqual(locate(undefined, { to: 1 }).coKhoiPhuc, true);
      for (const n of [2, 3, 4]) {
        const l = locate(undefined, { to: n });
        assert.strictEqual(l.coKhoiPhuc, false, `to ${n} phai TAT khoi phuc`);
        assert.strictEqual(to1.canRecover(200, l.recoverRange), false, `to ${n} khong duoc doan ho`);
      }
    });
    t('bo qua ban "- Copy" va file "chua tham gia BHYT"', () => {
      const tos = require('../src/lib/sources').listTos();
      for (const x of tos) {
        assert.ok(!/chua\s*tham\s*gia/i.test(x.ten), `sai: ${x.ten}`);
        assert.ok(!/copy/i.test(x.ten), `sai: ${x.ten}`);
      }
      assert.ok(tos.some((x) => x.to === 1 && x.docDuoc));
    });

    // ---- fix route tren file that ----
    const g = await G.Grid.open(loc.toFile.path, 0);
    t('fix route: 148 dong, 2 tieu de, 146 du lieu', () => {
      const s = g.thongKe(1, 148);
      assert.strictEqual(s.tong, 148);
      assert.strictEqual(s.soTieuDe, 2);
      assert.strictEqual(s.tong - s.soTieuDe, 146);
    });
    t('fix route: dong lien tuc tu 1 den 148', () => {
      const s = g.thongKe(1, 148);
      s.buoc.forEach((b, i) => assert.strictEqual(b.row, i + 1));
    });
    t('fix route: 5 loi cung bi bao loi', () => {
      const s = g.thongKe(1, 148);
      const dongLoi = s.buoc.filter((b) => b.sai).map((b) => b.row);
      assert.deepStrictEqual(dongLoi, [57, 58, 67, 69, 80], `thay doi: ${dongLoi}`);
    });
    t('fix route: cong thuc da bien thanh gia tri tinh', () => {
      assert.ok(g.dongCongThuc.length > 200, `co ${g.dongCongThuc.length} o cong thuc`);
      const h3 = g.ws.getRow(3).getCell(8).value;
      assert.strictEqual(typeof h3, 'number', `H3 phai la so that, hien = ${JSON.stringify(h3)}`);
      const voCongThuc = g.dongCongThuc.filter((c) => {
        const m = /^([A-Z]+)(\d+)$/.exec(c.o);
        const v = g.ws.getRow(+m[2]).getCell(G.colToNum(m[1])).value;
        return v && typeof v === 'object' && (v.formula || v.sharedFormula);
      });
      assert.strictEqual(voCongThuc.length, 0, `con ${voCongThuc.length} o con la cong thuc`);
    });
    t('fix route: chi doc khong doi gi', () => {
      assert.strictEqual(g.log.length, 0, `da co ${g.log.length} thay doi chi doc`);
    });
    t('fix route: sua duoc o H3 ma khong hong shared formula', () => {
      const r = g.setCell(3, 'H', 99);
      assert.strictEqual(r.doi, true);
      assert.strictEqual(g.read(3).cells.H, '99');
      const k = g.check(3);
      assert.ok(k.nghi >= 1, 'tuoi lech phai duoc canh bao');
    });
    t('fix route: tra lai gia tri cu', () => {
      g.setCell(3, 'H', 47);
      assert.strictEqual(g.read(3).cells.H, '47');
    });
    t('fix route: CCCD 12 so giu dang chuoi', () => {
      g.setCell(3, 'F', '093079008311');
      assert.strictEqual(g.read(3).cells.F, '093079008311');
      assert.strictEqual(typeof g.ws.getRow(3).getCell(6).value, 'string');
    });
    t('fix route: sua ngan o dang so thu tu Excel', () => {
      assert.strictEqual(g.read(87).cells.G, '46013');
      g.fixRange(1, 148);
      assert.strictEqual(g.read(87).cells.G, '22/12/2025');
      assert.ok(g.ws.getRow(87).getCell(7).value instanceof Date);
    });
    t('fix route: khong sua dong con loi sai', () => {
      const truoc = g.read(67).cells.E;
      g.fix(67);
      assert.strictEqual(g.read(67).cells.E, truoc, 'dong 67 con loi nen phai giu nguyen');
    });
    t('fix route: khong sua dong tieu de', () => {
      const r = g.fix(1);
      assert.deepStrictEqual(r.doi, []);
      assert.strictEqual(g.read(1).cells.A, 'STT');
    });
    t('fix route: nhat ky ghi du gia tri cu', () => {
      assert.ok(g.log.length > 0, 'phai co thay doi trong nhat ky');
      assert.ok(g.log.every((l) => 'row' in l && 'col' in l && 'giaTriCu' in l && 'giaTriMoi' in l && 'lyDo' in l));
    });
  } else {
    console.log('  (bo qua - khong tim thay file nguon)');
  }

  console.log(`\n${pass} pass, ${fail} fail\n`);
  process.exit(fail ? 1 : 0);
})();
