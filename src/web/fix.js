'use strict';
/** Trang fix route: xem va sua tung dong cua file Excel goc. */

const $ = (s) => document.querySelector(s);
const api = (p, opt) => fetch(p, opt).then((r) => r.json());

let DU = { buoc: [], from: 1, to: 148, tieuDe: {}, file: '', sheet: '' };
let COT = []; // danh sach cot theo thu tu A..Z
let hang = 1; // dong dang chon
let dangSua = null; // {row, col}

const CHI_DUNG = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'];

function chuoiCot(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** 1 dong -> trang thai + danh sach loi o theo cot */
function thongTin(b) {
  if (b.laTieuDe) return { cls: 'r-tieude', tt: 'tieude', chu: 'tiêu đề', loi: {} };
  const loi = {};
  for (const l of b.loi || []) loi[l.cot] = l;
  const sai = b.sai || 0;
  const nghi = b.nghi || 0;
  return {
    loi,
    sai,
    nghi,
    cls: sai ? 'r-sai' : nghi ? 'r-nghi' : 'r-ok',
    tt: sai ? 'tt-sai' : nghi ? 'tt-nghi' : 'tt-ok',
    chu: sai ? `${sai} lỗi` : nghi ? `${nghi} xem` : 'ok',
  };
}

function ve() {
  const chiLoi = $('#chiLoi').checked;
  const ds = DU.buoc.filter((b) => !chiLoi || (b.sai || b.nghi));

  $('#tongDong').textContent = DU.to - DU.from + 1;
  $('#count').textContent = `${ds.length} dòng hiện / ${DU.buoc.length}`;

  if (!ds.length) {
    $('#view').innerHTML = '<div class="empty">Không có dòng nào cần xem.</div>';
    return;
  }

  let h = '<div style="overflow-x:auto"><table class="route"><thead><tr><th class="dong">Dòng</th>';
  for (const c of COT) {
    const ten = DU.tieuDe[c] || '';
    h += `<th class="ten">${ten || '&nbsp;'}<i>${c}</i></th>`;
  }
  h += '<th class="ten" style="width:300px">Vấn đề / gợi ý<i>kiểm tra</i></th></tr></thead><tbody>';

  for (const b of ds) {
    const t = thongTin(b);
    const cls = b.row === hang ? `${t.cls} hienTai` : t.cls;
    h += `<tr class="${cls}" data-row="${b.row}">`;
    h += `<td class="dong">${b.row}</td>`;
    for (const c of COT) {
      const v = b.cells[c] || '';
      const l = t.loi[c];
      const cls = `o${l ? ' ' + l.muc : ''}${v === '' ? ' rong' : ''}`;
      const data = dangSua && dangSua.row === b.row && dangSua.col === c;
      h += data
        ? `<td class="sua"><input value="${String(v).replace(/"/g, '&quot;')}" /></td>`
        : `<td class="${cls}" data-row="${b.row}" data-col="${c}">${String(v).replace(/</g, '&lt;')}</td>`;
    }
    h += '<td class="tt">';
    if (t.cls === 'r-tieude') h += '<span class="tag trangthai tt-tieude">tiêu đề</span>';
    else if (b.sai || b.nghi) {
      h += `<ul class="loi">${(b.loi || [])
        .map((l) => `<li class="${l.muc}"><b>${l.cot}</b> ${loi(l.loi)}${l.goiY ? ` <em>→ ${loi(l.goiY)}</em>` : ''}</li>`)
        .join('')}</ul>`;
    } else h += '<span class="tag trangthai tt-ok">ok</span>';
    h += '</td></tr>';
  }
  h += '</tbody></table></div>';
  $('#view').innerHTML = h;
}

const loi = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

function diemDinh() {
  const ds = DU.buoc.filter((b) => !$('#chiLoi').checked || b.sai || b.nghi);
  const i = ds.findIndex((b) => b.row === hang);
  const tr = $('#view')?.querySelector(`tr[data-row="${hang}"]`);
  if (tr) tr.scrollIntoView({ block: 'center', behavior: 'smooth' });
  $('#dongHienTai').textContent = hang;
  $('#truoc').disabled = i <= 0;
  $('#sau').disabled = i < 0 || i >= ds.length - 1;
  return i;
}

function nhac(n) {
  $('#stats').innerHTML = n;
}

async function tai(from, to) {
  $('#stats').textContent = 'Đang tải…';
  const d = await api(`/api/fix-route?from=${from}&to=${to}`);
  if (d.loi) {
    $('#loiChung').hidden = false;
    $('#loiChung').textContent = d.loi;
    $('#view').innerHTML = '';
    return;
  }
  $('#loiChung').hidden = true;
  DU = d;
  COT = Object.keys(d.tieuDe || {});
  if (!COT.includes('I')) COT.push('I');
  hang = d.from;
  ve();
  diemDinh();
  nhac(
    `<span>File: <b>${loi(d.file)}</b> — sheet "${loi(d.sheet)}"</span>` +
      `<span>Dòng <b>${d.from}–${d.to}</b></span>` +
      `<span>ok <b>${d.ok}</b></span><span>nên xem <b>${d.nghi}</b></span>` +
      `<span>lỗi <b>${d.sai}</b></span><span>đã sửa <b>${d.daSua}</b> ô</span>`
  );
}

// --- sua o ---
function batDauSua(tr, td) {
  const row = +td.dataset.row;
  const col = td.dataset.col;
  if (row <= 2) return; // khong sua tieu de
  dangSua = { row, col };
  ve();
  const input = $('#view')?.querySelector(`tr[data-row="${row}"] td.sua input`);
  if (input) {
    input.focus();
    input.select();
  }
}

async function luuO(row, col, gia) {
  dangSua = null;
  const d = await api('/api/fix-route/set', {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ row, col, value: gia }),
  });
  if (d.loi) return alert(d.loi);
  await tai(DU.from, DU.to);
  // nhay den dong vua sua
  hang = row;
  diemDinh();
}

document.addEventListener('click', (e) => {
  const nut = e.target.closest('.nut');
  if (nut && nut.id === 'load') {
    tai(+$('#from').value || 1, +$('#to').value || 148);
    return;
  }
  if (nut && nut.id === 'fixAll') {
    if (!confirm('Sửa tự động những gì chắc chắn (bỏ khoảng trắng, chuyển ngày)?\nDòng còn lỗi sẽ không bị đụng tới.')) return;
    return api('/api/fix-route/fix', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ from: DU.from, to: DU.to }),
    }).then(() => tai(DU.from, DU.to));
  }
  if (nut && nut.id === 'save') {
    return api('/api/fix-route/save', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ from: DU.from, to: DU.to }),
    }).then((d) => {
      if (d.loi) return alert(d.loi);
      alert(`Đã lưu:\n${d.file}\n\n${d.soDoi} thay đổi\nNhật ký: ${d.nhatKy}\n\nFile gốc không bị thay đổi.`);
    });
  }
  if (nut && nut.id === 'log') return moNhatKy();
  if (nut && nut.id === 'truoc') {
    const ds = DU.buoc.filter((b) => !$('#chiLoi').checked || b.sai || b.nghi);
    const i = ds.findIndex((b) => b.row === hang);
    if (i > 0) {
      hang = ds[i - 1].row;
      ve();
      diemDinh();
    }
    return;
  }
  if (nut && nut.id === 'sau') {
    const ds = DU.buoc.filter((b) => !$('#chiLoi').checked || b.sai || b.nghi);
    const i = ds.findIndex((b) => b.row === hang);
    if (i >= 0 && i < ds.length - 1) {
      hang = ds[i + 1].row;
      ve();
      diemDinh();
    }
    return;
  }
  if (nut && nut.classList.contains('close')) {
    $('#panel').hidden = true;
    return;
  }

  const td = e.target.closest('td.o');
  if (td) return batDauSua(e.target.closest('tr'), td);
});

document.addEventListener('keydown', (e) => {
  const input = e.target.closest('td.sua input');
  if (input) {
    if (e.key === 'Enter') {
      e.preventDefault();
      luuO(dangSua.row, dangSua.col, input.value);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      dangSua = null;
      ve();
    }
    return;
  }
  if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
    e.preventDefault();
    $('#sau').click();
    if (e.key === 'ArrowUp') $('#truoc').click();
  }
  if (e.key === 'Enter') e.preventDefault();
});

$('#chiLoi').addEventListener('change', () => {
  ve();
  diemDinh();
});

async function moNhatKy() {
  const d = await api('/api/fix-route/doi-chieu');
  $('#panel').hidden = false;
  if (d.loi) {
    $('#panel').innerHTML = `<header><h2>Nhật ký</h2><button class="close">&times;</button></header><div class="body">${loi(d.loi)}</div>`;
    return;
  }
  const rows = (d.nhatKy || [])
    .map(
      (l) =>
        `<tr><td>${l.row}</td><td>${l.col}</td><td class="cu">${loi(l.giaTriCu) || '(trống)'}</td>` +
        `<td class="moi">${loi(l.giaTriMoi) || '(trống)'}</td><td>${loi(l.lyDo)}</td></tr>`
    )
    .join('');
  $('#panel').innerHTML =
    `<header><h2>Nhật ký sửa — ${(d.nhatKy || []).length} thay đổi</h2><button class="close">&times;</button></header>` +
    `<div class="body">${
      rows
        ? '<h3>Danh sách thay đổi</h3><table><thead><tr><th>Dòng</th><th>Cột</th><th>Cũ</th><th>Mới</th><th>Lý do</th></tr></thead><tbody>' +
          rows +
          '</tbody></table>'
        : '<div class="rong">Chưa có thay đổi nào.</div>'
    }</div>`;
}

tai(+$('#from').value, +$('#to').value);
