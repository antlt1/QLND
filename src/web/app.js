'use strict';
const $ = (s) => document.querySelector(s);

const K = {
  1: ['k1', 'Đã sửa xong'],
  2: ['k2', 'Thành viên bị chồn dòng'],
  3: ['k3', 'Chưa cập nhật'],
  S: ['kr', 'Cột A rác'],
};
const kClass = (khung) =>
  khung === '1_DA_FIX' ? K[1][0] : khung === 'STT_HO_RAC' ? K.S[0] : khung === '3_CHUA_CAP_NHAT' ? K[3][0] : K[2][0];
const kText = (khung) =>
  khung === '1_DA_FIX' ? K[1][1] : khung === 'STT_HO_RAC' ? K.S[1] : khung === '3_CHUA_CAP_NHAT' ? K[3][1] : K[2][1];

let D = null;
let tab = 'ho';

const esc = (s) =>
  String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);

function load() {
  fetch('api/du-lieu')
    .then((r) => r.json())
    .then((d) => {
      D = d;
      renderStats();
      draw();
    })
    .catch((e) => {
      $('#view').innerHTML = `<p class="empty">Không tải được dữ liệu: ${esc(e.message)}</p>`;
    });
}

function renderStats() {
  const t = D.thongKe || {};
  const el = $('#stats');
  el.innerHTML = [
    `<span><b>${D.tongHo}</b> hộ</span>`,
    `<span><b>${D.tongNguoi}</b> người</span>`,
    `<span>khôi phục <b>${t.recover?.khoiPhuc ?? 0}</b> dòng (dòng ${t.recoverRange?.from ?? '-'}–${t.recoverRange?.to ?? '-'})</span>`,
    `<span>cảnh báo <b>${(t.soCanhBao ?? D.canhBao.length)}</b></span>`,
    `<span>ghi chú <b>${(t.soGhiChu ?? D.ghiChu.length)}</b></span>`,
  ].join('');
}

const q = () => $('#q').value.trim().toLowerCase();
const khungF = () => $('#khung').value;
const onlyCb = () => $('#onlyCb').checked;

function matchHouse(h, s) {
  if (khungF() && h.khung !== khungF()) return false;
  if (onlyCb() && !(h.canhBao || []).length) return false;
  if (!s) return true;
  if (String(h.sttHo || '') === s) return true;
  if ((h.chuHo || '').toLowerCase().includes(s)) return true;
  return h.nguoi.some(
    (m) =>
      m.hoTen.toLowerCase().includes(s) ||
      m.cccd.includes(s) ||
      String(m.dienThoai || '').includes(s)
  );
}

function viewHo() {
  const s = q();
  const rows = D.ho.filter((h) => matchHouse(h, s));
  $('#count').textContent = `${rows.length} / ${D.ho.length} hộ`;
  if (!rows.length) return empty();
  let html = `<table><thead><tr>
    <th class="num">TT</th><th class="num">STT hộ</th><th>Chủ hộ</th>
    <th class="num">Số người</th><th class="num">Dòng</th><th>Khối</th><th>Cảnh báo</th>
  </tr></thead><tbody>`;
  for (const h of rows) {
    html += `<tr data-no="${h.no}">
      <td class="num">${h.no}</td>
      <td class="num">${h.sttHo || '—'}</td>
      <td>${esc(h.chuHo) || '<i>trống</i>'}</td>
      <td class="num">${h.soNguoi}</td>
      <td class="num">${h.rowMin}${h.rowMax !== h.rowMin ? '–' + h.rowMax : ''}</td>
      <td><span class="tag ${kClass(h.khung)}">${kText(h.khung)}</span></td>
      <td class="cb">${(h.canhBao || []).length || ''}</td>
    </tr>`;
  }
  return html + '</tbody></table>';
}

function viewNguoi() {
  const s = q();
  const all = [];
  for (const h of D.ho) {
    if (khungF() && h.khung !== khungF()) continue;
    for (const m of h.nguoi) {
      if (onlyCb() && !(m.canhBao || []).length && !(h.canhBao || []).length) continue;
      if (s && !(m.hoTen.toLowerCase().includes(s) || m.cccd.includes(s) || String(h.sttHo) === s)) continue;
      all.push({ h, m });
    }
  }
  $('#count').textContent = `${all.length} / ${D.tongNguoi} người`;
  if (!all.length) return empty();
  let html = `<table><thead><tr>
    <th class="num">STT</th><th class="num">Hộ</th><th>Họ tên</th><th>CCCD</th>
    <th class="num">Năm sinh</th><th>Vai trò</th><th>Điện thoại</th><th>Nguồn STT hộ</th><th>Cảnh báo</th>
  </tr></thead><tbody>`;
  for (const { h, m } of all) {
    html += `<tr data-no="${h.no}">
      <td class="num">${m.sttNguoi}</td>
      <td class="num">${h.sttHo || '—'}</td>
      <td>${esc(m.hoTen)}</td>
      <td>${esc(m.cccd) || '—'}</td>
      <td class="num">${m.ngaySinh ? m.ngaySinh.slice(0, 4) : '—'}</td>
      <td>${m.quanHe === 'chuHo' ? 'chủ hộ' : 'thành viên'}</td>
      <td>${esc(m.dienThoai) || '—'}</td>
      <td>${esc(m.nguonHo)}</td>
      <td class="cb">${(m.canhBao || []).length || ''}</td>
    </tr>`;
  }
  return html + '</tbody></table>';
}

function viewIssues(list, ten) {
  const s = q();
  const rows = list.filter((i) => !s || i.ten.toLowerCase().includes(s) || i.code.toLowerCase().includes(s) || i.moTa.toLowerCase().includes(s));
  $('#count').textContent = `${rows.length} / ${list.length} ${ten}`;
  if (!rows.length) return empty();
  let html = `<table><thead><tr>
    <th class="num">Dòng Excel</th><th>Họ tên</th><th>Mã</th><th>Mô tả</th>
  </tr></thead><tbody>`;
  for (const i of rows) {
    html += `<tr>
      <td class="num">${i.row}</td>
      <td>${esc(i.ten) || '—'}</td>
      <td>${esc(i.code)}</td>
      <td>${esc(i.moTa)}</td>
    </tr>`;
  }
  return html + '</tbody></table>';
}

const empty = () => '<p class="empty">Không có dòng nào khớp.</p>';

function draw() {
  const v =
    tab === 'ho' ? viewHo() : tab === 'nguoi' ? viewNguoi() : tab === 'canhbao' ? viewIssues(D.canhBao, 'cảnh báo') : viewIssues(D.ghiChu, 'ghi chú');
  $('#view').innerHTML = v;
}

function openHouse(no) {
  const h = D.ho.find((x) => x.no === no);
  if (!h) return;
  const dd = $('#detail');
  let html = `<header>
      <h2>Hộ ${h.sttHo || '(cột A rác)'} — ${esc(h.chuHo) || 'chưa rõ chủ hộ'}</h2>
      <button class="close" onclick="document.getElementById('detail').hidden=true">&times;</button>
    </header><div class="body">
      <table>
        <tr><td>TT</td><td>${h.no} / ${D.tongHo}</td></tr>
        <tr><td>Khối</td><td><span class="tag ${kClass(h.khung)}">${kText(h.khung)}</span></td></tr>
        <tr><td>Số người</td><td>${h.soNguoi}</td></tr>
        <tr><td>Dòng Excel</td><td>${h.rowMin}${h.rowMax !== h.rowMin ? ' – ' + h.rowMax : ''}</td></tr>
        <tr><td>Cột A gốc</td><td>${esc(JSON.stringify(h.sttHo ? String(h.sttHo) : '.'))}</td></tr>
      </table>`;
  if (h.goiY) html += `<h3>Gợi ý</h3><p class="goiy">${esc(h.goiY)}</p>`;
  if ((h.canhBao || []).length) {
    html += `<h3>Cảnh báo (${h.canhBao.length})</h3><ul>${h.canhBao.map((c) => `<li class="cb">${esc(c)}</li>`).join('')}</ul>`;
  }
  html += '<h3>Thành viên</h3><table><tr><th class="num">STT</th><th>Họ tên</th><th class="num">Năm sinh</th><th class="num">Dòng</th></tr>';
  for (const m of h.nguoi) {
    html += `<tr>
      <td class="num">${m.sttNguoi}</td>
      <td>${esc(m.hoTen)} ${m.quanHe === 'chuHo' ? '<b>(chủ hộ)</b>' : ''}${(m.canhBao || []).length ? ` <span class="cb">⚠ ${m.canhBao.length}</span>` : ''}</td>
      <td class="num">${m.ngaySinh ? m.ngaySinh.slice(0, 4) : '—'}</td>
      <td class="num">${m.row}</td>
    </tr>`;
  }
  html += '</table></div>';
  dd.innerHTML = html;
  dd.hidden = false;
}

document.addEventListener('click', (e) => {
  const t = e.target.closest('.tabs button');
  if (t) {
    tab = t.dataset.tab;
    document.querySelectorAll('.tabs button').forEach((b) => b.classList.toggle('on', b === t));
    draw();
    return;
  }
  const tr = e.target.closest('tbody tr[data-no]');
  if (tr) openHouse(+tr.dataset.no);
});

let t;
$('#q').addEventListener('input', () => {
  clearTimeout(t);
  t = setTimeout(draw, 120);
});
$('#khung').addEventListener('change', draw);
$('#onlyCb').addEventListener('change', draw);
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') $('#detail').hidden = true;
});

load();
