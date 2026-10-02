import { useEffect, useMemo, useState } from 'react';
import { ref, onValue } from 'firebase/database';
import { db } from './firebase.js';

const KCLASS = (khung) =>
  khung === '1_DA_FIX' ? 'k1' : khung === 'STT_HO_RAC' ? 'kr' : khung === '3_CHUA_CAP_NHAT' ? 'k3' : 'k2';
const KTEXT = (khung) =>
  khung === '1_DA_FIX' ? 'Đã sửa xong' : khung === 'STT_HO_RAC' ? 'Cột A rác' : khung === '3_CHUA_CAP_NHAT' ? 'Chưa cập nhật' : 'Thành viên bị chồn dòng';

const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

export default function App() {
  const [D, setD] = useState(null);
  const [err, setErr] = useState('');
  const [tab, setTab] = useState('ho');
  const [q, setQ] = useState('');
  const [khung, setKhung] = useState('');
  const [onlyCb, setOnlyCb] = useState(false);
  const [openNo, setOpenNo] = useState(null);

  useEffect(() => {
    const unsub = onValue(
      ref(db, 'du-lieu'),
      (snap) => {
        const v = snap.val();
        if (!v) setErr('Chưa có dữ liệu tại nhánh "du-lieu" trong RTDB.');
        else setD(v);
      },
      (e) => setErr(e.message)
    );
    return () => unsub();
  }, []);

  const s = q.trim().toLowerCase();
  const scope = useMemo(() => {
    if (!D) return [];
    return D.ho.filter((h) => {
      if (khung && h.khung !== khung) return false;
      if (onlyCb && !(h.canhBao || []).length) return false;
      if (!s) return true;
      if (String(h.sttHo || '') === s) return true;
      if ((h.chuHo || '').toLowerCase().includes(s)) return true;
      return h.nguoi.some((m) => m.hoTen.toLowerCase().includes(s) || m.cccd.includes(s) || String(m.dienThoai || '').includes(s));
    });
  }, [D, s, khung, onlyCb]);

  const nguoi = useMemo(() => {
    if (!D) return [];
    const out = [];
    for (const h of D.ho) {
      if (khung && h.khung !== khung) continue;
      for (const m of h.nguoi) {
        if (onlyCb && !(m.canhBao || []).length && !(h.canhBao || []).length) continue;
        if (s && !(m.hoTen.toLowerCase().includes(s) || m.cccd.includes(s) || String(h.sttHo) === s)) continue;
        out.push({ h, m });
      }
    }
    return out;
  }, [D, s, khung, onlyCb]);

  const issues = useMemo(() => {
    if (!D) return [];
    const list = tab === 'canhbao' ? D.canhBao : D.ghiChu;
    return (list || []).filter((i) => !s || i.ten.toLowerCase().includes(s) || i.code.toLowerCase().includes(s) || i.moTa.toLowerCase().includes(s));
  }, [D, s, tab]);

  const open = openNo != null && D ? D.ho.find((x) => x.no === openNo) : null;

  return (
    <>
      <header>
        <h1>Danh sách hộ — Ấp Trường Bình</h1>
        <div className="stats">
          {D && (
            <>
              <span><b>{D.tongHo}</b> hộ</span>
              <span><b>{D.tongNguoi}</b> người</span>
              <span>cảnh báo <b>{D.thongKe?.soCanhBao ?? D.canhBao?.length ?? 0}</b></span>
              <span>ghi chú <b>{D.thongKe?.soGhiChu ?? D.ghiChu?.length ?? 0}</b></span>
            </>
          )}
        </div>
      </header>

      <nav className="tabs">
        {[['ho', 'Danh sách hộ'], ['nguoi', 'Danh sách người'], ['canhbao', 'Cảnh báo'], ['ghichu', 'Ghi chú']].map(([k, t]) => (
          <button key={k} className={tab === k ? 'on' : ''} onClick={() => setTab(k)}>{t}</button>
        ))}
      </nav>

      <div className="bar">
        <input type="search" value={q} onChange={(e) => setQ(e.target.value)} placeholder="Tìm tên, STT hộ, CCCD, số điện thoại…" />
        <select value={khung} onChange={(e) => setKhung(e.target.value)}>
          <option value="">Tất cả khối</option>
          <option value="1_DA_FIX">1 — Đã sửa xong</option>
          <option value="2_CHUA_FIX_THANH_VIEN_BI_CHON">2 — Thành viên bị chồn dòng</option>
          <option value="3_CHUA_CAP_NHAT">3 — Chưa cập nhật</option>
          <option value="STT_HO_RAC">Cột A rác (dấu .)</option>
        </select>
        <label className="chk"><input type="checkbox" checked={onlyCb} onChange={(e) => setOnlyCb(e.target.checked)} /> Chỉ hiện dòng có cảnh báo</label>
        <span className="count">
          {D && tab === 'ho' && `${scope.length} / ${D.ho.length} hộ`}
          {D && tab === 'nguoi' && `${nguoi.length} / ${D.tongNguoi} người`}
          {D && (tab === 'canhbao' || tab === 'ghichu') && `${issues.length} dòng`}
        </span>
      </div>

      <main id="view">
        {err && <p className="empty">Lỗi: {err}</p>}
        {!D && !err && <p className="empty">Đang tải…</p>}
        {D && tab === 'ho' && <TableHo rows={scope} onOpen={setOpenNo} />}
        {D && tab === 'nguoi' && <TableNguoi rows={nguoi} onOpen={setOpenNo} />}
        {D && (tab === 'canhbao' || tab === 'ghichu') && <TableIssue rows={issues} />}
      </main>

      {open && (
        <div id="detail" onClick={(e) => e.target.id === 'detail' && setOpenNo(null)}>
          <header>
            <h2>Hộ {open.sttHo || '(cột A rác)'} — {open.chuHo || 'chưa rõ chủ hộ'}</h2>
            <button className="close" onClick={() => setOpenNo(null)}>&times;</button>
          </header>
          <div className="body">
            <p>Số người: {open.soNguoi} — Dòng Excel {open.rowMin}{open.rowMax !== open.rowMin ? `–${open.rowMax}` : ''} — <span className={`tag ${KCLASS(open.khung)}`}>{KTEXT(open.khung)}</span></p>
            {open.goiY && <p className="goiy">{open.goiY}</p>}
            <h3>Thành viên</h3>
            <table>
              <thead><tr><th>STT</th><th>Họ tên</th><th>Năm sinh</th><th>Dòng</th></tr></thead>
              <tbody>
                {open.nguoi.map((m) => (
                  <tr key={m.row}>
                    <td>{m.sttNguoi}</td>
                    <td>{m.hoTen} {m.quanHe === 'chuHo' && <b>(chủ hộ)</b>}</td>
                    <td>{m.ngaySinh ? m.ngaySinh.slice(0, 4) : '—'}</td>
                    <td>{m.row}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </>
  );
}

function TableHo({ rows, onOpen }) {
  if (!rows.length) return <p className="empty">Không có dòng nào khớp.</p>;
  return (
    <table>
      <thead><tr><th>TT</th><th>STT hộ</th><th>Chủ hộ</th><th>Số người</th><th>Dòng</th><th>Khối</th><th>Cảnh báo</th></tr></thead>
      <tbody>
        {rows.map((h) => (
          <tr key={h.no} onClick={() => onOpen(h.no)}>
            <td>{h.no}</td>
            <td>{h.sttHo || '—'}</td>
            <td>{h.chuHo || <i>trống</i>}</td>
            <td>{h.soNguoi}</td>
            <td>{h.rowMin}{h.rowMax !== h.rowMin ? '–' + h.rowMax : ''}</td>
            <td><span className={`tag ${KCLASS(h.khung)}`}>{KTEXT(h.khung)}</span></td>
            <td className="cb">{(h.canhBao || []).length || ''}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function TableNguoi({ rows, onOpen }) {
  if (!rows.length) return <p className="empty">Không có dòng nào khớp.</p>;
  return (
    <table>
      <thead><tr><th>STT</th><th>Hộ</th><th>Họ tên</th><th>CCCD</th><th>Năm sinh</th><th>Vai trò</th><th>Điện thoại</th><th>Nguồn STT hộ</th><th>Cảnh báo</th></tr></thead>
      <tbody>
        {rows.map(({ h, m }) => (
          <tr key={`${h.no}-${m.row}`} onClick={() => onOpen(h.no)}>
            <td>{m.sttNguoi}</td>
            <td>{h.sttHo || '—'}</td>
            <td>{m.hoTen}</td>
            <td>{m.cccd || '—'}</td>
            <td>{m.ngaySinh ? m.ngaySinh.slice(0, 4) : '—'}</td>
            <td>{m.quanHe === 'chuHo' ? 'chủ hộ' : 'thành viên'}</td>
            <td>{m.dienThoai || '—'}</td>
            <td>{m.nguonHo}</td>
            <td className="cb">{(m.canhBao || []).length || ''}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function TableIssue({ rows }) {
  if (!rows.length) return <p className="empty">Không có dòng nào khớp.</p>;
  return (
    <table>
      <thead><tr><th>Dòng Excel</th><th>Họ tên</th><th>Mã</th><th>Mô tả</th></tr></thead>
      <tbody>
        {rows.map((i, idx) => (
          <tr key={idx}>
            <td>{i.row}</td>
            <td>{i.ten || '—'}</td>
            <td>{i.code}</td>
            <td>{i.moTa}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}
