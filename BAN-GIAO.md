# BAN GIAO VIEC — QLND (to mau 1)

> File nay de AI khac tiep tuc viec. Bat dau voi dieu nay.

## 0. UU TRENH BOI (doc truoc het)

**File goc co du lieu CCCD/ho ten thuc cua nguoi thuc.** Thu muc `source/fileexcel/` va cac file `.xlsx` trong repo CONG CAC so CCCD 12 chu, ho ten, ngay sinh, so dien thoai cua 600-700 nguoi thuc. Commit `6af29a0` da PUSH LEN `https://github.com/antlt1/QLND`. Neu remote co the truy cong, du lieu ca nhan da bi mat. Can KIEM TRA ngay: doi repo thanh private hoặc xoa commit/remote đó.

## 1. Trang thai git

- Branch `main`, commit gan nhat: `6af29a0` "upload all file new project" (da push, chua tao PR).
- File DA DOI but chua commit:
  - `src/lib/arrange.js`
  - `src/lib/sources.js`
  - `src/lib/to1.js`
  - `test/run.js`
  - `tools/kiem-tra-xep.js`
- File MOI chua commit: `tools/khao-sat-to.js`, `tools/so-sanh-schema.js`, `tools/tim-vung-chen.js`
- Khong co PR. `git status --short` de kiem tra lai truoc khi push.

## 2. Muc tieu hien tai

Them menu chon to + mapping file tra, debug `qlnd.cmd`, chay `xep` cho To 2 va To 4 (To 1 & 4 dung `Hộ ấp Trường Bình sử dụng để tra.xlsx`; To 2 dung `Hộ ấp Trường Bình A sử dụng để tra.xlsx`).

## 3. Ket qua da lam duoc

### 3a. Chon to / mapping (hoàn thành)
- `src/lib/sources.js`: `listTos()` (`TRA_THEO_TO = {1:'tra', 4:'tra', 2:'traA'}`, loại file "chưa tham gia BHYT" và "- Copy"; `listTra()`, `locate(to)` trả về `toFile, traChon, ap, recoverRange, coKhoiPhuc, headerRows`), `scanUnreadable()`.
- `src/lib/chon.js`: `resolve(args)`, `assertChay(cfg)`, `inNguon(cfg)`, `menu()` (hội interactive), `docChon/luuChon/xoaChon`.
- `src/cli/chon.js`: `qlnd.cmd chon [--to N] [--tra traA] [--xoa] [--chay]`.
- `--to` chạy qua toàn bộ pipeline (extract/build/excel/xep/serve/test).
- **Phát hiện quan trọng:** mỗi tổ có layout cột khác nhau:
  - To 1: ten=c[3], mabh=c[4], ns=c[5], cccd=c[6], ngaycap=c[7], tuoi=c[8], quanhe=c[9], dt=c[10], ngh=c[14], ghiChu=c[25].
  - To 2: ten=c[2], mabh=c[3], ns=c[5], cccd=c[6], ngaycap=c[7], tuoi=c[8], quanhe=c[9], dt=c[10], ngh=c[14], ghiChu=c[25].
  - To 4: ten=c[2], mabh=c[3], ns=c[5], cccd=c[6], ngaycap=c[7], quane=c[8], tuoi=c[9], dt=c[10], ngh=c[15], ghiChu=c[26].
- `src/lib/to1.js`: `COT_THEO_TO` (bản đồ cột theo tổ), `parseTo1(rows, to)` dùng `C`. Xuất `COT_THEO_TO`.

### 3b. Xếp lại danh sách (hoàn thành cho cả 3 tổ)
- `src/lib/arrange.js`: `xep({srcFile,traFile,ap,to,recoverRange,headerRows,outFile})`, dùng `C` theo tổ; cột B chỉ ghi STT người nếu `C.ten !== 2` (tránh ghi đè tên To 2/4).
- `src/cli/xep.js`: dùng `resolve` thay vì hard-code.
- `src/cli/extract.js`, `src/cli/build.js`, `src/cli/excel.js`, `src/cli/serve.js`: wire `--to`, output `to${to}-ho.json`...
- `src/lib/xlsx.js`: `isJunkName` lọc `~$`, `.~lock.`, `~`; `isOle2/isRealXlsx` kiểm tra magic; `writeXlsx` báo lỗi dễ đọc khi file bị Excel mở.
- `qlnd.cmd`: `pause` khi bấm đúp (`QLND_NO_PAUSE=1` để tắt), lỗi `exit /b %ERRORLEVEL%` sau pause (cần lưu `ERR` trước pause).
- `tools/kiem-tra-xep.js`: nhận `--to`, đọc đúng cột theo `COT_THEO_TO`.
- `test/run.js`: 52 tests pass (0 fail).

### 3c. Kết quả chạy `qlnd.cmd xep --to N` (đã xác minh)

| Tổ | Hộ | Người | STT người cuối | Chưa xác định hộ | Lỗi checker |
|---|---|---|---|---|---|
| **1** | 151 | 598 | 598 | 9 | 6 (có sẵn trong nguồn) |
| **2** | 155 | 654 | 0 | 1 | ~237 "chưa vào hộ nào" (cấu trúc) |
| **4** | 260 | 765 | 158 | 93 | ~86 (cấu trúc) |

File output:
- `output/Danh sách Tổ 1 - đã xếp.xlsx` (sheet "ấp Trường Bình", 601 dòng)
- `output/Danh sách Tổ 2 - đã xếp.xlsx` (sheet "ấp Trường Bình", 674 dòng)
- `output/Danh sách Tổ 4 - đã xếp.xlsx` (sheet "ấp Trường Bình", 767 dòng)

Các file source **KHÔNG bị thay đổi**.

## 4. Phan tich va loi can gap

### 4a. Vùng không xác định hộ (nguyên nhân: thành viên đứng trước chủ hộ trong file)
- **To 4:** dòng 3–132 (130 người) toàn "thành viên", chủ hộ bắt đầu từ dòng 133. Bật `RECOVER_THEO_TO[4] = {from:3, to:132}` → giảm 130 → 93 chưa xác định (phần còn lại là các nhóm nhỏ rời).
- **To 2:** vùng lớn nhất là dòng 3–160 (158 người). Đã thêm `RECOVER_THEO_TO[2] = {from:3, to:160}` trong `src/lib/sources.js`, **CHƯA CHẠY XONG để xác nhận** (lệnh bị user abort khi đang chạy `--to 2`).
- **To 1:** `RECOVER_THEO_TO[1] = {from:149, to:325}` (đã xác minh, 151 hộ/598 người/9 unknown).
- **To 3:** OLE2 `.xls` → đọc không được, cần chuyển sang .xlsx trong Excel trước.
- **To 5:** chỉ có file "DS NGƯỜI CHƯA THAM GIA BHYT TỔ 5" → không phải danh sách tổ.

### 4b. Phần còn lại của To 2 / To 4 sau khi bật recovery (cần chạy lại)
- Còn nhiều nhóm nhỏ "chưa vào hộ nào" vì phân tán khắp file (dòng 163-251, 294-320, ... ở To 2; các dòng rời ở To 4). Đây là vấn đề cấu trúc nguồn (file có nhiều khối dữ liệu riêng lẻ), không phải lỗi code.
- Checker `tools/kiem-tra-xep.js` cũng chỉ quét tuần tự nên báo thừa cho các trường hợp đó.

### 4c. Các file diagnostic (không cần commit)
- `tools/khao-sat-to.js` — khảo sát cấu trúc 1 tổ.
- `tools/so-sanh-schema.js` — so sánh layout cột các tổ.
- `tools/tim-vung-chen.js` — tìm vùng dòng mà sttHo=0.

## 5. Buoc tiep theo cho AI khac

1. **Chạy lại `qlnd.cmd xep --to 2`** (với `RECOVER_THEO_TO[2]={3,160}` vừa thêm) rồi `qlnd.cmd xep --to 4`, xem số "chưa xác định hộ" giảm bao nhiêu.
2. Chạy `qlnd.cmd test` → phải còn 52 pass, 0 fail.
3. **Kiểm tra privacy:** đổi repo `antlt1/QLND` thành private hoặc xóa commit `6af29a0` trên remote (vì chứa CCCD thật).
4. Có thể bỏ commit/push các `tools/` diagnostic nếu không cần.
5. Commit tất cả file đổi + push.
6. Tùy chọn: thêm `qlnd.cmd xep --to 3` sau khi file To 3 được chuyển sang .xlsx; thêm `--tra` cho các tổ khác; thêm `output/.chon.json` cho config mặc định.

## 6. File can tham khao

- `src/lib/to1.js` — `COT_THEO_TO`, `parseTo1(rows, to)`, `build({to1Rows, lookupRows, ap, to, recoverRange})`.
- `src/lib/sources.js` — `locate(to)`, `RECOVER_THEO_TO`, `TRA_THEO_TO`.
- `src/lib/arrange.js` — `xep(...)`, dùng `C = to1.COT_THEO_TO[to]`.
- `tools/kiem-tra-xep.js` — `--to`, kiểm tra file xếp.
- `test/run.js` — 52 tests.

## 7. Ghi chu

- File nguồn (`source/fileexcel/*.xlsx`) **KHÔNG** được sửa.
- Không guess hộ/không sửa dữ liệu; chỉ sắp xếp lại vị trí dòng và gán lại STT hộ bằng CCCD trong vùng đã xác định.
- Output ghi nhận mọi trường hợp chưa xác định, chưa khớp, trùng CCCD để người dùng duyệt.
