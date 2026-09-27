# Quản lý danh sách hộ cán bộ Ấp Trường Bình — Tổ 1

Tool đọc file Excel gốc, gom người về đúng hộ, xuất ra **Excel sạch** và **web xem nhanh**.

Nguyên tắc: **không xoá dòng nào, không tự khử trùng, không tự đoán.** Mọi chỗ sai hoặc chưa chắc chắn đều được giữ nguyên và ghi cảnh báo.

---

## 1. Chạy chương trình

Mở **Command Prompt** hoặc **PowerShell** tại thư mục dự án, gõ:

```
qlnd.cmd help
```

| Lệnh | Việc làm | Kết quả |
|------|----------|---------|
| **`web.cmd`** | **Bấm để xem trên web — dễ nhất** | Mở trình duyệt tự động |
| **`run-server.bat`** | **Chỉ chạy server, không mở trình duyệt** | Server chạy trong cửa sổ này |
| `qlnd.cmd build` | Đọc Excel → gom hộ → xuất tất cả | `output\*.json` + `output\*.xlsx` |
| `qlnd.cmd extract` | Chỉ đọc và gom hộ | `output\*.json` |
| `qlnd.cmd excel` | Chỉ xuất Excel từ JSON có sẵn | `output\Danh sách Tổ 1 - đã gom.xlsx` |
| **`qlnd.cmd xep`** | **Xếp lại từng hộ: chủ hộ rồi tới thành viên** | `output\Danh sách Tổ 1 - đã xếp.xlsx` |
| `qlnd.cmd fixroute` | Xem / kiểm tra từng dòng 1–148 của file gốc | In ra bảng + danh sách lỗi |
| `qlnd.cmd fixroute --sua` | Sửa những gì chắc chắn rồi xuất file mới | `output\Danh sách Tổ 1 - đã sửa dòng 1-148.xlsx` |
| `qlnd.cmd serve` | Mở web server (tự chạy, không mở trình duyệt) | <http://localhost:5173> |
| `qlnd.cmd test` | Chạy test | 45 test |

### Chạy server: `run-server.bat`

```
run-server.bat            bấm đúp — chạy server ở cổng 5173
run-server.bat 5200       chạy server ở cổng 5200
```

Khác với `web.cmd`, file này **giữ server trong cửa sổ đang chạy** nên bạn thấy log và bấm **Ctrl+C** để dừng. Nó cũng **không tự mở trình duyệt**.

| Địa chỉ | Nội dung |
|----------|----------|
| <http://localhost:5173/> | Danh sách hộ |
| <http://localhost:5173/fix.html> | Sửa Excel dòng 1–148 |

Nếu cổng đã bị chiếm, file sẽ báo và gợi ý chạy lại với cổng khác. Nếu chưa có dữ liệu trong `output\`, nó tự chạy `build` trước.

### Cách nhanh nhất: bấm đúp `web.cmd`

File nằm ngay trong thư mục dự án. **Bấm đúp chuột trái** vào `web.cmd`:

1. Tự động đọc lại file Excel nguồn (~3 giây)
2. Tự động mở trình duyệt tại <http://localhost:5173>
3. In ra địa chỉ và hướng dẫn dùng

Muốn đổi cổng (ví dụ khi cổng 5173 bị chiếm):

```
web.cmd 5200
```

Tắt web:

```
taskkill /FI "WINDOWTITLE eq QLND web*" /T
```

### Thao tác trên web

| Việc | Cách làm |
|------|----------|
| Xem chi tiết một hộ | Bấm vào hộ bất kỳ trong bảng |
| Tìm người | Gõ tên / STT hộ / CCCD / số điện thoại vào ô tìm kiếm |
| Lọc theo khối dữ liệu | Chọn trong danh sách "Tất cả khối" |
| Chỉ xem dòng có cảnh báo | Tích ô "Chỉ hiện dòng có cảnh báo" |
| Xem lỗi dữ liệu | Bấm tab **Cảnh báo** |
| Xem trạng thái xử lý | Bấm tab **Ghi chú** |
| Đóng khung chi tiết | Bấm **Esc** |
| **Sửa Excel từng dòng** | **Bấm nút "Sửa Excel dòng 1–148 →" ở thanh trên cùng** |

## 1b. Fix route — xem và sửa Excel từng dòng (dòng 1 → 148)

Đây là phần để **duyệt và sửa trực tiếp file Excel gốc**, từng dòng một, không xóa dòng nào.

### Mở trên web (dễ nhất)

```
qlnd.cmd serve
```

rồi vào <http://localhost:5173/fix.html> (hoặc bấm nút ở trang chính).

| Việc | Cách làm |
|------|----------|
| Xem từng dòng | Bảng hiện đủ dòng 1 → 148, màu viền trái báo trạng thái: xanh = ok, vàng = nên xem, đỏ = lỗi |
| Đi từng dòng | Nút **◀ Dòng trước** / **Dòng sau ▶**, hoặc phím **↑ ↓** |
| Sửa một ô | Bấm vào ô đó, gõ giá trị, nhấn **Enter** (Esc để hủy) |
| Xem vấn đề của dòng | Cột cuối bên phải liệt kê lỗi + gợi ý sửa |
| Chỉ xem dòng có vấn đề | Tích ô "Chỉ hiện dòng có vấn đề" |
| Sửa tự động phần chắc chắn | Nút **Sửa tự động chắc chắn** — chỉ bỏ khoảng trắng, chuyển ngày, định dạng ngày |
| Xem đã sửa gì | Nút **Nhật ký sửa** |
| Xuất file đã sửa | Nút **Lưu file Excel** → file mới trong `output\`, **file gốc không bị đổi** |

### Cách thông thường (gõ lệnh)

```
qlnd.cmd fixroute
qlnd.cmd fixroute --sua
qlnd.cmd fixroute --from 1 --to 148 --sua
```

### Cột nào được kiểm tra

| Cột | Ý nghĩa | Kiểm tra |
|------|---------|----------|
| A | STT hộ | phải là số hoặc trống |
| B | STT người | phải là số |
| C | Họ và tên | không được trống, không thừa khoảng trắng |
| E | Ngày sinh | phải đọc được, hoặc là số thứ tự ngày Excel |
| F | CCCD | đúng 12 chữ số |
| G | Ngày cấp | phải đọc được, hoặc là số thứ tự ngày Excel |
| H | Tuổi | phải là số, và không lệch quá 2 năm so với năm sinh |
| I | Quan hệ | chỉ `chủ hộ` hoặc `thành viên`; chỉ dòng đầu hộ được ghi `chủ hộ` |

### Kết quả kiểm tra dòng 1–148

| Trạng thái | Số dòng | Chi tiết |
|------------|---------|----------|
| Tiêu đề | 2 | dòng 1, 2 — không sửa |
| ok | 141 | không cần làm gì |
| Nên xem | 2 | dòng 63 (thiếu tuổi), dòng 87 (ngày cấp ghi dạng số `46013` = 22/12/2025) |
| Lỗi | 5 | dòng 57, 58, 69, 80 (cột A là dấu `.`) + dòng 67 (CCCD lệch sang cột E) |

Những dòng lỗi ở trên **không bị sửa tự động** — cần người kiểm tra quyết định.

### Hai điều cần biết

1. **Cột B và H trong file gốc là công thức** (`SUBTOTAL` đếm số, `DATEDIF` tính tuổi). Tool đổi chúng thành **số tĩnh** khi mở file, vì:
   - ExcelJS không ghi được workbook nếu sửa ô làm master của công thức dùng chung;
   - để nguyên công thức thì Excel sẽ tự tính lại và ghi đè giá trị bạn vừa sửa tay.

   Ô **H63** trong file gốc có công thức nhưng không lưu kết quả, nên sau khi đổi thành số tĩnh nó bị trống — tool báo **"Thiếu tuổi"** để bạn điền.

2. **File gốc không bao giờ bị sửa.** Mọi thay đổi đều ghi ra file mới trong `output\`, kèm file nhật ký `.json` ghi lại từng thay đổi kèm giá trị cũ.

### Cách thông thường (gõ lệnh)

```
qlnd.cmd build
```

Sau đó mở file: `output\Danh sách Tổ 1 - đã gom.xlsx`

### Chạy với thư mục khác

```
qlnd.cmd build --src D:\du_lieu --out D:\ket_qua
```

---

## 1c. Xếp lại danh sách — `qlnd.cmd xep`

Sao chép nguyên vẹn file Tổ 1 rồi **chỉ đổi thứ tự dòng và điền cột A / cột B**, để mỗi hộ nằm liền một khối: chủ hộ rồi tới thành viên, giống hàng 1–97 của file gốc.

Vì sao cần: trong file gốc, hộ 28–78 (dòng 98–148) chỉ có chủ hộ, cột B trống; thành viên của họ nằm rải rác ở dòng 149–325 không theo thứ tự. Cột A = STT hộ (chỉ dòng chủ hộ), cột B = STT người chạy liền từ 1.

```
qlnd.cmd xep
```

Xuất ra `output\Danh sách Tổ 1 - đã xếp.xlsx` với 3 sheet:

| Sheet | Nội dung |
|-------|----------|
| sheet đầu | danh sách đã xếp, giữ nguyên tên sheet và toàn bộ cột của file gốc |
| **Nhật ký xếp** | từng người: STT người, STT hộ, vai trò, họ tên, CCCD, **dòng cũ → dòng mới**, trạng thái, nguồn |
| **Tóm tắt** | số hộ, số người, số dòng giữ nguyên / di chuyển, số người chưa xác định hộ |

Quy tắc:

- **Không mất dòng nào** — mọi dòng có dữ liệu ở file gốc đều có trong file mới (đã kiểm: 599 dòng → 599 dòng).
- Người **chưa xác định được hộ** xếp cuối file, cột A ghi `.` theo đúng quy ước của file gốc, và có `gợi ý` trong Nhật ký để bạn xử lý.
- Cột B và H vốn là công thức → đổi thành **số tĩnh** (416 ô), lý do xem mục 1b.
- **File gốc không bị đụng tới.**

Kiểm tra lại kết quả:

```
node tools\kiem-tra-xep.js
```

---

## 2. Node.js ở đâu

Chương trình tự thêm Node vào `PATH`, **không cần cài gì thêm**. Node nằm ở:

```
%USERPROFILE%\.qlnd\node
```

Nếu thư mục này bị xoá (ví dụ dọn dọn máy), tải lại Node 24 từ <https://nodejs.org> rồi chép vào đúng đường dẫn trên.

Muốn dùng `npm` (cài thêm thư viện):

```
%USERPROFILE%\.qlnd\node\npm.cmd install
```

---

## 3. File nguồn

Tool tự tìm trong `source\fileexcel\`:

| File | Vai trò |
|------|---------|
| `Danh sách Tổ 1.xlsx` | Dữ liệu chính — 598 người, dòng 3–602 |
| `Hộ ấp Trường Bình sử dụng để tra.xlsx` | Danh sách tra cứu — 1623 người / 382 hộ |

Đặt file vào `source\fileexcel\` rồi chạy `qlnd.cmd build`. Tên file có dấu tiếng Việt không quan trọng, tool tự nhận.

### File không đọc được

Bốn file sau là định dạng `.xls` cũ (OLE2) nhưng đuôi lại đặt `.xlsx`, ExcelJS không đọc được. Tool **tự động bỏ qua** và in ra danh sách:

- `Danh sách Tổ 2.xlsx`, `Danh sách Tổ 2 - Copy.xlsx`
- `Danh sách Tổ 3.xlsx`, `Danh sách Tổ 4.xlsx`

Muốn dùng các tổ này thì phải mở bằng Excel và **Save As → .xlsx** thật.

---

## 4. File kết quả

Trong `output\`:

| File | Nội dung |
|------|----------|
| `Danh sách Tổ 1 - đã gom.xlsx` | **File chính** — 7 sheet, dùng để in/gửi |
| `to1-ho.json` | Dữ liệu đã gom nhóm (cho web + Excel) |
| `to1-nguoi.json` | 598 người, một dòng mỗi người |
| `to1-loi.json` | Toàn bộ cảnh báo + ghi chú |
| `tra-cuoc.json` | 1623 người trong danh sách tra cứu |
| `Danh sách Tổ 1 - đã sửa dòng 1-148.xlsx` | File **đã sửa** dòng 1–148 (tạo bởi `fixroute --sua`) |
| `fixroute-1-148.json` | Nhật ký từng thay đổi của fix route, kèm giá trị cũ |

### Cấu trúc file Excel

| Sheet | Dòng | Nội dung |
|-------|------|----------|
| Tổng hợp | 16 | Số liệu tổng, phân bố theo khối |
| Danh sách hộ | 142 | Mỗi hộ một dòng, kèm danh sách thành viên |
| Danh sách người | 598 | Mỗi người một dòng (mỗi hộ bắt đầu bằng dòng chủ hộ) |
| Cảnh báo | 122 | Lỗi trong dữ liệu — cần đối chiếu |
| Ghi chú | 287 | Trạng thái xử lý — không phải lỗi |
| DS Trường Bình (tra) | 1623 | Danh sách tra cứu gốc |
| Quy tắc | 7 | Giải thích cách tool xử lý |

**Màu nền trong Excel:**

| Màu | Ý nghĩa |
|-----|---------|
| 🟢 Xanh | Hộ 1–27, đã sửa xong |
| 🟡 Vàng | Hộ 28–78, thành viên còn bị chồn dòng |
| 🟣 Tím | Hộ 79–140, chưa cập nhật |
| 🔴 Đỏ | Cột A gõ dấu `.` hoặc chữ — cần kiểm tra lại |

---

## 5. Cách tool xử lý dữ liệu

### Cột A = STT hộ

Chỉ dòng **chủ hộ** mới ghi STT hộ ở cột A. Dòng thành viên để trống.

### Khôi phục STT hộ

Trong khu vực **dòng 149–325** (vùng bị chồn giữa hộ 20 và hộ 78), các dòng thành viên bị thiếu STT hộ. Tool đối chiếu **CCCD** với danh sách tra cứu, rồi suy ra STT hộ từ dòng chủ hộ tương ứng.

- **Ngoài khu vực 149–325: giữ nguyên**, không đoán.
- 177 người trong vùng: 173 khớp CCCD → **168 được khôi phục**, 9 giữ nguyên.

### Không đoán bừa hộ trong vùng 149–325

Trong vùng này các dòng thành viên nằm tán manh nên **vị trí không nghĩa là thuộc hộ nào**. File tra là nguồn chuẩn; tra không ra thì để trống (`sttHo = 0`) và tách thành nhóm riêng, **không gộp vào chủ hộ đang mở**.

Trước đây 9 người này bị đổ nhầm vào hộ 78 (chủ hộ ngay trước đó) — hộ 78 trong file tra chỉ có 7 người, xếp vào 18 người là sai. Nay hộ 78 còn **9 người**.

9 người này (xếp cuối file, cột A = `.`):

| Dòng | Họ tên | CCCD | Lý do |
|------|--------|------|-------|
| 158 | Lê Thị Hồng Ca | 093163002668 | khớp hộ #74 của file tra, nhưng hộ đó không có chủ hộ trong Tổ 1 |
| 212 | Nguyễn Ngọc Thống | 093089005740 | như trên |
| 220 | Ngụy Thị Minh Thư | 093192005610 | như trên |
| 305 | Nguyễn Ngọc Minh Khôi | 093216001690 | như trên |
| 314 | Nguyễn Minh Mẫn | 093219004023 | như trên |
| 186 | Nguyễn Thành Hôn | 093083010542 | không có trong file tra |
| 190 | Nguyễn Qúi Tiền | 093084007770 | không có trong file tra |
| 228 | Lê Thị Kiều Linh | 089194013856 | không có trong file tra |
| 255 | Châu Thúy Vy | 093303010632 | không có trong file tra |

### Dấu `.` ở cột A

Giữ nguyên, tách thành **3 nhóm riêng** (tổng 12 người), kèm gợi ý hộ lân cận. **Không tự gộp vào hộ nào.**

### CCCD trùng

Giữ toàn bộ các dòng, chỉ ghi cảnh báo `CCCD_TRUNG`.

### Dòng bị lệch cột

Một số dòng có CCCD nằm nhầm ở cột E, quan hệ ở cột G. Tool khôi phục CCCD nhưng **vẫn ghi cảnh báo** để bạn kiểm tra lại file gốc.

### Số liệu hiện tại

```
151 hộ / 598 người
  Hộ  1– 27:  25 hộ    85 người   (xanh  — đã sửa xong)
  Hộ 28– 78:  51 hộ   219 người   (vàng  — còn chồn dòng)
  Hộ 79–140:  63 hộ   273 người   (tím   — chưa cập nhật)
  Cột A rác:   3 hộ    12 người   (đỏ)
  Chưa xác định hộ:
              9 hộ     9 người   (9 nhóm 1 người, xếp cuối file)
```

STT hộ thiếu: **2** và **4**. STT hộ **105 xuất hiện 2 lần** — giữ nguyên cả hai.

---

## 6. Cần bạn kiểm tra

Ba điểm sau tool **không tự quyết**, cần đối chiếu với file gốc:

### a) `r160 Nguyễn Thị Tím` → hộ 36

Hộ tra cứu số 201 khớp với chủ hộ hộ 36 (`r106 Nguyễn Ngọc Anh`, CCCD khớp). Nếu thực tế `r160` thuộc hộ 20 thì cần sửa lại bản đồ hộ.

### b) Hộ `105` xuất hiện 2 lần

`r458 Phạm Văn Bé` và `r462 Hà Văn Tuyết` đều ghi cột A = 105. Giữ thành 2 hộ riêng. Cần xác nhận có phải gộp làm một không.

### c) 4 người không tra được

Không tìm thấy trong danh sách tra cứu. Nay **không còn bị gộp vào hộ 78** mà tách riêng, xếp cuối file (xem bảng ở mục 5):

| Dòng | Họ tên | CCCD |
|------|--------|------|
| 186 | Nguyễn Thành Hôn | 093083010542 |
| 190 | Nguyễn Qúi Tiền | 093084007770 |
| 228 | Lê Thị Kiều Linh | 089194013856 |
| 255 | Châu Thúy Vy | 093303010632 |

---

## 7. Cấu trúc thư mục

```
qlnd.cmd                  Lệnh chạy (điểm khởi động duy nhất)
web.cmd                   Bấm đúp để mở web + trình duyệt (dễ nhất)
run-server.bat            Bấm đúp để chỉ chạy server (Ctrl+C để dừng)
package.json
source\fileexcel\         File Excel nguồn (đặt vào đây)
output\                   Kết quả
src\lib\xlsx.js           Đọc Excel, chuẩn hóa ngày / số / tên
src\lib\to1.js           Nghiệp vụ: gom hộ, khôi phục STT hộ, cảnh báo
src\lib\grid.js          Fix route: đọc / kiểm tra / sửa từng dòng Excel
src\lib\sources.js       Tự tìm file nguồn, phát hiện file OLE2
src\lib\arrange.js       Xếp lại từng hộ: chủ hộ → thành viên + nhật ký dòng
src\cli\build.js         Chạy extract + excel
src\cli\extract.js        Xuất JSON
src\cli\excel.js          Xuất .xlsx
src\cli\xep.js            Xếp lại danh sách từ dòng lệnh
src\cli\fixroute.js       Duyệt dòng 1–148 từ dòng lệnh
src\cli\serve.js          Web server + API
src\web\                  Giao diện web (HTML/CSS/JS, không framework)
test\run.js               45 test
tools\                    Script kiểm tra dữ liệu
```

Các script trong `tools\` chạy độc lập để kiểm tra:

```
node tools/check-nhom.js          In ra từng nhóm hộ + STT bị thiếu/trùng
node tools/check-vung.js          Chi tiết vùng bị chồn dòng 149–325
node tools/check-colA.js          In cột A thô của từng dòng
node tools/check-map.js           Bản đồ hộ tra cứu → STT hộ
node tools/inspect-1-148.js       In cấu trúc dòng 1–148 của file gốc
node tools/kiem-tra-fixroute.js   Kiểm tra vòng tròn: sửa → lưu → đối chiếu file gốc
node tools/kiem-tra-xep.js        Kiểm tra file đã xếp: khối hộ liền, STT người chạy đều
node tools/so-tra.js              Đối chiếu từng hộ Tổ 1 với nhóm tương ứng trong file tra
node tools/xem-tra.js             In một đoạn file tra để xác định ranh giới mỗi hộ
node tools/tim-nguoi.js           Tìm CCCD / tên trong cả Tổ 1 và file tra
node tools/xem-xuat.js            In bất kỳ sheet nào trong file output
```

---

## 8. Khi cần sửa logic

Mọi quy tắc nghiệp vụ nằm trong **`src/lib/to1.js`**. Sửa xong nhớ chạy lại test:

```
qlnd.cmd test
```

Test phải đủ **44 pass, 0 fail** trước khi xuất lại Excel.

Một số hằng số cần biết:

| Hằng số | Vị trí | Ý nghĩa |
|----------|--------|----------|
| `RECOVER_RANGE` | `to1.js` | Phạm vi dòng được khôi phục STT hộ (149–325). Đặt `null` để khôi phục toàn bộ |
| `BLOCK_BOUNDARY` | `to1.js` | Dòng 328, chia làm 2 khối cấu trúc dữ liệu |
| `KHUNG` | `to1.js` | Tên 4 khối dữ liệu và ngưỡng STT |

Sau khi sửa phạm vi khôi phục, cần sửa lại số liệu kỳ vọng trong `test\run.js` cho khớp.

---

## 9. Xử lý sự cố

**`Khong tim thay file To 1`**
Kiểm tra `source\fileexcel\` có file Excel không. Tên file phải chứa "Tổ 1".

**`node is not recognized`**
Thư mục `%USERPROFILE%\.qlnd\node` đã bị xoá. Xem mục 2.

**`npm.ps1 cannot be loaded`**
PowerShell chặn chạy script. Dùng `npm.cmd` thay vì `npm`:
```
%USERPROFILE%\.qlnd\node\npm.cmd install
```

**Excel mở ra chậm hoặc bị treo**
File `DS Trường Bình (tra)` có 1623 dòng. Bình thường mở khoảng 2–3 giây.

**Cổng 5173 đã bị chiếm**
```
web.cmd 5200
```

**Trình duyệt không tự mở**
Mở tay: <http://localhost:5173>. Nếu vẫn không được, web chưa khởi động — xem cửa sổ lệnh có báo lỗi không.

**Muốn dừng web**
```
taskkill /FI "WINDOWTITLE eq QLND web*" /T
```

Hoặc dùng `web.cmd 5200`.

**Muốn debug**
```
qlnd.cmd extract
```
In ra từng bước: file nguồn tìm được, file bị bỏ qua, số dòng đọc được, số hộ/người, phân bố theo khối.
