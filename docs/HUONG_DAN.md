# Hướng dẫn sử dụng — phieu-xaydung-excel

## 0. Dùng trên web hay trên máy?

- **Trên web (không cần cài):** mở link app Streamlit, chọn chế độ
  **Tải file lên**, upload PDF + Excel nguồn của đợt, bấm Xử lý, duyệt
  từng hộ, rồi tải file Excel kết quả về máy. Có nút **Thử ngay với dữ
  liệu giả** để xem tool chạy thế nào. Tiến độ chỉ giữ trong lúc tab
  còn mở — tải kết quả về trước khi đóng tab. Cách triển khai bản web:
  xem [TRIEN_KHAI_WEB.md](TRIEN_KHAI_WEB.md).
- **Trên máy Windows:** làm theo các mục dưới đây (chọn thư mục hồ sơ,
  tiến độ lưu SQLite, mở lại tiếp tục đúng vị trí).

Các mục 1–8 dưới đây mô tả chung cho cả hai chế độ; khác biệt duy
nhất là cách nạp hồ sơ (mục 1) và cách lấy file kết quả (mục 5).

## 1. Chuẩn bị thư mục hồ sơ

Mỗi đợt phiếu là một thư mục, chứa PDF quét và Excel nguồn (nếu có).
Đặt tên file theo ấp để tool tự ghép, ví dụ:

```
ho-so-q4/
  Dong-Nam.pdf
  Dong-Nam.xlsx
  Vinh-Chinh.pdf
  Vinh-Chinh.xlsx
  Thanh-Tinh.xlsx        (không có PDF — vẫn xử lý được phần Excel)
  My-Phuoc.xlsx
```

Tên ấp, số trang, số hộ và số dòng mỗi trang lấy theo hồ sơ từng đợt —
không hard-code danh sách ấp hay số hộ của đợt trước.

Thư mục này có thể là thư mục Google Drive đã đồng bộ xuống máy
(Drive for Desktop). Tool chỉ đọc file nguồn, không sửa file gốc.

## 2. Mở công cụ

Chạy `run.bat` (lần đầu chạy `setup.bat` trước). Trình duyệt mở tại
http://localhost:8501.

Nhập đường dẫn thư mục hồ sơ, ví dụ:

```
C:\Users\Nguyen\Google Drive\lan 2
```

Bấm **Quét thư mục**. Bảng liệt kê từng ấp kèm file PDF/Excel đã ghép.
Kiểm tra nhanh:

- Tên ấp đúng chưa? Nếu tên file và địa bàn trong hồ sơ khác nhau
  (vd file tên "lan-bac" nhưng trong phiếu là "Lân Bắc"), sửa tên ấp
  trong bảng trước khi xử lý.
- Thiếu PDF hay thiếu Excel? Thiếu PDF thì tool chỉ xử lý theo Excel
  nguồn, chữ ký để trống và ghi rõ "thiếu PDF" trong file kết quả.
- PDF quét ngược (trang cuối của ấp nằm ở đầu file)? Tick ô
  **PDF quét ngược** của ấp đó.

## 3. Xử lý

Bấm **Xử lý tất cả**. Tool làm, theo thứ tự:

1. Đọc Excel nguồn: tự điền STT, tên, tiền, mã R, SĐT đọc được.
   Phát hiện lệch cột (SĐT ở R, mã 1/2 ở Q) nhưng không tự dịch dãy
   tiền; tên mẫu ("Chủ hộ 1") và SĐT sai đưa vào Cần đối chiếu.
2. Phân loại trang PDF: danh sách hộ / tổng hợp / trống / xác nhận.
   Dò đường kẻ lưới để xác định ranh giới dòng và cột chữ ký.
3. Ghép Excel với PDF theo STT + tên + vị trí trang. STT trùng,
   thiếu STT, tên không khớp → đưa vào duyệt, không tự đoán.
4. Cắt chữ ký theo cột T của từng dòng. Ô chữ ký trống (không có mực)
   → cột T để trống thật.
5. Kiểm tra số liệu: lệch tổng trong kỳ, tổng ghi tay vs cộng dòng,
   SĐT trùng, mã R mâu thuẫn với số tiền, giá trị lặp.

Kết quả mỗi ấp hiện số hộ, số mục cần đối chiếu, số lần gọi AI (= 0).

## 4. Duyệt từng hộ

Chọn ấp, tool nhảy tới hộ chưa xác nhận đầu tiên. Mỗi hộ hiển thị:

- **Dòng trước – dòng hiện tại – dòng sau**: để phát hiện ghép lệch hộ
  (chữ ký/tên của hộ khác lọt sang).
- **Ảnh đúng dòng trên phiếu** (cắt từ PDF) bên cạnh giá trị dự kiến.
- **Ảnh chữ ký** (cột T). Nếu nét chữ vượt đường kẻ, dùng thanh trượt
  **Nới khung chữ ký** để kéo lại khung trên ảnh gốc.

Sửa trực tiếp các ô: tên, tổng 3 kỳ, mã R, SĐT. Lưu ý:

- Tổng năm là số độc lập — không tự tính bằng QIII + QIV.
- SĐT khó đọc: để trống ô S, ghi lý do vào ô ghi chú đối chiếu.
  Phân biệt "phiếu bỏ trống" (không ghi gì) với "có ghi nhưng chưa
  đọc chắc".
- Bấm **Lưu & sang hộ tiếp theo** (Enter) để xác nhận, hoặc
  **Đánh dấu cần đối chiếu** để quay lại sau.

Tiến độ lưu ngay vào SQLite. Đóng trình duyệt giữa chừng rồi mở lại,
tool tiếp tục đúng hộ đang làm dở.

## 5. Xuất Excel

Bấm **Xuất Excel + kiểm tra**. Mỗi ấp ra 1 file trong thư mục xuất:

```
Phieu_Thu_Thap_Xay_Dung_Ap_Dong_Nam_Q4_2026.xlsx
```

- **Trên web:** bấm nút tải từng file (hoặc tải tất cả dạng ZIP) ngay
  sau khi xuất — đóng tab là mất file.
- **Trên máy:** file nằm trong thư mục xuất đã chọn.

File có 4 sheet: **Danh sách hộ**, **Bản quét** (ảnh toàn trang),
**Cần đối chiếu**, **Tổng hợp theo tờ**. Tool tự kiểm tra trước khi
giao: 0 ảnh ở A:S, tổng dùng SUM, S là text, R chỉ 1/2/trống.
File nào không đạt sẽ báo lỗi và không ghi đè file cũ.

## 6. Chạy lại cùng hồ sơ

- Ô đã xác nhận được giữ nguyên, không xử lý lại.
- Sửa khung cắt hoặc thay PDF: chỉ xử lý lại phần bị ảnh hưởng.
- Nguồn thay đổi dưới một ô đã duyệt: ô đó bị đánh dấu
  **cần kiểm tra lại** để duyệt lại tay.
- Muốn làm lại từ đầu một ấp: xóa file `.phieu-xaydung/state.db`
  trong thư mục hồ sơ (hoặc xóa riêng các dòng của ấp đó trong giao
  diện — sẽ bổ sung).

## 7. Chạy hàng loạt bằng lệnh (không cần mở giao diện)

```bat
.venv\Scripts\python -m phieu_xaydung.cli scan data\ho-so-q4
.venv\Scripts\python -m phieu_xaydung.cli process data\ho-so-q4 --out data\ket-qua
.venv\Scripts\python -m phieu_xaydung.cli process data\ho-so-q4 --out data\ket-qua --hamlet "Dong Nam"
.venv\Scripts\python -m phieu_xaydung.cli verify data\ket-qua\Phieu_....xlsx
```

Lệnh `process` in ra thống kê lượt chạy, trong đó `ai_calls` luôn là 0
ở chế độ cục bộ.

## 8. Xử lý sự cố

| Hiện tượng | Cách xử lý |
|---|---|
| `run.bat` báo chưa cài đặt | Chạy `setup.bat` trước, kiểm tra Python 3.11+ đã có trong PATH. |
| Tool không nhận file trong thư mục | Chỉ nhận `.pdf`, `.xlsx`, `.xls` ở thư mục gốc (không quét thư mục con). Đặt file đúng tên ấp. |
| Trang PDF bị nhận diện sai loại | Dùng tick "PDF quét ngược" nếu cần; trang trắng/tổng hợp sai loại không ảnh hưởng dữ liệu hộ. |
| Ranh giới dòng lệch (chữ ký ghép nhầm hộ) | Ở màn hình duyệt, xem 3 dòng liên tiếp; sửa khung chữ ký bằng thanh trượt; ghi chú để kiểm tra lại. |
| Số tiền không khớp tổng ghi tay | Đọc ghi chú trong sheet Cần đối chiếu (ghi rõ chênh lệch + trang nguồn). Chỉ sửa số khi xác định được ô nhập sai. |
| Muốn thêm OCR đọc chữ viết tay | Cài Tesseract (https://github.com/UB-Mannheim/tesseract/wiki) rồi bật tùy chọn OCR trong cấu hình (đang phát triển — hiện mặc định tắt). |
