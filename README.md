# phieu-xaydung-excel

Chuyển phiếu thu thập thông tin xây dựng (PDF quét + Excel nguồn) thành
file Excel chuẩn theo mẫu Phú Hiệp — **chạy cục bộ trên Windows, không
gọi API AI**.

Đóng gói từ quy trình đã dùng cho 9 ấp đợt Q4/2026 (127 trang quét,
3.721 hộ). Phần tái sử dụng nhiều nhất: nhận diện bảng, ghép đúng dòng
hộ, nhập đúng cột tiền, cắt chữ ký, kiểm tra file Excel trước khi xuất.

## Chạy trên Windows (lần đầu)

1. Cài Python 3.11+ từ https://www.python.org/downloads/
   (tích chọn **Add Python to PATH** khi cài).
2. Tải repo về (Code → Download ZIP, hoặc `git clone`), giải nén.
3. Chạy **`setup.bat`** — chờ cài xong thư viện (cần mạng, 1 lần duy nhất).
4. Chạy **`run.bat`** — trình duyệt mở giao diện công cụ.

Sau đó chỉ cần `run.bat` mỗi lần dùng.

## Luồng dùng

1. **Chọn thư mục hồ sơ** — thư mục chứa PDF + Excel nguồn của đợt
   (có thể là thư mục Google Drive đã đồng bộ xuống máy). Tool liệt kê
   từng ấp, ghép PDF/Excel theo tên; tên file và địa bàn khác nhau thì
   chọn lại tay. Tick **PDF quét ngược** nếu trang cuối nằm ở đầu file
   (như Đông Nam đợt trước).
2. **Bấm Xử lý** — tool tự điền phần đọc được từ Excel nguồn và từ
   text layer của PDF (nếu có). Phần chưa chắc đưa vào **Cần đối chiếu**.
3. **Duyệt từng hộ** — màn hình hiển thị ảnh đúng dòng/ô trên phiếu
   bên cạnh giá trị dự kiến. Sửa được dữ liệu, kéo lại khung chữ ký,
   xác nhận hoặc đánh dấu cần đối chiếu. Có chế độ xem dòng trước –
   dòng hiện tại – dòng sau để phát hiện ghép lệch hộ.
4. **Bấm Xuất Excel** — tạo 1 file cho từng ấp, kèm kết quả kiểm tra.
   Đóng rồi mở lại vẫn tiếp tục đúng vị trí đã làm (tiến độ lưu trong
   SQLite tại `.phieu-xaydung/state.db` trong thư mục hồ sơ).

Chạy hàng loạt không cần giao diện:

```bat
.venv\Scripts\python -m phieu_xaydung.cli scan data\ho-so-q4
.venv\Scripts\python -m phieu_xaydung.cli process data\ho-so-q4 --out data\ket-qua
.venv\Scripts\python -m phieu_xaydung.cli verify data\ket-qua\Phieu_....xlsx
```

## Định dạng đầu ra

Theo mẫu `Phu-Hiep-Quy-IV-2026.xlsx`, sheet **Danh sách hộ**:

| Cột | Nội dung |
|---|---|
| A – STT | Ghi theo phiếu. Giữ được STT trống, trùng, bị gạch, thiếu số. Mỗi hộ có mã dòng nội bộ riêng nên trùng STT vẫn là 2 bản ghi độc lập. |
| B – Họ tên | Tiếng Việt có dấu. Cảnh báo tên mẫu ("Chủ hộ 1"...). |
| C:G – Quý III | C tổng kỳ; D:G nhà ở, SXKD, nông lâm TS, XD khác. |
| H:L – Ước Quý IV | H tổng kỳ; I:L 4 nhóm tương ứng. |
| M:Q – Ước cả năm | M tổng năm; N:Q 4 nhóm tương ứng. |
| R – Hộ có XD | 1 / 2 / trống. Mã khác được cảnh báo và lưu giá trị gốc. |
| S – Điện thoại | Text, giữ số 0 đầu. Khó đọc → để trống + ghi Cần đối chiếu. Phân biệt "phiếu bỏ trống" với "có ghi nhưng chưa đọc chắc". |
| T – Chữ ký | Chỉ ảnh chữ ký gốc của đúng hộ. Phiếu không ký → để trống. |

Quy tắc cứng: **A:S có 0 ảnh** (kể cả dòng tổng); ô tiền tính được bằng
SUM; ô nguồn trống xuất thành ô trống thật. Sheet **Bản quét** chứa ảnh
trang gốc; **Cần đối chiếu** chứa các mục chưa chắc; **Tổng hợp theo
tờ** chứa tổng từng trang.

## Kiểm tra số liệu (theo lỗi thật đã gặp)

1. So sánh trong từng kỳ: C vs D+E+F+G, H vs I+J+K+L, M vs N+O+P+Q.
   **Không mặc định cả năm = Quý III + Quý IV** — tổng năm ghi tay là
   số độc lập.
2. Giữ riêng 3 loại số: tổng ghi trên phiếu, tổng cộng từ các dòng hộ,
   chênh lệch (triệu đồng, kèm trang nguồn). Không tự cân số.
3. Phát hiện lệch cột (SĐT ở R, mã 1/2 ở Q) — chỉ cảnh báo, không tự
   dịch toàn bộ dãy tiền.
4. Giá trị lặp (vd C=D) chỉ ghi chú, không tự xóa — có thể hợp lệ
   (hộ chỉ xây nhà ở). Dòng tổng ghi rõ "tổng tạm tính, chưa gồm X hộ"
   khi có dòng chưa xác định bị loại khỏi phép cộng.
5. Phân biệt phạm vi: toàn ấp / chỉ hộ có XD / số dòng thực tế đã nhập.
6. Đếm rõ: số mục ghi chú, số ô, số hộ bị ảnh hưởng — một ghi chú
   chung cho 40 hộ truy ra được 40 dòng cụ thể.

## Tiết kiệm token / thời gian

- **Mặc định 0 lần gọi AI.** Đọc Excel trực tiếp, dùng text layer của
  PDF, xử lý ảnh và cắt chữ ký bằng code. Chữ viết tay duyệt thủ công
  theo ô. (OCR cục bộ qua Tesseract là tùy chọn, chưa bật mặc định.)
- Tiến độ lưu SQLite theo dấu vân tay file: chạy lại cùng hồ sơ dùng
  kết quả đã lưu, giữ nguyên ô đã xác nhận; sửa khung cắt/thay PDF chỉ
  xử lý lại phần bị ảnh hưởng; nguồn đổi dưới ô đã duyệt → đánh dấu
  kiểm tra lại.
- Cuối mỗi lượt báo: số trang đã xử lý, số ô dùng lại, số ô cần duyệt,
  số lần gọi AI (= 0 ở chế độ cục bộ) và token thực dùng (nếu có).

## Dữ liệu giả & kiểm thử

- `tools/make_fake_data.py` tạo hồ sơ giả ấp "An Bình" (32 hộ, PDF
  quét ngược, STT trùng, thiếu STT, tên mẫu, lệch cột, SĐT sai, tổng
  năm khác QIII+QIV) + `ground-truth.json`.
- `tests/` — 10 test pytest theo bảng nghiệm thu. Chạy:
  `python -m pytest tests/ -q`
- `examples/sample-output/` — file Excel mẫu xuất từ dữ liệu giả.
- **Không có dữ liệu thật trong repo**: PDF thật, Excel thật, ảnh chữ
  ký, SĐT thật, file đã xử lý và khóa API đều nằm ngoài Git
  (xem `.gitignore`). Thư mục dữ liệu cục bộ: `data/`.

> **Clone mới:** file PDF/Excel giả và file Excel mẫu là file nhị phân,
> không có sẵn ngay sau khi clone. Chạy 2 lệnh để tạo lại:
>
> ```bat
> .venv\Scripts\python tools\make_fake_data.py --out tests\fixtures\fake-dossier
> .venv\Scripts\python -m phieu_xaydung.cli process tests\fixtures\fake-dossier --out examples\sample-output
> ```
>
> Sau đó `python -m pytest tests/ -q` chạy xanh.

## Cấu trúc

```
src/phieu_xaydung/
  layout.py        bố cục mẫu Phú Hiệp (header, độ rộng cột, style)
  pdf_tools.py     phân loại trang, dò lưới, trích text, cắt chữ ký
  excel_source.py  đọc Excel nguồn, phát hiện lệch cột / tên mẫu / SĐT sai
  matcher.py       ghép Excel–PDF theo STT + tên + vị trí
  validate.py      kiểm tra trong kỳ, tổng, phạm vi, trùng lặp
  store.py         SQLite: tiến độ, dấu vân tay, khung cắt, trạng thái duyệt
  export.py        xuất Excel 4 sheet + kiểm tra trước khi giao
  pipeline.py      quy trình chung cho CLI và giao diện
  cli.py           lệnh chạy hàng loạt
app/streamlit_app.py   giao diện duyệt
tools/make_fake_data.py  tạo dữ liệu giả
docs/              hướng dẫn chi tiết + kết quả test
```

Xem thêm: [docs/HUONG_DAN.md](docs/HUONG_DAN.md),
[docs/KET_QUA_TEST.md](docs/KET_QUA_TEST.md).

## Giới hạn đã biết

- Chữ viết tay trên PDF quét **vẫn cần người đọc tay** ở màn hình duyệt;
  tool chỉ tự điền từ Excel nguồn và từ text layer (PDF form điện tử).
- PDF quét nghiêng mạnh, lưới mờ hoặc mất nét có thể dò sai ranh giới
  dòng — màn hình duyệt cho chỉnh thứ tự trang và ranh giới dòng tay.
- Chưa kiểm tra thực tế trên Windows (phát triển và test trên Linux);
  `setup.bat`/`run.bat` viết theo chuẩn Windows nhưng chưa chạy thử
  trên máy Windows thật.
