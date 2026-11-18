# Kết quả kiểm thử

## 1. Test tự động (dữ liệu giả)

Chạy `python -m pytest tests/ -q` — **10/10 test PASS**.

| # | Test | Nội dung |
|---|---|---|
| 1 | template giống mẫu | Bố cục file xuất khớp mẫu Phú Hiệp (header, độ rộng cột, style) |
| 2 | PDF ngược + STT trùng + thiếu STT | Giữ đủ 32 dòng, truy được trang gốc, không ghép nhầm chữ ký vì trùng STT |
| 3 | SĐT text / 0 đầu | Phân biệt SĐT có 0 đầu, khó đọc, bỏ trống |
| 4 | Excel thiếu PDF + tên mẫu | Vẫn xử lý phần có căn cứ, thể hiện đúng phạm vi |
| 5 | Tổng năm độc lập | Tổng năm khác QIII+QIV vẫn giữ số nguồn, không tự cân |
| 6 | Giá trị lặp + lệch cột | Không tự xóa giá trị lặp; lệch cột chỉ cảnh báo |
| 7 | Đóng/mở lại | Giữ dữ liệu đã duyệt + khung chữ ký (SQLite) |
| 8 | Chạy lại 0 gọi AI | `ai_calls=0`, `ai_tokens=0` |
| 9 | File xuất verify đạt | 32 dòng, dòng tổng =SUM, S format "@", 0 ảnh A:S |
| 10 | Matcher không ghép sai | Tên khác nhau không ghép chung dù cùng STT |

File mẫu đầu ra: `examples/sample-output/Phieu_Thu_Thap_Xay_Dung_Ap_An_Binh_Q4_2026.xlsx`
(32 hộ, 32 chữ ký ở cột T, 0 ảnh ở A:S, dòng tổng dùng SUM).
Thống kê lượt chạy CLI trên hồ sơ giả: 5 trang, 150 ô tự điền,
14 mục cần đối chiếu, `ai_calls=0`, `ai_tokens=0`.

## 2. Chạy thử riêng bằng dữ liệu thật Q4/2026 (chỉ thống kê ẩn danh)

Đối chiếu với file Excel đã kiểm tra xong của đợt Q4/2026.
**Không có tên, SĐT hay ảnh chữ ký thật nào trong repo** — chỉ số lượng.

### 2.1. Một trang có chữ viết tay (20 hộ)

| Chỉ số | Kết quả |
|---|---|
| Dải dòng phát hiện | 20 / 21 (20 hộ + 1 dòng CỘNG) — thiếu dải dòng CỘNG cuối trang |
| Khung chữ ký có mực | 19 / 19 khung kiểm tra được (1 dòng chưa kiểm tra do thiếu dải) |
| Excel nguồn tự điền đúng | **0 / 20** — toàn bộ tên trong Excel nguồn là tên mẫu ("Chủ hộ 1"...), phải đọc tay hết |

### 2.2. Hai trang danh sách liên tiếp (31 + 30 hộ), PDF quét ngược

| Chỉ số | Trang 1 (31 hộ) | Trang 2 (30 hộ) |
|---|---|---|
| Dải dòng phát hiện | 36 / 32 — **thừa 4 dải giả** ở vùng tiêu đề cột (cần xóa tay ở màn hình duyệt) | 31 / 31 |
| Khung chữ ký có mực | 35 / 35 khung (gồm 4 dải giả chứa chữ tiêu đề cột; 31 dòng thật đều có mực) | 30 / 30 |
| Ranh giới chuyển trang | STT cuối trang 1 và STT đầu trang 2 liên tục (82 → 83), không trùng, không thiếu | — |

Lưu ý: thứ tự trang trong file PDF quét ngược **không phải** thứ tự đọc —
trang cuối của ấp nằm ở đầu file, và có trang đầu tờ mới xen giữa
(trang STT 1–20 nằm sau trang STT 421–451 trong file). Cần tick
"PDF quét ngược" hoặc chỉnh thứ tự trang tay ở màn hình duyệt.

### 2.3. Excel nguồn tự điền (61 hộ trên 2 trang Đông Nam)

| Kết quả | Số hộ |
|---|---|
| Tự điền đúng (tên + 3 kỳ tiền + mã R + SĐT khớp file đã kiểm tra) | **0** |
| Có trong Excel nguồn nhưng sai khác (phải sửa tay) | 22 |
| Không có trong Excel nguồn (phải đọc tay từ phiếu) | 39 |

Kết luận: với đợt này, Excel nguồn chỉ là gợi ý — phần lớn phải đọc tay
từ ảnh phiếu. Tool không tự đoán số liệu chưa có căn cứ.

## 3. Giới hạn đã biết từ chạy thử thật

1. **Dò lưới trên bản scan:** PDF đợt này là ảnh thuần raster (không text
   layer, không vector). Dò đường kẻ bằng run-length + chia khối tìm được
   20–21/21 dải dòng trên trang tốt, nhưng thừa dải giả ở tiêu đề và
   thiếu dải ở trang mờ; dò cột dọc có trang chỉ được 0–2 cột (trang tốt
   được 20 cột). Tool có dự phòng cột chữ ký mặc định (11,5% cuối trang)
   nên vẫn cắt được chữ ký; màn hình duyệt cho chỉnh ranh giới dòng tay.
2. **Chữ viết tay không tự đọc được** — phải duyệt tay từng ô (đúng thiết
   kế: chế độ cục bộ 0 gọi AI).
3. **Chưa kiểm tra trên Windows thật** — phát triển và test trên Linux;
   `setup.bat`/`run.bat` chưa chạy thử trên máy Windows.
4. **Chưa test hiệu năng toàn bộ 127 trang** — chạy thử 3 trang mất ~13
   giây (chủ yếu dò lưới raster ~1,5 giây/trang).

## 4. Môi trường kiểm thử

- Python 3.12.3, PyMuPDF 2.1, openpyxl, Pillow, numpy, streamlit 1.67.0,
  pytest (Linux VM).
- Giao diện Streamlit smoke-test: `healthz` trả 200, trang chủ tải được.
