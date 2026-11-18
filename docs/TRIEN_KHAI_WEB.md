# Triển khai bản web (Streamlit Cloud)

Bản web của tool chạy trên Streamlit Community Cloud (miễn phí), nối
thẳng vào repo GitHub này. User mở link là dùng được, không cần cài
Python hay tải repo về.

## 1. Triển khai (làm 1 lần)

1. Mở https://share.streamlit.io/ và đăng nhập bằng tài khoản GitHub
   `emailbuikhanhnguyen` (tài khoản đang giữ repo này).
2. Bấm **Create app** (hoặc **New app**).
3. Điền:
   - **Repository:** `emailbuikhanhnguyen/phieu-xaydung-excel`
   - **Branch:** `main`
   - **Main file path:** `app/streamlit_app.py`
   - **App URL:** chọn tên gợi nhớ, ví dụ `phieu-xaydung-excel`
     (link sẽ là `https://phieu-xaydung-excel.streamlit.app`)
4. Bấm **Deploy!** Chờ 2–5 phút cho máy chủ cài thư viện
   (`requirements.txt`) và khởi động app.

Từ lần sau, mỗi lần push code mới lên nhánh `main`, Streamlit Cloud
tự khởi động lại app với bản mới (có thể mất 1–2 phút).

> **Lưu ý (lỗi gặp 18/11, phiếu PXD-001):** Streamlit Cloud chỉ git pull
> rồi chạy lại script trong cùng tiến trình Python, nên module trong
> `src/` có thể vẫn là bản cũ trong `sys.modules` — app đã tự nạp lại
> module (xem khối `importlib.reload` đầu `app/streamlit_app.py`). Sau
> khi merge một PR sửa code, bấm **Reboot app** một lần (trang quản trị
> app → menu ⋮ → **Reboot app**), rồi kiểm tra số phiên bản hiện trên
> app (đầu trang + sidebar) khớp với sha của `main`.

## 2. Kiểm tra sau khi triển khai

1. Mở link app, chọn chế độ **Tải file lên (dùng trên web)**.
2. Bấm **Thử ngay với dữ liệu giả (ấp An Bình, 32 hộ)**.
3. Bấm **Xử lý tất cả** — phải thấy `32 hộ, 14 mục cần đối chiếu`.
4. Bấm **Xuất Excel + kiểm tra** — phải thấy `ĐẠT` và nút tải file.
5. Tải file Excel về, mở kiểm tra: 32 hộ, chữ ký ở cột T, 0 ảnh ở A:S.

Nếu bước nào lỗi, xem log tại trang quản trị app
(share.streamlit.io → app → **Manage app** → **Logs**).

## 3. Giới hạn của bản web (so với bản local)

| | Bản web (Streamlit Cloud) | Bản local (run.bat) |
|---|---|---|
| Nạp hồ sơ | Upload file qua trình duyệt (mỗi file tối đa 200MB) | Chọn thư mục trên máy (kể cả thư mục Drive đã đồng bộ) |
| Tiến độ duyệt | Giữ trong lúc tab còn mở; đóng tab là mất | Lưu SQLite trong thư mục hồ sơ, mở lại tiếp tục đúng vị trí |
| Dữ liệu | File upload lên máy chủ Streamlit Cloud trong lúc xử lý | Chỉ ở trên máy mình |
| Tài nguyên | Gói miễn phí: 1 GB RAM, app ngủ khi không ai dùng một lúc (mở lại mất ~1 phút) | Tùy máy |
| Xử lý hồ sơ lớn | Chậm hơn (dò lưới raster ~1,5 giây/trang); đợt 127 trang nên chia nhỏ từng ấp | Nhanh hơn |

## 4. Gỡ / dừng bản web

Vào share.streamlit.io → app → **Settings** → **Delete app**.
Repo GitHub và bản local không bị ảnh hưởng.
