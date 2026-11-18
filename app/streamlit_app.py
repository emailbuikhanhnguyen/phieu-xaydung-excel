"""Giao diện Streamlit — chạy cục bộ trên Windows (run.bat).

Luồng: 1) chọn thư mục hồ sơ -> 2) kiểm tra ghép file theo ấp ->
3) bấm Xử lý -> 4) duyệt từng ô kèm ảnh đúng dòng/ô trên phiếu ->
5) Xuất Excel kèm kết quả kiểm tra. Đóng rồi mở lại vẫn tiếp tục đúng
vị trí đã làm (tiến độ lưu trong SQLite).
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import fitz  # noqa: E402
import streamlit as st  # noqa: E402

from phieu_xaydung import __version__  # noqa: E402
from phieu_xaydung import pdf_tools  # noqa: E402
from phieu_xaydung.export import verify_workbook  # noqa: E402
from phieu_xaydung.models import RowStatus  # noqa: E402
from phieu_xaydung.pipeline import (  # noqa: E402
    export_hamlet, process_hamlet, scan_dossier)
from phieu_xaydung.store import Store  # noqa: E402

st.set_page_config(page_title="Phiếu xây dựng -> Excel", layout="wide")
st.title("Phiếu xây dựng -> Excel  (chạy cục bộ, không gọi AI)")
st.caption(f"Phiên bản {__version__} — mọi xử lý chạy trên máy, 0 lần gọi AI.")

if "dossier_dir" not in st.session_state:
    st.session_state.dossier_dir = ""
if "results" not in st.session_state:
    st.session_state.results = {}
if "review_idx" not in st.session_state:
    st.session_state.review_idx = {}

folder = st.text_input(
    "Thư mục hồ sơ (có thể là thư mục Google Drive đã đồng bộ xuống máy)",
    value=st.session_state.dossier_dir,
    placeholder=r"C:\Users\Nguyen\Drive\lan 2")
if st.button("Quét thư mục"):
    st.session_state.dossier_dir = folder
    st.session_state.dossiers = scan_dossier(folder)
    st.session_state.results = {}

dossiers = st.session_state.get("dossiers", [])
if not dossiers:
    st.info("Nhập thư mục hồ sơ rồi bấm Quét thư mục để bắt đầu.")
    st.stop()

st.subheader("1. Danh sách ấp — kiểm tra ghép file")
for d in dossiers:
    c1, c2, c3, c4 = st.columns([2, 3, 3, 2])
    c1.write(f"**{d.name}**")
    c2.write(f"PDF: {d.pdf_path.name if d.pdf_path else '— thiếu —'}")
    c3.write(f"Excel: {d.excel_path.name if d.excel_path else '— thiếu —'}")
    d.reversed_pdf = c4.checkbox("PDF quét ngược", value=d.reversed_pdf,
                                 key=f"rev-{d.name}")

out_dir = st.text_input("Thư mục xuất kết quả",
                        value=str(Path(folder) / "ket-qua") if folder else "")
if st.button("Xử lý tất cả", type="primary"):
    store = Store(Path(folder) / ".phieu-xaydung" / "state.db")
    progress = st.progress(0.0)
    for i, d in enumerate(dossiers):
        result = process_hamlet(d, store=store)
        st.session_state.results[d.name] = result
        progress.progress((i + 1) / len(dossiers))
        st.write(f"[{d.name}] {len(result.rows)} hộ, "
                 f"{len(result.notes)} mục cần đối chiếu, "
                 f"AI: 0 lần gọi.")
    store.close()

results = st.session_state.results
if not results:
    st.info("Bấm Xử lý tất cả để tự điền phần đọc được từ Excel/PDF.")
    st.stop()

st.subheader("2. Duyệt từng hộ")
hamlet = st.selectbox("Chọn ấp", list(results.keys()))
result = results[hamlet]
rows = result.rows
idx = st.session_state.review_idx.get(hamlet, 0)
# Nhảy tới hộ chưa xác nhận đầu tiên khi mở lại.
if st.session_state.review_idx.get(hamlet) is None:
    for i, r in enumerate(rows):
        if r.status != RowStatus.CONFIRMED:
            idx = i
            break
    st.session_state.review_idx[hamlet] = idx

if rows:
    c_prev, c_cur, c_next = st.columns(3)
    for col, j, label in ((c_prev, idx - 1, "Dòng trước"),
                          (c_cur, idx, "Dòng hiện tại"),
                          (c_next, idx + 1, "Dòng sau")):
        if 0 <= j < len(rows):
            r = rows[j]
            col.write(f"{label}: STT {r.stt} — {r.name or '(chưa có tên)'}")
    row = rows[idx]
    st.write(f"Hộ {idx + 1}/{len(rows)} — STT {row.stt} — "
             f"trạng thái: {row.status.value}")

    # Ảnh đúng dòng trên phiếu (nếu có PDF + vị trí dòng).
    if result.dossier.pdf_path and row.pdf_page:
        doc = fitz.open(result.dossier.pdf_path)
        page = doc[row.pdf_page - 1]
        bands = pdf_tools.row_bands(page)
        if row.pdf_row is not None and row.pdf_row < len(bands):
            y0, y1 = bands[row.pdf_row]
            rect = fitz.Rect(0, max(0, y0 - 4), page.rect.width,
                             min(page.rect.height, y1 + 4))
            st.image(pdf_tools.crop_png(doc, row.pdf_page - 1, rect, dpi=180),
                     caption=f"PDF trang {row.pdf_page}, dòng {row.pdf_row + 1}")
            sig = pdf_tools.signature_rect(page, y0, y1)
            if sig is not None:
                st.image(pdf_tools.crop_png(doc, row.pdf_page - 1, sig, dpi=200),
                         caption="Chữ ký (cột T) — kéo khung bên dưới nếu nét vượt đường kẻ")
                pad = st.slider("Nới khung chữ ký (điểm)", 0, 12, 2,
                                key=f"pad-{row.row_id}")
                if pad != 2:
                    sig2 = fitz.Rect(sig.x0 - pad, sig.y0 - pad,
                                     sig.x1 + pad, sig.y1 + pad)
                    st.image(pdf_tools.crop_png(doc, row.pdf_page - 1, sig2,
                                                dpi=200),
                             caption="Khung đã nới")
        doc.close()
    elif result.dossier.pdf_path:
        st.warning("Chưa xác định được dòng của hộ này trên PDF — "
                   "mở trang PDF để đối chiếu tay.")
    else:
        st.warning("Ấp này thiếu PDF — chỉ duyệt được theo Excel nguồn.")

    with st.form(f"edit-{row.row_id}"):
        name = st.text_input("Họ tên chủ hộ", value=row.name or "")
        c1, c2, c3 = st.columns(3)
        q3 = c1.number_input("Tổng Quý III", value=row.q3.total or 0.0, step=5.0)
        q4 = c2.number_input("Tổng Quý IV", value=row.q4.total or 0.0, step=5.0)
        year = c3.number_input("Tổng cả năm", value=row.year.total or 0.0, step=5.0)
        code = st.selectbox("Hộ có XD (R)", [None, 1, 2],
                            index=[None, 1, 2].index(row.code))
        phone = st.text_input("Số điện thoại (S)", value=row.phone or "")
        note = st.text_input("Ghi chú đối chiếu (để trống nếu không có)",
                             value="")
        col_save, col_flag = st.columns(2)
        saved = col_save.form_submit_button("Lưu & sang hộ tiếp theo (Enter)")
        flagged = col_flag.form_submit_button("Đánh dấu cần đối chiếu")
        if saved or flagged:
            row.name = name or None
            row.q3.total = q3 or None
            row.q4.total = q4 or None
            row.year.total = year or None
            row.code = code
            row.phone = phone or None
            row.status = (RowStatus.NEEDS_REVIEW if flagged
                          else RowStatus.CONFIRMED)
            store = Store(Path(folder) / ".phieu-xaydung" / "state.db")
            store.save_row(hamlet, row)
            store.close()
            if note:
                st.write(f"Đã ghi chú: {note}")
            st.session_state.review_idx[hamlet] = min(idx + 1, len(rows) - 1)
            st.rerun()

st.subheader("3. Xuất Excel")
if st.button("Xuất Excel + kiểm tra", type="primary"):
    store = Store(Path(folder) / ".phieu-xaydung" / "state.db")
    for name, res in results.items():
        path = export_hamlet(res, out_dir)
        check = verify_workbook(path)
        st.write(f"**{name}** -> {path.name} — "
                 f"{'ĐẠT' if check['ok'] else 'CÓ LỖI: ' + '; '.join(check['errors'])}")
    store.close()
    st.success("Xong. Mở thư mục xuất kết quả để xem file.")
