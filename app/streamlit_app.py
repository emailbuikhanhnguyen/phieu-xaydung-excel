"""Giao diện Streamlit — chạy cục bộ trên Windows (run.bat) hoặc trên web
(Streamlit Cloud: mở link là dùng được).

Hai chế độ:
  - "Tải file lên (dùng trên web)": upload PDF + Excel nguồn qua trình
    duyệt, xử lý ngay trong phiên làm việc, tải file kết quả về máy.
    Tiến độ duyệt được giữ trong lúc tab còn mở; đóng tab là mất
    (tải file kết quả về trước khi đóng).
  - "Thư mục trên máy (chạy local)": chọn thư mục hồ sơ trên máy
    (có thể là thư mục Google Drive đã đồng bộ), tiến độ lưu trong
    SQLite tại thư mục hồ sơ, đóng/mở lại tiếp tục đúng vị trí.

Luồng chung: 1) nạp hồ sơ -> 2) kiểm tra ghép file theo ấp ->
3) bấm Xử lý -> 4) duyệt từng ô kèm ảnh đúng dòng/ô trên phiếu ->
5) Xuất Excel kèm kết quả kiểm tra. Không gọi AI.
"""

from __future__ import annotations

import io
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT / "src"))

import fitz  # noqa: E402
import streamlit as st  # noqa: E402

# Streamlit Cloud chỉ git pull rồi chạy lại script trong cùng tiến trình
# Python: module trong src/ đã nằm trong sys.modules từ lần deploy trước
# nên import thường sẽ lấy bản CŨ (lỗi ImportError sau deploy 684e52b5).
# ensure_fresh_modules() chỉ nạp lại khi code trên đĩa thật sự đổi — nạp
# lại vô điều kiện ở mỗi rerun sẽ tạo enum RowStatus mới trong khi
# session_state vẫn giữ RowStatus cũ (so sánh enum theo identity).
import phieu_xaydung  # noqa: E402
_CODE_CHANGED = phieu_xaydung.ensure_fresh_modules()  # noqa: E402

from phieu_xaydung import build_version  # noqa: E402
from phieu_xaydung import pdf_tools  # noqa: E402
from phieu_xaydung.export import verify_workbook  # noqa: E402
from phieu_xaydung.models import HouseholdRow, Provenance, RowStatus  # noqa: E402
from phieu_xaydung.pipeline import (  # noqa: E402
    export_hamlet, process_hamlet, scan_dossier)
from phieu_xaydung.store import Store  # noqa: E402

_VERSION = build_version()  # chạy mỗi rerun (chỉ vài lệnh git, rẻ) để số
# phiên bản luôn đúng ngay sau deploy; không cache vì tiến trình không
# khởi động lại khi Streamlit Cloud deploy bản mới.

st.set_page_config(page_title="Phiếu xây dựng -> Excel", layout="wide")
st.title("Phiếu xây dựng -> Excel")
st.caption(f"Phiên bản {_VERSION} — mọi xử lý chạy bằng code, 0 lần gọi AI.")
st.sidebar.write(f"Phiên bản {_VERSION}")

for key, default in (("work_dir", ""), ("results", {}),
                     ("review_idx", {}), ("dossiers", []),
                     ("exported", [])):
    if key not in st.session_state:
        st.session_state[key] = default

if _CODE_CHANGED:
    # Code vừa đổi dưới chân kết quả đang duyệt: các object cũ (vd enum
    # RowStatus) không còn đúng nữa — xóa kết quả, giữ lại hồ sơ đã nạp
    # để người dùng chỉ cần bấm Xử lý tất cả chạy lại.
    st.session_state["results"] = {}
    st.session_state["review_idx"] = {}
    st.session_state["exported"] = []
    st.warning("Code trên máy chủ vừa cập nhật — bấm Xử lý tất cả để "
               "chạy lại với bản mới.")

mode = st.radio(
    "Chế độ",
    ["Tải file lên (dùng trên web)", "Thư mục trên máy (chạy local)"],
    horizontal=True)
is_web = mode.startswith("Tải file lên")

if is_web:
    uploads = st.file_uploader(
        "Chọn PDF và Excel nguồn của đợt (nhiều file, mỗi file tối đa 200MB)",
        type=["pdf", "xlsx", "xls"], accept_multiple_files=True)
    c1, c2 = st.columns(2)
    if c1.button("Nạp file đã chọn", type="primary", disabled=not uploads):
        tmp = Path(tempfile.mkdtemp(prefix="phieu-web-"))
        for f in uploads:
            (tmp / f.name).write_bytes(f.getbuffer())
        st.session_state.work_dir = str(tmp)
        st.session_state.dossiers = scan_dossier(tmp)
        st.session_state.results = {}
        st.session_state.exported = []
        st.rerun()
    if c2.button("Thử ngay với dữ liệu giả (ấp An Bình, 32 hộ)"):
        tmp = Path(tempfile.mkdtemp(prefix="phieu-demo-"))
        fake = tmp / "fake-dossier"
        fake.mkdir()
        subprocess.run(
            [sys.executable,
             str(REPO_ROOT / "tools" / "make_fake_data.py"),
             "--out", str(fake)], check=True)
        st.session_state.work_dir = str(fake)
        st.session_state.dossiers = scan_dossier(fake)
        st.session_state.results = {}
        st.session_state.exported = []
        st.rerun()
    st.info("Lưu ý khi dùng trên web: file được xử lý ngay trong phiên làm "
            "việc này và không lưu lại sau khi đóng tab — hãy tải file kết "
            "quả về máy trước khi đóng. Đừng tải lên hồ sơ thật có SĐT/chữ "
            "ký nếu chưa chắc chắn về nơi lưu trữ của máy chủ web.")
else:
    folder = st.text_input(
        "Thư mục hồ sơ (có thể là thư mục Google Drive đã đồng bộ xuống máy)",
        value=st.session_state.work_dir,
        placeholder=r"C:\Users\Nguyen\Drive\lan 2")
    if st.button("Quét thư mục"):
        st.session_state.work_dir = folder
        st.session_state.dossiers = scan_dossier(folder)
        st.session_state.results = {}
        st.session_state.exported = []

work_dir = Path(st.session_state.work_dir) if st.session_state.work_dir else None
dossiers = st.session_state.get("dossiers", [])
if not dossiers:
    if is_web:
        st.info("Tải file PDF/Excel lên rồi bấm Nạp file đã chọn — hoặc bấm "
                "Thử ngay với dữ liệu giả để xem tool chạy thế nào.")
    else:
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

if is_web:
    out_dir = work_dir / "ket-qua"
else:
    out_dir = st.text_input("Thư mục xuất kết quả",
                            value=str(Path(work_dir) / "ket-qua"))

if st.button("Xử lý tất cả", type="primary"):
    store = Store(work_dir / ".phieu-xaydung" / "state.db")
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
    if st.button("Thêm hộ mới (trang ít hộ, dòng bị thiếu)"):
        new_row = HouseholdRow(status=RowStatus.NEEDS_REVIEW)
        new_row.sources["pdf"] = Provenance(
            "nhap-tay", "Thêm tay ở màn hình duyệt")
        rows.append(new_row)
        store = Store(work_dir / ".phieu-xaydung" / "state.db")
        store.save_row(hamlet, new_row)
        store.close()
        st.session_state.review_idx[hamlet] = len(rows) - 1
        st.rerun()

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
        stt_val = st.number_input("STT theo phiếu (0 = phiếu không ghi STT)",
                                  min_value=0, step=1,
                                  value=row.stt if row.stt is not None else 0)
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
        col_save, col_flag, col_del = st.columns(3)
        saved = col_save.form_submit_button("Lưu & sang hộ tiếp theo (Enter)")
        flagged = col_flag.form_submit_button("Đánh dấu cần đối chiếu")
        deleted = col_del.form_submit_button("Xóa dòng này")
        if deleted:
            store = Store(work_dir / ".phieu-xaydung" / "state.db")
            store.delete_row(row.row_id, hamlet)
            store.close()
            result.rows.pop(idx)
            st.session_state.review_idx[hamlet] = max(
                0, min(idx, len(result.rows) - 1))
            st.rerun()
        if saved or flagged:
            row.stt = int(stt_val) if stt_val > 0 else None
            row.name = name or None
            row.q3.total = q3 or None
            row.q4.total = q4 or None
            row.year.total = year or None
            row.code = code
            row.phone = phone or None
            row.status = (RowStatus.NEEDS_REVIEW if flagged
                          else RowStatus.CONFIRMED)
            store = Store(work_dir / ".phieu-xaydung" / "state.db")
            store.save_row(hamlet, row)
            store.close()
            if note:
                st.write(f"Đã ghi chú: {note}")
            st.session_state.review_idx[hamlet] = min(idx + 1, len(rows) - 1)
            st.rerun()

st.subheader("3. Xuất Excel")
if st.button("Xuất Excel + kiểm tra", type="primary"):
    store = Store(work_dir / ".phieu-xaydung" / "state.db")
    exported = []
    for name, res in results.items():
        path = export_hamlet(res, out_dir)
        check = verify_workbook(path)
        st.write(f"**{name}** -> {path.name} — "
                 f"{'ĐẠT' if check['ok'] else 'CÓ LỖI: ' + '; '.join(check['errors'])}")
        exported.append(str(path))
    store.close()
    st.session_state.exported = exported
    if is_web:
        st.success("Xong — tải file kết quả về máy bên dưới trước khi đóng tab.")
    else:
        st.success("Xong. Mở thư mục xuất kết quả để xem file.")

if st.session_state.exported:
    st.subheader("Tải file kết quả")
    paths = [Path(p) for p in st.session_state.exported if Path(p).exists()]
    if len(paths) > 1:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for p in paths:
                zf.write(p, arcname=p.name)
        st.download_button("Tải tất cả (ZIP)", buf.getvalue(),
                           file_name="phieu-xaydung-ket-qua.zip",
                           mime="application/zip")
    for p in paths:
        st.download_button(f"Tải {p.name}", p.read_bytes(),
                           file_name=p.name,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key=f"dl-{p.name}")
