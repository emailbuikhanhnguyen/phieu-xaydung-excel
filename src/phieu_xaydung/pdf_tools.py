"""Đọc PDF phiếu quét: phân loại trang, dò đường kẻ -> ranh giới dòng,
trích text layer khi có, cắt chữ ký.

Không gọi AI/OCR ở đây. Với PDF quét (không có text layer), việc đọc chữ
viết tay để cho màn hình duyệt thủ công; hàm chỉ trả về hình ảnh từng dòng
và khung chữ ký để người dùng đọc.
"""

from __future__ import annotations

import hashlib
import io
import re

import fitz  # PyMuPDF

from .models import PageInfo

# Từ khóa phân loại trang (chữ không dấu cũng gặp vì font quét).
SUMMARY_WORDS = ("tổng hợp", "tổng cộng", "cộng", "tổng số hộ")
CONFIRM_WORDS = ("xác nhận", "chủ tịch", "ký tên", "điều tra viên")
TITLE_WORDS = ("phiếu thu thập", "công trình xây dựng")


def fingerprint(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()[:16]


def open_pdf(path: str) -> fitz.Document:
    return fitz.open(path)


def page_text(page: fitz.Page) -> str:
    return page.get_text("text") or ""


def _norm(text: str) -> str:
    import unicodedata
    s = unicodedata.normalize("NFD", text.lower())
    return "".join(c for c in s if unicodedata.category(c) != "Mn")


def classify_page(page: fitz.Page) -> str:
    """danh_sach | tong_hop | trang_trong | xac_nhan | quet.

    "quet" = trang scan thuần raster (không text layer, không vector):
    không đọc được bằng code, cần người dùng nhập loại trang + khoảng STT
    ở màn hình duyệt.
    """
    text = _norm(page_text(page))
    drawings = page.get_drawings()
    has_grid = len(drawings) > 20
    if not text.strip() and not has_grid:
        if page.get_images(full=True):
            return "quet"
        return "trang_trong"
    if any(w in text for w in CONFIRM_WORDS) and "stt" not in text:
        return "xac_nhan"
    if any(w in text for w in SUMMARY_WORDS) and "họ tên" not in text:
        return "tong_hop"
    if "stt" in text or "họ tên" in text or has_grid:
        return "danh_sach"
    if not text.strip():
        return "trang_trong"
    return "tong_hop"


def horizontal_lines(page: fitz.Page) -> list[float]:
    """Tung độ (y) của các đường kẻ ngang dài, đã gộp trùng và sắp xếp."""
    ys: list[float] = []
    for d in page.get_drawings():
        r = d["rect"]
        if r.width > page.rect.width * 0.5 and r.height < 2.5:
            ys.append(round((r.y0 + r.y1) / 2, 1))
    ys.sort()
    merged: list[float] = []
    for y in ys:
        if not merged or y - merged[-1] > 2.0:
            merged.append(y)
    return merged


def vertical_lines(page: fitz.Page) -> list[float]:
    """Hoành độ (x) của các đường kẻ dọc dài."""
    xs: list[float] = []
    for d in page.get_drawings():
        r = d["rect"]
        if r.height > page.rect.height * 0.4 and r.width < 2.5:
            xs.append(round((r.x0 + r.x1) / 2, 1))
    xs.sort()
    merged: list[float] = []
    for x in xs:
        if not merged or x - merged[-1] > 2.0:
            merged.append(x)
    return merged


def _long_run_rows(dark: np.ndarray, min_run: int, axis: int = 1) -> np.ndarray:
    """Với mỗi hàng (axis=1) hoặc cột (axis=0): tỉ lệ pixel nằm trong
    một đoạn mực ngang/dọc liên tục dài >= min_run.

    Lọc run-length này phân biệt đường kẻ (chạy dài liên tục) với chữ
    viết tay (nét ngắn, đứt đoạn) — hình chiếu mực thô không làm được
    vì trang nhiều chữ viết tay.
    """
    import numpy as np
    a = dark if axis == 1 else dark.T
    h, w = a.shape
    frac = np.zeros(h)
    for i in range(h):
        row = a[i]
        d = np.diff(np.concatenate(([0], row.view(np.int8), [0])))
        starts = np.where(d == 1)[0]
        ends = np.where(d == -1)[0]
        if len(starts) == 0:
            continue
        lengths = ends - starts
        long_mask = lengths >= min_run
        if long_mask.any():
            frac[i] = lengths[long_mask].sum() / w
    return frac


def _block_line_points(dark: np.ndarray, n_blocks: int = 8,
                       frac_threshold: float = 0.55) -> list[tuple[float, float]]:
    """Tìm các điểm (x_tâm_khối, y) thuộc đường kẻ ngang.

    Chia trang thành n_blocks khối dọc: trong mỗi khối hẹp, đường kẻ
    nghiêng nhẹ vẫn nằm gọn trong 1–2 hàng pixel nên hình chiếu hàng
    của khối đạt tỉ lệ cao (>0.55). Mỗi đỉnh hình chiếu là một điểm
    ứng viên của đường kẻ tại tâm khối đó.
    """
    import numpy as np
    h, w = dark.shape
    bw = w // n_blocks
    points: list[tuple[float, float]] = []
    for b in range(n_blocks):
        block = dark[:, b * bw:(b + 1) * bw]
        prof = block.mean(axis=1)
        ys = np.where(prof > frac_threshold)[0]
        if len(ys) == 0:
            continue
        start = prev = int(ys[0])
        for y in ys[1:]:
            y = int(y)
            if y - prev > 3:
                points.append(((b + 0.5) * bw, (start + prev) / 2))
                start = y
            prev = y
        points.append(((b + 0.5) * bw, (start + prev) / 2))
    return points


def _fit_lines_from_points(points: list[tuple[float, float]],
                           page_w: float, tol: float = 18.0,
                           min_blocks: int = 3) -> list[float]:
    """Gom các điểm ứng viên thành đường kẻ: các điểm cùng một đường
    nằm gần nhau theo y (cho phép nghiêng). Mỗi đường cần điểm ở ít
    nhất min_blocks khối. Trả về y của đường tại tâm trang.
    """
    import numpy as np
    if not points:
        return []
    ys = sorted(p[1] for p in points)
    clusters: list[list[tuple[float, float]]] = []
    for p in sorted(points, key=lambda t: t[1]):
        if clusters and p[1] - clusters[-1][-1][1] <= tol:
            clusters[-1].append(p)
        else:
            clusters.append([p])
    lines: list[float] = []
    for cl in clusters:
        if len({round(p[0]) for p in cl}) < min_blocks:
            continue
        xs = np.array([p[0] for p in cl], dtype=float)
        yy = np.array([p[1] for p in cl], dtype=float)
        if len(cl) >= 4 and xs.max() - xs.min() > page_w * 0.4:
            a, b = np.polyfit(xs, yy, 1)
            lines.append(float(a * page_w / 2 + b))
        else:
            lines.append(float(np.median(yy)))
    return sorted(lines)


def raster_lines(page: fitz.Page, dpi: int = 150,
                 threshold: float = 0.2) -> tuple[list[float], list[float]]:
    """Dò đường kẻ trên trang scan thuần raster (ảnh, không vector).

    Phương pháp khối: chia trang thành 8 khối dọc, tìm đường kẻ trong
    từng khối (chịu được nghiêng nhẹ 0.3–1 độ làm nhòe hình chiếu
    toàn trang), rồi gom các điểm thành đường bằng fit tuyến tính.
    Trả về (y các đường ngang, x các đường dọc) theo tọa độ điểm PDF.
    Dọc dùng run-length toàn trang (đường dọc ít bị nghiêng làm nhòe).
    """
    import numpy as np
    pix = page.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY)
    arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w)
    dark = arr < 160
    scale = 72.0 / dpi

    points = _block_line_points(dark)
    h_lines = [y * scale for y in
               _fit_lines_from_points(points, float(pix.w))]
    # Đường dọc: run-length toàn trang.
    min_run = max(100, int(pix.h * 0.12))
    v_frac = _long_run_rows(dark, min_run, axis=0)
    v_lines = []
    ys = np.where(v_frac > threshold)[0]
    if len(ys):
        start = prev = int(ys[0])
        for y in ys[1:]:
            y = int(y)
            if y - prev > 3:
                v_lines.append((start + prev) / 2 * scale)
                start = y
            prev = y
        v_lines.append((start + prev) / 2 * scale)
    return h_lines, v_lines


def _merge_close(vals: list[float], gap: float = 2.0) -> list[float]:
    merged: list[float] = []
    for v in sorted(vals):
        if not merged or v - merged[-1] > gap:
            merged.append(v)
    return merged


def _fill_by_pitch(ys: list[float]) -> list[float]:
    """Điền các đường kẻ bị mất bằng bước dòng trung vị.

    Trên bản photocopy, nhiều đường kẻ ngang mờ/đứt nên không dò được;
    nhưng các đường còn lại cách nhau đúng bội số của bước dòng (~22pt).
    Khoảng nào gần bằng k*bước (k>=2) thì chèn k-1 đường cách đều.
    """
    import statistics
    if len(ys) < 4:
        return ys
    diffs = [b - a for a, b in zip(ys, ys[1:]) if 12 <= b - a <= 45]
    if not diffs:
        return ys
    pitch = statistics.median(diffs)
    out: list[float] = [ys[0]]
    for a, b in zip(ys, ys[1:]):
        gap = b - a
        k = round(gap / pitch)
        if k >= 2 and abs(gap - k * pitch) <= 0.4 * pitch:
            for i in range(1, k):
                out.append(a + i * pitch)
        out.append(b)
    return out


def data_band_range(bands: list[tuple[float, float]]) -> tuple[int, int]:
    """Chỉ số (start, end) các dải thuộc vùng dữ liệu của bảng.

    Vùng dữ liệu là đoạn dài nhất ở cuối trang mà các dải kề nhau cách
    nhau 18–32pt (bước dòng của bảng: phiếu thật ~21pt, dữ liệu giả
    ~28pt). Dùng để tách dải hộ khỏi các
    đường kẻ tiêu đề phía trên — các đường này cũng cách nhau ~21pt
    nên một vài dải giả ở tiêu đề có thể lọt vào (người dùng xóa tay
    ở màn hình duyệt).
    Trả về (start, len(bands)) để cắt bands[start:]; (0, 0) khi không
    tìm được đoạn nào.
    """
    if not bands:
        return (0, 0)
    ys = [b[0] for b in bands]
    best_start, best_len = len(ys) - 1, 1
    start, length = len(ys) - 1, 1
    for i in range(len(ys) - 2, -1, -1):
        if 18 <= ys[i + 1] - ys[i] <= 32:
            length += 1
            start = i
        else:
            if length > best_len:
                best_start, best_len = start, length
            start, length = i, 1
    if length > best_len:
        best_start, best_len = start, length
    return (best_start, len(bands))


def row_bands(page: fitz.Page, min_height: float = 12.0) -> list[tuple[float, float]]:
    """Các dải dòng (y0, y1) suy từ đường kẻ ngang.

    Ưu tiên vector; trang scan thuần raster thì dò bằng ảnh (raster_lines)
    rồi điền đường kẻ mất bằng bước dòng. Trả về rỗng khi không dò được
    lưới — màn hình duyệt cho chỉnh tay.
    """
    ys = horizontal_lines(page)
    if len(ys) < 3:
        ys, _ = raster_lines(page)
        ys = _merge_close(ys)
        ys = _fill_by_pitch(ys)
    bands = [(ys[i], ys[i + 1]) for i in range(len(ys) - 1)
             if ys[i + 1] - ys[i] >= min_height]
    return bands


def column_bounds(page: fitz.Page) -> list[tuple[float, float]]:
    """Các dải cột (x0, x1) suy từ đường kẻ dọc (vector, rồi raster)."""
    xs = vertical_lines(page)
    if len(xs) < 3:
        _, xs = raster_lines(page)
        xs = _merge_close(xs)
    return [(xs[i], xs[i + 1]) for i in range(len(xs) - 1)]


def text_in_rect(page: fitz.Page, rect: fitz.Rect) -> str:
    return page.get_text("text", clip=rect).strip()


def extract_text_grid(page: fitz.Page) -> list[dict]:
    """Trích text theo lưới: mỗi dòng -> dict {row, y0, y1, cells:[text...]}.

    Chỉ dùng được khi PDF có text layer. Với PDF quét, cells rỗng và
    pipeline đánh dấu trang là 'cần duyệt tay'.
    """
    bands = row_bands(page)
    cols = column_bounds(page)
    out: list[dict] = []
    for i, (y0, y1) in enumerate(bands):
        cells = []
        for (x0, x1) in cols:
            cells.append(text_in_rect(page, fitz.Rect(x0, y0, x1, y1)))
        out.append({"row": i, "y0": y0, "y1": y1, "cells": cells})
    return out


def render_page_png(doc: fitz.Document, pageno: int, dpi: int = 150) -> bytes:
    page = doc[pageno]
    pix = page.get_pixmap(dpi=dpi)
    return pix.tobytes("png")


def crop_png(doc: fitz.Document, pageno: int,
             rect: fitz.Rect, dpi: int = 200) -> bytes:
    """Cắt một vùng trên trang thành PNG (dùng cho ô chữ ký, ô nghi vấn)."""
    page = doc[pageno]
    pix = page.get_pixmap(dpi=dpi, clip=rect)
    return pix.tobytes("png")


def signature_rect(page: fitz.Page, y0: float, y1: float,
                   pad: float = 2.0) -> fitz.Rect | None:
    """Khung chữ ký = cột cuối cùng của lưới, trong dải dòng (y0, y1).

    Cột chữ ký là cột phải nhất; nét chữ vượt đường kẻ vẫn nằm trong khung
    nhờ pad, người dùng có thể kéo lại khung trên ảnh gốc.

    Dự phòng: nếu không dò được đường dọc (bản scan mờ), dùng cột
    chữ ký mặc định của mẫu phiếu — 11,5% chiều rộng cuối trang
    (mẫu phiếu cố định: chữ ký luôn là cột phải cùng).
    """
    cols = column_bounds(page)
    if len(cols) >= 2:
        x0, x1 = cols[-1]
    else:
        x0, x1 = page.rect.width * 0.885, page.rect.width
    rect = fitz.Rect(x0 + pad, y0 + pad, x1 - pad, y1 - pad)
    if rect.width <= 0 or rect.height <= 0:
        return None
    return rect


def analyze_pdf(path: str, reversed_order: bool = False) -> list[PageInfo]:
    """Phân loại toàn bộ trang. reversed_order=True khi PDF quét ngược
    (trang cuối của ấp nằm ở đầu file, như Đông Nam): thứ tự hiển thị
    đảo lại, số trang gốc vẫn giữ để truy vết.
    """
    data = open(path, "rb").read()
    doc = fitz.open(stream=data, filetype="pdf")
    infos: list[PageInfo] = []
    order = list(range(len(doc)))
    if reversed_order:
        order = order[::-1]
    for pos, pageno in enumerate(order):
        page = doc[pageno]
        kind = classify_page(page)
        bands = row_bands(page)
        info = PageInfo(number=pageno + 1, kind=kind, rows=len(bands),
                        fingerprint=fingerprint(data))
        if kind == "danh_sach" and bands:
            grid = extract_text_grid(page)
            stts = []
            for g in grid:
                if g["cells"]:
                    m = re.search(r"\d+", g["cells"][0])
                    if m:
                        stts.append(int(m.group()))
            if stts:
                info.first_stt, info.last_stt = min(stts), max(stts)
                info.rows = len(stts)  # chỉ đếm dòng có STT, không đếm dòng header
        infos.append(info)
    doc.close()
    return infos


def hamlet_total_from_summary(page: fitz.Page) -> int | None:
    """Đọc 'Tổng số hộ của ấp: N hộ' trên trang tổng hợp (nếu có)."""
    text = page_text(page)
    m = re.search(r"[Tt]ổng số hộ[^0-9]*(\d+)\s*hộ", text)
    return int(m.group(1)) if m else None


def written_totals_from_summary(page: fitz.Page) -> dict[str, float | None]:
    """Đọc các tổng ghi tay trên trang tổng hợp (nếu ghi rõ số)."""
    text = page_text(page)
    out: dict[str, float | None] = {"q3": None, "q4": None, "year": None}
    for key, pat in (
            ("q3", r"[Qq]uý\s*III[^0-9]*([\d.,]+)"),
            ("q4", r"[Qq]uý\s*IV[^0-9]*([\d.,]+)"),
            ("year", r"cả năm[^0-9]*([\d.,]+)")):
        m = re.search(pat, text)
        if m:
            try:
                out[key] = float(m.group(1).replace(".", "").replace(",", "."))
            except ValueError:
                pass
    return out


def ink_bbox(page: fitz.Page, rect: fitz.Rect) -> fitz.Rect | None:
    """Hộp bao nét mực trong một vùng (dùng để kiểm tra chữ ký có thật
    hay ô trống). Trả về None khi vùng gần như trắng."""
    pix = page.get_pixmap(dpi=150, clip=rect)
    samples = pix.samples
    n = pix.w * pix.h
    dark = 0
    for i in range(0, len(samples), pix.n):
        if samples[i] < 100:
            dark += 1
            if dark > max(20, n // 200):
                return rect
    return None
