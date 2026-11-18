"""Kiểm tra số liệu theo các lỗi đã gặp thực tế ở đợt Q4/2026.

1. So sánh trong từng kỳ: C vs D+E+F+G, H vs I+J+K+L, M vs N+O+P+Q.
   Thiếu thành phần hoặc có ô chưa đọc được -> thể hiện đúng tình trạng
   thiếu. KHÔNG mặc định cả năm = Quý III + Quý IV.
2. Giữ riêng 3 loại số: tổng ghi trên phiếu, tổng cộng từ các dòng hộ,
   chênh lệch (kèm đơn vị triệu đồng, quy ước dấu, trang nguồn).
3. Phát hiện lệch cột (SĐT ở R, mã 1/2 ở Q) — chỉ cảnh báo, không tự dịch.
4. Giá trị lặp (vd C=D) chỉ ghi chú, không tự xóa: có thể hợp lệ
   (hộ chỉ xây nhà ở).
5. Phân biệt phạm vi: toàn ấp / chỉ hộ có XD / số dòng thực tế đã nhập.
6. Đếm rõ: số mục ghi chú, số ô, số hộ bị ảnh hưởng (một ghi chú chung
   cho 40 hộ phải truy ra được 40 dòng cụ thể).
"""

from __future__ import annotations

from collections import Counter

from .excel_source import is_template_name
from .models import HouseholdRow, Money, ReviewNote


def check_period(m: Money, label: str, stt, name) -> ReviewNote | None:
    """Một kỳ: tổng ghi tay khác tổng thành phần -> ghi chú, giữ cả hai số."""
    if m.total is None:
        return None
    comp = m.component_sum()
    if comp is None:
        return None  # không có thành phần nào -> không kết luận được
    if abs(m.total - comp) > 1e-9:
        return ReviewNote(
            stt_text="" if stt is None else str(stt),
            name_text=name or "",
            column=label,
            reason=(f"Tổng {label} ghi tay {m.total:g} khác tổng thành phần "
                    f"{comp:g} (chênh {m.total - comp:+g} triệu đồng) — "
                    f"giữ cả hai số, chỉ sửa khi xác định được ô nhập sai."),
            kind="lech_tong")
    return None


def check_row(row: HouseholdRow) -> list[ReviewNote]:
    notes: list[ReviewNote] = []
    for money, label in ((row.q3, "C"), (row.q4, "H"), (row.year, "M")):
        n = check_period(money, label, row.stt, row.name)
        if n is not None:
            n.row_ids = [row.row_id]
            n.pdf_page = row.pdf_page
            notes.append(n)
    if is_template_name(row.name):
        notes.append(ReviewNote(
            row_ids=[row.row_id], stt_text=str(row.stt or ""),
            name_text=row.name or "", column="B",
            reason="Tên mẫu trong Excel nguồn (vd 'Chủ hộ 1') — không phải tên thật, cần đọc lại từ phiếu.",
            pdf_page=row.pdf_page, kind="chua_chac"))
    if row.code_raw:
        notes.append(ReviewNote(
            row_ids=[row.row_id], stt_text=str(row.stt or ""),
            name_text=row.name or "", column="R",
            reason=(f"Mã hộ có XD không hợp lệ: '{row.code_raw}' "
                    f"(chỉ nhận 1, 2 hoặc trống) — đã lưu giá trị gốc để đối chiếu."),
            pdf_page=row.pdf_page, kind="chua_chac"))
    if row.has_money() and row.code == 2:
        notes.append(ReviewNote(
            row_ids=[row.row_id], stt_text=str(row.stt or ""),
            name_text=row.name or "", column="R",
            reason="Hộ có số tiền xây dựng nhưng mã R=2 (không XD) — mâu thuẫn, cần đối chiếu ô C trên phiếu.",
            pdf_page=row.pdf_page, kind="chua_chac"))
    if row.code == 1 and not row.has_money():
        notes.append(ReviewNote(
            row_ids=[row.row_id], stt_text=str(row.stt or ""),
            name_text=row.name or "", column="R",
            reason="Mã R=1 (có XD) nhưng không có số tiền nào ở cả ba kỳ.",
            pdf_page=row.pdf_page, kind="chua_chac"))
    return notes


def check_duplicate_values(rows: list[HouseholdRow]) -> list[ReviewNote]:
    """Giá trị lặp ở hai cột kề nhau (vd C=D, K=L): ghi chú, không xóa."""
    notes: list[ReviewNote] = []
    pairs = [(("q3", "total"), ("q3", "nha_o"), "C=D"),
             (("q4", "total"), ("q4", "nha_o"), "H=I"),
             (("year", "total"), ("year", "nha_o"), "M=N"),
             (("q4", "nong_lam"), ("q4", "xd_khac"), "K=L"),
             (("q3", "nha_o"), ("q3", "sxkd"), "D=E")]
    groups: dict[str, list[HouseholdRow]] = {label: [] for _, _, label in pairs}
    for r in rows:
        for (m1, f1), (m2, f2), label in pairs:
            v1 = getattr(getattr(r, m1), f1)
            v2 = getattr(getattr(r, m2), f2)
            if v1 is not None and v1 == v2 and v1 != 0:
                groups[label].append(r)
    for label, rs in groups.items():
        if len(rs) >= 3:  # lặp hệ thống mới ghi chú chung
            notes.append(ReviewNote(
                row_ids=[r.row_id for r in rs],
                stt_text=f"{len(rs)} hộ",
                column=label,
                reason=(f"Giá trị lặp {label} ở {len(rs)} hộ — có thể hợp lệ "
                        f"(vd hộ chỉ xây nhà ở) nên không tự xóa; kiểm tra lại "
                        f"khi có bản giấy."),
                kind="trung_lap"))
    return notes


def check_phones(rows: list[HouseholdRow]) -> list[ReviewNote]:
    """SĐT: trùng số giữa các hộ, số khó đọc (đã để trống S)."""
    notes: list[ReviewNote] = []
    by_phone: dict[str, list[HouseholdRow]] = {}
    for r in rows:
        if r.phone:
            by_phone.setdefault(r.phone, []).append(r)
    for phone, rs in by_phone.items():
        if len(rs) > 1:
            notes.append(ReviewNote(
                row_ids=[r.row_id for r in rs],
                stt_text=", ".join(str(r.stt) for r in rs if r.stt is not None),
                column="S",
                reason=f"SĐT {phone} trùng ở {len(rs)} hộ — kiểm tra lại trên phiếu.",
                kind="trung_lap"))
    for r in rows:
        if r.phone_state.value == "co_ghi_chua_doc_chac":
            notes.append(ReviewNote(
                row_ids=[r.row_id], stt_text=str(r.stt or ""),
                name_text=r.name or "", column="S",
                reason=(f"SĐT ghi tay khó đọc (đọc được: '{r.phone_raw}') — "
                        f"đã để trống ô S, cần đối chiếu bản giấy."),
                pdf_page=r.pdf_page, kind="chua_chac"))
    return notes


def sum_period(rows: list[HouseholdRow], which: str) -> dict:
    """Tổng cộng từ các dòng hộ cho một kỳ + số dòng có tổng kỳ đó."""
    total = 0.0
    counted = 0
    missing = 0
    for r in rows:
        m: Money = getattr(r, which)
        if m.total is not None:
            total += m.total
            counted += 1
        else:
            missing += 1
    return {"tong": total, "so_dong_co_tong": counted,
            "so_dong_chua_co_tong": missing}


def check_totals(rows: list[HouseholdRow],
                 written: dict[str, float | None],
                 hamlet_total_households: int | None) -> list[ReviewNote]:
    """So sánh tổng ghi trên phiếu với tổng cộng từ các dòng hộ."""
    notes: list[ReviewNote] = []
    for which, label in (("q3", "Quý III"), ("q4", "Quý IV"),
                         ("year", "cả năm")):
        w = written.get(which)
        if w is None:
            continue
        s = sum_period(rows, which)
        if abs(w - s["tong"]) > 1e-9:
            notes.append(ReviewNote(
                column=f"tổng {label}",
                reason=(f"Tổng {label} ghi trên phiếu {w:g} triệu đồng, "
                        f"cộng từ {s['so_dong_co_tong']} dòng hộ được "
                        f"{s['tong']:g} (chênh {w - s['tong']:+g}). "
                        f"Giữ cả hai số, không tự cân."),
                kind="lech_tong"))
    if hamlet_total_households is not None and len(rows) != hamlet_total_households:
        notes.append(ReviewNote(
            column="số hộ",
            reason=(f"Phiếu tổng hợp ghi {hamlet_total_households} hộ, "
                    f"đã nhập {len(rows)} dòng — chưa đủ, không tự thêm hộ."),
            kind="pham_vi"))
    return notes


def count_impact(notes: list[ReviewNote]) -> dict:
    """Đếm rõ: số mục, số ô, số hộ bị ảnh hưởng."""
    households = set()
    for n in notes:
        households.update(n.row_ids)
    kinds = Counter(n.kind for n in notes)
    return {"so_muc": len(notes),
            "so_ho_bi_anh_huong": len(households),
            "theo_loai": dict(kinds)}
