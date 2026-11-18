"""Ghép Excel nguồn với PDF theo STT + tên + vị trí trên trang.

Nguyên tắc:
  - Khớp đúng: cùng STT và tên trùng (sau chuẩn hóa) -> ghép.
  - STT trùng, thiếu STT, tên không khớp -> đưa vào duyệt, KHÔNG tự đoán.
  - Không tự thêm hộ để đạt số tổng ghi trên phiếu.
  - Thứ tự hiển thị cuối cùng theo STT trên phiếu (PDF là chuẩn về
    thứ tự khi hai nguồn lệch nhau).
"""

from __future__ import annotations

import unicodedata

from .excel_source import normalize_name
from .models import HouseholdRow, Provenance, ReviewNote


def _fold(s: str | None) -> str:
    if not s:
        return ""
    s = unicodedata.normalize("NFD", s.lower())
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return " ".join(s.split())


def match_rows(excel_rows: list[HouseholdRow],
               pdf_rows: list[HouseholdRow]) -> tuple[
                   list[tuple[HouseholdRow, HouseholdRow]],
                   list[HouseholdRow], list[HouseholdRow], list[ReviewNote]]:
    """Trả về (cặp đã ghép, excel chưa ghép, pdf chưa ghép, ghi chú duyệt).

    pdf_rows: các dòng đọc được từ PDF (tối thiểu có STT + vị trí trang/dòng;
    tên có thể None khi PDF quét chưa đọc tay).
    """
    notes: list[ReviewNote] = []
    by_stt: dict[int, list[HouseholdRow]] = {}
    by_name_no_stt: dict[str, list[HouseholdRow]] = {}
    for r in excel_rows:
        if r.stt is not None:
            by_stt.setdefault(r.stt, []).append(r)
        elif r.name:
            by_name_no_stt.setdefault(_fold(r.name), []).append(r)
    used_excel: set[str] = set()
    pairs: list[tuple[HouseholdRow, HouseholdRow]] = []
    pdf_unmatched: list[HouseholdRow] = []
    noted_dup_stt: set[int | None] = set()

    for p in pdf_rows:
        cands = by_stt.get(p.stt, []) if p.stt is not None else []
        cands = [c for c in cands if c.row_id not in used_excel]
        chosen = None
        if p.stt is None and p.name:
            # Hộ thiếu STT: ghép theo tên khi tên duy nhất ở cả hai phía.
            same = [c for c in by_name_no_stt.get(_fold(p.name), [])
                    if c.row_id not in used_excel]
            if len(same) == 1:
                chosen = same[0]
            elif len(same) > 1:
                notes.append(ReviewNote(
                    row_ids=[p.row_id] + [c.row_id for c in same],
                    stt_text="(trống)", name_text=p.name or "",
                    column="A",
                    reason=(f"Hộ thiếu STT trùng tên {len(same)} hộ trong "
                            f"Excel nguồn — chưa ghép tự động."),
                    pdf_page=p.pdf_page, kind="trung_lap"))
        elif len(cands) == 1:
            chosen = cands[0]
            if p.name and cands[0].name and _fold(p.name) != _fold(cands[0].name):
                notes.append(ReviewNote(
                    row_ids=[p.row_id, cands[0].row_id],
                    stt_text=str(p.stt), name_text=p.name or "",
                    column="B",
                    reason=(f"Tên trên PDF ({p.name}) khác Excel nguồn "
                            f"({cands[0].name}) — giữ tên theo PDF, cần đối chiếu."),
                    pdf_page=p.pdf_page, kind="chua_chac"))
        elif len(cands) > 1:
            # STT trùng: thử khớp theo tên, không được thì đưa vào duyệt.
            same = [c for c in cands
                    if p.name and c.name and _fold(p.name) == _fold(c.name)]
            if len(same) == 1:
                chosen = same[0]
            else:
                if p.stt not in noted_dup_stt:
                    noted_dup_stt.add(p.stt)
                    notes.append(ReviewNote(
                        row_ids=[p.row_id] + [c.row_id for c in cands],
                        stt_text=str(p.stt), name_text=p.name or "",
                        column="A",
                        reason=(f"STT {p.stt} trùng {len(cands)} hộ trong Excel "
                                f"nguồn — chưa ghép tự động."),
                        pdf_page=p.pdf_page, kind="trung_lap"))
        if chosen is not None:
            used_excel.add(chosen.row_id)
            pairs.append((chosen, p))
        else:
            pdf_unmatched.append(p)

    excel_unmatched = [r for r in excel_rows if r.row_id not in used_excel]
    for r in excel_unmatched:
        if r.stt is None:
            notes.append(ReviewNote(
                row_ids=[r.row_id], stt_text="(trống)",
                name_text=r.name or "", column="A",
                reason="Hộ trong Excel nguồn không có STT — cần đối chiếu vị trí trên phiếu.",
                kind="chua_chac"))
    return pairs, excel_unmatched, pdf_unmatched, notes


def merge_pair(excel_row: HouseholdRow, pdf_row: HouseholdRow) -> HouseholdRow:
    """Gộp một cặp đã ghép: Excel cho số liệu, PDF cho thứ tự/vị trí.

    Tên lấy theo PDF khi PDF có tên (PDF là bản gốc); số liệu lấy theo
    Excel khi PDF chưa đọc tay. Nguồn của từng trường được ghi lại.
    """
    out = HouseholdRow(
        stt=pdf_row.stt if pdf_row.stt is not None else excel_row.stt,
        stt_raw=pdf_row.stt_raw or excel_row.stt_raw,
        name=pdf_row.name or excel_row.name,
        q3=excel_row.q3, q4=excel_row.q4, year=excel_row.year,
        code=excel_row.code, code_raw=excel_row.code_raw,
        phone=excel_row.phone, phone_state=excel_row.phone_state,
        phone_raw=excel_row.phone_raw,
        pdf_page=pdf_row.pdf_page, pdf_row=pdf_row.pdf_row,
        excel_row=excel_row.excel_row,
    )
    out.sources = dict(excel_row.sources)
    if pdf_row.pdf_page is not None:
        out.sources["pdf"] = Provenance(
            "pdf-grid", f"PDF tr.{pdf_row.pdf_page}, dòng {pdf_row.pdf_row}")
    return out
