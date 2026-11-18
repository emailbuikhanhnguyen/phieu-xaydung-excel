"""Đọc Excel nguồn của từng ấp.

Excel nguồn rất đa dạng: có ấp là danh sách đầy đủ, có ấp chỉ có hộ xây
dựng, có ấp dùng tên mẫu ("Chủ hộ 1..."), có ấp lệch cột (SĐT nằm ở R,
mã 1/2 nằm ở Q). Module này chỉ ĐỌC và MÔ TẢ, không tự sửa:
  - phát hiện lệch cột, nhưng KHÔNG tự dịch toàn bộ dãy tiền
    (nhận ra 2 trường chưa đủ để biết cột tiền gốc);
  - tên mẫu, SĐT sai, mã R lạ -> đưa vào danh sách duyệt;
  - ghi rõ phạm vi: danh sách toàn ấp hay chỉ hộ có xây dựng.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

import openpyxl

from .models import HouseholdRow, Money, PhoneState, Provenance

TEMPLATE_NAME_RE = re.compile(r"^chủ hộ\s*\d+\s*$", re.IGNORECASE)
PHONE_RE = re.compile(r"^0\d{9}$")


def normalize_name(name: str | None) -> str | None:
    if name is None:
        return None
    s = " ".join(str(name).split())
    return s or None


def is_template_name(name: str | None) -> bool:
    return bool(name) and bool(TEMPLATE_NAME_RE.match(name.strip()))


def phone_state(value) -> tuple[str | None, PhoneState, str]:
    """Trả về (số đã chuẩn hóa, trạng thái, chuỗi gốc)."""
    if value is None or (isinstance(value, str) and not value.strip()):
        return None, PhoneState.BLANK_ON_FORM, ""
    raw = str(value).strip().replace(" ", "").replace(".", "")
    if PHONE_RE.match(raw):
        return raw, PhoneState.PRESENT, raw
    return None, PhoneState.UNREADABLE, raw


def parse_money(value) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def parse_code(value) -> tuple[int | None, str]:
    if value is None or value == "":
        return None, ""
    if isinstance(value, (int, float)) and int(value) in (1, 2):
        return int(value), ""
    s = str(value).strip()
    if s in ("1", "2"):
        return int(s), ""
    return None, s


@dataclass
class SourceSheet:
    name: str
    header_row: int
    columns: dict[str, int]  # vai trò -> chỉ số cột (1-based)
    rows: list[HouseholdRow] = field(default_factory=list)
    shifted: bool = False    # SĐT ở R, mã 1/2 ở Q
    scope_hint: str = "unknown"  # full | construction_only | unknown


HEADER_ALIASES = {
    "stt": ("stt", "số tt", "sott"),
    "name": ("họ tên", "ho ten", "họ và tên", "tên chủ hộ", "chu ho"),
    "code": ("hộ có xd", "có xd", "hộ có xây dựng"),
    "phone": ("số điện thoại", "sdt", "điện thoại", "so dien thoai"),
}


def _find_header(ws) -> tuple[int, dict[str, int]] | None:
    """Tìm dòng header có chứa STT + họ tên. Trả về (dòng, {vai trò: cột})."""
    for r in range(1, min(12, ws.max_row + 1)):
        cols: dict[str, int] = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if not isinstance(v, str):
                continue
            t = v.strip().lower()
            for role, aliases in HEADER_ALIASES.items():
                if any(a in t for a in aliases) and role not in cols:
                    cols[role] = c
        if "stt" in cols and "name" in cols:
            return r, cols
    return None


def _detect_money_columns(ws, header_row: int, name_col: int) -> dict[str, int]:
    """Cột tiền C:Q theo vị trí tương đối sau cột tên trong file mẫu chuẩn.

    File nguồn thường giữ đúng thứ tự 15 cột tiền sau STT+tên; khi không
    chắc, trả về rỗng để pipeline đưa vào duyệt thay vì đoán.
    """
    money: dict[str, int] = {}
    labels = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L",
              "M", "N", "O", "P", "Q"]
    start = name_col + 1
    # Kiểm tra dòng header: phải có chữ "tổng" ở cột tổng kỳ.
    ok = 0
    for i, lab in enumerate(labels):
        v = ws.cell(header_row, start + i).value
        if isinstance(v, str) and ("tổng" in v.lower() or "nhà ở" in v.lower()
                                   or "sxkd" in v.lower() or "nông" in v.lower()):
            ok += 1
    if ok >= 3:
        for i, lab in enumerate(labels):
            money[lab] = start + i
    return money


def detect_shift(rows_raw: list[dict]) -> bool:
    """Lệch cột kiểu Tây Lộc/Thanh Tịnh: SĐT nằm ở cột R, mã 1/2 ở cột Q."""
    phone_in_r = 0
    code_in_q = 0
    for r in rows_raw:
        if isinstance(r.get("R"), str) and re.match(r"^0\d{8,10}$",
                                                     r["R"].strip()):
            phone_in_r += 1
        if str(r.get("Q", "")).strip() in ("1", "2"):
            code_in_q += 1
    return phone_in_r >= 3 and code_in_q >= 3


def read_source_excel(path: str) -> list[SourceSheet]:
    """Đọc mọi sheet có bảng hộ trong Excel nguồn."""
    wb = openpyxl.load_workbook(path, data_only=False)
    out: list[SourceSheet] = []
    for ws in wb.worksheets:
        found = _find_header(ws)
        if not found:
            continue
        header_row, cols = found
        money = _detect_money_columns(ws, header_row, cols["name"])
        sheet = SourceSheet(name=ws.title, header_row=header_row,
                            columns=cols)
        raw_rows: list[dict] = []
        for r in range(header_row + 1, ws.max_row + 1):
            stt_v = ws.cell(r, cols["stt"]).value
            name_v = ws.cell(r, cols["name"]).value
            if stt_v is None and (name_v is None or str(name_v).strip() == ""):
                continue
            raw = {openpyxl.utils.get_column_letter(c): ws.cell(r, c).value
                   for c in range(1, ws.max_column + 1)}
            raw_rows.append(raw)
            row = HouseholdRow(excel_row=r)
            row.stt = int(stt_v) if isinstance(stt_v, (int, float)) else None
            row.stt_raw = "" if stt_v is None else str(stt_v).strip()
            row.name = normalize_name(name_v)
            if money:
                for lab, col in money.items():
                    v = parse_money(ws.cell(r, col).value)
                    _set_money(row, lab, v)
            if "code" in cols:
                row.code, row.code_raw = parse_code(ws.cell(r, cols["code"]).value)
            if "phone" in cols:
                row.phone, row.phone_state, row.phone_raw = phone_state(
                    ws.cell(r, cols["phone"]).value)
            row.sources["excel"] = Provenance(
                "excel", f"sheet {ws.title}, dòng {r}")
            sheet.rows.append(row)
        sheet.shifted = detect_shift(raw_rows)
        # Gợi ý phạm vi: tên sheet hoặc số dòng so với tổng ấp (nếu có).
        if "xay_dung" in ws.title.lower() or "co_xd" in ws.title.lower():
            sheet.scope_hint = "construction_only"
        out.append(sheet)
    wb.close()
    return out


def _set_money(row: HouseholdRow, lab: str, value: float | None) -> None:
    m = {"C": ("q3", "total"), "D": ("q3", "nha_o"), "E": ("q3", "sxkd"),
         "F": ("q3", "nong_lam"), "G": ("q3", "xd_khac"),
         "H": ("q4", "total"), "I": ("q4", "nha_o"), "J": ("q4", "sxkd"),
         "K": ("q4", "nong_lam"), "L": ("q4", "xd_khac"),
         "M": ("year", "total"), "N": ("year", "nha_o"),
         "O": ("year", "sxkd"), "P": ("year", "nong_lam"),
         "Q": ("year", "xd_khac")}[lab]
    setattr(getattr(row, m[0]), m[1], value)
