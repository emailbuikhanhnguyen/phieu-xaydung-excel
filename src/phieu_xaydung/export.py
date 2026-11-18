"""Xuất file Excel kết quả theo bố cục mẫu Phú Hiệp.

Sheet "Danh sách hộ": A:S chỉ có giá trị (số/text/công thức SUM), 0 ảnh;
chữ ký chỉ ở cột T. Sheet "Bản quét": ảnh toàn trang. Sheet "Cần đối
chiếu": ghi chú + ảnh cắt minh họa (nếu có). Sheet "Tổng hợp theo tờ":
tổng từng trang + đối chiếu 3 loại số (ghi tay / cộng dòng / chênh lệch).

Tổng kỳ C/H/M: dùng =SUM(...) khi có thành phần; giữ nguyên số ghi tay
khi tổng kỳ khác tổng thành phần (không tự cân số). Dòng tổng cuối:
"TỔNG CỘNG (N HỘ)" — khi có dòng chưa xác định bị loại khỏi phép cộng,
ghi rõ "tổng tạm tính, chưa gồm X hộ" trong nhãn.
"""

from __future__ import annotations

import io
from pathlib import Path

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment

from . import layout
from .layout import (
    BODY_FONT, CELL_BORDER, CENTER, DATA_ROW_HEIGHT, FIRST_DATA_ROW,
    LEFT, RIGHT, TOTAL_FONT,
)
from .models import HouseholdRow, Money, ReviewNote

MONEY_LETTERS = {
    "q3": ("C", ["D", "E", "F", "G"]),
    "q4": ("H", ["I", "J", "K", "L"]),
    "year": ("M", ["N", "O", "P", "Q"]),
}


def _money_cells(row: HouseholdRow, which: str) -> dict[str, float | None]:
    m: Money = getattr(row, which)
    total_col, comp_cols = MONEY_LETTERS[which]
    out = {total_col: m.total}
    for col, val in zip(comp_cols, m.components()):
        out[col] = val
    return out


def _write_row(ws, excel_row: int, row: HouseholdRow):
    ws.cell(excel_row, 1).value = row.stt  # None -> ô trống thật
    ws.cell(excel_row, 1).alignment = CENTER
    ws.cell(excel_row, 2).value = row.name
    ws.cell(excel_row, 2).alignment = LEFT
    for which in ("q3", "q4", "year"):
        total_col, comp_cols = MONEY_LETTERS[which]
        cells = _money_cells(row, which)
        for col in [total_col, *comp_cols]:
            c = ws[f"{col}{excel_row}"]
            c.value = cells[col]
            c.number_format = "#,##0"
            c.alignment = RIGHT
    ws.cell(excel_row, 18).value = row.code
    ws.cell(excel_row, 18).alignment = CENTER
    s = ws.cell(excel_row, 19)
    s.value = row.phone  # text, giữ số 0 đầu
    s.number_format = "@"
    s.alignment = CENTER
    for col in range(1, 21):
        c = ws.cell(excel_row, col)
        c.font = BODY_FONT
        c.border = CELL_BORDER
    ws.row_dimensions[excel_row].height = DATA_ROW_HEIGHT


def _total_label(n_rows: int, excluded: int) -> str:
    label = f"TỔNG CỘNG ({n_rows} HỘ)"
    if excluded:
        label += f" — tổng tạm tính, chưa gồm {excluded} hộ chưa xác định"
    return label


def build_workbook(hamlet: str, rows: list[HouseholdRow],
                   notes: list[ReviewNote],
                   page_images: list[tuple[int, bytes]] | None = None,
                   signatures: list[tuple[int, bytes]] | None = None,
                   page_totals: list[dict] | None = None,
                   excluded_from_total: int = 0,
                   **template_kwargs) -> openpyxl.Workbook:
    wb = layout.build_template(hamlet, **template_kwargs)
    ws = wb[layout.SHEET_LIST]
    r = FIRST_DATA_ROW
    for row in rows:
        _write_row(ws, r, row)
        r += 1
    # Chữ ký: chỉ ảnh gốc, neo đúng ô T của từng hộ.
    if signatures:
        from PIL import Image as PILImage
        for excel_row, png in signatures:
            img = XLImage(io.BytesIO(png))
            with PILImage.open(io.BytesIO(png)) as pil:
                w0, h0 = pil.size
            img.width = 190
            img.height = min(48, max(20, int(190 * h0 / max(w0, 1))))
            ws.add_image(img, f"T{excel_row}")
    total_row = r
    ws.cell(total_row, 2).value = _total_label(len(rows), excluded_from_total)
    ws.cell(total_row, 2).font = TOTAL_FONT
    ws.cell(total_row, 2).alignment = LEFT
    for which in ("q3", "q4", "year"):
        total_col, comp_cols = MONEY_LETTERS[which]
        for col in [total_col, *comp_cols]:
            c = ws[f"{col}{total_row}"]
            if total_row > FIRST_DATA_ROW:
                c.value = f"=SUM({col}{FIRST_DATA_ROW}:{col}{total_row - 1})"
            else:
                # 0 hộ: không có dải nào để cộng — ghi 0 thay vì công thức
                # =SUM(C7:C6) (vừa ngược dải vừa tự tham chiếu tới chính ô).
                c.value = 0
            c.number_format = "#,##0"
            c.font = TOTAL_FONT
            c.alignment = RIGHT
            c.border = CELL_BORDER
    for col in range(1, 21):
        ws.cell(total_row, col).border = CELL_BORDER
    ws.row_dimensions[total_row].height = DATA_ROW_HEIGHT

    # --- Bản quét: ảnh toàn trang (không giới hạn ảnh ở đây) ---
    ws_scan = wb.create_sheet(layout.SHEET_SCAN)
    if page_images:
        for i, (pageno, png) in enumerate(page_images):
            img = XLImage(io.BytesIO(png))
            img.width = 1100
            img.height = int(1100 * img.height / img.width)
            ws_scan.add_image(img, f"A{i * 60 + 1}")
            ws_scan.cell(i * 60 + 1, 1).value = f"Trang {pageno} (bản quét gốc)"

    # --- Cần đối chiếu ---
    ws_rev = wb.create_sheet(layout.SHEET_REVIEW)
    ws_rev.append(["STT", "Họ tên", "Cột", "Lý do", "Trang PDF", "Số hộ liên quan"])
    for n in notes:
        ws_rev.append([n.stt_text, n.name_text, n.column, n.reason,
                       n.pdf_page, len(n.row_ids)])
    for col, w in zip("ABCDEF", [10, 28, 10, 110, 10, 14]):
        ws_rev.column_dimensions[col].width = w
    ws_rev.column_dimensions["D"].alignment = Alignment(wrap_text=True,
                                                        vertical="top")

    # --- Tổng hợp theo tờ ---
    ws_sum = wb.create_sheet(layout.SHEET_SUMMARY)
    ws_sum.append(["Trang", "STT từ–đến", "Số hộ", "Tổng QIII (trđ)",
                   "Tổng QIV (trđ)", "Tổng cả năm (trđ)", "Ghi chú"])
    if page_totals:
        for t in page_totals:
            ws_sum.append([t.get("trang"), t.get("stt_range"), t.get("so_ho"),
                           t.get("q3"), t.get("q4"), t.get("year"),
                           t.get("ghi_chu", "")])
    return wb


def verify_workbook(path: str | Path) -> dict:
    """Kiểm tra file đã xuất theo 4 tiêu chí bàn giao.

    Trả về dict {ok, errors[], stats{}}. Dùng trước khi giao file.
    """
    wb = openpyxl.load_workbook(path)
    errors: list[str] = []
    ws = wb[layout.SHEET_LIST]
    images = ws._images
    bad_cols = sorted({img.anchor._from.col + 1 for img in images
                       if img.anchor._from.col + 1 <= 19})
    if bad_cols:
        errors.append(f"Còn {len(images)} ảnh ở cột A:S "
                      f"(cột {bad_cols}) — phải bằng 0.")
    sig_rows = {img.anchor._from.row + 1 for img in images
                if img.anchor._from.col + 1 == 20}
    n_data = 0
    total_row = None
    for r in range(FIRST_DATA_ROW, ws.max_row + 1):
        v = ws.cell(r, 2).value
        if isinstance(v, str) and v.startswith("TỔNG CỘNG"):
            total_row = r
            break
        n_data += 1
    if total_row is None:
        errors.append("Không tìm thấy dòng TỔNG CỘNG.")
    else:
        for col in layout.MONEY_COLS:
            v = ws[f"{col}{total_row}"].value
            is_sum = isinstance(v, str) and v.startswith("=SUM(")
            # 0 hộ: dòng tổng là số 0, không phải công thức.
            if not is_sum and not (n_data == 0 and v == 0):
                errors.append(f"Dòng tổng cột {col} không phải công thức SUM.")
            if is_sum:
                # Công thức tổng không được tham chiếu tới chính dòng tổng
                # (lỗi circular reference gặp ngày 18/11 với file 0 hộ).
                ref = v[len("=SUM("):].rstrip(")")
                lo, _, hi = ref.partition(":")
                if int(hi[1:]) >= total_row or int(lo[1:]) >= total_row:
                    errors.append(
                        f"Dòng tổng cột {col} tự tham chiếu ({v}).")
    bad_r = []
    bad_s = []
    for r in range(FIRST_DATA_ROW, (total_row or ws.max_row + 1)):
        rv = ws.cell(r, 18).value
        if rv is not None and rv not in (1, 2):
            bad_r.append((r, rv))
        sv = ws.cell(r, 19).value
        if sv is not None and (not isinstance(sv, str)
                               or not sv.startswith("0")
                               or len(sv) != 10):
            bad_s.append((r, sv))
    if bad_r:
        errors.append(f"Cột R có {len(bad_r)} giá trị ngoài 1/2/trống: {bad_r[:5]}")
    if bad_s:
        errors.append(f"Cột S có {len(bad_s)} số không đúng dạng text 10 số: {bad_s[:5]}")
    wb.close()
    return {"ok": not errors, "errors": errors,
            "stats": {"so_hang_du_lieu": n_data,
                      "so_anh_chu_ky_cot_T": len(sig_rows),
                      "dong_tong": total_row}}


def save_workbook(wb: openpyxl.Workbook, path: str | Path):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
