"""Bố cục file Excel đầu ra — sao theo mẫu Phu-Hiep-Quy-IV-2026.xlsx.

Sheet chính "Danh sách hộ":
  dòng 1-3  tiêu đề (merge A:S), dòng 4 trống,
  dòng 5      header, dòng 6 sub-header, dòng 7... dữ liệu, cuối cùng dòng tổng.
  Cột A STT, B họ tên, C:G Quý III, H:L ước Quý IV, M:Q ước cả năm,
  R hộ có XD (1/2/trống), S điện thoại (text), T chữ ký (ảnh duy nhất).
Quy tắc cứng (không đổi theo ấp):
  - A:S không chứa ảnh nào, kể cả dòng tổng.
  - C/H/M là công thức =SUM(...) khi có thành phần; giữ nguyên số ghi tay
    khi tổng kỳ khác tổng thành phần (ghi chênh lệch vào Cần đối chiếu).
  - Ô nguồn trống -> ô trống thật (không 0, không chuỗi rỗng).
  - S luôn là text, giữ số 0 đầu.
"""

from __future__ import annotations

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

SHEET_LIST = "Danh sách hộ"
SHEET_SCAN = "Bản quét"
SHEET_REVIEW = "Cần đối chiếu"
SHEET_SUMMARY = "Tổng hợp theo tờ"

HEADER_ROW = 5
SUBHEADER_ROW = 6
FIRST_DATA_ROW = 7

# Cột tiền: (cột tổng, [cột thành phần]) theo từng kỳ.
PERIODS = {
    "q3": ("C", ["D", "E", "F", "G"]),
    "q4": ("H", ["I", "J", "K", "L"]),
    "year": ("M", ["N", "O", "P", "Q"]),
}
MONEY_COLS = list("CDEFGHIJKLMNOPQ")
COL_STT, COL_NAME, COL_CODE, COL_PHONE, COL_SIGN = "A", "B", "R", "S", "T"

COLUMN_WIDTHS = {
    "A": 14, "B": 28,
    **{c: 14 for c in "CDEFGHIJKLMNOPQR"},
    "S": 16, "T": 25,
}
DATA_ROW_HEIGHT = 27.0

HEADER_FILL = PatternFill("solid", fgColor="1F497D", bgColor="1F497D")
HEADER_FONT = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
TITLE_FONT = Font(name="Times New Roman", size=13, bold=True)
BODY_FONT = Font(name="Times New Roman", size=11)
TOTAL_FONT = Font(name="Times New Roman", size=11, bold=True)
THIN = Side(style="thin", color="BFBFBF")
CELL_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")

GROUP_HEADERS = [
    ("C", "G", "Thực hiện Quý III/2026 (Triệu đồng)"),
    ("H", "L", "Ước Quý IV/2026 (Triệu đồng)"),
    ("M", "Q", "Ước cả năm 2026 (Triệu đồng)"),
]
SUB_HEADERS = {
    "C": "Tổng cộng", "D": "Nhà ở", "E": "Cơ sở SXKD",
    "F": "Nông, lâm, TS", "G": "XDV khác",
    "H": "Tổng cộng", "I": "Nhà ở", "J": "Cơ sở SXKD",
    "K": "Nông, lâm, TS", "L": "XDV khác",
    "M": "Tổng cộng", "N": "Nhà ở", "O": "Cơ sở SXKD",
    "P": "Nông, lâm, TS", "Q": "XDV khác",
}


def build_template(hamlet: str, xa: str = "XÃ VĨNH THÀNH",
                   tinh: str = "TỈNH VĨNH LONG",
                   quyet_dinh: str = "6616/QĐ-UBND ngày 16/10/2026") -> openpyxl.Workbook:
    """Tạo workbook trống đúng bố cục mẫu, chỉ có tiêu đề + header.

    Không chứa dữ liệu hộ nào. Tên ấp/kỳ lấy theo hồ sơ từng đợt,
    không hard-code 9 ấp của đợt Q4/2026.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_LIST
    for col, w in COLUMN_WIDTHS.items():
        ws.column_dimensions[col].width = w
    ws.merge_cells("A1:S1")
    ws["A1"] = f"ỦY BAN NHÂN DÂN {tinh} - {xa} (ẤP {hamlet.upper()})"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A2:S2")
    ws["A2"] = ("PHIẾU THU THẬP THÔNG TIN VỀ CÔNG TRÌNH XÂY DỰNG "
                "CỦA HỘ DÂN CƯ - QUÝ IV NĂM 2026")
    ws["A2"].font = TITLE_FONT
    ws.merge_cells("A3:S3")
    ws["A3"] = (f"(Theo Quyết định số {quyet_dinh} "
                f"của UBND tỉnh Vĩnh Long)")
    ws["A3"].font = BODY_FONT

    for col, text in [("A", "STT"), ("B", "Họ tên chủ hộ"),
                      ("R", "Hộ có XD (1:Có, 2:Không)"),
                      ("S", "Số điện thoại"), ("T", "Chữ ký")]:
        c = ws[f"{col}{HEADER_ROW}"]
        c.value = text
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = CENTER
        c.border = CELL_BORDER
    for c1, c2, text in GROUP_HEADERS:
        ws.merge_cells(f"{c1}{HEADER_ROW}:{c2}{HEADER_ROW}")
        c = ws[f"{c1}{HEADER_ROW}"]
        c.value = text
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = CENTER
        c.border = CELL_BORDER
    for col, text in SUB_HEADERS.items():
        c = ws[f"{col}{SUBHEADER_ROW}"]
        c.value = text
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = CENTER
        c.border = CELL_BORDER
    for col in ["A", "B", "R", "S", "T"]:
        ws.merge_cells(f"{col}{HEADER_ROW}:{col}{SUBHEADER_ROW}")
    ws.row_dimensions[HEADER_ROW].height = 30
    ws.row_dimensions[SUBHEADER_ROW].height = 30
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    return wb
