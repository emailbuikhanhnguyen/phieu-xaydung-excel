"""Mô hình dữ liệu nội bộ.

Mỗi hộ là một bản ghi độc lập có ``row_id`` riêng (UUID) — hai hộ trùng STT
vẫn là hai bản ghi khác nhau. Mọi giá trị quan trọng đều giữ ``nguồn``
(provenance): lấy từ đâu, trang/dòng hay sheet/ô nào, để màn hình duyệt
mở đúng vị trí kiểm tra.
"""

from __future__ import annotations

import uuid
from dataclasses import dataclass, field
from enum import Enum


class PhoneState(str, Enum):
    """Trạng thái số điện thoại — phân biệt rõ 3 trường hợp."""
    PRESENT = "co_so"          # đọc được đủ 10 chữ số, bắt đầu bằng 0
    BLANK_ON_FORM = "phieu_bo_trong"   # trên phiếu không ghi gì
    UNREADABLE = "co_ghi_chua_doc_chac"  # có ghi nhưng chưa đọc chắc -> S trống + ghi chú


class RowStatus(str, Enum):
    AUTO = "tu_dien"          # tự điền từ Excel/PDF, chưa duyệt
    CONFIRMED = "da_xac_nhan"  # người dùng đã duyệt
    NEEDS_REVIEW = "can_doi_chieu"
    NEEDS_RECHECK = "can_kiem_tra_lai"  # nguồn đổi dưới ô đã duyệt


@dataclass
class Provenance:
    """Nguồn của một giá trị."""
    kind: str  # 'excel' | 'pdf-text' | 'pdf-grid' | 'manual' | 'ocr'
    detail: str = ""  # vd: "sheet Ho_Co_Xay_Dung, ô S12" | "PDF tr.3, dòng 5"


@dataclass
class Money:
    """Số tiền một kỳ: tổng + 4 thành phần. None = ô trống thật."""
    total: float | None = None
    nha_o: float | None = None
    sxkd: float | None = None
    nong_lam: float | None = None
    xd_khac: float | None = None

    def components(self) -> list[float | None]:
        return [self.nha_o, self.sxkd, self.nong_lam, self.xd_khac]

    def component_sum(self) -> float | None:
        vals = [v for v in self.components() if v is not None]
        return sum(vals) if vals else None


@dataclass
class HouseholdRow:
    row_id: str = field(default_factory=lambda: uuid.uuid4().hex[:12])
    stt: int | None = None          # None = phiếu không ghi STT
    stt_raw: str = ""               # giữ nguyên chuỗi gốc (vd "130a", bị gạch)
    name: str | None = None
    q3: Money = field(default_factory=Money)
    q4: Money = field(default_factory=Money)
    year: Money = field(default_factory=Money)
    code: int | None = None         # R: 1 | 2 | None
    code_raw: str = ""              # giá trị gốc khi mã không hợp lệ (vd "4")
    phone: str | None = None        # text, giữ số 0 đầu
    phone_state: PhoneState = PhoneState.BLANK_ON_FORM
    phone_raw: str = ""
    status: RowStatus = RowStatus.AUTO
    sources: dict[str, Provenance] = field(default_factory=dict)
    pdf_page: int | None = None     # trang PDF chứa hộ này (số trang trong file)
    pdf_row: int | None = None      # dòng trên trang đó
    excel_row: int | None = None    # dòng trong Excel nguồn

    def has_money(self) -> bool:
        for m in (self.q3, self.q4, self.year):
            if m.total or any(v for v in m.components() if v):
                return True
        return False


@dataclass
class Signature:
    """Chữ ký của một hộ — chỉ ảnh gốc, cắt từ PDF."""
    row_id: str
    pdf_page: int
    # Khung cắt theo tọa độ PDF (điểm), trong cột chữ ký của trang.
    x0: float = 0.0
    y0: float = 0.0
    x1: float = 0.0
    y1: float = 0.0
    present: bool = True            # False = phiếu không ký -> T trống
    adjusted: bool = False          # người dùng đã kéo lại khung
    image_bytes: bytes | None = None


@dataclass
class ReviewNote:
    """Một mục trong sheet Cần đối chiếu."""
    note_id: str = field(default_factory=lambda: uuid.uuid4().hex[:8])
    row_ids: list[str] = field(default_factory=list)  # truy ra từng hộ cụ thể
    stt_text: str = ""
    name_text: str = ""
    column: str = ""                # vd "S", "C", "tổng QIV"
    reason: str = ""
    pdf_page: int | None = None
    kind: str = "chua_chac"         # chua_chac | lech_tong | trung_lap | thieu_chu_ky | pham_vi


@dataclass
class PageInfo:
    """Một trang PDF sau khi phân loại."""
    number: int                     # số trang trong file (1-based)
    kind: str = "danh_sach"         # danh_sach | tong_hop | trang_trong | xac_nhan
    rows: int = 0                   # số dòng hộ phát hiện được
    first_stt: int | None = None
    last_stt: int | None = None
    rotated: bool = False
    fingerprint: str = ""


@dataclass
class RunStats:
    """Thống kê cuối mỗi lượt chạy — dùng để đối chiếu, không gọi AI."""
    pages_processed: int = 0
    pages_reused: int = 0
    cells_autofilled: int = 0
    cells_reused: int = 0
    cells_needs_review: int = 0
    signatures_cropped: int = 0
    ai_calls: int = 0
    ai_tokens: int = 0

    def as_dict(self) -> dict:
        return {
            "pages_processed": self.pages_processed,
            "pages_reused": self.pages_reused,
            "cells_autofilled": self.cells_autofilled,
            "cells_reused": self.cells_reused,
            "cells_needs_review": self.cells_needs_review,
            "signatures_cropped": self.signatures_cropped,
            "ai_calls": self.ai_calls,
            "ai_tokens": self.ai_tokens,
        }
