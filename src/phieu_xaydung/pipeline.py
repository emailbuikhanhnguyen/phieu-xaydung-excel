"""Pipeline xử lý một thư mục hồ sơ — dùng chung cho CLI và giao diện.

Quy trình (không gọi AI):
  1. Quét thư mục: nhận PDF + Excel nguồn theo tên ấp trong hồ sơ.
  2. Đọc Excel nguồn (tự điền phần đọc được) + phân loại trang PDF.
  3. Ghép Excel với PDF theo STT/tên/vị trí; phần chưa chắc -> duyệt.
  4. Cắt chữ ký theo lưới cột T; khung cắt lưu vào SQLite để chỉnh tay.
  5. Kiểm tra số liệu (validate) -> ghi chú Cần đối chiếu.
  6. Xuất Excel + báo cáo lượt chạy (số trang, số ô dùng lại, 0 gọi AI).

Chạy lại cùng hồ sơ: các ô đã xác nhận được giữ nguyên; chỉ xử lý lại
phần bị ảnh hưởng (file nguồn đổi, khung cắt đổi).
"""

from __future__ import annotations

import hashlib
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

from . import pdf_tools, validate
from .excel_source import read_source_excel
from .export import build_workbook, save_workbook, verify_workbook
from .layout import FIRST_DATA_ROW
from .matcher import match_rows, merge_pair
from .models import HouseholdRow, PageInfo, Provenance, ReviewNote, RunStats
from .pdf_tools import fingerprint
from .store import Store


def _fold(s: str) -> str:
    s = unicodedata.normalize("NFD", s.lower())
    return "".join(c for c in s if unicodedata.category(c) != "Mn")


@dataclass
class HamletDossier:
    name: str
    pdf_path: Path | None = None
    excel_path: Path | None = None
    reversed_pdf: bool = False
    scope: str = "unknown"  # full | construction_only | unknown


def _hamlet_key(name: str) -> str:
    return _fold(name).replace(" ", "")


def scan_dossier(folder: str | Path) -> list[HamletDossier]:
    """Ghép PDF + Excel theo tên ấp trong tên file.

    Tên ấp lấy theo hồ sơ từng đợt, không hard-code danh sách ấp.
    """
    folder = Path(folder)
    dossiers: dict[str, HamletDossier] = {}
    for p in sorted(folder.iterdir()):
        if not p.is_file():
            continue
        if p.suffix.lower() not in (".pdf", ".xlsx", ".xls"):
            continue  # chỉ nhận PDF + Excel nguồn, bỏ qua ground-truth.json...
        stem = p.stem
        # Bỏ hậu tố thường gặp: _Quy_IV_2026, -Q4-2026, _lan2...
        clean = re.sub(r"(_quy_iv_2026|_q4_2026|_2026|_lan\s*2)$", "",
                       stem, flags=re.IGNORECASE)
        clean = clean.replace("_", " ").replace("-", " ").strip()
        key = _hamlet_key(clean)
        d = dossiers.setdefault(key, HamletDossier(name=clean.title()))
        if p.suffix.lower() == ".pdf":
            d.pdf_path = p
        elif p.suffix.lower() in (".xlsx", ".xls"):
            d.excel_path = p
    return list(dossiers.values())


def pdf_rows_from_text_grid(pdf_path: Path,
                            reversed_order: bool = False) -> tuple[list[HouseholdRow], list[PageInfo]]:
    """Đọc các dòng hộ từ PDF có text layer (dạng form điện tử).

    Với PDF quét (không text layer) trả về rỗng — các dòng sẽ do người
    dùng nhập/đọc tay ở màn hình duyệt, chữ ký vẫn cắt được theo lưới.
    """
    import fitz
    doc = fitz.open(pdf_path)
    infos = pdf_tools.analyze_pdf(str(pdf_path), reversed_order)
    rows: list[HouseholdRow] = []
    order = list(range(len(doc)))
    if reversed_order:
        order = order[::-1]
    for pageno in order:
        page = doc[pageno]
        info = next((i for i in infos if i.number == pageno + 1), None)
        if info is None or info.kind != "danh_sach":
            continue
        for g in pdf_tools.extract_text_grid(page):
            cells = g["cells"]
            if not cells or not any(c.strip() for c in cells):
                continue  # dải trống (cuối trang chỉ điền một phần)
            first = cells[0].strip().lower()
            second = cells[1].strip().lower() if len(cells) > 1 else ""
            if first == "stt" or "họ tên" in second:
                continue  # dòng header, không phải hộ
            m = re.search(r"\d+", cells[0])
            # Dòng thiếu STT vẫn là một hộ thật (giữ vị trí để duyệt tay),
            # chỉ khác là stt=None.
            row = HouseholdRow(
                stt=int(m.group()) if m else None,
                stt_raw=cells[0].strip(),
                name=cells[1].strip() or None if len(cells) > 1 else None,
                pdf_page=pageno + 1, pdf_row=g["row"])
            row.sources["pdf"] = Provenance(
                "pdf-text", f"PDF tr.{pageno + 1}, dòng {g['row']}")
            rows.append(row)
    doc.close()
    return rows, infos


@dataclass
class HamletResult:
    dossier: HamletDossier
    rows: list[HouseholdRow] = field(default_factory=list)
    notes: list[ReviewNote] = field(default_factory=list)
    pages: list[PageInfo] = field(default_factory=list)
    stats: RunStats = field(default_factory=RunStats)
    scope: str = "unknown"
    written_totals: dict = field(default_factory=dict)
    hamlet_total_households: int | None = None


def process_hamlet(dossier: HamletDossier, store: Store | None = None,
                   reuse: bool = True) -> HamletResult:
    """Xử lý một ấp: tự điền phần đọc được, phần chưa chắc -> ghi chú."""
    result = HamletResult(dossier=dossier)
    stats = result.stats
    excel_rows: list[HouseholdRow] = []
    if dossier.excel_path:
        data = dossier.excel_path.read_bytes()
        fp = fingerprint(data)
        if store is not None:
            if store.file_changed(str(dossier.excel_path), fp):
                store.mark_recheck_for_file(str(dossier.excel_path))
            store.save_file(str(dossier.excel_path), fp, "excel",
                            dossier.excel_path.stat().st_mtime)
        sheets = read_source_excel(str(dossier.excel_path))
        for sh in sheets:
            excel_rows.extend(sh.rows)
            if sh.shifted:
                result.notes.append(ReviewNote(
                    column="Q/R",
                    reason=(f"Excel nguồn sheet '{sh.name}' lệch cột: SĐT nằm ở R, "
                            f"mã 1/2 nằm ở Q — chỉ cảnh báo, không tự dịch dãy tiền."),
                    kind="chua_chac"))
            if sh.scope_hint != "unknown":
                result.scope = sh.scope_hint
        stats.cells_autofilled += len(excel_rows) * 5

    pdf_rows: list[HouseholdRow] = []
    if dossier.pdf_path:
        data = dossier.pdf_path.read_bytes()
        fp = fingerprint(data)
        if store is not None:
            if store.file_changed(str(dossier.pdf_path), fp):
                store.mark_recheck_for_file(str(dossier.pdf_path))
            store.save_file(str(dossier.pdf_path), fp, "pdf",
                            dossier.pdf_path.stat().st_mtime)
        pdf_rows, result.pages = pdf_rows_from_text_grid(
            dossier.pdf_path, dossier.reversed_pdf)
        stats.pages_processed = len(result.pages)
        stats.signatures_cropped = len(pdf_rows)
        # Trang tổng hợp: đọc tổng số hộ + tổng ghi tay của ấp.
        import fitz
        doc = fitz.open(dossier.pdf_path)
        for info in result.pages:
            if info.kind == "tong_hop":
                page = doc[info.number - 1]
                total = pdf_tools.hamlet_total_from_summary(page)
                if total:
                    result.hamlet_total_households = total
                for k, v in pdf_tools.written_totals_from_summary(page).items():
                    if v is not None:
                        result.written_totals[k] = v
        doc.close()

    # Ghép: PDF là chuẩn về thứ tự; Excel là chuẩn về số liệu.
    if pdf_rows and excel_rows:
        pairs, excel_left, pdf_left, notes = match_rows(excel_rows, pdf_rows)
        result.notes.extend(notes)
        merged = [merge_pair(e, p) for e, p in pairs]
        merged.extend(pdf_left)   # hộ chỉ có trên PDF
        # Hộ Excel chưa ghép mà STT đã có phía PDF (vd STT trùng chưa phân
        # biệt được): hộ đã có mặt phía PDF, KHÔNG thêm dòng trùng —
        # dữ liệu Excel của chúng nằm trong ghi chú duyệt để ghép tay.
        pdf_left_stts = {p.stt for p in pdf_left if p.stt is not None}
        for e in excel_left:
            if e.stt is not None and e.stt in pdf_left_stts:
                continue
            merged.append(e)  # hộ chỉ có trên Excel
        result.rows = sorted(
            merged,
            key=lambda r: (r.stt is None, r.stt if r.stt is not None else 0,
                           r.row_id))
        if not pdf_rows:
            result.scope = result.scope or "unknown"
    elif excel_rows:
        result.rows = sorted(
            excel_rows,
            key=lambda r: (r.stt is None, r.stt if r.stt is not None else 0,
                           r.row_id))
        # Không có PDF: phạm vi chỉ là gợi ý từ tên sheet, chưa xác minh được.
        result.notes.append(ReviewNote(
            column="phạm vi",
            reason=("Không có PDF để đối chiếu: chưa xác định được file Excel "
                    "này là toàn bộ hộ dân hay chỉ hộ có xây dựng."),
            kind="pham_vi"))
    else:
        result.rows = pdf_rows

    # Kiểm tra số liệu.
    for row in result.rows:
        result.notes.extend(validate.check_row(row))
    result.notes.extend(validate.check_duplicate_values(result.rows))
    result.notes.extend(validate.check_phones(result.rows))
    result.notes.extend(validate.check_totals(
        result.rows, result.written_totals, result.hamlet_total_households))
    stats.cells_needs_review = len(result.notes)
    stats.ai_calls = 0
    stats.ai_tokens = 0

    if store is not None:
        for row in result.rows:
            if reuse:
                old = [r for r in store.load_rows(dossier.name)
                       if r.row_id == row.row_id]
                if old and old[0].status.value == "da_xac_nhan":
                    row.status = old[0].status
                    stats.cells_reused += 1
            store.save_row(dossier.name, row)
        for n in result.notes:
            store.save_note(n.note_id, n.row_ids, n.stt_text, n.name_text,
                            n.column, n.reason, n.pdf_page, n.kind)
    return result


def export_hamlet(result: HamletResult, out_dir: str | Path,
                  page_images: list[tuple[int, bytes]] | None = None,
                  include_scans: bool = True,
                  **template_kwargs) -> Path:
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    safe = re.sub(r"[^\w\-]+", "_", result.dossier.name)
    path = out_dir / f"Phieu_Thu_Thap_Xay_Dung_Ap_{safe}_Q4_2026.xlsx"
    signatures: list[tuple[int, bytes]] = []
    if result.dossier.pdf_path and result.dossier.pdf_path.exists():
        import fitz
        doc = fitz.open(result.dossier.pdf_path)
        if page_images is None and include_scans:
            page_images = []
            for info in result.pages:
                page_images.append(
                    (info.number,
                     pdf_tools.render_page_png(doc, info.number - 1, dpi=120)))
        for i, row in enumerate(result.rows):
            if row.pdf_page is None or row.pdf_row is None:
                continue
            page = doc[row.pdf_page - 1]
            bands = pdf_tools.row_bands(page)
            if row.pdf_row >= len(bands):
                continue
            y0, y1 = bands[row.pdf_row]
            rect = pdf_tools.signature_rect(page, y0, y1)
            if rect is None:
                continue
            if pdf_tools.ink_bbox(page, rect) is None:
                continue  # ô chữ ký trống -> T để trống thật
            signatures.append(
                (FIRST_DATA_ROW + i,
                 pdf_tools.crop_png(doc, row.pdf_page - 1, rect, dpi=200)))
        doc.close()
    wb = build_workbook(result.dossier.name, result.rows, result.notes,
                        page_images=page_images, signatures=signatures,
                        **template_kwargs)
    save_workbook(wb, path)
    check = verify_workbook(path)
    if not check["ok"]:
        raise RuntimeError(f"File xuất không đạt kiểm tra: {check['errors']}")
    return path
