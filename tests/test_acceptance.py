"""Bảng nghiệm thu (mục 7 yêu cầu): mỗi test ánh xạ một hàng trong bảng."""

from __future__ import annotations

import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from phieu_xaydung import pdf_tools, validate
from phieu_xaydung.excel_source import (
    detect_shift, is_template_name, phone_state, read_source_excel)
from phieu_xaydung.export import build_workbook, verify_workbook
from phieu_xaydung.layout import build_template
from phieu_xaydung.matcher import match_rows
from phieu_xaydung.models import HouseholdRow, Money, PhoneState
from phieu_xaydung.pipeline import export_hamlet, process_hamlet, scan_dossier
from phieu_xaydung.store import Store

from conftest import FIX


def test_template_layout_giong_mau_phu_hiep():
    wb = build_template("An Bình")
    ws = wb["Danh sách hộ"]
    assert ws["A5"].value == "STT"
    assert ws["B5"].value == "Họ tên chủ hộ"
    assert ws["C5"].value == "Thực hiện Quý III/2026 (Triệu đồng)"
    assert ws["C6"].value == "Tổng cộng" and ws["D6"].value == "Nhà ở"
    assert ws["R5"].value.startswith("Hộ có XD")
    assert ws["S5"].value == "Số điện thoại"
    assert ws["T5"].value == "Chữ ký"
    assert ws.column_dimensions["B"].width == 28
    assert ws.column_dimensions["T"].width == 25
    assert ws.max_row == 6  # trống, chưa có dữ liệu hộ


def test_pdf_quyet_nguoc_stt_trung_thieu_stt_giu_du_dong():
    """PDF quét ngược + STT trùng + dòng thiếu STT: giữ đủ dòng nguồn,
    truy được trang gốc, không ghép nhầm chữ ký chỉ vì trùng STT."""
    dossiers = scan_dossier(FIX)
    assert len(dossiers) == 1
    d = dossiers[0]
    d.reversed_pdf = True
    result = process_hamlet(d)
    assert len(result.rows) == 32  # đủ 32 hộ nguồn
    stts = [r.stt for r in result.rows]
    assert stts.count(5) == 2      # STT 5 trùng 2 hộ
    assert sum(1 for s in stts if s is None) == 1  # 1 hộ thiếu STT
    # Ghi chú trùng STT phải có, và truy được row_id của cả 2 hộ.
    dup = [n for n in result.notes if n.kind == "trung_lap"
           and "STT 5" in n.reason]
    assert dup and len(dup[0].row_ids) >= 2
    # Trang gốc truy được: hộ đầu phải biết mình ở trang PDF nào.
    assert all(r.pdf_page is not None for r in result.rows
               if r.stt is not None)


def test_sdt_giu_so_0_dang_text_phan_biet_trong_va_kho_doc():
    rows = read_source_excel(str(FIX / "An-Binh.xlsx"))[0].rows
    by_stt = {r.stt: r for r in rows if r.stt is not None}
    # SĐT 9 số -> chưa đọc chắc, S trống + ghi chú (không giữ số sai).
    bad = next(r for r in rows if r.phone_raw == "091234567")
    assert bad.phone is None and bad.phone_state == PhoneState.UNREADABLE
    # SĐT đọc được giữ nguyên dạng text, bắt đầu bằng 0.
    good = by_stt[1]
    assert good.phone and good.phone.startswith("0") and len(good.phone) == 10
    # Phiếu bỏ trống phân biệt với có ghi nhưng chưa đọc chắc.
    blank = next(r for r in rows if r.phone_state == PhoneState.BLANK_ON_FORM)
    assert blank.phone is None
    notes = validate.check_phones(rows)
    assert any("khó đọc" in n.reason for n in notes)


def test_excel_thieu_pdf_ten_mau_the_hien_pham_vi():
    """Excel chỉ hộ có XD + tên mẫu: vẫn xử lý được phần có căn cứ,
    thể hiện đúng phạm vi và phần cần bổ sung."""
    dossiers = scan_dossier(FIX)
    d = dossiers[0]
    d.pdf_path = None  # giả lập thiếu PDF
    result = process_hamlet(d)
    assert result.scope == "construction_only"
    assert any(n.kind == "pham_vi" for n in result.notes)
    assert any("Chủ hộ" in n.reason for n in result.notes)  # tên mẫu
    # Không có PDF -> chữ ký trống, ghi rõ.
    assert all(r.pdf_page is None for r in result.rows)


def test_tong_nam_khac_tong_hai_quy_giu_nguyen_so_nguon():
    """Tổng năm khác QIII+QIV và tổng ghi tay khác tổng cộng dòng:
    giữ nguyên số nguồn, báo đúng chênh lệch, không tự cân số."""
    row = HouseholdRow(stt=1, name="Nguyễn Văn An")
    row.q3.total, row.q3.nha_o = 30, 30
    row.q4.total, row.q4.nha_o = 15, 15
    row.year.total, row.year.nha_o = 100, 100  # 100 != 30+15
    notes = validate.check_row(row)
    # Không có ghi chú nào so sánh năm với QIII+QIV.
    assert not any("QIII+QIV" in n.reason or "quý III + quý IV" in n.reason
                   for n in notes)
    # Trong kỳ lệch thì có ghi chú, giữ cả hai số.
    row.year.total = 90
    notes = validate.check_row(row)
    assert any("khác tổng thành phần" in n.reason and "90" in n.reason
               for n in notes)
    # Tổng toàn ấp: ghi tay vs cộng dòng.
    rows = [HouseholdRow(stt=i, name=f"Hộ {i}") for i in range(1, 4)]
    for r in rows:
        r.q3.total, r.q3.nha_o = 10, 10
    notes = validate.check_totals(rows, {"q3": 40}, None)
    assert any("chênh +10" in n.reason for n in notes)


def test_gia_tri_lap_khong_tu_xoa_va_lech_cot_chi_canh_bao():
    rows = []
    for i in range(1, 6):
        r = HouseholdRow(stt=i, name=f"Hộ {i}")
        r.q3.total, r.q3.nha_o = 20, 20  # C=D: hợp lệ khi chỉ xây nhà ở
        rows.append(r)
    notes = validate.check_duplicate_values(rows)
    assert notes  # có ghi chú hệ thống
    assert "không tự xóa" in notes[0].reason
    assert all(r.q3.total == 20 for r in rows)  # giá trị còn nguyên
    # Lệch cột: phát hiện nhưng không tự dịch.
    assert detect_shift([{"Q": "1", "R": "0912345678"},
                         {"Q": "2", "R": "0987654321"},
                         {"Q": "1", "R": "0901234567"}])


def test_dong_lai_giu_nguyen_du_lieu_da_duyet_va_khung_chu_ky(tmp_path):
    store = Store(tmp_path / "state.db")
    row = HouseholdRow(stt=1, name="Nguyễn Văn An")
    row.q3.total = 30
    from phieu_xaydung.models import RowStatus
    row.status = RowStatus.CONFIRMED
    store.save_row("An Bình", row)
    store.save_signature(row.row_id, 1, (100.0, 200.0, 180.0, 240.0),
                         present=True, adjusted=True)
    loaded = store.load_rows("An Bình")
    assert len(loaded) == 1
    assert loaded[0].status == RowStatus.CONFIRMED
    assert loaded[0].q3.total == 30
    sig = store.load_signature(row.row_id)
    assert sig and sig["adjusted"] and sig["box"] == (100.0, 200.0, 180.0, 240.0)
    # Nguồn đổi dưới ô đã duyệt -> đánh dấu kiểm tra lại.
    store.save_file("x.xlsx", "aaa", "excel", 0.0)
    store.mark_recheck_for_file("x.xlsx")
    store.close()


def test_chay_lai_cung_ho_so_0_goi_ai(tmp_path):
    dossiers = scan_dossier(FIX)
    d = dossiers[0]
    d.reversed_pdf = True
    store = Store(tmp_path / "state.db")
    r1 = process_hamlet(d, store=store)
    r2 = process_hamlet(d, store=store)  # chạy lại
    assert r1.stats.ai_calls == 0 and r2.stats.ai_calls == 0
    assert r1.stats.ai_tokens == 0 and r2.stats.ai_tokens == 0
    assert len(r2.rows) == len(r1.rows)
    store.close()


def test_kiem_tra_file_excel_sau_khi_luu(tmp_path):
    dossiers = scan_dossier(FIX)
    d = dossiers[0]
    d.reversed_pdf = True
    result = process_hamlet(d)
    path = export_hamlet(result, tmp_path)
    assert path.exists()
    check = verify_workbook(path)
    assert check["ok"], check["errors"]
    assert check["stats"]["so_hang_du_lieu"] == 32
    # Mở lại bằng openpyxl: S là text, C:Q là số.
    wb = openpyxl.load_workbook(path)
    ws = wb["Danh sách hộ"]
    assert ws.cell(7, 19).number_format == "@"
    assert isinstance(ws.cell(7, 3).value, (int, float, type(None)))
    # Dòng tổng dùng SUM.
    total_row = check["stats"]["dong_tong"]
    assert str(ws.cell(total_row, 3).value).startswith("=SUM(")
    wb.close()


def test_matcher_khong_ghep_sai_khi_ten_khac_nhau():
    excel = [HouseholdRow(stt=1, name="Nguyễn Văn An"),
             HouseholdRow(stt=1, name="Trần Văn Bình")]
    pdf = [HouseholdRow(stt=1, name="Nguyễn Văn An", pdf_page=1, pdf_row=0)]
    pairs, _, _, notes = match_rows(excel, pdf)
    assert len(pairs) == 1
    assert pairs[0][0].name == "Nguyễn Văn An"


def _rasterize_pdf(src: Path, dst: Path, dpi: int = 150):
    """Biến PDF có text layer thành PDF quét thuần ảnh (mô phỏng file scan
    thật từ máy quét: không text, không vector, chỉ ảnh trang)."""
    import fitz
    doc = fitz.open(src)
    out = fitz.open()
    for page in doc:
        pix = page.get_pixmap(dpi=dpi)
        p = out.new_page(width=page.rect.width, height=page.rect.height)
        p.insert_image(page.rect, pixmap=pix)
    out.save(dst)
    out.close()
    doc.close()


def test_workbook_0_ho_khong_circular_reference(tmp_path):
    """Hồi quy lỗi 18/11: file xuất từ hồ sơ 0 hộ (chỉ PDF quét, tool chưa
    tạo được dòng nào) báo circular reference vì dòng tổng ghi
    =SUM(C7:C6) — vừa ngược dải vừa tự tham chiếu chính ô."""
    from phieu_xaydung.export import save_workbook
    wb = build_workbook("Tây Lộc", [], [])
    path = tmp_path / "rong.xlsx"
    save_workbook(wb, path)
    check = verify_workbook(path)
    assert check["ok"], check["errors"]
    assert check["stats"]["so_hang_du_lieu"] == 0
    wb = openpyxl.load_workbook(path)
    ws = wb["Danh sách hộ"]
    total_row = check["stats"]["dong_tong"]
    assert ws.cell(total_row, 2).value == "TỔNG CỘNG (0 HỘ)"
    for col in ("C", "D", "H", "M", "Q"):
        assert ws[f"{col}{total_row}"].value == 0
    # Không còn công thức nào tự tham chiếu chính ô của nó.
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if isinstance(v, str) and v.startswith("=SUM("):
                ref = v[len("=SUM("):].rstrip(")")
                _, _, hi = ref.partition(":")
                assert int(hi[1:]) < c.row
    wb.close()


def test_pdf_quet_chi_co_anh_tao_dong_giu_cho(tmp_path):
    """Hồi quy lỗi 18/11: hồ sơ chỉ có PDF quét (không Excel nguồn) ra
    0 hộ. Nay mỗi dải dòng trong vùng dữ liệu thành một dòng giữ chỗ
    (nguồn pdf-band) để nhập tay ở màn hình duyệt."""
    from phieu_xaydung.pipeline import HamletDossier
    scan_pdf = tmp_path / "An-Binh-scan.pdf"
    _rasterize_pdf(FIX / "An-Binh.pdf", scan_pdf)
    d = HamletDossier(name="An Binh", pdf_path=scan_pdf, reversed_pdf=True)
    result = process_hamlet(d)
    # 3 trang danh sách, mỗi trang 11 hộ + 1 dòng CỘNG + 1 dải tiêu đề
    # lọt vào vùng dữ liệu = 13 dòng giữ chỗ/trang.
    assert len(result.rows) >= 32
    assert all(r.sources["pdf"].kind == "pdf-band" for r in result.rows)
    assert all(r.pdf_page is not None and r.pdf_row is not None
               for r in result.rows)
    # Thứ tự theo trang/dòng, không xáo theo row_id ngẫu nhiên.
    keys = [(r.pdf_page, r.pdf_row) for r in result.rows]
    assert keys == sorted(keys)
    # Có ghi chú hướng dẫn nhập tay + xóa dải giả/dòng CỘNG.
    assert any("dòng giữ chỗ" in n.reason and "dòng CỘNG" in n.reason
               for n in result.notes)
    path = export_hamlet(result, tmp_path)
    check = verify_workbook(path)
    assert check["ok"], check["errors"]
    assert check["stats"]["so_hang_du_lieu"] == len(result.rows)


def test_xoa_dong_giu_cho_chay_lai_khong_tao_lai(tmp_path):
    """Dòng giữ chỗ đã xóa tay (dải giả, dòng CỘNG) không xuất hiện lại
    khi bấm Xử lý tất cả lần nữa."""
    from phieu_xaydung.pipeline import HamletDossier
    scan_pdf = tmp_path / "An-Binh-scan.pdf"
    _rasterize_pdf(FIX / "An-Binh.pdf", scan_pdf)
    d = HamletDossier(name="An Binh", pdf_path=scan_pdf, reversed_pdf=True)
    store = Store(tmp_path / "state.db")
    r1 = process_hamlet(d, store=store)
    victim = r1.rows[0]
    store.delete_row(victim.row_id, "An Binh")
    assert (victim.pdf_page, victim.pdf_row) in store.deleted_bands("An Binh")
    r2 = process_hamlet(d, store=store)
    assert len(r2.rows) == len(r1.rows) - 1
    assert all((r.pdf_page, r.pdf_row) != (victim.pdf_page, victim.pdf_row)
               for r in r2.rows)
    store.close()


def test_phien_ban_3_truong_hop(monkeypatch):
    """build_version ở 3 trường hợp (mock subprocess.run):
    git đầy đủ, shallow clone, không có git. PXD-001."""
    from datetime import datetime, timedelta, timezone
    import phieu_xaydung
    from phieu_xaydung import __version__, build_version

    # 18/11/2026 15:32 giờ VN, cố định để kiểm tra chuỗi ngày giờ.
    ts = int(datetime(2026, 11, 18, 15, 32,
                      tzinfo=timezone(timedelta(hours=7))).timestamp())
    when = "18/11/2026 15:32"

    class R:
        def __init__(self, out, code=0):
            self.stdout = out
            self.returncode = code

    def make_fake_run(case):
        def fake_run(cmd, **kwargs):
            if case == "no_git":
                raise FileNotFoundError("git not found")
            if cmd[1] == "rev-parse" and "--short" in cmd:
                return R("abc1234\n")
            if cmd[1] == "rev-parse" and "--is-shallow-repository" in cmd:
                return R("true\n" if case == "shallow" else "false\n")
            if cmd[1] == "rev-list":
                return R("7\n")
            if cmd[1] == "log":
                return R(f"{ts}\n")
            raise AssertionError(f"lenh git khong duoc mock: {cmd}")
        return fake_run

    monkeypatch.setattr(phieu_xaydung.subprocess, "run",
                        make_fake_run("full"))
    assert build_version() == f"0.1.7 (abc1234) — {when}"

    monkeypatch.setattr(phieu_xaydung.subprocess, "run",
                        make_fake_run("shallow"))
    # Shallow: không hiện số đếm sai, chỉ sha + ngày giờ.
    assert build_version() == f"abc1234 — {when}"

    monkeypatch.setattr(phieu_xaydung.subprocess, "run",
                        make_fake_run("no_git"))
    assert build_version() == __version__


def test_phien_ban_that_khop_git_that():
    """Kiểm tra nhẹ với git thật (không mock): số phiên bản bắt đầu bằng
    0.1. và chứa đúng sha ngắn của HEAD."""
    import subprocess
    from phieu_xaydung import build_version
    sha = subprocess.run(["git", "rev-parse", "--short", "HEAD"],
                         cwd=Path(__file__).resolve().parents[1],
                         capture_output=True, text=True,
                         check=True).stdout.strip()
    v = build_version()
    assert v.startswith("0.1.") and sha in v


import contextlib


@contextlib.contextmanager
def _package_copy(tmp_path, monkeypatch):
    """Copy src/phieu_xaydung sang thư mục tạm rồi import bản copy, để
    mô phỏng deploy mới trong cùng tiến trình mà không chạm code thật."""
    import shutil
    import sys
    src = Path(__file__).resolve().parents[1] / "src" / "phieu_xaydung"
    copy = tmp_path / "phieu_xaydung"
    shutil.copytree(src, copy)
    saved = {k: v for k, v in sys.modules.items()
             if k == "phieu_xaydung" or k.startswith("phieu_xaydung.")}
    for k in saved:
        del sys.modules[k]
    monkeypatch.syspath_prepend(str(tmp_path))
    try:
        import phieu_xaydung
        yield phieu_xaydung, copy
    finally:
        for k in list(sys.modules):
            if k == "phieu_xaydung" or k.startswith("phieu_xaydung."):
                del sys.modules[k]
        sys.modules.update(saved)


def test_ensure_khong_reload_khi_file_khong_doi(tmp_path, monkeypatch):
    """PXD-001 (1): chạy khối nạp 2 lần khi file không đổi thì RowStatus
    vẫn là cùng object (is) — không reload gì."""
    with _package_copy(tmp_path, monkeypatch) as (pkg, _):
        from phieu_xaydung import models
        rs_before = models.RowStatus
        assert pkg.ensure_fresh_modules() is False
        assert pkg.ensure_fresh_modules() is False
        from phieu_xaydung import models as models2
        assert models2.RowStatus is rs_before


def test_ensure_reload_dung_1_lan_khi_file_doi(tmp_path, monkeypatch):
    """PXD-001 (2): đổi file thì reload đúng 1 lần và ra __version__ mới;
    gọi lại lần nữa thì không reload thêm."""
    with _package_copy(tmp_path, monkeypatch) as (pkg, copy):
        from phieu_xaydung import models
        rs_before = models.RowStatus
        assert pkg.ensure_fresh_modules() is False  # ghi nhận dấu lần đầu
        init = copy / "__init__.py"
        init.write_text(init.read_text(encoding="utf-8").replace(
            '__version__ = "0.1.0"', '__version__ = "9.9.9"'),
            encoding="utf-8")
        assert pkg.ensure_fresh_modules() is True  # reload 1 lần
        assert pkg.__version__ == "9.9.9"
        assert pkg.build_version().startswith("9.9.")
        from phieu_xaydung import models as models2
        assert models2.RowStatus is not rs_before  # enum đã nạp lại
        assert pkg.ensure_fresh_modules() is False  # lần sau không reload
