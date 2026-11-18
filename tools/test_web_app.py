"""Kiem tra giao dien web (che do upload) bang AppTest, khong can trinh duyet.

Chay: PYTHONPATH=src .venv/bin/python tools/test_web_app.py
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from streamlit.testing.v1 import AppTest


def main():
    at = AppTest.from_file(
        str(Path(__file__).resolve().parents[1] / "app" / "streamlit_app.py"),
        default_timeout=180)
    at.run()
    assert not at.exception, f"loi khi mo app: {at.exception}"
    assert at.radio, "khong co radio chon che do"
    print("1. App mo len, co radio che do:", [str(o) for o in at.radio[0].options])

    # Che do web: bam nut thu du lieu gia.
    demo = [b for b in at.button if "Thử ngay" in str(b.label)]
    assert demo, "khong tim thay nut du lieu gia"
    demo[0].click().run()
    assert not at.exception, f"loi khi tao du lieu gia: {at.exception}"
    print("2. Da nap du lieu gia; so dossier:", len(at.session_state.get("dossiers", [])))
    assert len(at.session_state.get("dossiers", [])) == 1, "phai co 1 ap"

    # Upload that: nap lai bang file uploader voi file gia that.
    at2 = AppTest.from_file(
        str(Path(__file__).resolve().parents[1] / "app" / "streamlit_app.py"),
        default_timeout=180)
    at2.run()
    assert not at2.exception
    fx = Path(__file__).resolve().parents[1] / "tests" / "fixtures" / "fake-dossier"
    files = [
        ("An-Binh.pdf", (fx / "An-Binh.pdf").read_bytes(), "application/pdf"),
        ("An-Binh.xlsx", (fx / "An-Binh.xlsx").read_bytes(),
         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    ]
    at2.file_uploader[0].set_value(files).run()
    assert not at2.exception, f"loi khi upload: {at2.exception}"
    nap = [b for b in at2.button if "Nạp file đã chọn" in str(b.label)]
    assert nap, "khong tim thay nut Nap file"
    nap[0].click().run()
    assert not at2.exception, f"loi khi nap file: {at2.exception}"
    print("2b. Upload 2 file gia; so dossier:",
          len(at2.session_state.get("dossiers", [])))
    assert len(at2.session_state.get("dossiers", [])) == 1, "upload phai ra 1 ap"

    # Hoi quy loi 18/11: upload CHI PDF quet (khong Excel) — truoc day ra
    # 0 ho + file xuat bi circular reference. Nay phai co dong giu cho.
    import fitz
    src = fitz.open(fx / "An-Binh.pdf")
    out = fitz.open()
    for page in src:
        pix = page.get_pixmap(dpi=150)
        p = out.new_page(width=page.rect.width, height=page.rect.height)
        p.insert_image(page.rect, pixmap=pix)
    import tempfile
    scan_path = Path(tempfile.mkdtemp(prefix="phieu-scan-")) / "An-Binh-scan.pdf"
    out.save(scan_path)
    out.close()
    src.close()
    at3 = AppTest.from_file(
        str(Path(__file__).resolve().parents[1] / "app" / "streamlit_app.py"),
        default_timeout=180)
    at3.run()
    assert not at3.exception
    at3.file_uploader[0].set_value([
        ("An-Binh-scan.pdf", scan_path.read_bytes(), "application/pdf"),
    ]).run()
    assert not at3.exception, f"loi khi upload PDF quet: {at3.exception}"
    nap3 = [b for b in at3.button if "Nạp file đã chọn" in str(b.label)]
    nap3[0].click().run()
    assert not at3.exception, f"loi khi nap PDF quet: {at3.exception}"
    proc3 = [b for b in at3.button if "Xử lý tất cả" in str(b.label)]
    proc3[0].click().run()
    assert not at3.exception, f"loi khi xu ly PDF quet: {at3.exception}"
    results3 = at3.session_state.get("results", {})
    rows3 = list(results3.values())[0].rows
    print(f"2c. Upload chi PDF quet: {len(rows3)} dong giu cho")
    assert len(rows3) >= 32, f"PDF quet phai co dong giu cho, co {len(rows3)}"
    exp3 = [b for b in at3.button if "Xuất Excel" in str(b.label)]
    exp3[0].click().run()
    assert not at3.exception, f"loi khi xuat PDF quet: {at3.exception}"
    exported3 = at3.session_state.get("exported", [])
    assert exported3 and Path(exported3[0]).exists(), "khong co file xuat"
    # File xuat khong duoc co cong thuc tu tham chieu (circular reference).
    import openpyxl
    wb3 = openpyxl.load_workbook(exported3[0])
    ws3 = wb3["Danh sách hộ"]
    for row in ws3.iter_rows():
        for c in row:
            v = c.value
            if isinstance(v, str) and v.startswith("=SUM("):
                ref = v[len("=SUM("):].rstrip(")")
                _, _, hi = ref.partition(":")
                assert int(hi[1:]) < c.row, f"cong thuc tu tham chieu: {v}"
    wb3.close()
    print("2d. Xuat tu PDF quet OK, khong circular reference:", exported3)

    # Xu ly tat ca.
    proc = [b for b in at.button if "Xử lý tất cả" in str(b.label)]
    assert proc, "khong tim thay nut Xu ly"
    proc[0].click().run()
    assert not at.exception, f"loi khi xu ly: {at.exception}"
    results = at.session_state.get("results", {})
    assert results, "khong co ket qua xu ly"
    rows = list(results.values())[0].rows
    print(f"3. Xu ly xong: {len(rows)} ho")
    assert len(rows) == 32, f"phai co 32 ho, co {len(rows)}"

    # Duyet: luu 1 ho roi sang ho tiep theo.
    save = [b for b in at.button if "Lưu & sang" in str(b.label)]
    assert save, "khong tim thay nut Luu"
    save[0].click().run()
    assert not at.exception, f"loi khi luu ho: {at.exception}"
    print("4. Luu 1 ho, chi so duyet hien tai:", at.session_state["review_idx"])

    # Xuat Excel + kiem tra nut tai ve.
    exp = [b for b in at.button if "Xuất Excel" in str(b.label)]
    assert exp, "khong tim thay nut Xuat"
    exp[0].click().run()
    assert not at.exception, f"loi khi xuat: {at.exception}"
    exported = at.session_state.get("exported", [])
    assert exported and Path(exported[0]).exists(), "khong co file xuat"
    print("5. Da xuat:", exported)
    dl = [b for b in at.download_button] if hasattr(at, "download_button") else []
    print("   nut tai ve:", len(dl))
    assert dl, "phai co nut tai file ket qua"
    print("WEB APPTEST: PASS")


if __name__ == "__main__":
    main()
