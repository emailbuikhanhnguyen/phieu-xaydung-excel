"""Tạo bộ dữ liệu GIẢ để chạy thử và kiểm thử (không dùng dữ liệu thật).

Tạo một thư mục hồ sơ giả cho ấp "An Bình" gồm:
  - PDF form có text layer + đường kẻ lưới, 3 trang danh sách + 1 trang
    tổng hợp + 1 trang trắng; trang danh sách quét ngược (trang cuối lên
    đầu) để test xử lý PDF ngược như Đông Nam;
  - Excel nguồn chỉ có hộ xây dựng, có tên mẫu, lệch cột ở vài dòng,
    SĐT sai, mã R lạ, tổng năm khác QIII+QIV;
  - ground-truth.json: dữ liệu đúng để test đối chiếu.

Tên, SĐT, chữ ký đều là giả (chữ ký vẽ bằng nét ngẫu nhiên).
Chạy: python tools/make_fake_data.py --out tests/fixtures/fake-dossier
"""

from __future__ import annotations

import argparse
import json
import random
from pathlib import Path

import fitz
import openpyxl

FONT = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
FONT_B = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"

FAKE_NAMES = [
    "Nguyễn Văn An", "Trần Thị Bình", "Lê Văn Cường", "Phạm Thị Dung",
    "Hoàng Văn Em", "Võ Thị Phượng", "Đặng Văn Giang", "Bùi Thị Hạnh",
    "Đỗ Văn Khang", "Hồ Thị Lan", "Ngô Văn Minh", "Dương Thị Ngọc",
    "Lý Văn Phúc", "Mai Thị Quỳnh", "Trịnh Văn Sơn", "Lâm Thị Thảo",
    "Cao Văn Tuấn", "Đinh Thị Uyên", "Tạ Văn Vinh", "Hà Thị Xuân",
    "Vũ Văn Yên", "Phan Thị Ánh", "Châu Văn Bảo", "Tô Thị Chi",
]
FAKE_PHONES = [f"09{random.randint(10000000, 99999999)}" for _ in range(40)]


def make_households(n: int = 32, seed: int = 7) -> list[dict]:
    rnd = random.Random(seed)
    out = []
    for i in range(1, n + 1):
        q3 = rnd.choice([0, 10, 20, 30, 50])
        q4 = rnd.choice([0, 15, 25, 40])
        # Tổng năm ĐỘC LẬP — không mặc định bằng QIII+QIV.
        year = rnd.choice([0, q3 + q4, q3 + q4 + 10, 45])
        code = 1 if (q3 or q4 or year) else 2
        out.append({
            "stt": i,
            "name": FAKE_NAMES[(i - 1) % len(FAKE_NAMES)],
            "q3_total": q3, "q3_nha_o": q3, "q3_sxkd": 0,
            "q3_nong_lam": 0, "q3_xd_khac": 0,
            "q4_total": q4, "q4_nha_o": q4, "q4_sxkd": 0,
            "q4_nong_lam": 0, "q4_xd_khac": 0,
            "year_total": year, "year_nha_o": year, "year_sxkd": 0,
            "year_nong_lam": 0, "year_xd_khac": 0,
            "code": code,
            "phone": FAKE_PHONES[(i - 1) % len(FAKE_PHONES)],
        })
    # Trường hợp đặc biệt để test:
    # STT 5 trùng 2 hộ CÙNG TÊN ở cả PDF lẫn Excel -> không thể phân biệt
    # bằng tên, buộc đưa vào duyệt (không ghép nhầm chữ ký vì trùng STT).
    out[4]["stt"] = 5
    out[4]["name"] = "Lê Văn Cường"
    out[4]["q3_total"] = 20
    out[4]["q3_nha_o"] = 20
    out[4]["code"] = 1
    out[5]["stt"] = 5
    out[5]["name"] = "Lê Văn Cường"
    out[5]["q3_total"] = 30
    out[5]["q3_nha_o"] = 30
    out[5]["code"] = 1
    out[9]["stt"] = None         # 1 hộ thiếu STT
    out[9]["name"] = "Lê Thị Mai"
    out[14]["phone"] = "091234567"   # SĐT 9 số (sai)
    out[19]["phone"] = ""            # phiếu bỏ trống SĐT
    return out


def draw_pdf(path: Path, households: list[dict], reversed_order: bool = True):
    doc = fitz.open()
    W, H = 1190, 842  # A3 ngang
    cols = [30, 60, 300, 360, 420, 480, 540, 600, 660, 720, 780, 840, 900,
            960, 1020, 1080, 1110, 1140, 1190]
    # STT, tên, C..Q (15 cột), R, S, T — rút gọn: dùng 19 mốc.
    headers = ["STT", "Họ tên chủ hộ", "C", "D", "E", "F", "G", "H", "I",
               "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T"]
    per_page = 12
    pages: list[list[dict]] = [households[i:i + per_page]
                               for i in range(0, len(households), per_page)]
    shapes: list[fitz.Page] = []

    def draw_list_page(rows: list[dict], pageno_label: str):
        page = doc.new_page(width=W, height=H)
        page.insert_text((30, 30), "PHIẾU THU THẬP THÔNG TIN VỀ CÔNG TRÌNH "
                         "XÂY DỰNG CỦA HỘ DÂN CƯ - QUÝ IV NĂM 2026",
                         fontname="F0", fontfile=FONT_B, fontsize=13)
        page.insert_text((30, 52), f"ẤP AN BÌNH — {pageno_label}",
                         fontname="F0", fontfile=FONT, fontsize=11)
        y_top, y_bot = 70, 70 + 13 * 28
        # Lưới: kẻ ngang + dọc.
        for y in range(y_top, y_bot + 1, 28):
            page.draw_line((30, y), (W - 10, y), color=(0, 0, 0), width=0.7)
        xs = [30, 70, 300]
        x = 300
        for _ in range(15):
            x += 42
            xs.append(min(x, 930))
        xs += [960, 1080, W - 10]
        for x in xs:
            page.draw_line((x, y_top), (x, y_bot), color=(0, 0, 0), width=0.7)
        # Header.
        for j, h in enumerate(headers):
            page.insert_text((xs[j] + 3, y_top + 18), h, fontname="F0",
                             fontfile=FONT_B, fontsize=8)
        rnd = random.Random(99)
        for i, hh in enumerate(rows):
            y = y_top + (i + 1) * 28
            vals = ["" if hh["stt"] is None else str(hh["stt"]), hh["name"],
                    str(hh["q3_total"]), str(hh["q3_nha_o"]), "0", "0", "0",
                    str(hh["q4_total"]), str(hh["q4_nha_o"]), "0", "0", "0",
                    str(hh["year_total"]), str(hh["year_nha_o"]), "0", "0",
                    "0", str(hh["code"]), hh["phone"]]
            for j, v in enumerate(vals):
                if v:
                    page.insert_text((xs[j] + 3, y + 18), v, fontname="F0",
                                     fontfile=FONT, fontsize=8)
            # Chữ ký giả: vài nét nguệch ngoạc trong cột T (1080-1180).
            sx = xs[-2] + 6
            for _ in range(3):
                x1 = sx + rnd.randint(0, 50)
                page.draw_line((x1, y + 6), (x1 + rnd.randint(8, 25),
                                             y + rnd.randint(6, 22)),
                               color=(0.1, 0.1, 0.5), width=0.8)
        shapes.append(page)
        return page

    list_pages = []
    for i, chunk in enumerate(pages):
        list_pages.append((f"Tờ {i + 1}", chunk))
    order = list_pages[::-1] if reversed_order else list_pages
    for label, chunk in order:
        draw_list_page(chunk, label)
    # Trang tổng hợp.
    p = doc.new_page(width=W, height=H)
    p.insert_text((30, 30), "TỔNG HỢP THEO TỜ — ẤP AN BÌNH",
                  fontname="F0", fontfile=FONT_B, fontsize=13)
    p.insert_text((30, 60), f"Tổng số hộ của ấp: {len(households)} hộ",
                  fontname="F0", fontfile=FONT, fontsize=11)
    p.insert_text((30, 90),
                  f"Tổng Quý III: {sum(h['q3_total'] for h in households)} "
                  f"triệu đồng",
                  fontname="F0", fontfile=FONT, fontsize=11)
    # Trang trắng.
    doc.new_page(width=W, height=H)
    doc.save(path)
    return len(doc)


def make_excel(path: Path, households: list[dict]):
    """Excel nguồn: chỉ hộ có XD + tên mẫu + lệch cột vài dòng + SĐT sai."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Ho_Co_Xay_Dung"
    ws.append(["STT", "Họ tên chủ hộ", "Tổng QIII", "Nhà ở", "Cơ sở SXKD",
               "Nông, lâm, TS", "XDV khác", "Tổng QIV", "Nhà ở", "Cơ sở SXKD",
               "Nông, lâm, TS", "XDV khác", "Tổng cả năm", "Nhà ở",
               "Cơ sở SXKD", "Nông, lâm, TS", "XDV khác", "Hộ có XD",
               "Số điện thoại"])
    for hh in households:
        if hh["code"] != 1:
            continue  # chỉ hộ có XD -> phạm vi hẹp
        name = hh["name"]
        if hh["stt"] in (2, 3):
            name = f"Chủ hộ {hh['stt']}"  # tên mẫu
        ws.append([hh["stt"], name, hh["q3_total"], hh["q3_nha_o"], 0, 0, 0,
                   hh["q4_total"], hh["q4_nha_o"], 0, 0, 0,
                   hh["year_total"], hh["year_nha_o"], 0, 0, 0,
                   hh["code"], hh["phone"]])
    # Lệch cột ở 3 dòng cuối: SĐT sang cột R, mã sang cột Q.
    for r in range(ws.max_row - 2, ws.max_row + 1):
        phone = ws.cell(r, 19).value
        ws.cell(r, 17).value = ws.cell(r, 18).value  # mã -> Q
        ws.cell(r, 18).value = phone                # SĐT -> R
        ws.cell(r, 19).value = None
    wb.save(path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    args = ap.parse_args()
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    hh = make_households()
    n_pages = draw_pdf(out / "An-Binh.pdf", hh, reversed_order=True)
    make_excel(out / "An-Binh.xlsx", hh)
    truth = {
        "hamlet": "An Bình",
        "n_households": len(hh),
        "n_pages_pdf": n_pages,
        "reversed_pdf": True,
        "duplicate_stt": 5,
        "missing_stt_count": 1,
        "households": hh,
        "totals": {
            "q3": sum(h["q3_total"] for h in hh),
            "q4": sum(h["q4_total"] for h in hh),
            "year": sum(h["year_total"] for h in hh),
        },
        "excel_scope": "construction_only",
        "excel_rows": sum(1 for h in hh if h["code"] == 1),
    }
    (out / "ground-truth.json").write_text(
        json.dumps(truth, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Đã tạo hồ sơ giả tại {out} ({n_pages} trang PDF, "
          f"{len(hh)} hộ, {truth['excel_rows']} dòng Excel).")


if __name__ == "__main__":
    main()
