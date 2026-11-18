"""Lệnh chạy hàng loạt (không giao diện).

Ví dụ:
  python -m phieu_xaydung.cli scan data/ho-so-q4
  python -m phieu_xaydung.cli process data/ho-so-q4 --out data/ket-qua
  python -m phieu_xaydung.cli verify data/ket-qua/Phieu_....xlsx
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from phieu_xaydung import __version__, build_version
from phieu_xaydung.export import verify_workbook
from phieu_xaydung.pipeline import export_hamlet, process_hamlet, scan_dossier
from phieu_xaydung.store import Store


def cmd_scan(args) -> int:
    dossiers = scan_dossier(args.folder)
    for d in dossiers:
        print(f"- {d.name}: PDF={'có' if d.pdf_path else 'không'} "
              f"({d.pdf_path.name if d.pdf_path else '-'}) | "
              f"Excel={'có' if d.excel_path else 'không'} "
              f"({d.excel_path.name if d.excel_path else '-'})")
    print(f"Tổng: {len(dossiers)} ấp.")
    return 0


def cmd_process(args) -> int:
    folder = Path(args.folder)
    out_dir = Path(args.out)
    store = Store(folder / ".phieu-xaydung" / "state.db")
    dossiers = scan_dossier(folder)
    if args.hamlet:
        dossiers = [d for d in dossiers if args.hamlet.lower() in d.name.lower()]
    grand = {"pages_processed": 0, "cells_autofilled": 0,
             "cells_needs_review": 0, "ai_calls": 0, "ai_tokens": 0}
    for d in dossiers:
        result = process_hamlet(d, store=store)
        path = export_hamlet(result, out_dir)
        s = result.stats.as_dict()
        for k in grand:
            grand[k] += s[k]
        print(f"[{d.name}] {len(result.rows)} hộ, "
              f"{len(result.notes)} mục cần đối chiếu -> {path.name}")
    store.save_run(datetime.now().isoformat(), grand)
    print("Tổng lượt chạy:", json.dumps(grand, ensure_ascii=False))
    print("Chế độ cục bộ: 0 lần gọi AI, 0 token.")
    store.close()
    return 0


def cmd_verify(args) -> int:
    res = verify_workbook(args.file)
    print(json.dumps(res, ensure_ascii=False, indent=2))
    return 0 if res["ok"] else 1


def main(argv=None) -> int:
    p = argparse.ArgumentParser(
        prog="phieu-xaydung",
        description="Chuyển phiếu xây dựng PDF/Excel thành Excel chuẩn (chạy cục bộ, không gọi AI).")
    p.add_argument("--version", action="version",
                   version=f"%(prog)s {build_version()}")
    sub = p.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("scan", help="Liệt kê các ấp trong thư mục hồ sơ.")
    s.add_argument("folder")
    s.set_defaults(func=cmd_scan)

    pr = sub.add_parser("process", help="Xử lý + xuất Excel cho mọi ấp.")
    pr.add_argument("folder")
    pr.add_argument("--out", required=True)
    pr.add_argument("--hamlet", help="Chỉ xử lý ấp có tên chứa chuỗi này.")
    pr.set_defaults(func=cmd_process)

    v = sub.add_parser("verify", help="Kiểm tra một file Excel đã xuất.")
    v.add_argument("file")
    v.set_defaults(func=cmd_verify)

    args = p.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
