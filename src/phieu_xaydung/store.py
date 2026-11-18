"""Lưu tiến độ bằng SQLite (file .phieu-xaydung/state.db trong thư mục dữ
liệu cục bộ — không đưa lên Git).

Theo dõi: dấu vân tay file nguồn, trang, vùng cắt chữ ký, trạng thái duyệt
từng ô, phiên bản cấu hình. Chạy lại cùng hồ sơ thì dùng kết quả đã lưu,
giữ nguyên các ô đã xác nhận, không gọi lại OCR/AI. Nếu nguồn thay đổi
dưới một ô đã duyệt -> đánh dấu cần kiểm tra lại. Chỉ xử lý lại phần bị
ảnh hưởng khi sửa khung cắt hoặc thay PDF.
"""

from __future__ import annotations

import json
import sqlite3
from pathlib import Path

from .models import HouseholdRow, Money, PhoneState, Provenance, RowStatus

SCHEMA = """
CREATE TABLE IF NOT EXISTS files(
  path TEXT PRIMARY KEY, fingerprint TEXT, kind TEXT, mtime REAL);
CREATE TABLE IF NOT EXISTS pages(
  file_path TEXT, number INTEGER, kind TEXT, rows INTEGER,
  first_stt INTEGER, last_stt INTEGER,
  PRIMARY KEY(file_path, number));
CREATE TABLE IF NOT EXISTS rows(
  row_id TEXT PRIMARY KEY, hamlet TEXT, stt INTEGER, stt_raw TEXT,
  name TEXT, q3 TEXT, q4 TEXT, year TEXT, code INTEGER, code_raw TEXT,
  phone TEXT, phone_state TEXT, phone_raw TEXT, status TEXT,
  pdf_page INTEGER, pdf_row INTEGER, excel_row INTEGER,
  sources TEXT);
CREATE TABLE IF NOT EXISTS signatures(
  row_id TEXT PRIMARY KEY, pdf_page INTEGER,
  x0 REAL, y0 REAL, x1 REAL, y1 REAL, present INTEGER, adjusted INTEGER);
CREATE TABLE IF NOT EXISTS notes(
  note_id TEXT PRIMARY KEY, row_ids TEXT, stt_text TEXT, name_text TEXT,
  column_name TEXT, reason TEXT, pdf_page INTEGER, kind TEXT);
CREATE TABLE IF NOT EXISTS runs(
  id INTEGER PRIMARY KEY AUTOINCREMENT, started_at TEXT, stats TEXT);
CREATE TABLE IF NOT EXISTS meta(key TEXT PRIMARY KEY, value TEXT);
"""


def _money_json(m: Money) -> str:
    return json.dumps({"total": m.total, "nha_o": m.nha_o, "sxkd": m.sxkd,
                       "nong_lam": m.nong_lam, "xd_khac": m.xd_khac},
                      ensure_ascii=False)


def _money_from_json(s: str) -> Money:
    d = json.loads(s)
    return Money(total=d.get("total"), nha_o=d.get("nha_o"),
                 sxkd=d.get("sxkd"), nong_lam=d.get("nong_lam"),
                 xd_khac=d.get("xd_khac"))


class Store:
    def __init__(self, db_path: str | Path):
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(str(self.db_path))
        self.conn.executescript(SCHEMA)
        self.conn.commit()

    def close(self):
        self.conn.close()

    # -- file fingerprint -------------------------------------------------
    def file_fingerprint(self, path: str) -> str | None:
        cur = self.conn.execute("SELECT fingerprint FROM files WHERE path=?",
                                (path,))
        row = cur.fetchone()
        return row[0] if row else None

    def save_file(self, path: str, fingerprint: str, kind: str, mtime: float):
        self.conn.execute(
            "INSERT OR REPLACE INTO files VALUES (?,?,?,?)",
            (path, fingerprint, kind, mtime))
        self.conn.commit()

    def file_changed(self, path: str, fingerprint: str) -> bool:
        old = self.file_fingerprint(path)
        return old is not None and old != fingerprint

    # -- rows --------------------------------------------------------------
    def save_row(self, hamlet: str, row: HouseholdRow):
        self.conn.execute(
            "INSERT OR REPLACE INTO rows VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (row.row_id, hamlet, row.stt, row.stt_raw, row.name,
             _money_json(row.q3), _money_json(row.q4), _money_json(row.year),
             row.code, row.code_raw, row.phone, row.phone_state.value,
             row.phone_raw, row.status.value, row.pdf_page, row.pdf_row,
             row.excel_row,
             json.dumps({k: [v.kind, v.detail] for k, v in row.sources.items()},
                        ensure_ascii=False)))
        self.conn.commit()

    def load_rows(self, hamlet: str) -> list[HouseholdRow]:
        cur = self.conn.execute(
            "SELECT * FROM rows WHERE hamlet=? ORDER BY "
            "CASE WHEN stt IS NULL THEN 1 ELSE 0 END, stt, row_id",
            (hamlet,))
        out: list[HouseholdRow] = []
        for r in cur.fetchall():
            (row_id, _, stt, stt_raw, name, q3, q4, year, code, code_raw,
             phone, phone_state, phone_raw, status, pdf_page, pdf_row,
             excel_row, sources) = r
            row = HouseholdRow(
                row_id=row_id, stt=stt, stt_raw=stt_raw or "", name=name,
                q3=_money_from_json(q3), q4=_money_from_json(q4),
                year=_money_from_json(year), code=code, code_raw=code_raw or "",
                phone=phone,
                phone_state=PhoneState(phone_state),
                phone_raw=phone_raw or "",
                status=RowStatus(status),
                pdf_page=pdf_page, pdf_row=pdf_row, excel_row=excel_row)
            row.sources = {k: Provenance(kind=v[0], detail=v[1])
                           for k, v in json.loads(sources).items()}
            out.append(row)
        return out

    def mark_recheck_for_file(self, path: str):
        """Nguồn đổi: các ô đã duyệt lấy từ file đó -> cần kiểm tra lại."""
        cur = self.conn.execute("SELECT row_id, sources FROM rows")
        for row_id, sources_json in cur.fetchall():
            sources = json.loads(sources_json)
            for kind, detail in sources.values():
                if path in detail:
                    self.conn.execute(
                        "UPDATE rows SET status=? WHERE row_id=?",
                        (RowStatus.NEEDS_RECHECK.value, row_id))
        self.conn.commit()

    # -- signatures ---------------------------------------------------------
    def save_signature(self, row_id: str, pdf_page: int, box: tuple[float, float, float, float],
                       present: bool, adjusted: bool = False):
        self.conn.execute(
            "INSERT OR REPLACE INTO signatures VALUES (?,?,?,?,?,?,?,?)",
            (row_id, pdf_page, box[0], box[1], box[2], box[3],
             1 if present else 0, 1 if adjusted else 0))
        self.conn.commit()

    def load_signature(self, row_id: str) -> dict | None:
        cur = self.conn.execute("SELECT * FROM signatures WHERE row_id=?",
                                (row_id,))
        r = cur.fetchone()
        if not r:
            return None
        return {"row_id": r[0], "pdf_page": r[1], "box": (r[2], r[3], r[4], r[5]),
                "present": bool(r[6]), "adjusted": bool(r[7])}

    # -- notes ---------------------------------------------------------------
    def save_note(self, note_id: str, row_ids: list[str], stt_text: str,
                  name_text: str, column: str, reason: str,
                  pdf_page: int | None, kind: str):
        self.conn.execute(
            "INSERT OR REPLACE INTO notes VALUES (?,?,?,?,?,?,?,?)",
            (note_id, json.dumps(row_ids, ensure_ascii=False), stt_text,
             name_text, column, reason, pdf_page, kind))
        self.conn.commit()

    def load_notes(self) -> list[dict]:
        cur = self.conn.execute("SELECT * FROM notes")
        return [{"note_id": r[0], "row_ids": json.loads(r[1]),
                 "stt_text": r[2], "name_text": r[3], "column": r[4],
                 "reason": r[5], "pdf_page": r[6], "kind": r[7]}
                for r in cur.fetchall()]

    # -- runs / meta ----------------------------------------------------------
    def save_run(self, started_at: str, stats: dict):
        self.conn.execute("INSERT INTO runs(started_at, stats) VALUES (?,?)",
                          (started_at, json.dumps(stats, ensure_ascii=False)))
        self.conn.commit()

    def set_meta(self, key: str, value: str):
        self.conn.execute("INSERT OR REPLACE INTO meta VALUES (?,?)",
                          (key, value))
        self.conn.commit()

    def get_meta(self, key: str) -> str | None:
        cur = self.conn.execute("SELECT value FROM meta WHERE key=?", (key,))
        r = cur.fetchone()
        return r[0] if r else None
