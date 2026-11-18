"""phieu-xaydung-excel — chuyển phiếu thu thập thông tin xây dựng (PDF/Excel)
thành file Excel chuẩn theo mẫu Phú Hiệp.

Chạy cục bộ, không gọi API AI. Đọc trực tiếp Excel nguồn, dùng text layer
của PDF khi có, xử lý hình ảnh bằng code; chữ viết tay duyệt thủ công theo ô.
"""

from __future__ import annotations

import subprocess
from datetime import datetime, timedelta, timezone
from pathlib import Path

__version__ = "0.1.0"

_REPO_ROOT = Path(__file__).resolve().parents[2]
# Giờ VN cố định UTC+7 (không phụ thuộc tzdata của máy chủ).
_VN_TZ = timezone(timedelta(hours=7))


def _git_output(args: list[str]) -> str | None:
    """Chạy một lệnh git trong thư mục repo, trả về stdout đã strip.

    None khi không có git, không có .git, hoặc lệnh lỗi.
    """
    try:
        r = subprocess.run(["git", *args], cwd=_REPO_ROOT,
                           capture_output=True, text=True, timeout=10)
        if r.returncode == 0:
            return r.stdout.strip()
    except Exception:
        pass
    return None


def _git_build() -> dict | None:
    """Thông tin bản dựng từ git: sha ngắn, số commit, ngày giờ commit.

    Trả về None khi không đọc được git. Shallow clone (Streamlit Cloud có
    thể chỉ clone 1 commit): `rev-list --count` luôn ra 1 nên không dùng
    số đếm — chỉ lấy sha + ngày giờ commit.
    """
    sha = _git_output(["rev-parse", "--short", "HEAD"])
    if not sha:
        return None
    shallow = _git_output(["rev-parse", "--is-shallow-repository"]) == "true"
    when = None
    ts = _git_output(["log", "-1", "--format=%ct"])
    if ts:
        try:
            when = datetime.fromtimestamp(int(ts), tz=_VN_TZ).strftime(
                "%d/%m/%Y %H:%M")
        except (ValueError, OSError, OverflowError):
            when = None
    count = None
    if not shallow:
        c = _git_output(["rev-list", "--count", "HEAD"])
        if c is not None:
            try:
                count = int(c)
            except ValueError:
                count = None
    return {"sha": sha, "shallow": shallow, "when": when, "count": count}


def build_version() -> str:
    """Số phiên bản hiển thị — tự tăng theo mỗi lần deploy.

    Git đầy đủ: ``0.1.<số commit> (<sha ngắn>) — <ngày giờ commit giờ VN>``
    (mỗi lần push lên main thêm ít nhất 1 commit nên số này tăng theo).
    Shallow clone: ``<sha ngắn> — <ngày giờ commit giờ VN>`` (không hiện
    số đếm sai). Không đọc được git: trả về ``__version__``.
    """
    info = _git_build()
    if info is None:
        return __version__
    base = __version__.rsplit(".", 1)[0]  # "0.1"
    when = f" — {info['when']}" if info.get("when") else ""
    if info.get("count") is not None:
        return f"{base}.{info['count']} ({info['sha']}){when}"
    return f"{info['sha']}{when}"


# --- Nạp lại module khi code trên đĩa đổi (PXD-001) -----------------------
#
# Streamlit Cloud chỉ git pull rồi chạy lại script trong cùng tiến trình
# Python: module trong src/ đã nằm trong sys.modules từ lần deploy trước
# nên import thường sẽ lấy bản CŨ. Nhưng nạp lại VÔ ĐIỀU KIỆN ở mỗi rerun
# cũng sai: reload(models) tạo ra enum RowStatus MỚI, trong khi
# st.session_state.results vẫn giữ HouseholdRow có status thuộc RowStatus
# CŨ — enum so sánh theo identity nên `r.status != RowStatus.CONFIRMED`
# luôn đúng. Vì vậy chỉ nạp lại khi code trên đĩa thật sự đổi.

_RELOAD_ORDER = ("models", "pdf_tools", "layout", "excel_source",
                 "matcher", "validate", "export", "store", "pipeline")

# Dấu của code đã nạp: (tên file, mtime_ns, size) của src/phieu_xaydung/*.py.
_loaded_sig: tuple | None = None


def _source_signature() -> tuple | None:
    """Dấu hiện tại của code trên đĩa — None khi không đọc được thư mục."""
    try:
        pkg_dir = Path(__file__).resolve().parent
        return tuple(
            (p.name, p.stat().st_mtime_ns, p.stat().st_size)
            for p in sorted(pkg_dir.glob("*.py")))
    except OSError:
        return None


def ensure_fresh_modules() -> bool:
    """Nạp lại module src/ NẾU code trên đĩa đã đổi so với lần nạp trước.

    Trả về True khi vừa nạp lại (app cần xóa results/review_idx trong
    session_state vì các object cũ — vd RowStatus — không còn đúng nữa),
    False khi code không đổi (không reload gì, enum giữ nguyên identity).
    Lần đầu chạy trong một tiến trình mới chỉ ghi nhận dấu, không reload
    vì mọi import sau đó đều đọc code mới từ đĩa.
    """
    import importlib
    import sys

    sig = _source_signature()
    if sig is None:
        return False
    loaded = getattr(sys.modules.get(__name__), "_loaded_sig", None)
    if loaded is None:
        sys.modules[__name__]._loaded_sig = sig
        return False
    if sig == loaded:
        return False
    for name in _RELOAD_ORDER:
        mod = sys.modules.get(f"phieu_xaydung.{name}")
        if mod is not None:
            importlib.reload(mod)
    importlib.reload(sys.modules[__name__])
    sys.modules[__name__]._loaded_sig = sig
    return True
