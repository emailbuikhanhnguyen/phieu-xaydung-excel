"""phieu-xaydung-excel — chuyển phiếu thu thập thông tin xây dựng (PDF/Excel)
thành file Excel chuẩn theo mẫu Phú Hiệp.

Chạy cục bộ, không gọi API AI. Đọc trực tiếp Excel nguồn, dùng text layer
của PDF khi có, xử lý hình ảnh bằng code; chữ viết tay duyệt thủ công theo ô.
"""

__version__ = "0.1.0"


def _git_build() -> tuple[int, str] | None:
    """(số commit, SHA ngắn) của HEAD — None khi không đọc được git.

    Chạy `git` trên chính thư mục repo đang chứa code (bản deploy trên
    Streamlit Cloud là một bản clone của repo nên vẫn có .git).
    """
    try:
        import subprocess
        from pathlib import Path
        root = Path(__file__).resolve().parents[2]
        count = subprocess.run(
            ["git", "rev-list", "--count", "HEAD"], cwd=root,
            capture_output=True, text=True, timeout=10)
        sha = subprocess.run(
            ["git", "rev-parse", "--short", "HEAD"], cwd=root,
            capture_output=True, text=True, timeout=10)
        if count.returncode == 0 and sha.returncode == 0:
            return int(count.stdout.strip()), sha.stdout.strip()
    except Exception:
        pass
    return None


def build_version() -> str:
    """Số phiên bản hiển thị — tự tăng theo mỗi lần deploy.

    Mỗi lần push lên main thêm ít nhất 1 commit nên số commit tăng theo
    từng lần deploy: ``0.1.<số commit> (<sha ngắn>)``, vd ``0.1.5
    (020a3d3)``. Khi không đọc được git (vd chạy từ bản copy không có
    .git) thì trả về ``__version__``.
    """
    info = _git_build()
    if info is None:
        return __version__
    count, sha = info
    base = __version__.rsplit(".", 1)[0]  # "0.1"
    return f"{base}.{count} ({sha})"
