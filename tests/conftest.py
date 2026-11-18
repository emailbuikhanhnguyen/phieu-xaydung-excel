"""Test bộ lõi trên dữ liệu GIẢ (tests/fixtures/fake-dossier).

Chạy: .venv/bin/python -m pytest tests/ -q
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

FIX = Path(__file__).resolve().parent / "fixtures" / "fake-dossier"


@pytest.fixture(scope="session")
def truth():
    return json.loads((FIX / "ground-truth.json").read_text(encoding="utf-8"))
