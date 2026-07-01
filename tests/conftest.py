"""
pytest configuration — auto-generate fixture workbooks before the test session.
"""

from __future__ import annotations

from pathlib import Path

import pytest

FIXTURES_DIR = Path(__file__).parent / "fixtures"


def pytest_configure(config):
    """Generate fixture workbooks if they don't exist yet."""
    good = FIXTURES_DIR / "good_model.xlsx"
    broken = FIXTURES_DIR / "broken_model.xlsx"

    if not good.exists() or not broken.exists():
        import sys
        sys.path.insert(0, str(FIXTURES_DIR))
        from make_fixtures import make_good_model, make_broken_model
        FIXTURES_DIR.mkdir(parents=True, exist_ok=True)
        make_good_model(good)
        make_broken_model(broken)


@pytest.fixture
def good_model(tmp_path) -> Path:
    """Return path to the good_model fixture (copied to tmp_path for isolation)."""
    import shutil
    src = FIXTURES_DIR / "good_model.xlsx"
    dst = tmp_path / "good_model.xlsx"
    shutil.copy2(src, dst)
    return dst


@pytest.fixture
def broken_model(tmp_path) -> Path:
    """Return path to the broken_model fixture (copied to tmp_path)."""
    import shutil
    src = FIXTURES_DIR / "broken_model.xlsx"
    dst = tmp_path / "broken_model.xlsx"
    shutil.copy2(src, dst)
    return dst


@pytest.fixture
def pycel_evaluator():
    from testsheet.evaluator.pycel_eval import PycelEvaluator
    return PycelEvaluator()
