"""
M1 evaluator tests — accuracy assertions against known fixture values.

FormulasEvaluator: full accuracy tests (selected evaluator for M2+).
PycelEvaluator:    xfail accuracy tests documenting why pycel was rejected
                   (broken on Python 3.12+ due to ast.Str removal).
LibreOfficeEvaluator: skipped if soffice not on PATH.
"""

from __future__ import annotations

import importlib.util
import math
import shutil as _shutil
import sys
from pathlib import Path

import pytest

FIXTURES_DIR = Path(__file__).parent / "fixtures"
sys.path.insert(0, str(FIXTURES_DIR))


# ── Shared fixtures ───────────────────────────────────────────────────────────

@pytest.fixture(scope="module")
def m1_fixtures():
    """Build all 3 M1 benchmark models and return (paths, expected) dict."""
    from make_m1_fixtures import build_all, ALL_MODELS
    expected = build_all(FIXTURES_DIR)
    paths = {fname: FIXTURES_DIR / fname for fname, _, _ in ALL_MODELS}
    return paths, expected


# ── Helpers ───────────────────────────────────────────────────────────────────

def _close(a, b, tol=1e-6):
    if a is None and b is None:
        return True
    try:
        return math.isclose(float(a), float(b), rel_tol=tol, abs_tol=tol)
    except (TypeError, ValueError):
        return str(a) == str(b)


def _assert_cells(computed: dict, expected: dict, evaluator_name: str):
    failures = []
    for sheet, cells in expected.items():
        for addr, exp in cells.items():
            got = (computed.get(sheet) or {}).get(addr)
            if not _close(got, exp):
                failures.append(f"{sheet}!{addr}: expected={exp!r}, got={got!r}")
    if failures:
        msg = f"[{evaluator_name}] {len(failures)} mismatch(es):\n" + "\n".join(failures)
        pytest.fail(msg)


# ── FormulasEvaluator tests (M1 winner) ──────────────────────────────────────

formulas_available = pytest.mark.skipif(
    not importlib.util.find_spec("formulas"),
    reason="formulas not installed — run: pip install formulas",
)


@formulas_available
class TestFormulasEvaluator:
    @pytest.fixture(autouse=True)
    def evaluator(self):
        from testsheet.evaluator.formulas_eval import FormulasEvaluator
        self.ev = FormulasEvaluator()

    def test_simple_model(self, m1_fixtures):
        paths, expected = m1_fixtures
        computed = self.ev.compute(paths["model_simple.xlsx"])
        _assert_cells(computed, expected["model_simple.xlsx"], "formulas/simple")

    def test_intermediate_model(self, m1_fixtures):
        paths, expected = m1_fixtures
        computed = self.ev.compute(paths["model_intermediate.xlsx"])
        _assert_cells(computed, expected["model_intermediate.xlsx"], "formulas/intermediate")

    def test_complex_model(self, m1_fixtures):
        paths, expected = m1_fixtures
        computed = self.ev.compute(paths["model_complex.xlsx"])
        _assert_cells(computed, expected["model_complex.xlsx"], "formulas/complex")

    def test_returns_dict_structure(self, m1_fixtures):
        paths, _ = m1_fixtures
        computed = self.ev.compute(paths["model_simple.xlsx"])
        assert isinstance(computed, dict)
        assert "Calc" in computed
        assert isinstance(computed["Calc"], dict)

    def test_sum_formula(self, m1_fixtures):
        """SUM(A1:A4) where A1:A4 = 10,20,30,40 must equal 100."""
        paths, _ = m1_fixtures
        computed = self.ev.compute(paths["model_simple.xlsx"])
        assert _close(computed.get("Calc", {}).get("B1"), 100), \
            f"SUM expected 100, got {computed.get('Calc', {}).get('B1')}"

    def test_cross_sheet_ref(self, m1_fixtures):
        """Summary!A1 = Data!B1 = 25.0 — cross-sheet ref must resolve."""
        paths, _ = m1_fixtures
        computed = self.ev.compute(paths["model_intermediate.xlsx"])
        got = (computed.get("Summary") or {}).get("A1")
        assert _close(got, 25.0), f"Cross-sheet ref expected 25.0, got {got}"

    def test_iferror_catches_div_by_zero(self, m1_fixtures):
        """IFERROR(100/0, 'DIV_ERR') must return 'DIV_ERR', not raise."""
        paths, _ = m1_fixtures
        computed = self.ev.compute(paths["model_complex.xlsx"])
        got = (computed.get("Model") or {}).get("B1")
        assert got == "DIV_ERR", f"IFERROR expected 'DIV_ERR', got {got!r}"


# ── PycelEvaluator tests (xfail — eliminated, broken on Python 3.12+) ────────
#
# These document exactly why pycel was rejected in M1. They are marked xfail
# so they appear in the test report as expected failures, not noise.
# Root cause: ast.Str was removed in Python 3.12; pycel uses it internally.

pycel_available = pytest.mark.skipif(
    not importlib.util.find_spec("pycel"),
    reason="pycel not installed",
)

_PYCEL_BROKEN = pytest.mark.xfail(
    reason="pycel broken on Python 3.12+: ast.Str removed (see M1 findings)",
    strict=False,
)


@pycel_available
class TestPycelEvaluator:
    @pytest.fixture(autouse=True)
    def evaluator(self):
        from testsheet.evaluator.pycel_eval import PycelEvaluator
        self.ev = PycelEvaluator()

    def test_returns_dict_structure(self, m1_fixtures):
        """pycel can still load a workbook and return a dict."""
        paths, _ = m1_fixtures
        computed = self.ev.compute(paths["model_simple.xlsx"])
        assert isinstance(computed, dict)

    def test_sum_formula(self, m1_fixtures):
        """pycel evaluates simple SUM (no comparison operators)."""
        paths, _ = m1_fixtures
        computed = self.ev.compute(paths["model_simple.xlsx"])
        assert _close(computed.get("Calc", {}).get("B1"), 100)

    def test_cross_sheet_ref(self, m1_fixtures):
        """pycel resolves cross-sheet refs for non-IF formulas."""
        paths, _ = m1_fixtures
        computed = self.ev.compute(paths["model_intermediate.xlsx"])
        got = (computed.get("Summary") or {}).get("A1")
        assert _close(got, 25.0)

    @_PYCEL_BROKEN
    def test_simple_model(self, m1_fixtures):
        """Full simple model — fails on chained formula refs."""
        paths, expected = m1_fixtures
        computed = self.ev.compute(paths["model_simple.xlsx"])
        _assert_cells(computed, expected["model_simple.xlsx"], "pycel/simple")

    @_PYCEL_BROKEN
    def test_intermediate_model(self, m1_fixtures):
        """IF formulas crash pycel on Python 3.12+ (ast.Str)."""
        paths, expected = m1_fixtures
        computed = self.ev.compute(paths["model_intermediate.xlsx"])
        _assert_cells(computed, expected["model_intermediate.xlsx"], "pycel/intermediate")

    @_PYCEL_BROKEN
    def test_complex_model(self, m1_fixtures):
        """IFERROR / nested IF crash pycel on Python 3.12+ (ast.Str)."""
        paths, expected = m1_fixtures
        computed = self.ev.compute(paths["model_complex.xlsx"])
        _assert_cells(computed, expected["model_complex.xlsx"], "pycel/complex")

    @_PYCEL_BROKEN
    def test_iferror_catches_div_by_zero(self, m1_fixtures):
        """IFERROR crashes pycel on Python 3.12+ (ast.Str)."""
        paths, _ = m1_fixtures
        computed = self.ev.compute(paths["model_complex.xlsx"])
        got = (computed.get("Model") or {}).get("B1")
        assert got == "DIV_ERR"


# ── LibreOfficeEvaluator tests (skipped if not installed) ────────────────────

lo_available = pytest.mark.skipif(
    not any(_shutil.which(c) for c in [
        "soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]),
    reason="LibreOffice (soffice) not on PATH",
)


@lo_available
class TestLibreOfficeEvaluator:
    @pytest.fixture(autouse=True)
    def evaluator(self):
        from testsheet.evaluator.libreoffice import LibreOfficeEvaluator
        self.ev = LibreOfficeEvaluator()

    def test_simple_model(self, m1_fixtures):
        paths, expected = m1_fixtures
        computed = self.ev.compute(paths["model_simple.xlsx"])
        _assert_cells(computed, expected["model_simple.xlsx"], "libreoffice/simple")

    def test_intermediate_model(self, m1_fixtures):
        paths, expected = m1_fixtures
        computed = self.ev.compute(paths["model_intermediate.xlsx"])
        _assert_cells(computed, expected["model_intermediate.xlsx"], "libreoffice/intermediate")

    def test_complex_model(self, m1_fixtures):
        paths, expected = m1_fixtures
        computed = self.ev.compute(paths["model_complex.xlsx"])
        _assert_cells(computed, expected["model_complex.xlsx"], "libreoffice/complex")
