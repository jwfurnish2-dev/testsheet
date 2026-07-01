"""
Generate before/after workbook pairs for each drift classification.

Drift kinds tested:
  value_only        — same formula, input literal changed → different result
  formula_only      — formula string changed, result happens to be identical
  both              — formula string changed AND result changed
  error_introduced  — formula result becomes a #DIV/0! error
  new_cell          — cell exists in current but not in baseline
  deleted_cell      — cell existed in baseline but removed in current

Each pair: (before.xlsx, after.xlsx) + expected list of Drift dicts.

Usage:
  python tests/fixtures/make_drift_fixtures.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

FIXTURES_DIR = Path(__file__).parent


def _wb_with(cells: dict) -> Workbook:
    """Create a single-sheet workbook ('Sheet') with given {addr: value} cells."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet"
    for addr, val in cells.items():
        ws[addr] = val
    return wb


# ─────────────────────────────────────────────────────────────────────────────
# value_only: input literal A1 changes (10 → 99), formula =A1*10 unchanged
#   before: A1=10, B1=A1*10  → B1=100
#   after:  A1=99, B1=A1*10  → B1=990  (value drifted, formula same)
# ─────────────────────────────────────────────────────────────────────────────

def make_value_only(out_dir: Path) -> dict:
    before = _wb_with({"A1": 10, "B1": "=A1*10"})
    after  = _wb_with({"A1": 99, "B1": "=A1*10"})
    before.save(out_dir / "value_only_before.xlsx")
    after.save(out_dir / "value_only_after.xlsx")
    return {
        "expected_drifts": [
            {"sheet": "Sheet", "address": "A1", "kind": "value_only",
             "baseline_value": 10, "current_value": 99},
            {"sheet": "Sheet", "address": "B1", "kind": "value_only",
             "baseline_value": 100, "current_value": 990},
        ]
    }


# ─────────────────────────────────────────────────────────────────────────────
# formula_only: formula string changes, but result is identical
#   before: B1=A1+A2+A3  (A1=5, A2=10, A3=15) → B1=30
#   after:  B1=SUM(A1:A3)                       → B1=30  (same result!)
# ─────────────────────────────────────────────────────────────────────────────

def make_formula_only(out_dir: Path) -> dict:
    before = _wb_with({"A1": 5, "A2": 10, "A3": 15, "B1": "=A1+A2+A3"})
    after  = _wb_with({"A1": 5, "A2": 10, "A3": 15, "B1": "=SUM(A1:A3)"})
    before.save(out_dir / "formula_only_before.xlsx")
    after.save(out_dir / "formula_only_after.xlsx")
    return {
        "expected_drifts": [
            {"sheet": "Sheet", "address": "B1", "kind": "formula_only",
             "baseline_formula": "=A1+A2+A3", "current_formula": "=SUM(A1:A3)"},
        ]
    }


# ─────────────────────────────────────────────────────────────────────────────
# both: formula string changes AND result changes
#   before: B1=A1*2  (A1=10) → B1=20
#   after:  B1=A1+100        → B1=110
# ─────────────────────────────────────────────────────────────────────────────

def make_both(out_dir: Path) -> dict:
    before = _wb_with({"A1": 10, "B1": "=A1*2"})
    after  = _wb_with({"A1": 10, "B1": "=A1+100"})
    before.save(out_dir / "both_before.xlsx")
    after.save(out_dir / "both_after.xlsx")
    return {
        "expected_drifts": [
            {"sheet": "Sheet", "address": "B1", "kind": "both",
             "baseline_value": 20, "current_value": 110,
             "baseline_formula": "=A1*2", "current_formula": "=A1+100"},
        ]
    }


# ─────────────────────────────────────────────────────────────────────────────
# error_introduced: formula result becomes an error
#   before: A2=5,  B1=A1/A2 → B1=2.0
#   after:  A2=0,  B1=A1/A2 → B1=#DIV/0!
# ─────────────────────────────────────────────────────────────────────────────

def make_error_introduced(out_dir: Path) -> dict:
    before = _wb_with({"A1": 10, "A2": 5,  "B1": "=A1/A2"})
    after  = _wb_with({"A1": 10, "A2": 0,  "B1": "=A1/A2"})
    before.save(out_dir / "error_before.xlsx")
    after.save(out_dir / "error_after.xlsx")
    return {
        "expected_drifts": [
            {"sheet": "Sheet", "address": "A2", "kind": "value_only",
             "baseline_value": 5, "current_value": 0},
            {"sheet": "Sheet", "address": "B1", "kind": "error_introduced"},
        ]
    }


# ─────────────────────────────────────────────────────────────────────────────
# new_cell: cell appears in after that wasn't in before
# ─────────────────────────────────────────────────────────────────────────────

def make_new_cell(out_dir: Path) -> dict:
    before = _wb_with({"A1": 1, "A2": 2})
    after  = _wb_with({"A1": 1, "A2": 2, "A3": 99})
    before.save(out_dir / "new_cell_before.xlsx")
    after.save(out_dir / "new_cell_after.xlsx")
    return {
        "expected_drifts": [
            {"sheet": "Sheet", "address": "A3", "kind": "new"},
        ]
    }


# ─────────────────────────────────────────────────────────────────────────────
# deleted_cell: cell in before is gone in after
# ─────────────────────────────────────────────────────────────────────────────

def make_deleted_cell(out_dir: Path) -> dict:
    before = _wb_with({"A1": 1, "A2": 2, "A3": 3})
    after  = _wb_with({"A1": 1, "A2": 2})
    before.save(out_dir / "deleted_cell_before.xlsx")
    after.save(out_dir / "deleted_cell_after.xlsx")
    return {
        "expected_drifts": [
            {"sheet": "Sheet", "address": "A3", "kind": "deleted"},
        ]
    }


# ─────────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────────

ALL_DRIFT_FIXTURES = [
    ("value_only",       make_value_only),
    ("formula_only",     make_formula_only),
    ("both",             make_both),
    ("error_introduced", make_error_introduced),
    ("new_cell",         make_new_cell),
    ("deleted_cell",     make_deleted_cell),
]


def build_all_drift(out_dir: Path = FIXTURES_DIR) -> dict[str, dict]:
    out_dir.mkdir(parents=True, exist_ok=True)
    results = {}
    for name, builder in ALL_DRIFT_FIXTURES:
        results[name] = builder(out_dir)
        print(f"  Written: {name} (before + after)")
    return results


if __name__ == "__main__":
    print("Building drift fixtures...")
    build_all_drift()
    print("Done.")
