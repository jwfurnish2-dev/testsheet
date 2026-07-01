"""Drift detection — compare current workbook against a baseline."""

from __future__ import annotations

import math
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from testsheet.evaluator.base import BaseEvaluator
from testsheet.parser import parse_workbook


@dataclass
class StructuralChangeWarning:
    """
    Emitted when the fraction of drifted cells exceeds *threshold*.

    A high drift fraction usually means the workbook was reorganised
    (rows/columns inserted, sheet renamed, layout changed) rather than
    genuinely modified.  Comparing cell-by-cell is still possible but
    the results may be noisy.
    """
    drift_count: int
    baseline_cell_count: int
    drift_fraction: float   # drift_count / baseline_cell_count
    threshold: float        # the configured threshold

    @property
    def message(self) -> str:
        pct = f"{self.drift_fraction * 100:.0f}%"
        return (
            f"Structural change detected: {self.drift_count} of "
            f"{self.baseline_cell_count} baseline cells drifted ({pct}), "
            f"exceeding the {self.threshold * 100:.0f}% threshold. "
            "The workbook layout may have changed."
        )

    def as_dict(self) -> dict:
        return {
            "drift_count": self.drift_count,
            "baseline_cell_count": self.baseline_cell_count,
            "drift_fraction": self.drift_fraction,
            "threshold": self.threshold,
            "message": self.message,
        }


def detect_structural_change(
    drifts: list[Drift],
    baseline: dict,
    threshold: float = 0.5,
) -> StructuralChangeWarning | None:
    """
    Return a :class:`StructuralChangeWarning` if more than *threshold*
    fraction of baseline cells drifted; otherwise return *None*.
    """
    baseline_sheets = baseline.get("sheets", {})
    total = sum(len(cells) for cells in baseline_sheets.values())
    if total == 0:
        return None
    fraction = len(drifts) / total
    if fraction > threshold:
        return StructuralChangeWarning(
            drift_count=len(drifts),
            baseline_cell_count=total,
            drift_fraction=fraction,
            threshold=threshold,
        )
    return None

# Error strings Excel / openpyxl surface
_EXCEL_ERRORS = {"#REF!", "#DIV/0!", "#VALUE!", "#N/A", "#NAME?", "#NUM!", "#NULL!"}


def _is_error(value: Any) -> bool:
    """True if *value* represents an Excel error (string or error object)."""
    if value is None:
        return False
    return str(value).strip().upper() in _EXCEL_ERRORS


def _values_equal(a: Any, b: Any, rel_tol: float = 1e-9, abs_tol: float = 1e-12) -> bool:
    """Float-tolerant equality check.

    Booleans are compared exactly (bool is a subclass of int so we guard first).
    All other numeric types (int, float, numpy scalars) are compared with tolerance.
    Strings and None fall back to exact equality.
    """
    if a is None and b is None:
        return True
    # Guard booleans before numeric coercion (float(True)==1.0 would confuse things)
    if isinstance(a, bool) or isinstance(b, bool):
        return a == b
    # Normalise Excel error objects to strings before comparing
    sa, sb = str(a).strip().upper(), str(b).strip().upper()
    if sa in _EXCEL_ERRORS and sb in _EXCEL_ERRORS:
        return sa == sb
    # Try numeric tolerance for int/float/numpy etc.
    try:
        af, bf = float(a), float(b)  # type: ignore[arg-type]
        return math.isclose(af, bf, rel_tol=rel_tol, abs_tol=abs_tol)
    except (TypeError, ValueError):
        pass
    return a == b


@dataclass
class Drift:
    """A single cell that changed from baseline."""
    sheet: str
    address: str
    kind: str  # "value_only" | "formula_only" | "both" | "new" | "deleted" | "error_introduced"
    baseline_value: Any
    current_value: Any
    baseline_formula: str | None
    current_formula: str | None

    def as_dict(self) -> dict:
        return {
            "sheet": self.sheet,
            "address": self.address,
            "kind": self.kind,
            "baseline_value": self.baseline_value,
            "current_value": self.current_value,
            "baseline_formula": self.baseline_formula,
            "current_formula": self.current_formula,
        }


def diff_workbook(
    workbook_path: Path,
    baseline: dict,
    evaluator: BaseEvaluator,
    rel_tol: float = 1e-9,
    abs_tol: float = 1e-12,
) -> list[Drift]:
    """
    Compare *workbook_path* against *baseline* and return a list of Drifts.

    Parameters
    ----------
    workbook_path:
        Current version of the workbook under test.
    baseline:
        Dict loaded from .testsheet/baseline.json.
    evaluator:
        Evaluator instance for computing current values.
    rel_tol / abs_tol:
        Float comparison tolerances.
    """
    workbook_path = Path(workbook_path).resolve()
    computed = evaluator.compute(workbook_path)
    snapshot = parse_workbook(workbook_path, computed_values=computed)

    baseline_sheets: dict[str, dict[str, dict]] = baseline.get("sheets", {})
    drifts: list[Drift] = []

    all_sheets = set(baseline_sheets) | set(snapshot.cells)

    for sheet_name in all_sheets:
        base_cells = baseline_sheets.get(sheet_name, {})
        curr_cells = snapshot.cells.get(sheet_name, {})

        all_addresses = set(base_cells) | set(curr_cells)

        for addr in all_addresses:
            base = base_cells.get(addr)
            curr = curr_cells.get(addr)

            if base is None:
                # Cell is new in current version
                drifts.append(Drift(
                    sheet=sheet_name, address=addr,
                    kind="new",
                    baseline_value=None, current_value=curr.value,
                    baseline_formula=None, current_formula=curr.formula,
                ))
                continue

            if curr is None:
                # Cell existed in baseline but is gone
                drifts.append(Drift(
                    sheet=sheet_name, address=addr,
                    kind="deleted",
                    baseline_value=base["value"], current_value=None,
                    baseline_formula=base["formula"], current_formula=None,
                ))
                continue

            value_changed = not _values_equal(base["value"], curr.value, rel_tol, abs_tol)
            formula_changed = base["formula"] != curr.formula

            if not value_changed and not formula_changed:
                continue  # no drift

            # Classify drift kind
            if value_changed and _is_error(curr.value) and not _is_error(base["value"]):
                kind = "error_introduced"
            elif value_changed and formula_changed:
                kind = "both"
            elif value_changed:
                kind = "value_only"
            else:
                kind = "formula_only"

            drifts.append(Drift(
                sheet=sheet_name, address=addr,
                kind=kind,
                baseline_value=base["value"], current_value=curr.value,
                baseline_formula=base["formula"], current_formula=curr.formula,
            ))

    return drifts
