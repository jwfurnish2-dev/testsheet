"""Baseline capture and persistence (.testsheet/baseline.json)."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any

from testsheet.evaluator.base import BaseEvaluator
from testsheet.parser import WorkbookSnapshot, parse_workbook


# ── Schema version — bump when the JSON format changes in a breaking way ──
SCHEMA_VERSION = 1


def _hash_cell(value: Any, formula: str | None) -> str:
    """Stable SHA-256 hash of a cell's value + formula."""
    payload = json.dumps(
        {"v": str(value) if value is not None else None, "f": formula},
        sort_keys=True,
        ensure_ascii=True,
    ).encode()
    return hashlib.sha256(payload).hexdigest()[:16]


def capture_baseline(
    workbook_path: Path,
    evaluator: BaseEvaluator,
    sheets: list[str] | None = None,
) -> Path:
    """
    Compute values, parse the workbook, and write .testsheet/baseline.json.

    Returns the path of the written baseline file.
    """
    workbook_path = Path(workbook_path).resolve()
    computed = evaluator.compute(workbook_path, sheets=sheets)
    snapshot = parse_workbook(workbook_path, sheets=sheets, computed_values=computed)

    baseline_dir = workbook_path.parent / ".testsheet"
    baseline_dir.mkdir(exist_ok=True)
    baseline_path = baseline_dir / "baseline.json"

    data = _snapshot_to_dict(snapshot)
    baseline_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")

    return baseline_path


def load_baseline(baseline_path: Path) -> dict:
    """Load and return the raw baseline dict from disk."""
    return json.loads(Path(baseline_path).read_text(encoding="utf-8"))


def _snapshot_to_dict(snapshot: WorkbookSnapshot) -> dict:
    """Serialise a WorkbookSnapshot to the baseline JSON schema."""
    sheets_out: dict[str, dict[str, dict]] = {}
    for sheet_name, cells in snapshot.cells.items():
        cells_out: dict[str, dict] = {}
        for address, cell in cells.items():
            cells_out[address] = {
                "value": _serializable(cell.value),
                "formula": cell.formula,
                "hash": _hash_cell(cell.value, cell.formula),
            }
        sheets_out[sheet_name] = cells_out

    return {
        "schema_version": SCHEMA_VERSION,
        "workbook": snapshot.path.name,
        "named_ranges": snapshot.named_ranges,
        "sheets": sheets_out,
    }


def _serializable(value: Any) -> Any:
    """Convert a cell value to a JSON-safe type."""
    if value is None:
        return None
    if isinstance(value, (int, float, bool, str)):
        return value
    # datetime, date, timedelta — convert to ISO string
    try:
        return value.isoformat()
    except AttributeError:
        return str(value)
