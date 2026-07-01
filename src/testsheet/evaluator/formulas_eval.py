"""
M1 Spike — Evaluator C: 'formulas' library (pure Python, Python 3.12+ compatible).

https://github.com/vinci1it2000/formulas

Pros:  zero system deps, actively maintained, Python 3.12+ compatible,
       broad formula coverage including IF, IFERROR, ROUND, cross-sheet refs.
Cons:  slower first load (~10s) due to schedula compilation; some obscure
       functions missing.

Install:  pip install formulas

formulas API (confirmed via debug):
  - ExcelModel().loads(path).finish() compiles the workbook graph
  - solution = xl_model.calculate()  → schedula Solution (dict-like)
  - solution[ref] is a Ranges object; .value is a numpy array [[result]]
  - Keys are "'[filename.xlsx]SHEETNAME'!A1"  (sheet names UPPERCASED)
  - Range refs like A1:A4 also appear — skip non-scalar refs
"""

from __future__ import annotations

import logging
import math
import re
from pathlib import Path
from typing import Any

from testsheet.evaluator.base import BaseEvaluator

logger = logging.getLogger(__name__)

# Matches single-cell refs only: '[file.xlsx]SHEET'!A1
_SINGLE_CELL_RE = re.compile(
    r"'?\[.*?\]([^'!]+)'?!([A-Z]+)(\d+)$",
    re.IGNORECASE,
)


_EXCEL_ERRORS = {"#REF!", "#DIV/0!", "#VALUE!", "#N/A", "#NAME?", "#NUM!", "#NULL!"}


def _unwrap_ranges(val: Any) -> Any:
    """Extract a scalar Python value from a formulas Ranges object.

    Error values (ExcelError objects from the formulas library) are
    normalised to their canonical uppercase string form (e.g. ``"#REF!"``).
    """
    if val is None:
        return None

    # Ranges objects have a .value numpy ndarray
    arr = getattr(val, "value", val)

    try:
        import numpy as np
        if isinstance(arr, np.ndarray):
            flat = arr.flatten()
            if len(flat) == 0:
                return None
            item = flat[0]
            if isinstance(item, np.generic):
                item = item.item()
            # Normalise Excel error objects to their string representation
            s = str(item).strip().upper()
            if s in _EXCEL_ERRORS:
                return s
            if isinstance(item, float) and (math.isnan(item) or math.isinf(item)):
                return None
            return item
    except ImportError:
        pass

    # Normalise non-numpy error objects (e.g. formulas.tokens.operand.ExcelError)
    s = str(arr).strip().upper()
    if s in _EXCEL_ERRORS:
        return s

    # Fallback: plain Python scalar
    if isinstance(arr, (int, float, bool, str)):
        return arr
    return None


class FormulasEvaluator(BaseEvaluator):
    """Evaluate workbook using the 'formulas' library (pure Python, 3.12+ safe)."""

    @property
    def name(self) -> str:
        return "formulas"

    def is_available(self) -> bool:
        try:
            import formulas  # noqa: F401
            return True
        except ImportError:
            return False

    def compute(
        self,
        path: Path,
        sheets: list[str] | None = None,
    ) -> dict[str, dict[str, Any]]:
        """
        Compute all cell values via the formulas ExcelModel.

        Returns ``{sheet_name: {address: value}}`` using original-case sheet
        names (as openpyxl sees them).
        """
        if not self.is_available():
            logger.warning(
                "formulas not installed — returning empty value map. "
                "Run: pip install formulas"
            )
            return {}

        path = Path(path).resolve()

        try:
            import formulas as _formulas  # type: ignore[import]
            import openpyxl
        except ImportError:
            return {}

        # Build uppercase → original-case sheet name map using openpyxl
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            sheet_name_map = {s.upper(): s for s in wb.sheetnames}
            wb.close()
        except Exception as exc:
            logger.error("FormulasEvaluator: openpyxl failed on %s: %s", path.name, exc)
            return {}

        try:
            xl_model = _formulas.ExcelModel().loads(str(path)).finish()
            solution = xl_model.calculate()
        except Exception as exc:
            logger.error("FormulasEvaluator: failed on %s: %s", path.name, exc)
            return {}

        result: dict[str, dict[str, Any]] = {}

        for ref, val in solution.items():
            m = _SINGLE_CELL_RE.match(str(ref))
            if m is None:
                continue  # skip range refs (A1:A4 etc.)

            sheet_upper = m.group(1).upper()
            col = m.group(2).upper()
            row = m.group(3)
            address = f"{col}{row}"

            # Map back to original sheet name
            sheet_name = sheet_name_map.get(sheet_upper)
            if sheet_name is None:
                continue  # sheet not in workbook (shouldn't happen)

            # Filter to requested sheets
            if sheets and sheet_name not in sheets:
                continue

            result.setdefault(sheet_name, {})[address] = _unwrap_ranges(val)

        return result
