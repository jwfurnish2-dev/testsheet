"""Abstract base class for workbook evaluators."""

from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any


class BaseEvaluator(ABC):
    """
    Contract for an evaluator that computes cell values from a workbook.

    Implementations must handle formula recalculation; callers should not
    assume openpyxl cached values are accurate.
    """

    @abstractmethod
    def compute(
        self,
        path: Path,
        sheets: list[str] | None = None,
    ) -> dict[str, dict[str, Any]]:
        """
        Compute all cell values in *path*.

        Returns
        -------
        dict
            ``{sheet_name: {address: value}}``
            where *address* is ``"B5"`` style and *value* is Python-native.
            Cells with errors should surface the error string (e.g. ``"#REF!"``).
        """

    @property
    @abstractmethod
    def name(self) -> str:
        """Short identifier used in reports and CLI output."""

    def is_available(self) -> bool:
        """Return True if this evaluator's runtime dependencies are present."""
        return True
