"""Evaluator factory — select by name."""

from __future__ import annotations

from testsheet.evaluator.base import BaseEvaluator


def get_evaluator(name: str) -> BaseEvaluator:
    """Return an evaluator instance by name ('formulas', 'pycel', or 'libreoffice')."""
    name = name.lower().strip()
    if name in ("formulas", "formulas_eval"):
        from testsheet.evaluator.formulas_eval import FormulasEvaluator
        return FormulasEvaluator()
    if name == "pycel":
        from testsheet.evaluator.pycel_eval import PycelEvaluator
        return PycelEvaluator()
    if name in ("libreoffice", "lo"):
        from testsheet.evaluator.libreoffice import LibreOfficeEvaluator
        return LibreOfficeEvaluator()
    raise ValueError(f"Unknown evaluator '{name}'. Choose 'formulas', 'pycel', or 'libreoffice'.")
