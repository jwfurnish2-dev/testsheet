"""
M0 smoke tests — verify the scaffold is wired up correctly.

These tests use openpyxl cached values (no external evaluator) so they run
in any environment without LibreOffice or pycel installed.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from testsheet.baseline import capture_baseline, load_baseline, _snapshot_to_dict
from testsheet.diff import diff_workbook
from testsheet.evaluator.base import BaseEvaluator
from testsheet.parser import parse_workbook


# ── Null evaluator — returns empty map so openpyxl cached values are used ──

class NullEvaluator(BaseEvaluator):
    """Returns no computed values — falls back to openpyxl cached values."""

    @property
    def name(self) -> str:
        return "null"

    def compute(self, path, sheets=None):
        return {}


# ── Parser smoke ────────────────────────────────────────────────────────────

class TestParser:
    def test_parses_good_model(self, good_model):
        snap = parse_workbook(good_model)
        assert "Model" in snap.cells
        assert "Summary" in snap.cells
        assert "Output" in snap.cells

    def test_model_has_formula_cells(self, good_model):
        snap = parse_workbook(good_model)
        model_cells = snap.cells["Model"]
        # F2 should be =SUM(B2:E2)
        assert "F2" in model_cells
        cell = model_cells["F2"]
        assert cell.formula is not None
        assert "SUM" in cell.formula.upper()

    def test_literal_values_captured(self, good_model):
        snap = parse_workbook(good_model)
        # B2 = 100 (literal)
        b2 = snap.cells["Model"]["B2"]
        assert b2.formula is None
        assert b2.value == 100


# ── Baseline capture + load ──────────────────────────────────────────────────

class TestBaseline:
    def test_capture_creates_file(self, good_model, tmp_path):
        ev = NullEvaluator()
        baseline_path = capture_baseline(good_model, evaluator=ev)
        assert baseline_path.exists()

    def test_baseline_schema(self, good_model):
        ev = NullEvaluator()
        baseline_path = capture_baseline(good_model, evaluator=ev)
        data = load_baseline(baseline_path)
        assert "schema_version" in data
        assert "sheets" in data
        assert "Model" in data["sheets"]

    def test_cell_has_hash(self, good_model):
        ev = NullEvaluator()
        baseline_path = capture_baseline(good_model, evaluator=ev)
        data = load_baseline(baseline_path)
        cell = data["sheets"]["Model"]["B2"]
        assert "hash" in cell
        assert len(cell["hash"]) == 16  # 16-char hex


# ── Diff engine ─────────────────────────────────────────────────────────────

class TestDiff:
    def test_no_drift_on_same_file(self, good_model):
        """Baselining then running against the same file → zero drift."""
        ev = NullEvaluator()
        capture_baseline(good_model, evaluator=ev)
        baseline_path = good_model.parent / ".testsheet" / "baseline.json"
        data = load_baseline(baseline_path)
        drifts = diff_workbook(good_model, data, evaluator=ev)
        assert drifts == [], f"Expected zero drift, got: {drifts}"

    def test_drift_detected_on_broken_model(self, good_model, broken_model):
        """
        Baseline the good model, then run against the broken model.
        Expect at least one drift.
        """
        ev = NullEvaluator()
        good_baseline_path = capture_baseline(good_model, evaluator=ev)
        data = load_baseline(good_baseline_path)

        # diff_workbook takes the baseline dict directly — no copy needed
        drifts = diff_workbook(broken_model, data, evaluator=ev)
        assert len(drifts) >= 1, "Expected at least one drift on broken model"

    def test_drift_kind_value_only(self, good_model, broken_model):
        """The planted value change in Model!D2 (Q3 revenue) should be flagged."""
        ev = NullEvaluator()
        baseline_path = capture_baseline(good_model, evaluator=ev)
        data = load_baseline(baseline_path)

        drifts = diff_workbook(broken_model, data, evaluator=ev)
        kinds = {d.kind for d in drifts}
        addresses = {f"{d.sheet}!{d.address}" for d in drifts}

        # Revenue Q3 is in Model!D2 — should appear as a drift
        assert "Model!D2" in addresses


# ── CLI smoke (import only — no subprocess) ─────────────────────────────────

class TestCLI:
    def test_cli_importable(self):
        from testsheet.cli import app
        assert app is not None

    def test_version_string(self):
        from testsheet import __version__
        assert __version__ == "0.1.0"


# ── Rules engine stubs ──────────────────────────────────────────────────────

class TestRulesEngine:
    def test_load_empty_rules(self, tmp_path):
        from testsheet.rules.engine import load_rules
        rules_file = tmp_path / "rules.yaml"
        rules_file.write_text("rules: []\n")
        rules = load_rules(rules_file)
        assert rules == []

    def test_range_bound_pass(self):
        from testsheet.rules.engine import _rule_range_bound
        computed = {"Summary": {"B2": 460.0}}
        rule = {"id": "rev_positive", "type": "range_bound",
                "cell": "Summary!B2", "min": 0}
        result = _rule_range_bound(rule, computed, None)
        assert result.passed

    def test_range_bound_fail(self):
        from testsheet.rules.engine import _rule_range_bound
        computed = {"Summary": {"B2": -10.0}}
        rule = {"id": "rev_positive", "type": "range_bound",
                "cell": "Summary!B2", "min": 0}
        result = _rule_range_bound(rule, computed, None)
        assert not result.passed

    def test_no_error_pass(self):
        from testsheet.rules.engine import _rule_no_error
        computed = {"Output": {"B1": 123.0}}
        rule = {"id": "no_output_errors", "type": "no_error", "range": "Output!B1"}
        result = _rule_no_error(rule, computed, None)
        assert result.passed

    def test_no_error_fail(self):
        from testsheet.rules.engine import _rule_no_error
        computed = {"Output": {"B1": "#DIV/0!"}}
        rule = {"id": "no_output_errors", "type": "no_error", "range": "Output!B1"}
        result = _rule_no_error(rule, computed, None)
        assert not result.passed
