"""
M3 end-to-end rules test — full pipeline via CLI runner.

Verifies that `testsheet run` with a rules.yaml:
  - Returns PASS when workbook satisfies all rules
  - Returns FAIL when workbook violates rules
  - Rule results appear in the report alongside drift
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest
from typer.testing import CliRunner

from testsheet.cli import app

FIXTURES_DIR = Path(__file__).parent / "fixtures"
sys.path.insert(0, str(FIXTURES_DIR))

pytest.importorskip("formulas", reason="formulas not installed — run: pip install formulas")

runner = CliRunner()


@pytest.fixture(scope="module")
def rule_fixtures():
    from make_rules_fixtures import build_all_rules
    build_all_rules(FIXTURES_DIR)
    return {
        "pass_wb":   FIXTURES_DIR / "rules_pass.xlsx",
        "fail_wb":   FIXTURES_DIR / "rules_fail.xlsx",
        "pass_yaml": FIXTURES_DIR / "rules_pass.yaml",
        "fail_yaml": FIXTURES_DIR / "rules_fail.yaml",
    }


def _setup_testsheet_dir(workbook: Path, rules_yaml: Path, tmp_path: Path) -> Path:
    """
    Copy workbook and rules.yaml into tmp_path so baseline + run are isolated.
    Returns path to the workbook copy.
    """
    import shutil
    wb_copy = tmp_path / workbook.name
    shutil.copy2(workbook, wb_copy)
    ts_dir = tmp_path / ".testsheet"
    ts_dir.mkdir()
    shutil.copy2(rules_yaml, ts_dir / "rules.yaml")
    return wb_copy


class TestRulesEndToEnd:

    def test_run_with_rules_pass(self, rule_fixtures, tmp_path):
        """testsheet baseline + run on passing workbook → exit 0, PASS."""
        wb = _setup_testsheet_dir(
            rule_fixtures["pass_wb"], rule_fixtures["pass_yaml"], tmp_path
        )

        # Baseline
        result = runner.invoke(app, ["baseline", str(wb)])
        assert result.exit_code == 0, f"baseline failed: {result.output}"

        # Run
        result = runner.invoke(app, ["run", str(wb)])
        assert result.exit_code == 0, \
            f"Expected exit 0 (PASS), got {result.exit_code}.\nOutput:\n{result.output}"
        assert "NO CHANGES DETECTED" in result.output

    def test_run_with_rules_fail(self, rule_fixtures, tmp_path):
        """testsheet run on failing workbook → exit 1, FAIL, rule violations listed."""
        # Baseline from pass workbook, run against fail workbook
        import shutil

        # Set up baseline from the passing workbook
        pass_wb = _setup_testsheet_dir(
            rule_fixtures["pass_wb"], rule_fixtures["pass_yaml"], tmp_path
        )
        result = runner.invoke(app, ["baseline", str(pass_wb)])
        assert result.exit_code == 0, f"baseline failed: {result.output}"

        # Copy the fail workbook into same dir with same name for the run
        fail_wb_copy = tmp_path / rule_fixtures["pass_wb"].name
        # Overwrite with fail workbook (same filename, different content)
        shutil.copy2(rule_fixtures["fail_wb"], fail_wb_copy)

        # Also update rules.yaml to the fail variant
        shutil.copy2(rule_fixtures["fail_yaml"], tmp_path / ".testsheet" / "rules.yaml")

        result = runner.invoke(app, ["run", str(fail_wb_copy)])
        assert result.exit_code == 1, \
            f"Expected exit 1 (FAIL), got {result.exit_code}.\nOutput:\n{result.output}"
        assert "CHANGES DETECTED" in result.output

    def test_run_reports_rule_violations(self, rule_fixtures, tmp_path):
        """Rule failures appear in the console output."""
        import shutil

        pass_wb = _setup_testsheet_dir(
            rule_fixtures["pass_wb"], rule_fixtures["pass_yaml"], tmp_path
        )
        runner.invoke(app, ["baseline", str(pass_wb)])

        fail_wb_copy = tmp_path / rule_fixtures["pass_wb"].name
        shutil.copy2(rule_fixtures["fail_wb"], fail_wb_copy)
        shutil.copy2(rule_fixtures["fail_yaml"], tmp_path / ".testsheet" / "rules.yaml")

        result = runner.invoke(app, ["run", str(fail_wb_copy)])
        # At least one rule ID should appear in output
        assert any(
            rule_id in result.output
            for rule_id in [
                "revenues_positive", "no_formula_errors",
                "no_hardcoded_drivers", "total_ties_sum",
                "revenues_increasing", "gross_profit_nonneg",
            ]
        ), f"Expected rule IDs in output:\n{result.output}"

    def test_run_json_report_includes_rules(self, rule_fixtures, tmp_path):
        """--json output includes rule_results array."""
        import json, shutil

        wb = _setup_testsheet_dir(
            rule_fixtures["pass_wb"], rule_fixtures["pass_yaml"], tmp_path
        )
        runner.invoke(app, ["baseline", str(wb)])

        json_out = tmp_path / "report.json"
        result = runner.invoke(app, ["run", str(wb), "--json", str(json_out)])
        assert json_out.exists(), "JSON report not written"

        data = json.loads(json_out.read_text())
        assert "rule_results" in data
        assert isinstance(data["rule_results"], list)
        assert len(data["rule_results"]) == 6  # one per rule

    def test_no_fail_on_drift_flag(self, rule_fixtures, tmp_path):
        """--no-fail-on-drift exits 0 even when rules fail."""
        import shutil

        pass_wb = _setup_testsheet_dir(
            rule_fixtures["pass_wb"], rule_fixtures["pass_yaml"], tmp_path
        )
        runner.invoke(app, ["baseline", str(pass_wb)])

        fail_wb_copy = tmp_path / rule_fixtures["pass_wb"].name
        shutil.copy2(rule_fixtures["fail_wb"], fail_wb_copy)
        shutil.copy2(rule_fixtures["fail_yaml"], tmp_path / ".testsheet" / "rules.yaml")

        result = runner.invoke(
            app, ["run", str(fail_wb_copy), "--no-fail-on-drift"]
        )
        assert result.exit_code == 0, \
            f"Expected exit 0 with --no-fail-on-drift, got {result.exit_code}"
