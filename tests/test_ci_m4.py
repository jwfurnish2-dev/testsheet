"""
M4 CI gate tests — F4 acceptance criteria.

AC: `testsheet run --junit out.xml` exits non-zero on failure.
AC: Report uploaded as artifact (file exists and is non-empty).
AC: Passing model exits 0.
AC: --no-fail-on-drift exits 0 regardless.
"""

from __future__ import annotations

import json
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

import pytest
from typer.testing import CliRunner

from testsheet.cli import app

FIXTURES_DIR = Path(__file__).parent / "fixtures"
sys.path.insert(0, str(FIXTURES_DIR))

pytest.importorskip("formulas", reason="formulas not installed — run: pip install formulas")

runner = CliRunner()


# ── Helpers ───────────────────────────────────────────────────────────────────

def _baseline(wb_path: Path) -> None:
    result = runner.invoke(app, ["baseline", str(wb_path)])
    assert result.exit_code == 0, f"baseline failed:\n{result.output}"


# ── Exit code tests ───────────────────────────────────────────────────────────

class TestExitCodes:

    def test_pass_exits_zero(self, tmp_path):
        """Clean workbook → exit 0."""
        import shutil
        wb = tmp_path / "good_model.xlsx"
        shutil.copy2(FIXTURES_DIR / "good_model.xlsx", wb)
        _baseline(wb)
        result = runner.invoke(app, ["run", str(wb)])
        assert result.exit_code == 0, f"Expected 0, got {result.exit_code}\n{result.output}"

    def test_fail_exits_one(self, tmp_path):
        """Drifted workbook → exit 1."""
        import shutil
        good = tmp_path / "model.xlsx"
        shutil.copy2(FIXTURES_DIR / "good_model.xlsx", good)
        _baseline(good)
        # Overwrite with broken model (same filename so baseline path matches)
        shutil.copy2(FIXTURES_DIR / "broken_model.xlsx", good)
        result = runner.invoke(app, ["run", str(good)])
        assert result.exit_code == 1, f"Expected 1, got {result.exit_code}\n{result.output}"

    def test_no_fail_on_drift_exits_zero(self, tmp_path):
        """--no-fail-on-drift exits 0 even on drifted workbook."""
        import shutil
        good = tmp_path / "model.xlsx"
        shutil.copy2(FIXTURES_DIR / "good_model.xlsx", good)
        _baseline(good)
        shutil.copy2(FIXTURES_DIR / "broken_model.xlsx", good)
        result = runner.invoke(app, ["run", str(good), "--no-fail-on-drift"])
        assert result.exit_code == 0, f"Expected 0 with --no-fail-on-drift, got {result.exit_code}"

    def test_missing_baseline_exits_one(self, tmp_path):
        """No baseline → exit 1 with helpful message."""
        import shutil
        wb = tmp_path / "orphan.xlsx"
        shutil.copy2(FIXTURES_DIR / "good_model.xlsx", wb)
        result = runner.invoke(app, ["run", str(wb)])
        assert result.exit_code == 1
        assert "baseline" in result.output.lower()

    def test_missing_workbook_exits_one(self, tmp_path):
        result = runner.invoke(app, ["run", str(tmp_path / "nonexistent.xlsx")])
        assert result.exit_code == 1


# ── --junit flag ──────────────────────────────────────────────────────────────

class TestJunitFlag:

    def test_junit_file_created_on_pass(self, tmp_path):
        import shutil
        wb = tmp_path / "good_model.xlsx"
        shutil.copy2(FIXTURES_DIR / "good_model.xlsx", wb)
        _baseline(wb)
        junit_out = tmp_path / "results.xml"
        runner.invoke(app, ["run", str(wb), "--junit", str(junit_out)])
        assert junit_out.exists(), "--junit file not created on pass"

    def test_junit_file_created_on_fail(self, tmp_path):
        """F4 AC: report uploaded as artifact even on failure."""
        import shutil
        good = tmp_path / "model.xlsx"
        shutil.copy2(FIXTURES_DIR / "good_model.xlsx", good)
        _baseline(good)
        shutil.copy2(FIXTURES_DIR / "broken_model.xlsx", good)
        junit_out = tmp_path / "results.xml"
        runner.invoke(app, ["run", str(good), "--junit", str(junit_out)])
        assert junit_out.exists(), "--junit file not created on failure"

    def test_junit_is_valid_xml(self, tmp_path):
        import shutil
        wb = tmp_path / "good_model.xlsx"
        shutil.copy2(FIXTURES_DIR / "good_model.xlsx", wb)
        _baseline(wb)
        junit_out = tmp_path / "results.xml"
        runner.invoke(app, ["run", str(wb), "--junit", str(junit_out)])
        tree = ET.parse(str(junit_out))
        assert tree.getroot().tag == "testsuites"

    def test_junit_failures_nonzero_on_drift(self, tmp_path):
        import shutil
        good = tmp_path / "model.xlsx"
        shutil.copy2(FIXTURES_DIR / "good_model.xlsx", good)
        _baseline(good)
        shutil.copy2(FIXTURES_DIR / "broken_model.xlsx", good)
        junit_out = tmp_path / "results.xml"
        runner.invoke(app, ["run", str(good), "--junit", str(junit_out)])
        root = ET.parse(str(junit_out)).getroot()
        assert int(root.attrib["failures"]) > 0

    def test_junit_failures_zero_on_pass(self, tmp_path):
        import shutil
        wb = tmp_path / "good_model.xlsx"
        shutil.copy2(FIXTURES_DIR / "good_model.xlsx", wb)
        _baseline(wb)
        junit_out = tmp_path / "results.xml"
        runner.invoke(app, ["run", str(wb), "--junit", str(junit_out)])
        root = ET.parse(str(junit_out)).getroot()
        assert root.attrib["failures"] == "0"


# ── --json flag ───────────────────────────────────────────────────────────────

class TestJsonFlag:

    def test_json_file_created(self, tmp_path):
        import shutil
        wb = tmp_path / "good_model.xlsx"
        shutil.copy2(FIXTURES_DIR / "good_model.xlsx", wb)
        _baseline(wb)
        json_out = tmp_path / "report.json"
        runner.invoke(app, ["run", str(wb), "--json", str(json_out)])
        assert json_out.exists()

    def test_json_schema(self, tmp_path):
        import shutil
        wb = tmp_path / "good_model.xlsx"
        shutil.copy2(FIXTURES_DIR / "good_model.xlsx", wb)
        _baseline(wb)
        json_out = tmp_path / "report.json"
        runner.invoke(app, ["run", str(wb), "--json", str(json_out)])
        data = json.loads(json_out.read_text())
        for key in ("workbook", "passed", "drift_count", "rule_failure_count",
                    "drifts", "rule_results"):
            assert key in data

    def test_json_passed_true_on_clean(self, tmp_path):
        import shutil
        wb = tmp_path / "good_model.xlsx"
        shutil.copy2(FIXTURES_DIR / "good_model.xlsx", wb)
        _baseline(wb)
        json_out = tmp_path / "report.json"
        runner.invoke(app, ["run", str(wb), "--json", str(json_out)])
        data = json.loads(json_out.read_text())
        assert data["passed"] is True
        assert data["drift_count"] == 0

    def test_json_passed_false_on_drift(self, tmp_path):
        import shutil
        good = tmp_path / "model.xlsx"
        shutil.copy2(FIXTURES_DIR / "good_model.xlsx", good)
        _baseline(good)
        shutil.copy2(FIXTURES_DIR / "broken_model.xlsx", good)
        json_out = tmp_path / "report.json"
        runner.invoke(app, ["run", str(good), "--json", str(json_out)])
        data = json.loads(json_out.read_text())
        assert data["passed"] is False
        assert data["drift_count"] > 0


# ── --html flag ───────────────────────────────────────────────────────────────

class TestHtmlFlag:

    def test_html_file_created(self, tmp_path):
        import shutil
        wb = tmp_path / "good_model.xlsx"
        shutil.copy2(FIXTURES_DIR / "good_model.xlsx", wb)
        _baseline(wb)
        html_out = tmp_path / "report.html"
        runner.invoke(app, ["run", str(wb), "--html", str(html_out)])
        assert html_out.exists()

    def test_html_is_nonempty(self, tmp_path):
        import shutil
        wb = tmp_path / "good_model.xlsx"
        shutil.copy2(FIXTURES_DIR / "good_model.xlsx", wb)
        _baseline(wb)
        html_out = tmp_path / "report.html"
        runner.invoke(app, ["run", str(wb), "--html", str(html_out)])
        assert html_out.stat().st_size > 500

    def test_html_contains_pass_on_clean(self, tmp_path):
        import shutil
        wb = tmp_path / "good_model.xlsx"
        shutil.copy2(FIXTURES_DIR / "good_model.xlsx", wb)
        _baseline(wb)
        html_out = tmp_path / "report.html"
        runner.invoke(app, ["run", str(wb), "--html", str(html_out)])
        assert "NO CHANGES DETECTED" in html_out.read_text(encoding="utf-8")

    def test_html_contains_fail_on_drift(self, tmp_path):
        import shutil
        good = tmp_path / "model.xlsx"
        shutil.copy2(FIXTURES_DIR / "good_model.xlsx", good)
        _baseline(good)
        shutil.copy2(FIXTURES_DIR / "broken_model.xlsx", good)
        html_out = tmp_path / "report.html"
        runner.invoke(app, ["run", str(good), "--html", str(html_out)])
        assert "CHANGES DETECTED" in html_out.read_text(encoding="utf-8")

    def test_html_lists_drift_sheet(self, tmp_path):
        """HTML report groups drifts by sheet — sheet name must appear."""
        import shutil
        good = tmp_path / "model.xlsx"
        shutil.copy2(FIXTURES_DIR / "good_model.xlsx", good)
        _baseline(good)
        shutil.copy2(FIXTURES_DIR / "broken_model.xlsx", good)
        html_out = tmp_path / "report.html"
        runner.invoke(app, ["run", str(good), "--html", str(html_out)])
        content = html_out.read_text(encoding="utf-8")
        assert "Model" in content or "Summary" in content

    def test_all_three_reports_simultaneously(self, tmp_path):
        """--junit, --json, --html can all be specified at once."""
        import shutil
        wb = tmp_path / "good_model.xlsx"
        shutil.copy2(FIXTURES_DIR / "good_model.xlsx", wb)
        _baseline(wb)
        junit_out = tmp_path / "r.xml"
        json_out  = tmp_path / "r.json"
        html_out  = tmp_path / "r.html"
        result = runner.invoke(app, [
            "run", str(wb),
            "--junit", str(junit_out),
            "--json",  str(json_out),
            "--html",  str(html_out),
        ])
        assert result.exit_code == 0
        assert junit_out.exists()
        assert json_out.exists()
        assert html_out.exists()
