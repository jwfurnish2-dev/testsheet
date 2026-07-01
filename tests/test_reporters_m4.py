"""
M4 reporter tests — F5 acceptance criteria.

JSON:   schema documented; all required keys present.
JUnit:  valid XML; failure elements on failing tests; CI-parseable.
HTML:   standalone file; PASS/FAIL badge; drifts grouped by sheet.
Console: smoke only (rich output hard to assert on exactly).
"""

from __future__ import annotations

import json
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path

import pytest


# ── Shared test data helpers ──────────────────────────────────────────────────

@dataclass
class FakeDrift:
    sheet: str
    address: str
    kind: str
    baseline_value: object
    current_value: object
    baseline_formula: str | None = None
    current_formula: str | None = None

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


@dataclass
class FakeRuleResult:
    rule_id: str
    rule_type: str
    passed: bool
    message: str
    detail: dict = field(default_factory=dict)

    def as_dict(self) -> dict:
        return {
            "rule_id": self.rule_id,
            "rule_type": self.rule_type,
            "passed": self.passed,
            "message": self.message,
            "detail": self.detail,
        }


def _passing_report(workbook="model.xlsx") -> dict:
    return {
        "workbook": workbook,
        "passed": True,
        "drifts": [],
        "rule_results": [
            FakeRuleResult("rev_pos", "range_bound", True, "OK"),
            FakeRuleResult("no_err",  "no_error",    True, "OK"),
        ],
    }


def _failing_report(workbook="model.xlsx") -> dict:
    return {
        "workbook": workbook,
        "passed": False,
        "drifts": [
            FakeDrift("Summary", "B2", "value_only", 460.0, 999.0),
            FakeDrift("Model",   "D2", "error_introduced", 120, "#DIV/0!",
                      "=A1*10", "=A1/0"),
        ],
        "rule_results": [
            FakeRuleResult("rev_pos",  "range_bound", False, "Summary!B2=999 > max=500"),
            FakeRuleResult("no_err",   "no_error",    True,  "OK"),
            FakeRuleResult("hardcode", "no_hardcode_in_range", False, "Hardcoded: Model!C2=42"),
        ],
    }


# ── JSON reporter ─────────────────────────────────────────────────────────────

class TestJsonReporter:

    def test_creates_file(self, tmp_path):
        from testsheet.reporters.json_reporter import write_json
        out = tmp_path / "report.json"
        write_json(_passing_report(), out)
        assert out.exists()

    def test_valid_json(self, tmp_path):
        from testsheet.reporters.json_reporter import write_json
        out = tmp_path / "report.json"
        write_json(_passing_report(), out)
        data = json.loads(out.read_text(encoding="utf-8"))
        assert isinstance(data, dict)

    def test_required_top_level_keys(self, tmp_path):
        from testsheet.reporters.json_reporter import write_json
        out = tmp_path / "report.json"
        write_json(_failing_report(), out)
        data = json.loads(out.read_text(encoding="utf-8"))
        for key in ("workbook", "passed", "drift_count", "rule_failure_count",
                    "drifts", "rule_results"):
            assert key in data, f"Missing key: {key}"

    def test_pass_report_schema(self, tmp_path):
        from testsheet.reporters.json_reporter import write_json
        out = tmp_path / "report.json"
        write_json(_passing_report("q4.xlsx"), out)
        data = json.loads(out.read_text(encoding="utf-8"))
        assert data["workbook"] == "q4.xlsx"
        assert data["passed"] is True
        assert data["drift_count"] == 0
        assert data["rule_failure_count"] == 0
        assert data["drifts"] == []

    def test_fail_report_drift_count(self, tmp_path):
        from testsheet.reporters.json_reporter import write_json
        out = tmp_path / "report.json"
        write_json(_failing_report(), out)
        data = json.loads(out.read_text(encoding="utf-8"))
        assert data["passed"] is False
        assert data["drift_count"] == 2
        assert data["rule_failure_count"] == 2

    def test_drift_fields(self, tmp_path):
        from testsheet.reporters.json_reporter import write_json
        out = tmp_path / "report.json"
        write_json(_failing_report(), out)
        data = json.loads(out.read_text(encoding="utf-8"))
        drift = data["drifts"][0]
        for key in ("sheet", "address", "kind", "baseline_value", "current_value"):
            assert key in drift, f"Drift missing key: {key}"
        assert drift["sheet"] == "Summary"
        assert drift["address"] == "B2"
        assert drift["kind"] == "value_only"

    def test_rule_result_fields(self, tmp_path):
        from testsheet.reporters.json_reporter import write_json
        out = tmp_path / "report.json"
        write_json(_failing_report(), out)
        data = json.loads(out.read_text(encoding="utf-8"))
        rr = data["rule_results"][0]
        for key in ("rule_id", "rule_type", "passed", "message"):
            assert key in rr, f"Rule result missing key: {key}"


# ── JUnit XML reporter ────────────────────────────────────────────────────────

class TestJunitReporter:

    def test_creates_file(self, tmp_path):
        from testsheet.reporters.junit import write_junit
        out = tmp_path / "report.xml"
        write_junit(_passing_report(), out)
        assert out.exists()

    def test_valid_xml(self, tmp_path):
        from testsheet.reporters.junit import write_junit
        out = tmp_path / "report.xml"
        write_junit(_failing_report(), out)
        tree = ET.parse(str(out))   # raises if invalid
        assert tree.getroot() is not None

    def test_root_element_is_testsuites(self, tmp_path):
        from testsheet.reporters.junit import write_junit
        out = tmp_path / "report.xml"
        write_junit(_passing_report(), out)
        root = ET.parse(str(out)).getroot()
        assert root.tag == "testsuites"

    def test_passing_report_zero_failures(self, tmp_path):
        from testsheet.reporters.junit import write_junit
        out = tmp_path / "report.xml"
        write_junit(_passing_report(), out)
        root = ET.parse(str(out)).getroot()
        assert root.attrib["failures"] == "0"

    def test_failing_report_has_failures(self, tmp_path):
        from testsheet.reporters.junit import write_junit
        out = tmp_path / "report.xml"
        write_junit(_failing_report(), out)
        root = ET.parse(str(out)).getroot()
        assert int(root.attrib["failures"]) > 0

    def test_drift_testcase_present(self, tmp_path):
        from testsheet.reporters.junit import write_junit
        out = tmp_path / "report.xml"
        write_junit(_failing_report(), out)
        root = ET.parse(str(out)).getroot()
        cases = root.findall(".//testcase")
        names = [c.attrib.get("name", "") for c in cases]
        assert "cell_regression" in names

    def test_drift_failure_element_present(self, tmp_path):
        from testsheet.reporters.junit import write_junit
        out = tmp_path / "report.xml"
        write_junit(_failing_report(), out)
        root = ET.parse(str(out)).getroot()
        drift_case = next(
            c for c in root.findall(".//testcase")
            if c.attrib.get("name") == "cell_regression"
        )
        failure = drift_case.find("failure")
        assert failure is not None, "Drift testcase missing <failure> element"

    def test_passing_drift_no_failure_element(self, tmp_path):
        from testsheet.reporters.junit import write_junit
        out = tmp_path / "report.xml"
        write_junit(_passing_report(), out)
        root = ET.parse(str(out)).getroot()
        drift_case = next(
            c for c in root.findall(".//testcase")
            if c.attrib.get("name") == "cell_regression"
        )
        assert drift_case.find("failure") is None

    def test_rule_testcases_present(self, tmp_path):
        from testsheet.reporters.junit import write_junit
        out = tmp_path / "report.xml"
        write_junit(_failing_report(), out)
        root = ET.parse(str(out)).getroot()
        case_names = [c.attrib.get("name", "") for c in root.findall(".//testcase")]
        assert "rev_pos" in case_names
        assert "no_err" in case_names

    def test_failed_rule_has_failure_element(self, tmp_path):
        from testsheet.reporters.junit import write_junit
        out = tmp_path / "report.xml"
        write_junit(_failing_report(), out)
        root = ET.parse(str(out)).getroot()
        rev_case = next(
            c for c in root.findall(".//testcase")
            if c.attrib.get("name") == "rev_pos"
        )
        assert rev_case.find("failure") is not None

    def test_passed_rule_no_failure_element(self, tmp_path):
        from testsheet.reporters.junit import write_junit
        out = tmp_path / "report.xml"
        write_junit(_failing_report(), out)
        root = ET.parse(str(out)).getroot()
        no_err_case = next(
            c for c in root.findall(".//testcase")
            if c.attrib.get("name") == "no_err"
        )
        assert no_err_case.find("failure") is None


# ── HTML reporter ─────────────────────────────────────────────────────────────

class TestHtmlReporter:

    def test_creates_file(self, tmp_path):
        from testsheet.reporters.html_reporter import write_html
        out = tmp_path / "report.html"
        write_html(_passing_report(), out)
        assert out.exists()

    def test_standalone_html(self, tmp_path):
        """File must start with <!DOCTYPE html> — no external deps required."""
        from testsheet.reporters.html_reporter import write_html
        out = tmp_path / "report.html"
        write_html(_passing_report(), out)
        content = out.read_text(encoding="utf-8")
        assert "<!DOCTYPE html>" in content

    def test_pass_badge(self, tmp_path):
        from testsheet.reporters.html_reporter import write_html
        out = tmp_path / "report.html"
        write_html(_passing_report(), out)
        assert "NO CHANGES DETECTED" in out.read_text(encoding="utf-8")

    def test_fail_badge(self, tmp_path):
        from testsheet.reporters.html_reporter import write_html
        out = tmp_path / "report.html"
        write_html(_failing_report(), out)
        assert "CHANGES DETECTED" in out.read_text(encoding="utf-8")

    def test_drift_table_present(self, tmp_path):
        from testsheet.reporters.html_reporter import write_html
        out = tmp_path / "report.html"
        write_html(_failing_report(), out)
        content = out.read_text(encoding="utf-8")
        assert "Summary" in content          # sheet name
        assert "B2" in content               # address
        assert "value_only" in content       # kind
        assert "error_introduced" in content

    def test_rule_table_present(self, tmp_path):
        from testsheet.reporters.html_reporter import write_html
        out = tmp_path / "report.html"
        write_html(_failing_report(), out)
        content = out.read_text(encoding="utf-8")
        assert "rev_pos" in content
        assert "range_bound" in content

    def test_no_drift_message(self, tmp_path):
        from testsheet.reporters.html_reporter import write_html
        out = tmp_path / "report.html"
        write_html(_passing_report(), out)
        assert "No drift detected" in out.read_text(encoding="utf-8")

    def test_workbook_name_in_title(self, tmp_path):
        from testsheet.reporters.html_reporter import write_html
        out = tmp_path / "report.html"
        write_html(_passing_report("q4_forecast.xlsx"), out)
        assert "q4_forecast.xlsx" in out.read_text(encoding="utf-8")
