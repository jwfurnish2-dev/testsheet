"""JUnit XML reporter — for CI integration (GitHub Actions, Jenkins, etc.)."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from pathlib import Path


def write_junit(report: dict, output_path: Path) -> None:
    """Write a JUnit XML file to *output_path*."""
    drifts = report.get("drifts", [])
    rule_results = report.get("rule_results", [])

    # Total test cases = 1 drift suite + 1 per rule
    failures = len(drifts) + sum(1 for r in rule_results if not r.passed)
    total = 1 + len(rule_results)  # drift suite counts as 1

    root = ET.Element("testsuites", name="testsheet", failures=str(failures), tests=str(total))
    suite = ET.SubElement(root, "testsuite", name=report.get("workbook", "workbook"),
                          tests=str(total), failures=str(failures))

    # ── Drift test case ────────────────────────────────────────────────────
    drift_case = ET.SubElement(suite, "testcase",
                               classname="testsheet.drift",
                               name="cell_regression")
    if drifts:
        failure_el = ET.SubElement(drift_case, "failure",
                                   message=f"{len(drifts)} drift(s) detected",
                                   type="DriftError")
        lines = []
        for d in drifts:
            if hasattr(d, "as_dict"):
                d = d.as_dict()
            lines.append(
                f"{d['sheet']}!{d['address']} [{d['kind']}]  "
                f"{d['baseline_value']!r} → {d['current_value']!r}"
            )
        failure_el.text = "\n".join(lines)

    # ── Rule test cases ────────────────────────────────────────────────────
    for r in rule_results:
        rc = ET.SubElement(suite, "testcase",
                           classname=f"testsheet.rules.{r.rule_type}",
                           name=r.rule_id)
        if not r.passed:
            f_el = ET.SubElement(rc, "failure",
                                 message=r.message,
                                 type="RuleViolation")
            f_el.text = r.message

    tree = ET.ElementTree(root)
    ET.indent(tree, space="  ")
    tree.write(str(output_path), encoding="unicode", xml_declaration=True)
