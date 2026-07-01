"""JSON reporter — machine-readable output."""

from __future__ import annotations

import json
from pathlib import Path


def write_json(report: dict, output_path: Path) -> None:
    """Serialise *report* to *output_path* as JSON."""
    sc = report.get("structural_change")
    payload = {
        "workbook": report.get("workbook", ""),
        "passed": report.get("passed", False),
        "drift_count": len(report.get("drifts", [])),
        "rule_failure_count": sum(
            1 for r in report.get("rule_results", []) if not _rule_passed(r)
        ),
        "drifts": [_drift_dict(d) for d in report.get("drifts", [])],
        "rule_results": [_rule_dict(r) for r in report.get("rule_results", [])],
        "structural_change": sc.as_dict() if sc is not None and hasattr(sc, "as_dict") else sc,
    }
    Path(output_path).write_text(
        json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8"
    )


def _drift_dict(drift) -> dict:
    if hasattr(drift, "as_dict"):
        return drift.as_dict()
    return drift  # already a dict (e.g. in tests)


def _rule_dict(r) -> dict:
    if hasattr(r, "as_dict"):
        return r.as_dict()
    return r


def _rule_passed(r) -> bool:
    if hasattr(r, "passed"):
        return r.passed
    return r.get("passed", False)
