"""Standalone HTML reporter."""

from __future__ import annotations

from pathlib import Path


_TEMPLATE = """\
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>TestSheet Report — {workbook}</title>
<style>
  body {{ font-family: system-ui, sans-serif; max-width: 960px; margin: 2rem auto; color: #1a1a1a; }}
  h1 {{ font-size: 1.4rem; }}
  .badge {{ display: inline-block; padding: .2rem .6rem; border-radius: 4px; font-weight: bold; }}
  .pass {{ background: #d4edda; color: #155724; }}
  .fail {{ background: #fff3cd; color: #856404; }}
  .approved {{ background: #cce5ff; color: #004085; }}
  table {{ border-collapse: collapse; width: 100%; margin-bottom: 2rem; }}
  th, td {{ border: 1px solid #dee2e6; padding: .45rem .7rem; text-align: left; font-size: .9rem; }}
  th {{ background: #f8f9fa; }}
  .kind-error_introduced, .kind-both {{ color: #c0392b; font-weight: bold; }}
  .kind-value_only {{ color: #e67e22; }}
  .kind-formula_only {{ color: #2980b9; }}
  .kind-new {{ color: #27ae60; }}
  .kind-deleted {{ color: #7f8c8d; }}
  code {{ font-family: monospace; background: #f4f4f4; padding: 0 .3rem; border-radius: 3px; }}
  .warn-banner {{ background: #fff3cd; border: 1px solid #ffc107; border-radius: 4px;
                  padding: .6rem 1rem; margin-bottom: 1.2rem; color: #856404; }}
  .approve-banner {{ background: #e8f4fd; border: 1px solid #bee5eb; border-radius: 4px;
                     padding: .6rem 1rem; margin-bottom: 1.2rem; color: #0c5460; }}
</style>
</head>
<body>
<h1>TestSheet — <code>{workbook}</code>
  <span class="badge {status_class}">{status}</span>
</h1>

{structural_banner}
{approve_banner}

<h2>Cell Changes ({drift_count})</h2>
{drift_table}

<h2>Invariant Rules ({rule_count})</h2>
{rule_table}

</body>
</html>
"""


def write_html(report: dict, output_path: Path) -> None:
    """Write a standalone HTML report to *output_path*."""
    drifts = report.get("drifts", [])
    rule_results = report.get("rule_results", [])
    passed = report.get("passed", False)

    sc = report.get("structural_change")
    if passed:
        status, status_class = "NO CHANGES DETECTED", "pass"
    else:
        status, status_class = "CHANGES DETECTED — PENDING REVIEW", "fail"
    html = _TEMPLATE.format(
        workbook=report.get("workbook", "workbook"),
        status=status,
        status_class=status_class,
        drift_count=len(drifts),
        rule_count=len(rule_results),
        drift_table=_drift_table(drifts),
        rule_table=_rule_table(rule_results),
        structural_banner=_structural_banner(sc),
        approve_banner=_approve_banner(passed, report.get("workbook", "")),
    )
    Path(output_path).write_text(html, encoding="utf-8")


def _approve_banner(passed: bool, workbook: str) -> str:
    if passed:
        return ""
    wb_name = workbook.split("\\")[-1].split("/")[-1]
    return (
        f"<div class='approve-banner'>If these changes are intentional, run "
        f"<code>testsheet approve &quot;{_esc(wb_name)}&quot;</code> "
        f"to update the baseline and mark this version as approved.</div>"
    )


def _structural_banner(sc) -> str:
    if sc is None:
        return ""
    msg = sc.message if hasattr(sc, "message") else sc.get("message", "")
    return f"<div class='warn-banner'>⚠ <strong>Structural change detected</strong> — {_esc(msg)}</div>"


def _drift_table(drifts: list) -> str:
    if not drifts:
        return "<p>No drift detected.</p>"
    rows = []
    # Normalise to dicts first so sort key and field access are uniform
    drift_dicts = [d.as_dict() if hasattr(d, "as_dict") else d for d in drifts]
    for d in sorted(drift_dicts, key=lambda x: (x.get("sheet", ""), x.get("address", ""))):
        rows.append(
            f"<tr>"
            f"<td>{_esc(d['sheet'])}</td>"
            f"<td><code>{_esc(d['address'])}</code></td>"
            f"<td class='kind-{d['kind']}'>{d['kind']}</td>"
            f"<td><code>{_esc(str(d['baseline_value']))}</code></td>"
            f"<td><code>{_esc(str(d['current_value']))}</code></td>"
            f"</tr>"
        )
    header = "<tr><th>Sheet</th><th>Address</th><th>Kind</th><th>Baseline</th><th>Current</th></tr>"
    return f"<table>{header}{''.join(rows)}</table>"


def _rule_table(rule_results: list) -> str:
    if not rule_results:
        return "<p>No rules defined.</p>"
    rows = []
    for r in rule_results:
        icon = "✓" if r.passed else "✗"
        cls = "pass" if r.passed else "fail"
        rows.append(
            f"<tr>"
            f"<td><code>{_esc(r.rule_id)}</code></td>"
            f"<td>{r.rule_type}</td>"
            f"<td><span class='badge {cls}'>{icon}</span></td>"
            f"<td>{_esc(r.message)}</td>"
            f"</tr>"
        )
    header = "<tr><th>ID</th><th>Type</th><th>Result</th><th>Message</th></tr>"
    return f"<table>{header}{''.join(rows)}</table>"


def _esc(s: str) -> str:
    return (s.replace("&", "&amp;")
             .replace("<", "&lt;")
             .replace(">", "&gt;")
             .replace('"', "&quot;"))
