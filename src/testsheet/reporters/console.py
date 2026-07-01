"""Rich console reporter — pretty-prints a run report to the terminal."""

from __future__ import annotations

from rich.console import Console
from rich.table import Table
from rich import box

console = Console()

_KIND_STYLE = {
    "value_only": "yellow",
    "formula_only": "blue",
    "both": "red bold",
    "new": "green",
    "deleted": "dim red",
    "error_introduced": "red bold",
}


def print_report(report: dict) -> None:
    """Print a human-readable summary of a run report to stdout."""
    passed = report.get("passed", False)
    drifts = report.get("drifts", [])
    rule_results = report.get("rule_results", [])

    if passed:
        status = "[green]NO CHANGES DETECTED[/green]"
    else:
        status = "[yellow bold]CHANGES DETECTED — PENDING REVIEW[/yellow bold]"
    console.rule(f"TestSheet — {status}")

    # ── Drift table ────────────────────────────────────────────────────────
    if drifts:
        table = Table(
            "Sheet", "Address", "Kind", "Baseline", "Current",
            title=f"[bold]{len(drifts)} drift(s) detected[/bold]",
            box=box.SIMPLE_HEAVY,
            show_lines=False,
        )
        for d in sorted(drifts, key=lambda x: (x.sheet, x.address)):
            style = _KIND_STYLE.get(d.kind, "")
            table.add_row(
                d.sheet, d.address,
                f"[{style}]{d.kind}[/{style}]",
                _fmt(d.baseline_value),
                _fmt(d.current_value),
            )
        console.print(table)
    else:
        console.print("[green]✓[/green] No drift detected.")

    # ── Rule results ───────────────────────────────────────────────────────
    if rule_results:
        rtable = Table(
            "Rule ID", "Type", "Result", "Message",
            title="[bold]Invariant rules[/bold]",
            box=box.SIMPLE_HEAVY,
        )
        for r in rule_results:
            icon = "[green]✓ OK[/green]" if r.passed else "[yellow]⚠ VIOLATION[/yellow]"
            rtable.add_row(r.rule_id, r.rule_type, icon, r.message)
        console.print(rtable)

    console.rule()


def _fmt(value) -> str:
    if value is None:
        return "[dim]—[/dim]"
    s = str(value)
    return s[:60] + "…" if len(s) > 60 else s
