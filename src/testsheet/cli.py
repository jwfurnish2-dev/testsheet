"""TestSheet CLI — entry point for `testsheet` command."""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Annotated, Optional

import typer
from rich.console import Console

from testsheet import __version__
from testsheet.baseline import capture_baseline, load_baseline
from testsheet.config import load_config
from testsheet.diff import diff_workbook, detect_structural_change
from testsheet.evaluator import get_evaluator
from testsheet.reporters.console import print_report
from testsheet.reporters.json_reporter import write_json
from testsheet.reporters.junit import write_junit
from testsheet.reporters.html_reporter import write_html
from testsheet.rules.engine import load_rules, evaluate_rules

app = typer.Typer(
    name="testsheet",
    help="pytest for Excel — regression testing for spreadsheet models.",
    add_completion=False,
)
console = Console()


def version_callback(value: bool) -> None:
    if value:
        typer.echo(f"testsheet {__version__}")
        raise typer.Exit()


@app.callback()
def main(
    version: Annotated[
        Optional[bool],
        typer.Option("--version", "-V", callback=version_callback, is_eager=True),
    ] = None,
) -> None:
    """TestSheet — pytest for Excel."""


# ---------------------------------------------------------------------------
# baseline
# ---------------------------------------------------------------------------

@app.command()
def baseline(
    workbook: Annotated[Path, typer.Argument(help="Path to the .xlsx workbook.")],
    sheets: Annotated[
        Optional[str],
        typer.Option("--sheets", "-s", help="Comma-separated sheet names to include (default: all)."),
    ] = None,
    evaluator: Annotated[
        str,
        typer.Option("--evaluator", "-e", help="Value evaluator: 'formulas' (default), 'pycel', or 'libreoffice'."),
    ] = "formulas",
) -> None:
    """Capture a golden-master baseline from WORKBOOK."""
    workbook = workbook.resolve()
    if not workbook.exists():
        console.print(f"[red]Error:[/red] File not found: {workbook}")
        raise typer.Exit(1)

    sheet_list = [s.strip() for s in sheets.split(",")] if sheets else None

    console.print(f"[bold]TestSheet[/bold] — capturing baseline for [cyan]{workbook.name}[/cyan]")

    ev = get_evaluator(evaluator)
    capture_baseline(workbook, evaluator=ev, sheets=sheet_list)

    baseline_path = workbook.parent / ".testsheet" / "baseline.json"
    console.print(f"[green]✓[/green] Baseline written to {baseline_path}")


# ---------------------------------------------------------------------------
# approve
# ---------------------------------------------------------------------------

@app.command()
def approve(
    workbook: Annotated[Path, typer.Argument(help="Path to the .xlsx workbook to approve.")],
    evaluator: Annotated[
        str,
        typer.Option("--evaluator", "-e", help="Value evaluator: 'formulas' (default) or 'libreoffice'."),
    ] = "formulas",
) -> None:
    """Approve WORKBOOK as the new baseline — marks this version as the golden master.

    Run this after reviewing a 'CHANGES DETECTED' report and confirming the
    changes are intentional.  The current workbook state replaces the previous
    baseline so future runs compare against it.
    """
    workbook = workbook.resolve()
    if not workbook.exists():
        console.print(f"[red]Error:[/red] File not found: {workbook}")
        raise typer.Exit(1)

    baseline_path = workbook.parent / ".testsheet" / "baseline.json"
    had_baseline = baseline_path.exists()

    ev = get_evaluator(evaluator)
    capture_baseline(workbook, evaluator=ev)

    if had_baseline:
        console.print(
            f"[green]✓ Approved:[/green] [cyan]{workbook.name}[/cyan] is now the new baseline.\n"
            "  Future runs will compare against this version."
        )
    else:
        console.print(
            f"[green]✓ Baseline created:[/green] [cyan]{workbook.name}[/cyan]\n"
            "  Run [bold]testsheet run[/bold] to check future versions against it."
        )


# ---------------------------------------------------------------------------
# run
# ---------------------------------------------------------------------------

@app.command()
def run(
    workbook: Annotated[Path, typer.Argument(help="Path to the .xlsx workbook.")],
    evaluator: Annotated[
        str,
        typer.Option("--evaluator", "-e", help="Value evaluator: 'formulas' (default), 'pycel', or 'libreoffice'."),
    ] = "formulas",
    junit: Annotated[
        Optional[Path],
        typer.Option("--junit", help="Write JUnit XML report to this path."),
    ] = None,
    json_out: Annotated[
        Optional[Path],
        typer.Option("--json", help="Write JSON report to this path."),
    ] = None,
    html_out: Annotated[
        Optional[Path],
        typer.Option("--html", help="Write standalone HTML report to this path."),
    ] = None,
    fail_on_drift: Annotated[
        bool,
        typer.Option("--fail-on-drift/--no-fail-on-drift", help="Exit non-zero if any drift or rule failure found."),
    ] = True,
) -> None:
    """Run regression checks on WORKBOOK against its baseline."""
    workbook = workbook.resolve()
    if not workbook.exists():
        console.print(f"[red]Error:[/red] File not found: {workbook}")
        raise typer.Exit(1)

    testsheet_dir = workbook.parent / ".testsheet"
    baseline_path = testsheet_dir / "baseline.json"
    rules_path = testsheet_dir / "rules.yaml"

    if not baseline_path.exists():
        console.print(
            f"[red]Error:[/red] No baseline found at {baseline_path}. "
            "Run `testsheet baseline <workbook>` first."
        )
        raise typer.Exit(1)

    console.print(f"[bold]TestSheet[/bold] — running checks on [cyan]{workbook.name}[/cyan]")

    cfg = load_config(testsheet_dir)
    ev = get_evaluator(evaluator)
    baseline_data = load_baseline(baseline_path)

    drifts = diff_workbook(
        workbook, baseline_data, evaluator=ev,
        rel_tol=cfg["diff"]["rel_tol"],
        abs_tol=cfg["diff"]["abs_tol"],
    )

    structural_warning = detect_structural_change(
        drifts, baseline_data,
        threshold=cfg["structural_change_threshold"],
    )
    if structural_warning:
        console.print(f"[yellow]⚠ Warning:[/yellow] {structural_warning.message}")

    rules = load_rules(rules_path) if rules_path.exists() else []
    rule_results = evaluate_rules(workbook, rules, evaluator=ev)

    passed = not drifts and all(r.passed for r in rule_results)
    report = {
        "workbook": str(workbook),
        "passed": passed,
        "drifts": drifts,
        "rule_results": rule_results,
        "structural_change": structural_warning,
    }

    print_report(report)

    if json_out:
        write_json(report, json_out)
        console.print(f"JSON report → {json_out}")

    if junit:
        write_junit(report, junit)
        console.print(f"JUnit XML  → {junit}")

    if html_out:
        write_html(report, html_out)
        console.print(f"HTML report → {html_out}")

    if fail_on_drift and not passed:
        raise typer.Exit(1)
