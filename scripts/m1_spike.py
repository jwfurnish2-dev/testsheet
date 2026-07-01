"""
M1 Evaluator Spike — accuracy & timing benchmark.

Runs both evaluators (pycel, libreoffice) against 3 benchmark models and
compares computed values to known-good expected values.  Produces a
human-readable report and writes docs/m1-findings.md.

Usage (from repo root):
  python scripts/m1_spike.py                        # pycel only
  python scripts/m1_spike.py --evaluators all       # pycel + libreoffice
  python scripts/m1_spike.py --evaluators libreoffice

Output:
  Console summary table
  docs/m1-findings.md  (append-safe benchmark report)
"""

from __future__ import annotations

import argparse
import math
import sys
import time
from pathlib import Path

# Make sure src/ is on the path when run from repo root
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

FIXTURES_DIR = Path(__file__).parent.parent / "tests" / "fixtures"
DOCS_DIR = Path(__file__).parent.parent / "docs"


# ── Helpers ──────────────────────────────────────────────────────────────────

def _values_close(a, b, tol: float = 1e-6) -> bool:
    """Return True if a and b are equal (with float tolerance)."""
    if a is None and b is None:
        return True
    if type(a) != type(b):
        try:
            return math.isclose(float(a), float(b), rel_tol=tol, abs_tol=tol)
        except (TypeError, ValueError):
            return str(a) == str(b)
    if isinstance(a, float):
        return math.isclose(a, b, rel_tol=tol, abs_tol=tol)
    return a == b


def _check_accuracy(
    computed: dict[str, dict],
    expected: dict[str, dict],
) -> tuple[int, int, list[str]]:
    """
    Compare computed vs expected.
    Returns (correct_count, total_count, list_of_mismatches).
    """
    correct = 0
    total = 0
    mismatches = []

    for sheet, cells in expected.items():
        for addr, exp_val in cells.items():
            total += 1
            got = (computed.get(sheet) or {}).get(addr)
            if _values_close(got, exp_val):
                correct += 1
            else:
                mismatches.append(
                    f"  {sheet}!{addr}: expected={exp_val!r}  got={got!r}"
                )

    return correct, total, mismatches


# ── Benchmark runner ──────────────────────────────────────────────────────────

def run_benchmark(evaluator_names: list[str]) -> list[dict]:
    """Run all models through each evaluator. Return list of result dicts."""
    # Import fixtures builder
    sys.path.insert(0, str(FIXTURES_DIR))
    from make_m1_fixtures import build_all, ALL_MODELS

    print("\nBuilding benchmark fixtures...")
    all_expected = build_all(FIXTURES_DIR)

    from testsheet.evaluator import get_evaluator

    results = []

    for ev_name in evaluator_names:
        try:
            ev = get_evaluator(ev_name)
        except ValueError as e:
            print(f"  [skip] {e}")
            continue

        if not ev.is_available():
            print(f"\n[{ev_name}] Not available — skipping.")
            results.append({
                "evaluator": ev_name,
                "available": False,
                "models": [],
            })
            continue

        print(f"\n{'='*60}")
        print(f"Evaluator: {ev_name}")
        print(f"{'='*60}")

        ev_results = {"evaluator": ev_name, "available": True, "models": []}

        for filename, _, label in ALL_MODELS:
            path = FIXTURES_DIR / filename
            expected = all_expected[filename]

            print(f"\n  Model: {label}")

            t0 = time.perf_counter()
            try:
                computed = ev.compute(path)
                elapsed = time.perf_counter() - t0
                error = None
            except Exception as exc:
                elapsed = time.perf_counter() - t0
                computed = {}
                error = str(exc)

            correct, total, mismatches = _check_accuracy(computed, expected)
            accuracy = correct / total if total else 0.0

            status = "✓" if not mismatches and not error else "✗"
            print(f"  {status} Accuracy: {correct}/{total} ({accuracy:.0%})  Time: {elapsed*1000:.1f}ms")
            if error:
                print(f"    ERROR: {error}")
            for m in mismatches:
                print(m)

            ev_results["models"].append({
                "filename": filename,
                "label": label,
                "correct": correct,
                "total": total,
                "accuracy": accuracy,
                "elapsed_ms": round(elapsed * 1000, 1),
                "mismatches": mismatches,
                "error": error,
            })

        results.append(ev_results)

    return results


# ── Report writer ─────────────────────────────────────────────────────────────

def write_findings(results: list[dict]) -> Path:
    """Append findings to docs/m1-findings.md."""
    DOCS_DIR.mkdir(exist_ok=True)
    out = DOCS_DIR / "m1-findings.md"

    from datetime import datetime
    ts = datetime.now().strftime("%Y-%m-%d %H:%M")

    lines = [
        f"\n## M1 Spike Run — {ts}\n",
        "| Evaluator | Model | Accuracy | Time (ms) | Notes |",
        "|---|---|---|---|---|",
    ]

    recommendation = None

    for ev_result in results:
        ev_name = ev_result["evaluator"]
        if not ev_result["available"]:
            lines.append(f"| {ev_name} | — | N/A | — | Not installed |")
            continue

        total_correct = sum(m["correct"] for m in ev_result["models"])
        total_cells   = sum(m["total"]   for m in ev_result["models"])
        overall_acc   = total_correct / total_cells if total_cells else 0

        for m in ev_result["models"]:
            notes = "OK" if not m["mismatches"] and not m["error"] else (
                m["error"] or f"{len(m['mismatches'])} mismatch(es)"
            )
            lines.append(
                f"| {ev_name} | {m['label']} | "
                f"{m['correct']}/{m['total']} ({m['accuracy']:.0%}) | "
                f"{m['elapsed_ms']} | {notes} |"
            )

        lines.append(
            f"| **{ev_name} TOTAL** | — | "
            f"**{total_correct}/{total_cells} ({overall_acc:.0%})** | — | — |"
        )

        if recommendation is None and overall_acc >= 0.9:
            recommendation = ev_name

    lines.append("")
    if recommendation:
        lines.append(
            f"**Recommendation:** Use `{recommendation}` as the default evaluator for M2+."
        )
    else:
        lines.append(
            "**Recommendation:** Neither evaluator met 90% accuracy threshold — "
            "investigate formula gaps before proceeding to M2."
        )

    content = "\n".join(lines) + "\n"
    with out.open("a", encoding="utf-8") as f:
        f.write(content)

    return out


# ── CLI ───────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(description="TestSheet M1 evaluator spike")
    parser.add_argument(
        "--evaluators",
        default="pycel",
        help="Comma-separated list of evaluators to test, or 'all'. "
             "Default: pycel",
    )
    args = parser.parse_args()

    if args.evaluators == "all":
        ev_names = ["formulas", "pycel", "libreoffice"]
    else:
        ev_names = [e.strip() for e in args.evaluators.split(",")]

    results = run_benchmark(ev_names)
    report_path = write_findings(results)
    print(f"\nFindings written to: {report_path}")


if __name__ == "__main__":
    main()
