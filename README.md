# TestSheet — pytest for Excel

> Capture a golden-master baseline from your Excel model, then automatically flag any cell whose value or formula drifts in future versions.

TestSheet is a command-line regression-testing tool for spreadsheet models. It works the same way pytest works for code: you record what *should* be true, then run the suite on every change to verify nothing broke.

---

## Features

- **Golden-master baseline** — snapshot every cell's value and formula in one command
- **Drift detection** — six drift kinds: `value_only`, `formula_only`, `both`, `new`, `deleted`, `error_introduced`
- **Invariant rules** — six built-in rule types you describe in a YAML file (range bounds, error checks, totals tie-outs, monotonic sequences, cross-cell relationships, no-hardcode guards)
- **Multiple reporters** — rich console table, JSON, JUnit XML (for CI), standalone HTML
- **Structural-change heuristic** — warns when too many cells drift at once (likely a layout change, not real model drift)
- **Configurable float tolerance** — set `rel_tol` / `abs_tol` in a config file
- **GitHub Actions integration** — single composite action, reports uploaded as artifacts

---

## Installation

```bash
pip install testsheet
```

**Dependencies**: Python 3.10+, `openpyxl`, `typer`, `rich`, `pyyaml`, `jinja2`, `formulas`

---

## 5-minute quickstart

### 1. Capture a baseline

Navigate to the folder containing your workbook and run:

```bash
testsheet baseline models/q4_forecast.xlsx
```

This creates `.testsheet/baseline.json` next to your workbook. Commit this file to version control.

```
models/
  q4_forecast.xlsx
  .testsheet/
    baseline.json        ← commit this
```

### 2. Run the checks

After any edit to the workbook:

```bash
testsheet run models/q4_forecast.xlsx
```

TestSheet prints a table of drifted cells and exits `0` (pass) or `1` (fail).

```
TestSheet — running checks on q4_forecast.xlsx
┌──────────┬─────────┬───────────────────┬───────────┬───────────┐
│ Sheet    │ Address │ Kind              │ Baseline  │ Current   │
├──────────┼─────────┼───────────────────┼───────────┼───────────┤
│ Summary  │ B2      │ value_only        │ 460.0     │ 999.0     │
│ Model    │ D2      │ error_introduced  │ 120       │ #DIV/0!   │
└──────────┴─────────┴───────────────────┴───────────┴───────────┘
FAIL — 2 cells drifted
```

### 3. Add invariant rules (optional)

Create `.testsheet/rules.yaml`:

```yaml
rules:
  - id: revenues_positive
    type: range_bound
    range: "Model!A1:A12"
    min: 0

  - id: no_formula_errors
    type: no_error
    range: "Model!A1:Z200"

  - id: total_ties_sum
    type: totals_tie
    total_cell: "Summary!B10"
    sum_range:  "Summary!B1:B9"
    tolerance: 0.01

  - id: margins_increasing
    type: monotonic
    range: "Model!C1:C4"
    direction: increasing

  - id: no_hardcoded_drivers
    type: no_hardcode_in_range
    range: "Assumptions!B1:B20"

  - id: gross_profit_nonneg
    type: relationship
    expression: "Summary!B5 >= 0"
```

Rules are evaluated on every `testsheet run`. A rule failure counts the same as cell drift for the exit code.

---

## CLI reference

### `testsheet baseline <workbook>`

Captures a golden-master baseline.

| Option | Default | Description |
|---|---|---|
| `--sheets` | all | Comma-separated sheet names to include |
| `--evaluator` | `formulas` | `formulas` or `libreoffice` |

### `testsheet run <workbook>`

Runs regression checks against the baseline.

| Option | Default | Description |
|---|---|---|
| `--evaluator` | `formulas` | `formulas` or `libreoffice` |
| `--junit PATH` | — | Write JUnit XML report |
| `--json PATH` | — | Write JSON report |
| `--html PATH` | — | Write standalone HTML report |
| `--no-fail-on-drift` | — | Exit `0` even when drift is found |

---

## Rules reference

| Type | Required keys | Description |
|---|---|---|
| `range_bound` | `range` or `cell`, optionally `min`, `max` | Every cell in range must be within [min, max] |
| `no_error` | `range` | No Excel error string (`#REF!`, `#DIV/0!`, etc.) in range |
| `no_hardcode_in_range` | `range` | Every cell in range must contain a formula, not a literal |
| `totals_tie` | `total_cell`, `sum_range`, `tolerance` | `total_cell ≈ SUM(sum_range)` within tolerance |
| `monotonic` | `range`, `direction` (`increasing`\|`decreasing`), optionally `strict` | Values in range must be monotonically ordered |
| `relationship` | `expression` | Arbitrary cross-cell expression evaluates to `True` (e.g. `"Summary!B5 >= Summary!B4 * 0.9"`) |

---

## Configuration

Create `.testsheet/config.yaml` to override defaults:

```yaml
diff:
  rel_tol: 1.0e-6   # relative float tolerance (default 1e-9)
  abs_tol: 1.0e-9   # absolute float tolerance (default 1e-12)

structural_change_threshold: 0.5  # warn when this fraction of cells drift (default 0.5)
```

**Float tolerance** — useful for models with rounding differences between Excel versions. Set `rel_tol: 0.01` to ignore differences smaller than 1%.

**Structural change threshold** — when the fraction of drifted cells exceeds this value, TestSheet emits a warning that the workbook layout may have changed (rows/columns inserted, sheet renamed), rather than flagging hundreds of individual cell drifts.

---

## GitHub Actions

Add to your workflow:

```yaml
- name: Run TestSheet
  uses: ./action.yml         # or the published action path
  with:
    workbook: models/q4_forecast.xlsx
    junit-output: testsheet-report.xml
    html-output:  testsheet-report.html

- name: Publish test results
  uses: dorny/test-reporter@v1
  if: always()
  with:
    name: TestSheet
    path: testsheet-report.xml
    reporter: java-junit
```

Full example workflow (`.github/workflows/testsheet.yml`):

```yaml
name: Excel regression gate

on:
  pull_request:
    paths:
      - "models/**"

jobs:
  testsheet:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4

      - name: Run TestSheet
        uses: ./
        with:
          workbook: models/q4_forecast.xlsx

      - name: Publish JUnit results
        uses: dorny/test-reporter@v1
        if: always()
        with:
          name: TestSheet
          path: testsheet-report.xml
          reporter: java-junit
```

---

## How it works

1. **`baseline`** — loads the workbook twice (once with `data_only=False` for formulas, once with `data_only=True` for cached values), then runs the `formulas` pure-Python evaluator to recompute values from scratch. Writes `.testsheet/baseline.json`.

2. **`run`** — repeats the same parse + evaluate step on the current workbook, then diffs cell-by-cell against the baseline. Classifies each changed cell into one of six drift kinds. Evaluates any `rules.yaml` invariants. Exits non-zero if any drift or rule failure is found.

3. **Evaluator** — TestSheet uses the [`formulas`](https://github.com/vinci1it2000/formulas) pure-Python Excel evaluator (no LibreOffice required). A LibreOffice headless evaluator is also available via `--evaluator libreoffice` for models that use functions not yet supported by `formulas`.

---

## Project layout

```
testsheet/
├── src/testsheet/
│   ├── cli.py              # Typer CLI (baseline + run)
│   ├── parser.py           # openpyxl workbook parser
│   ├── baseline.py         # golden-master capture + load
│   ├── diff.py             # drift detection + structural-change heuristic
│   ├── config.py           # .testsheet/config.yaml loader
│   ├── evaluator/
│   │   ├── formulas_eval.py   # formulas library evaluator (default)
│   │   ├── pycel_eval.py      # pycel evaluator (legacy / Python ≤3.11)
│   │   └── libreoffice.py     # LibreOffice headless evaluator
│   ├── rules/
│   │   └── engine.py       # rules.yaml loader + 6 rule handlers
│   └── reporters/
│       ├── console.py      # rich terminal table
│       ├── json_reporter.py
│       ├── junit.py        # JUnit XML
│       └── html_reporter.py   # standalone HTML
├── tests/
│   ├── test_smoke.py
│   ├── test_evaluators.py
│   ├── test_diff_m2.py
│   ├── test_rules_m3.py
│   ├── test_rules_e2e.py
│   ├── test_reporters_m4.py
│   ├── test_ci_m4.py
│   └── test_hardening_m5.py
├── action.yml              # GitHub Actions composite action
└── pyproject.toml
```

---

## License

MIT
