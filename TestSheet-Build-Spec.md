# TestSheet — Build Spec (Item #2: Automated Testing & CI for Excel Models)

> **Purpose of this document.** This is a self-contained build brief intended to be handed to a fresh Claude / Claude Code session to start implementation. It assumes no prior conversation context. Read it top-to-bottom, then begin with the Milestone 0 scaffold.

---

## 1. Product in one sentence

**TestSheet** is "pytest for Excel": it captures a workbook's expected behavior as a **golden master**, then on every new version flags any cell whose **value or formula drifted**, enforces plain-language **invariant rules**, and runs as a **CI gate** so a broken model can't ship.

## 2. Why this exists (problem)

- 88–94% of business spreadsheets contain errors; ~50% of large-company models have material defects.
- Real-world losses: JPMorgan "London Whale" (~$6.2B), Fannie Mae ($1.1B equity misstatement), Citigroup ($900M mis-wire).
- SOX makes financial-reporting spreadsheets an explicit control/audit liability.
- **No standalone product tests spreadsheet logic.** Existing options are DIY Python (xlwings) or unrelated QA-tracking templates. This is the gap.

## 3. Target user (beachhead)

Primary: **model-risk / SOX & controls** teams at banks, insurers, public companies.
Secondary: **quant / FP&A** analysts who own models others depend on.

Design implication: the primary buyer cares about **auditability and a pass/fail gate**; the daily user is a **non-engineer analyst** who will not write code. Both must be served.

## 4. Core concepts (domain model)

| Concept | Definition |
|---|---|
| **Workbook** | The `.xlsx` under test. |
| **Baseline / Golden master** | A captured snapshot of the workbook's computed cell values + formula strings + named ranges at a known-good moment. |
| **Run** | An evaluation of a current workbook against its baseline + rule set, producing a Report. |
| **Drift** | A cell whose value and/or formula differs from baseline (classified: value-only, formula-only, both, new, deleted, error introduced). |
| **Invariant rule** | A user-defined assertion evaluated against the current workbook (range bound, no-hardcode, no-error, totals-tie, monotonic, cross-cell relationship). |
| **Report** | Structured result (pass/fail, list of drifts, rule outcomes) + human-readable + machine-readable (JSON/JUnit XML). |

## 5. Architecture (MVP)

Ship a **CLI + library core first** (portable, scriptable, CI-native). Add-in/UI comes later.

```
+------------------+      +--------------------+      +-------------------+
|  Excel workbook  | ---> |  Parser / Evaluator| ---> |  Diff + Rule Engine|
|  (.xlsx)         |      |  (values+formulas) |      |  -> Report          |
+------------------+      +--------------------+      +-------------------+
                                                              |
                                   +--------------------------+------------+
                                   |            Reporters                  |
                                   |  console | JSON | JUnit XML | HTML    |
                                   +---------------------------------------+
                                                  |
                                          CI integration (GitHub Action)
```

### Tech stack (recommended)
- **Language:** Python 3.11+ (richest xlsx ecosystem; easy CI packaging).
- **Parsing:** `openpyxl` (formulas + values; read `data_only=False` for formulas and a calculated copy for values). For reliable computed values, recalculate via **LibreOffice headless** (`soffice --headless --convert-to xlsx --calc`) or **formulas**/`pycel` for pure-Python evaluation. Decide per Milestone 1 spike.
- **CLI:** `typer` or `click`.
- **Config / baseline storage:** a `.testsheet/` directory next to the workbook — `baseline.json` (cell map + hashes) and `rules.yaml`.
- **Reports:** JSON + JUnit XML (for CI) + a small standalone HTML.
- **Packaging:** `pipx`-installable; plus a prebuilt **GitHub Action** wrapper.

### Key technical challenges (call these out, don't hand-wave)
1. **Getting computed values reliably.** `openpyxl` does not recalculate; cached values may be stale. Spike both LibreOffice-recalc and `pycel`/`formulas`; pick by accuracy vs. speed.
2. **Stable cell identity across versions.** Inserted/deleted rows shift `A1` addresses. MVP: compare by address but detect row/column insertions via a heuristic (named ranges, header matching) and flag "structural change" rather than thousands of false drifts. Document this as a known limitation in v1.
3. **Baseline size.** Large models = many cells. Store hashes + only materialize diffs. Allow scoping a baseline to specific sheets/ranges.
4. **Float comparison.** Use tolerance (relative + absolute epsilon), configurable per rule.

## 6. MVP feature set (with acceptance criteria)

### F1 — Auto-baseline
`testsheet baseline model.xlsx` captures values + formulas to `.testsheet/baseline.json`.
- **AC:** Re-running against an unchanged file yields **PASS, zero drift**.

### F2 — Regression diff
`testsheet run model.xlsx` compares current vs. baseline.
- **AC:** Changing one formula's result is reported as exactly one drift, classified correctly (value/formula/both), with sheet + address + before/after.
- **AC:** Introducing a `#REF!`/`#DIV/0!` is flagged as `error_introduced`.

### F3 — Invariant rules (`rules.yaml`)
Support at least: `range_bound` (min/max on a cell/range), `no_error`, `no_hardcode_in_range` (cell in range must be a formula, not a literal), `totals_tie` (cell == SUM(range) within tolerance), `relationship` (simple cross-cell expression).
- **AC:** Each rule type has a passing and a failing fixture test.

### F4 — CI gate
`testsheet run --junit out.xml` exits non-zero on failure; ship a `action.yml` GitHub Action.
- **AC:** A failing model fails the workflow; report uploaded as artifact.

### F5 — Reports
Console summary + JSON + JUnit XML + minimal HTML.
- **AC:** JSON schema documented; HTML opens standalone and lists drifts grouped by sheet/severity.

### Explicitly OUT of MVP scope
Excel ribbon add-in/GUI, SaaS dashboard, multi-user accounts, auth, the "send" button integration, AI rule suggestion. (These are fast-follows once the core is proven.)

## 7. Milestones

- **M0 — Scaffold.** Repo, CLI skeleton (`baseline`, `run`), config dir, test harness, sample fixture workbooks.
- **M1 — Value spike.** Compare LibreOffice-recalc vs. pycel/formulas on 3 sample models; pick the evaluator; document tradeoffs.
- **M2 — Diff engine (F1, F2).** Baseline capture + drift classification + console report.
- **M3 — Rules (F3).** `rules.yaml` parser + the six rule types + fixtures.
- **M4 — CI + reporters (F4, F5).** JUnit/JSON/HTML + GitHub Action.
- **M5 — Hardening.** Structural-change heuristic, float tolerance config, large-model perf pass, docs + quickstart.

## 8. Definition of done (v1)
A non-technical analyst can run `testsheet baseline` then `testsheet run` on a real FP&A model, get a correct, readable drift report, encode three invariant rules, and wire it into a GitHub Action that blocks a bad commit — all from the README, with no code changes to the workbook.

## 9. First task for the build session
Generate the M0 scaffold: project layout, `pyproject.toml` (Python 3.11, `openpyxl`, `typer`, `pyyaml`, `pytest`), a `testsheet` CLI with stub `baseline` and `run` commands, a `fixtures/` folder with 2 tiny sample `.xlsx` files (one "good", one with a planted error), and a passing smoke test. Then proceed to M1.

---

### Background sources (for the builder's context)
- Spreadsheet error rates & losses: CNBC "spreadsheet blunders costing business billions"; Golimelight "hidden cost of spreadsheets".
- SOX spreadsheet risk: Mitratech resource hub.
- DIY state of the art: Xlwings "unit tests for Microsoft Excel".
- Why no incumbent: Microsoft RSAT uses Excel only as a test-parameter sheet, not to test the workbook.
