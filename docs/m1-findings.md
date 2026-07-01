
# M1 Evaluator Spike — Decision Summary

**Winner: `formulas` library — 100% accuracy, Python 3.12+ safe, zero system deps.**

- `pycel` is eliminated: broken on Python 3.12+ (`ast.Str` removed, library unmaintained)
- `formulas` promoted to core dependency; CLI default changed from `pycel` to `formulas`
- LibreOffice remains available as `--evaluator libreoffice` for maximum fidelity
- First-load penalty ~1.5s (schedula graph compilation) is a fixed cost per workbook, acceptable for CI

---

## Raw Spike Runs

## M1 Spike Run — 2026-06-30 14:38

| Evaluator | Model | Accuracy | Time (ms) | Notes |
|---|---|---|---|---|
| pycel | Simple (SUM, arithmetic) | 5/9 (56%) | 70.1 | 4 mismatch(es) |
| pycel | Intermediate (IF, AVERAGE, cross-sheet) | 9/13 (69%) | 43.3 | 4 mismatch(es) |
| pycel | Complex (IFERROR, nested IF, ROUND, chained refs) | 3/11 (27%) | 49.4 | 8 mismatch(es) |
| **pycel TOTAL** | — | **17/33 (52%)** | — | — |

**Recommendation:** Neither evaluator met 90% accuracy threshold — investigate formula gaps before proceeding to M2.

## M1 Spike Run — 2026-06-30 14:50

| Evaluator | Model | Accuracy | Time (ms) | Notes |
|---|---|---|---|---|
| formulas | Simple (SUM, arithmetic) | 0/9 (0%) | 9349.9 | 9 mismatch(es) |
| formulas | Intermediate (IF, AVERAGE, cross-sheet) | 0/13 (0%) | 36.4 | 13 mismatch(es) |
| formulas | Complex (IFERROR, nested IF, ROUND, chained refs) | 0/11 (0%) | 50.7 | 11 mismatch(es) |
| **formulas TOTAL** | — | **0/33 (0%)** | — | — |

**Recommendation:** Neither evaluator met 90% accuracy threshold — investigate formula gaps before proceeding to M2.

## M1 Spike Run — 2026-06-30 15:10

| Evaluator | Model | Accuracy | Time (ms) | Notes |
|---|---|---|---|---|
| formulas | Simple (SUM, arithmetic) | 9/9 (100%) | 1501.0 | OK |
| formulas | Intermediate (IF, AVERAGE, cross-sheet) | 13/13 (100%) | 47.5 | OK |
| formulas | Complex (IFERROR, nested IF, ROUND, chained refs) | 11/11 (100%) | 42.1 | OK |
| **formulas TOTAL** | — | **33/33 (100%)** | — | — |

**Recommendation:** Use `formulas` as the default evaluator for M2+.
