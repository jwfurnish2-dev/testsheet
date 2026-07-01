"""
Diagnostic script — print the actual cell ref keys and values
that the formulas library produces, so we can fix the parser.

Usage (from repo root):
  python scripts/debug_formulas.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

FIXTURE = Path(__file__).parent.parent / "tests" / "fixtures" / "model_simple.xlsx"

import formulas

print(f"Loading: {FIXTURE}")
xl_model = formulas.ExcelModel().loads(str(FIXTURE)).finish()
xl_model.calculate()

print(f"\nTotal cells in xl_model.cells: {len(xl_model.cells)}")

# Check the return value of calculate()
print("\n--- Return value of calculate() ---")
solution = xl_model.calculate()
print(f"  type: {type(solution)}")
print(f"  len: {len(solution) if hasattr(solution, '__len__') else 'n/a'}")

# Try iterating the solution
print("\n--- Solution contents (first 20) ---")
try:
    for i, (k, v) in enumerate(solution.items()):
        print(f"  {k!r}: {v!r}")
        if i >= 19:
            print("  ...")
            break
except Exception as e:
    print(f"  Error iterating: {e}")
    print(f"  dir(solution): {[a for a in dir(solution) if not a.startswith('__')]}")
