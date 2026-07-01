"""Config loader — reads .testsheet/config.yaml if present."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml


_DEFAULTS: dict[str, Any] = {
    "diff": {
        "rel_tol": 1e-9,
        "abs_tol": 1e-12,
    },
    "structural_change_threshold": 0.5,
}


def load_config(testsheet_dir: Path) -> dict[str, Any]:
    """
    Load .testsheet/config.yaml and merge with defaults.

    Returns a dict with keys:
      diff.rel_tol                  float  (default 1e-9)
      diff.abs_tol                  float  (default 1e-12)
      structural_change_threshold   float  (default 0.5)
    """
    config_path = Path(testsheet_dir) / "config.yaml"
    if not config_path.exists():
        return _deep_copy(_DEFAULTS)

    raw = yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}
    merged = _deep_copy(_DEFAULTS)

    # Merge diff section
    if "diff" in raw:
        diff = raw["diff"]
        if "rel_tol" in diff:
            merged["diff"]["rel_tol"] = float(diff["rel_tol"])
        if "abs_tol" in diff:
            merged["diff"]["abs_tol"] = float(diff["abs_tol"])

    # Merge top-level keys
    if "structural_change_threshold" in raw:
        merged["structural_change_threshold"] = float(raw["structural_change_threshold"])

    return merged


def _deep_copy(d: dict) -> dict:
    import copy
    return copy.deepcopy(d)
