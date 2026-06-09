"""Generate the coverage FLOOR from the build_20260430 archive.

The floor is the minimum surface the standardized test method must keep
covering — derived from what the build_20260430 run actually exercised.
build_20260430 is NOT a quality standard (it was one run's findings); it
is only a *coverage floor* (do not regress below it).  See
`docs/standardized-testing-remediation-plan.md` task C0.

Sources (all under reports/archive/build_20260430/):
  - export_coverage_matrix.json  -> per (form, export-button) coverage DEPTH
  - pytest_marker_inventory.json -> collected-test counts (informational)
Plus the live `analysis/audit_*.py` set (every static audit must run) and
the appendix kinds that must be produced.

Output: docs/coverage-floor.json (committed).  Consumed by
analysis/check_coverage_floor.py.

Usage:
    python analysis/build_coverage_floor.py
"""
from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ARCHIVE = ROOT / "reports" / "archive" / "build_20260430"
EXPORT_MATRIX = ARCHIVE / "export_coverage_matrix.json"
MARKER_INV = ARCHIVE / "pytest_marker_inventory.json"
OUT = ROOT / "docs" / "coverage-floor.json"

# Depth ladder low->high; a current cell must reach >= its floor depth.
# Mirrors export_coverage_matrix.json summary.depth_ladder_low_to_high.
DEPTH_LADDER = ["none", "smoke", "structural", "manifest", "byte_golden"]

# audit_lib.py is a shared helper, not an audit; everything else under
# analysis/audit_*.py is a static audit that must run each build.
_AUDIT_LIB = "audit_lib.py"


def _export_cells() -> list[dict]:
    """Every (form, button) the archive exercised at depth > none."""
    data = json.loads(EXPORT_MATRIX.read_text(encoding="utf-8"))
    cells = []
    for r in data["rows"]:
        depth = r.get("depth", "none")
        if depth == "none":
            continue  # not covered in the floor build -> not part of the floor
        cells.append({
            "form": r["form"],
            "button": r["button"],
            "min_depth": depth,
        })
    # Deterministic order so the committed JSON is stable.
    cells.sort(key=lambda c: (c["form"], c["button"]))
    return cells


def _required_audits() -> list[str]:
    audits = sorted(
        p.name for p in (ROOT / "analysis").glob("audit_*.py")
        if p.name != _AUDIT_LIB
    )
    return audits


def _marker_summary() -> dict:
    if not MARKER_INV.exists():
        return {}
    d = json.loads(MARKER_INV.read_text(encoding="utf-8"))
    s = d.get("summary", {})
    return {
        "n_collected_default": s.get("n_collected_default"),
        "n_collected_with_include_vba": s.get("n_collected_with_include_vba"),
    }


def build_floor() -> dict:
    return {
        "_meta": {
            "purpose": "Coverage FLOOR (minimum surface to keep covering), "
                       "NOT a quality standard.  The current build must "
                       "meet-or-exceed every item here.",
            "source_build": "build_20260430",
            "generated_by": "analysis/build_coverage_floor.py",
            "generated_from": [
                "reports/archive/build_20260430/export_coverage_matrix.json",
                "reports/archive/build_20260430/pytest_marker_inventory.json",
                "analysis/audit_*.py (live set at generation time)",
            ],
            "depth_ladder_low_to_high": DEPTH_LADDER,
            "note_test_files_retired": {
                # Intentional removals — recorded so their absence is NOT a
                # silent coverage loss.  Each maps to what now covers it.
                "test_markdown_report.py":
                    "superseded by tests/test_report_triage_gate.py (report "
                    "structure/triage) + run_all_audits.py & "
                    "audit_report_code_labels.py / "
                    "audit_report_screenshot_consistency.py (run directly "
                    "post-rebuild per issue-report-maintainer skill).",
            },
        },
        # Hard checks (analysis/check_coverage_floor.py):
        "export_cells": _export_cells(),
        "required_audits": _required_audits(),
        # Appendix kinds build_20260430 produced.  TablesFields (B) and
        # ForeignKeys (C) appendices are a later addition gated by
        # remediation-plan tasks B7/C1, not part of this floor.
        "required_appendices": ["index_drift_classification"],
        # Context only (not a hard gate — marker inventory is top-N only):
        "informational": {
            "build_20260430_marker_inventory": _marker_summary(),
        },
    }


def main() -> int:
    floor = build_floor()
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(
        json.dumps(floor, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(f"wrote {OUT.relative_to(ROOT)}")
    print(f"  export_cells:        {len(floor['export_cells'])}")
    print(f"  required_audits:     {len(floor['required_audits'])}")
    print(f"  required_appendices: {floor['required_appendices']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
