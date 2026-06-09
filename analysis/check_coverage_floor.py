"""Check the current build against the coverage FLOOR.

Fails (exit 1, prints named gaps) if the current build covers LESS than
build_20260430 did — a regression below the floor.  See
`docs/standardized-testing-remediation-plan.md` task C0 and
`docs/coverage-floor.json` (produced by analysis/build_coverage_floor.py).

The pure functions (`*_gaps`) are unit-tested in
tests/test_coverage_floor.py with synthetic data (no MDB).  The CLI wires
them to the live artifacts.

Usage (run inside the standardized pipeline, AFTER the current build's
export coverage matrix + drift classifiers have been regenerated):
    python analysis/check_coverage_floor.py
    python analysis/check_coverage_floor.py --current reports/export_coverage_matrix.json
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FLOOR_JSON = ROOT / "docs" / "coverage-floor.json"
# Where analysis/export_coverage_matrix.py writes the CURRENT build's matrix.
DEFAULT_CURRENT = ROOT / "reports" / "export_coverage_matrix.json"
# Drift-classifier outputs that prove the Appendix-A classification ran.
APPENDIX_EVIDENCE = {
    "index_drift_classification": [
        ROOT / "reports" / "index_drift_classification.json",
        ROOT / "reports" / "index_addr_drift_classification.json",
    ],
}


# --------------------------------------------------------------------- #
# Pure functions (unit-tested; no file/MDB access)
# --------------------------------------------------------------------- #

def depth_rank(depth: str, ladder: list[str]) -> int:
    """Rank of a coverage depth on the ladder; -1 if unknown."""
    try:
        return ladder.index(depth)
    except ValueError:
        return -1


def export_gaps(current_rows: list[dict], floor_cells: list[dict],
                ladder: list[str]) -> list[str]:
    """Floor export cells not met by the current build.

    `current_rows`: [{form, button, depth}, ...] (current export matrix).
    `floor_cells`:  [{form, button, min_depth}, ...] (the floor).
    A gap = a floor cell whose current depth is missing or ranks BELOW
    its floor min_depth.
    """
    cur = {(str(r["form"]).strip(), str(r["button"]).strip()):
           r.get("depth", "none")
           for r in current_rows}
    gaps = []
    for cell in floor_cells:
        key = (str(cell["form"]).strip(), str(cell["button"]).strip())
        need = cell["min_depth"]
        if depth_rank(need, ladder) < 0:
            # A floor cell with an unknown depth (hand-edit typo) would
            # otherwise silently pass every current depth — flag it.
            gaps.append(
                f"export cell {key[0]}.{key[1]} has unknown floor "
                f"min_depth {need!r} (not on the depth ladder)"
            )
            continue
        have = cur.get(key)
        if have is None:
            gaps.append(
                f"export cell {key[0]}.{key[1]} missing from current build "
                f"(floor requires depth >= {need!r})"
            )
        elif depth_rank(have, ladder) < depth_rank(need, ladder):
            gaps.append(
                f"export cell {key[0]}.{key[1]} regressed: current depth "
                f"{have!r} < floor {need!r}"
            )
    return gaps


def audit_gaps(present_audits: list[str],
               required_audits: list[str]) -> list[str]:
    """Required static audits that are no longer present/runnable."""
    present = set(present_audits)
    return [
        f"required audit missing: {a}"
        for a in required_audits if a not in present
    ]


def appendix_gaps(present_appendices: list[str],
                  required_appendices: list[str]) -> list[str]:
    """Required appendix kinds not produced this build."""
    present = set(present_appendices)
    return [
        f"required appendix not produced: {a}"
        for a in required_appendices if a not in present
    ]


def all_gaps(current_rows: list[dict], present_audits: list[str],
             present_appendices: list[str], floor: dict) -> list[str]:
    ladder = floor.get("depth_ladder_low_to_high") or floor["_meta"][
        "depth_ladder_low_to_high"]
    return (
        export_gaps(current_rows, floor["export_cells"], ladder)
        + audit_gaps(present_audits, floor["required_audits"])
        + appendix_gaps(present_appendices, floor["required_appendices"])
    )


# --------------------------------------------------------------------- #
# CLI wiring (reads live artifacts)
# --------------------------------------------------------------------- #

def _present_audits() -> list[str]:
    return sorted(
        p.name for p in (ROOT / "analysis").glob("audit_*.py")
        if p.name != "audit_lib.py"
    )


def _present_appendices() -> list[str]:
    out = []
    for kind, evidence in APPENDIX_EVIDENCE.items():
        if all(p.exists() for p in evidence):
            out.append(kind)
    return out


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--current", default=str(DEFAULT_CURRENT),
                    help="current build's export_coverage_matrix.json")
    ap.add_argument("--floor", default=str(FLOOR_JSON))
    args = ap.parse_args(argv)

    floor = json.loads(Path(args.floor).read_text(encoding="utf-8"))
    # depth ladder lives under _meta in the generated file
    floor.setdefault("depth_ladder_low_to_high",
                     floor["_meta"]["depth_ladder_low_to_high"])

    current_path = Path(args.current)
    if not current_path.exists():
        print(
            f"[coverage-floor] CURRENT export matrix not found: {current_path}\n"
            f"  Regenerate it first (analysis/export_coverage_matrix.py) as part\n"
            f"  of the standardized run, then re-check.  Cannot verify the floor\n"
            f"  without it."
        )
        return 2
    current = json.loads(current_path.read_text(encoding="utf-8"))
    current_rows = current.get("rows", current if isinstance(current, list) else [])

    gaps = all_gaps(current_rows, _present_audits(),
                    _present_appendices(), floor)
    if gaps:
        print(f"[coverage-floor] {len(gaps)} GAP(S) below build_20260430:")
        for g in gaps:
            print(f"  - {g}")
        return 1
    print("[coverage-floor] OK — current build meets the build_20260430 floor.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
