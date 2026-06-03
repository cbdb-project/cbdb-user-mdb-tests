"""Build the form × button coverage matrix from a pytest JSON report.

Reads:
  - docs/test-matrix-mapping.json  — form/button → test node-id patterns
  - reports/pytest_report_*.json   — test outcomes (via --report flag)

Writes:
  - reports/coverage_matrix.json   — structured matrix for generate_report.py

Usage:
    python analysis/build_coverage_matrix.py --report reports/pytest_report_build20260518.json

Outcome priority (worst wins across all matching tests in a cell):
  FAIL > ERROR > PASS > SKIP > NOT_RUN

N/A cells are written as-is regardless of test matches.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
MAPPING_JSON = ROOT / "docs" / "test-matrix-mapping.json"
OUT_JSON = ROOT / "reports" / "coverage_matrix.json"

# Outcome priority: higher index = worse; determines cell outcome.
# FAIL > ERROR > SKIP > PASS > NOT_RUN
# SKIP outranks PASS: a skipped test leaves an unverified path; PASS means verified.
_PRIORITY = ["NOT_RUN", "PASS", "SKIP", "ERROR", "FAIL"]


def _priority(outcome: str) -> int:
    try:
        return _PRIORITY.index(outcome)
    except ValueError:
        return _PRIORITY.index("ERROR")  # unknown → treat as error


def _worst(a: str, b: str) -> str:
    return a if _priority(a) >= _priority(b) else b


def _pytest_outcome(outcome: str) -> str:
    """Map pytest outcome strings to matrix cell values."""
    return {
        "passed": "PASS",
        "failed": "FAIL",
        "error":  "ERROR",
        "skipped": "SKIP",
    }.get(outcome.lower(), "ERROR")


def build(report_path: Path, mapping_path: Path = MAPPING_JSON) -> dict:
    """Return the coverage matrix dict."""
    mapping = json.loads(mapping_path.read_text(encoding="utf-8"))
    report = json.loads(report_path.read_text(encoding="utf-8"))

    tests = report.get("tests", [])
    forms = mapping["forms"]
    buttons = mapping["buttons"]
    na_map = mapping["na"]
    cell_mapping = mapping["mapping"]

    matrix: dict[str, dict[str, str]] = {}

    for form in forms:
        matrix[form] = {}
        na_buttons = set(na_map.get(form, []))
        form_cells = cell_mapping.get(form, {})

        for button in buttons:
            if button in na_buttons:
                matrix[form][button] = "N/A"
                continue

            cell_def = form_cells.get(button)
            if cell_def is None:
                # Button applicable but no mapping entry → NOT RUN
                matrix[form][button] = "NOT_RUN"
                continue

            patterns = [p for p in cell_def.get("patterns", []) if p]
            if not patterns:
                matrix[form][button] = "NOT_RUN"
                continue

            # Find all tests matching any pattern
            cell_outcome = "NOT_RUN"
            for test in tests:
                node_id = test.get("nodeid", "")
                if any(pat in node_id for pat in patterns):
                    outcome = _pytest_outcome(test.get("outcome", "error"))
                    cell_outcome = _worst(cell_outcome, outcome)

            matrix[form][button] = cell_outcome

    return {
        "report": str(report_path),
        "mapping": str(mapping_path),
        "forms": forms,
        "buttons": buttons,
        "matrix": matrix,
        "summary": _summarise(matrix, forms, buttons),
    }


def _summarise(matrix: dict, forms: list, buttons: list) -> dict:
    counts: dict[str, int] = {o: 0 for o in _PRIORITY + ["N/A"]}
    for form in forms:
        for button in buttons:
            counts[matrix[form].get(button, "NOT_RUN")] += 1
    return counts


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "--report", required=True,
        help="Path to pytest JSON report (--json-report output)",
    )
    parser.add_argument(
        "--out", default=str(OUT_JSON),
        help=f"Output path (default: {OUT_JSON})",
    )
    args = parser.parse_args(argv)

    report_path = Path(args.report)
    if not report_path.exists():
        print(f"ERROR: report not found: {report_path}", file=sys.stderr)
        return 1

    data = build(report_path)

    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    summary = data["summary"]
    print(f"Coverage matrix written to {out_path}")
    print(f"  PASS={summary.get('PASS',0)}  FAIL={summary.get('FAIL',0)}"
          f"  ERROR={summary.get('ERROR',0)}  SKIP={summary.get('SKIP',0)}"
          f"  NOT_RUN={summary.get('NOT_RUN',0)}  N/A={summary.get('N/A',0)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
