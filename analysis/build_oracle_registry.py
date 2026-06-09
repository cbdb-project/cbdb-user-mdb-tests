"""Classify every tests/test_*.py by its ORACLE type (D0).

A test's "pass" is only evidence for what its oracle can actually prove.
Three oracle classes (see docs/standardized-testing-remediation-plan.md D0):

  A  real VBA  x  INDEPENDENT oracle  — drives the real VBA via COM and
     asserts against an independent source (HelpFile values / hand-derived
     SQL / the online snapshot / structural invariants).  CAN find VBA bugs;
     counts as VBA verification.
  B  real VBA  x  REPLAY oracle       — drives real VBA but compares to
     tests/cbdb_replay (a Python re-impl read from the same VBA).  A shared
     bug passes; only detects DIVERGENCE.  Weak evidence.
  C  replay    x  golden              — runs the Python replay and compares
     to a frozen golden.  Tests the replay, NOT the VBA.  Says nothing about
     the .mdb.
  NA infra / unit                     — pure data/contract tests (no Access,
     no replay): the gate, coverage-floor, build-stamp, timeout, golden-helper,
     refresh-decision tests, etc.

Classification is by import signature (deterministic, re-runnable):
  imports cbdb_replay  AND  cbdb_driver/VbaSession  -> B  (differential)
  imports cbdb_replay  (no VbaSession)              -> C
  imports cbdb_driver/VbaSession (no cbdb_replay)   -> A
  neither                                           -> NA

Output: docs/oracle-classification.json.  Pinned by tests/test_oracle_classification.py
(every test file must be classified; only A counts as VBA verification).

Usage:  python analysis/build_oracle_registry.py
"""
from __future__ import annotations

import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TESTS = ROOT / "tests"
OUT = ROOT / "docs" / "oracle-classification.json"

_IMPORT_LINE = re.compile(r"^\s*(from|import)\s")


def classify_source(src: str) -> str:
    # Only inspect IMPORT lines, never string literals in test bodies (a test
    # that merely *mentions* 'cbdb_replay'/'VbaSession' in an assertion string
    # must not be misclassified).
    imports = "\n".join(
        ln for ln in src.splitlines() if _IMPORT_LINE.match(ln))
    has_replay = "cbdb_replay" in imports
    # Real-VBA signal: a direct import OR use of a conftest COM fixture
    # (com_app / fresh_form), since those drive real Access but the
    # cbdb_driver import lives in conftest, invisible to per-file scanning.
    has_vba = (
        ("cbdb_driver" in imports) or ("VbaSession" in imports)
        or bool(re.search(r"\b(com_app|fresh_form)\b", src))
    )
    if has_replay and has_vba:
        return "B"
    if has_replay:
        return "C"
    if has_vba:
        return "A"
    return "NA"


def build() -> dict:
    entries = {}
    for p in sorted(TESTS.glob("test_*.py")):
        src = p.read_text(encoding="utf-8", errors="replace")
        entries[p.name] = classify_source(src)
    counts: dict[str, int] = {}
    for c in entries.values():
        counts[c] = counts.get(c, 0) + 1
    return {
        "_meta": {
            "purpose": "Oracle classification of each test file (D0).  Only "
                       "class A (real VBA x independent oracle) counts as VBA "
                       "verification; B/C passes are NOT evidence the .mdb's VBA "
                       "is correct.  See docs/standardized-testing-remediation-plan.md.",
            "generated_by": "analysis/build_oracle_registry.py",
            "classes": {
                "A": "real VBA x independent oracle (counts as VBA verification)",
                "B": "real VBA x cbdb_replay oracle (divergence-only; weak)",
                "C": "replay x golden (tests Python, not the VBA)",
                "NA": "infra / pure-unit (no Access, no replay)",
            },
            "counts": counts,
        },
        "tests": entries,
    }


def main() -> int:
    reg = build()
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(reg, ensure_ascii=False, indent=2) + "\n",
                   encoding="utf-8")
    print(f"wrote {OUT.relative_to(ROOT)}  counts={reg['_meta']['counts']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
