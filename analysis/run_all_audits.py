"""Aggregate runner for every static audit in `analysis/`.

Usage:
    python analysis/run_all_audits.py            # human-readable summary
    python analysis/run_all_audits.py --quiet    # only show non-clean
    python analysis/run_all_audits.py --ci       # exit non-zero ONLY on
                                                 # NEW findings beyond
                                                 # the recorded baseline
    python analysis/run_all_audits.py --update-baseline
                                                 # snapshot today's
                                                 # findings as the new
                                                 # known-OK baseline

Exit code:
  - default mode: number of audits that flagged anything (human view).
    Note: every documented bug currently flags — so default exit code
    is 6/20-something, NOT 'zero issues'.  Don't wire this directly
    into CI.
  - `--ci`:  exit code = (flagged count) - (baseline count).  Stays
    at 0 as long as no NEW findings appear.  This is the CI-safe
    mode.

Baseline lives at `analysis/audit_baseline.json` — small JSON dict
of `{audit_filename: int_finding_count}`.
"""
from __future__ import annotations

import argparse
import re
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ANALYSIS = ROOT / "analysis"

# Order matters only for output readability — clean audits should
# come last so the user sees "the bad stuff" first.
AUDITS = [
    # SQL column / table resolution (caught Bugs #5, #6 earlier)
    "audit_sql_columns.py",
    "audit_sql_table_names.py",
    "audit_insert_select_columns.py",
    "audit_saved_queries.py",
    # Control / form resolution (caught Bugs #4, #5)
    "audit_missing_controls.py",
    # View alias detection (caught Bug #1)
    "audit_view_aliases.py",
    "audit_duplicate_aliases.py",
    # Recordset field tracking (caught Bugs #7-#9)
    "audit_recordset_fields.py",
    "audit_recordset_sql_projection.py",
    # Sub-form / cross-form (caught Bugs #10-#14)
    "audit_subform_control_sources.py",
    "audit_cross_form_references.py",
    "audit_doc_md_open_form.py",
    # Domain aggregates
    "audit_dcount_where_columns.py",
    "audit_dlookup_fields.py",
    # Code-smell / orphan detectors (caught Bugs #15-#19)
    "audit_orphan_event_handlers.py",
    "audit_event_handlers_exist.py",
    "audit_error_label_targets.py",
    "audit_blocking_msgbox.py",
    "audit_control_row_sources.py",
    "audit_dynamic_record_source.py",
    "audit_view_to_view_columns.py",
]


_FINDING_COUNT_RE = __import__("re").compile(
    r"total[^:]*:\s*(\d+)", __import__("re").IGNORECASE
)


def _run_one(name: str) -> tuple[int, float, str, int]:
    """Return (exit_code, elapsed_seconds, summary_line, finding_count)."""
    t0 = time.time()
    proc = subprocess.run(
        [sys.executable, str(ANALYSIS / name)],
        capture_output=True, text=True,
        env={"PYTHONIOENCODING": "utf-8", **__import__("os").environ},
    )
    dt = time.time() - t0
    out = (proc.stdout or "") + (proc.stderr or "")
    summary = ""
    for line in reversed(out.splitlines()):
        line = line.strip()
        if line.startswith("==="):
            summary = line.strip("= ")
            break
    # Extract the integer finding count from the summary so --ci can
    # baseline-diff.  Fallback to exit code (1 means "had findings",
    # 0 means clean) when summary doesn't carry a count.
    n = 0
    m = _FINDING_COUNT_RE.search(summary)
    if m:
        n = int(m.group(1))
    elif proc.returncode != 0:
        n = 1
    return proc.returncode, dt, summary, n


BASELINE_PATH = ROOT / "analysis" / "audit_baseline.json"


def _load_baseline() -> dict[str, int]:
    import json
    if BASELINE_PATH.exists():
        try:
            return json.loads(BASELINE_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def _save_baseline(counts: dict[str, int]) -> None:
    import json
    BASELINE_PATH.write_text(
        json.dumps(counts, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--quiet", action="store_true",
                    help="only show audits with findings")
    ap.add_argument("--ci", action="store_true",
                    help="exit non-zero ONLY on findings ABOVE the "
                         "recorded baseline (analysis/audit_baseline.json)")
    ap.add_argument("--update-baseline", action="store_true",
                    help="snapshot today's findings as the new baseline "
                         "and exit 0")
    args = ap.parse_args()

    baseline = _load_baseline()
    print(f"Running {len(AUDITS)} static audits ...\n")
    if baseline:
        print(f"  (baseline loaded — {sum(baseline.values())} known "
              f"findings across {len(baseline)} audits)\n")

    total_t = 0.0
    counts: dict[str, int] = {}
    flagged: list[tuple[str, str, float, int]] = []
    clean: list[tuple[str, str, float]] = []
    for name in AUDITS:
        rc, dt, summary, n = _run_one(name)
        total_t += dt
        counts[name] = n
        if rc != 0:
            flagged.append((name, summary, dt, n))
        else:
            clean.append((name, summary, dt))

    # Baseline diff — per-audit deltas.
    new_per_audit: dict[str, int] = {}
    for name, n in counts.items():
        baseline_n = int(baseline.get(name, 0))
        if n > baseline_n:
            new_per_audit[name] = n - baseline_n
    new_total = sum(new_per_audit.values())

    if flagged:
        print("FLAGGED (audit reported findings):")
        for name, summary, dt, n in flagged:
            extra = ""
            if name in new_per_audit:
                extra = f"  [NEW: +{new_per_audit[name]} vs baseline]"
            print(f"  {name:50s}  {dt:5.1f}s  {summary}{extra}")
        print()
    if not args.quiet:
        print(f"CLEAN ({len(clean)} audits, no findings):")
        for name, summary, dt in clean:
            print(f"  {name:50s}  {dt:5.1f}s  {summary}")
        print()

    print(f"=== {len(flagged)}/{len(AUDITS)} audits flagged "
          f"({sum(counts.values())} total findings) in "
          f"{total_t:.1f}s ===")
    if baseline:
        print(f"=== {new_total} findings ABOVE baseline "
              f"({len(new_per_audit)} audits with new flags) ===")

    if args.update_baseline:
        _save_baseline(counts)
        print(f"\nbaseline updated → {BASELINE_PATH}")
        return 0
    if args.ci:
        # CI mode: only fail on regressions (findings > baseline).
        if new_total:
            print("\n[--ci] NEW findings above baseline; exiting non-zero.")
        return new_total
    # Default mode: human view; exit code = audits with any finding.
    return len(flagged)


if __name__ == "__main__":
    raise SystemExit(main())
