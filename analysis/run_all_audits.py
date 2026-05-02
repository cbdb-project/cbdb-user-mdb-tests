"""Aggregate runner for every static audit in `analysis/`.

Usage:
    python analysis/run_all_audits.py          # human-readable summary
    python analysis/run_all_audits.py --quiet  # only show non-clean

Exit code: number of audits that exited non-zero (i.e. had findings).
0 means every audit is clean.

Designed for the per-release workflow — `AGENTS.md` lists each audit
individually but in practice you want a single command that confirms
nothing regressed (or surfaces what did).
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
]


def _run_one(name: str) -> tuple[int, float, str]:
    """Return (exit_code, elapsed_seconds, summary_line)."""
    t0 = time.time()
    proc = subprocess.run(
        [sys.executable, str(ANALYSIS / name)],
        capture_output=True, text=True,
        env={"PYTHONIOENCODING": "utf-8", **__import__("os").environ},
    )
    dt = time.time() - t0
    # Pull the trailing "=== total ... ===" line if present.
    out = (proc.stdout or "") + (proc.stderr or "")
    summary = ""
    for line in reversed(out.splitlines()):
        line = line.strip()
        if line.startswith("==="):
            summary = line.strip("= ")
            break
    return proc.returncode, dt, summary


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--quiet", action="store_true",
                    help="only show audits with findings")
    args = ap.parse_args()

    print(f"Running {len(AUDITS)} static audits ...\n")
    findings = 0
    total_t = 0.0
    flagged: list[tuple[str, str, float]] = []
    clean: list[tuple[str, str, float]] = []
    for name in AUDITS:
        rc, dt, summary = _run_one(name)
        total_t += dt
        line = f"{name:50s}  {dt:5.1f}s  {summary}"
        if rc != 0:
            findings += 1
            flagged.append((name, summary, dt))
        else:
            clean.append((name, summary, dt))

    if flagged:
        print("FLAGGED (audit reported findings):")
        for name, summary, dt in flagged:
            print(f"  {name:50s}  {dt:5.1f}s  {summary}")
        print()
    if not args.quiet:
        print(f"CLEAN ({len(clean)} audits, no findings):")
        for name, summary, dt in clean:
            print(f"  {name:50s}  {dt:5.1f}s  {summary}")
        print()
    print(f"=== {findings}/{len(AUDITS)} audits flagged in "
          f"{total_t:.1f}s total ===")
    return findings


if __name__ == "__main__":
    raise SystemExit(main())
