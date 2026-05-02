"""Static audit (9th scanner): for each `On Error GoTo <label>` (or
`Resume <label>` / `GoTo <label>`), verify that `<label>:` is
actually defined inside the same Sub / Function.

Bug shape: a developer renames `Err_CmdGIS_Click:` to
`Err_CmdGIS_Click_Old:` (or just typos it) but forgets to update the
`On Error GoTo Err_CmdGIS_Click` line at the top.  At runtime, when
an error fires, VBA has nowhere to jump and crashes with "Invalid
procedure call" — masking the original error and confusing every
downstream debugger.

Walks each Sub linearly:
  1. Collect every `<name>:` label definition.
  2. Collect every `On Error GoTo <name>` / `Resume <name>` /
     `GoTo <name>` reference.
  3. Diff: any reference whose target isn't a defined label is a
     bug.

Special cases (skipped):
  - `On Error GoTo 0`  — VBA built-in, disables current handler.
  - `On Error GoTo -1` — VBA built-in (Office 2010+), clears handler.
  - `Resume Next` / `Resume` (no arg) — built-in.
  - References inside comments — already filtered by `'` skip.
"""
from __future__ import annotations

import re
from collections import defaultdict
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "analysis"))
from audit_lib import read_vba_lines

VBA_DIR = ROOT / "analysis" / "dump" / "vba"

SUB_START_RE = re.compile(
    r"^\s*(?:Private|Public)?\s*(?:Sub|Function)\s+([A-Za-z_]\w*)",
    re.IGNORECASE,
)
SUB_END_RE = re.compile(r"^\s*End\s+(?:Sub|Function)\b", re.IGNORECASE)
LABEL_RE = re.compile(r"^\s*([A-Za-z_]\w*)\s*:\s*(?:'.*)?$")
ON_ERROR_GOTO_RE = re.compile(
    r"\bOn\s+Error\s+GoTo\s+([A-Za-z_]\w*)\b", re.IGNORECASE
)
GOTO_RE = re.compile(r"\bGoTo\s+([A-Za-z_]\w*)\b", re.IGNORECASE)
RESUME_RE = re.compile(r"\bResume\s+([A-Za-z_]\w*)\b", re.IGNORECASE)


def _scan(form_path: Path) -> list[tuple[int, str, str, str]]:
    """Return (line_no, sub, kind, target) per missing-label reference."""
    lines = read_vba_lines(form_path)
    diagnostics: list[tuple[int, str, str, str]] = []
    cur_sub = ""
    labels: set[str] = set()
    refs: list[tuple[int, str, str]] = []  # (ln, kind, target)

    def _flush():
        for ln, kind, target in refs:
            # `Resume Next` is a special case — match excludes it
            # because RESUME_RE is `\bResume\s+(\w+)\b` and "Next" is
            # a word — we'd hit it.  Filter here.
            if target.lower() in ("next",):
                continue
            if target not in labels:
                diagnostics.append((ln, cur_sub, kind, target))

    for ln_no, line in enumerate(lines, 1):
        stripped = line.lstrip()
        if stripped.startswith("'"):
            continue

        m_sub = SUB_START_RE.match(line)
        if m_sub:
            if cur_sub:
                _flush()
            cur_sub = m_sub.group(1)
            labels = set()
            refs = []
            continue

        if SUB_END_RE.match(line):
            _flush()
            cur_sub = ""
            labels = set()
            refs = []
            continue

        if not cur_sub:
            continue

        m_label = LABEL_RE.match(line)
        if m_label:
            labels.add(m_label.group(1))
            continue

        for m in ON_ERROR_GOTO_RE.finditer(line):
            target = m.group(1)
            if target in ("0",):
                continue  # built-in
            refs.append((ln_no, "On Error GoTo", target))

        for m in RESUME_RE.finditer(line):
            target = m.group(1)
            refs.append((ln_no, "Resume", target))

        # Plain `GoTo <label>` (skip `On Error GoTo` matches above).
        for m in GOTO_RE.finditer(line):
            # Don't double-count On Error GoTo matches.
            if "On Error" in line[max(0, m.start() - 12):m.start()]:
                continue
            target = m.group(1)
            if target in ("0",):
                continue
            refs.append((ln_no, "GoTo", target))

    if cur_sub:
        _flush()
    return diagnostics


def main() -> int:
    forms = sorted(VBA_DIR.glob("Form_*.vb"))
    grand_total = 0
    print(f"=== Auditing {len(forms)} VBA modules for "
          f"missing label targets ===\n")
    for f in forms:
        diags = _scan(f)
        if not diags:
            continue
        grand_total += len(diags)
        print(f"\n[FLAG] {f.name}:")
        by_sub: dict[str, list[tuple[int, str, str]]] = defaultdict(list)
        for ln, sub, kind, target in diags:
            by_sub[sub].append((ln, kind, target))
        for sub, occurrences in sorted(by_sub.items()):
            for ln, kind, target in occurrences[:5]:
                print(f"  Sub {sub}: line {ln}: {kind} {target} "
                      f"— label not defined in this Sub")
    print(f"\n=== total flagged: {grand_total} ===")
    return 1 if grand_total else 0


if __name__ == "__main__":
    raise SystemExit(main())
