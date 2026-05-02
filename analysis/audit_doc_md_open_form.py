"""Static audit (14th scanner): for every `DoCmd.OpenForm "<form>"`,
verify `<form>` exists in `control_inventory.json`.

Bug shape: a `DoCmd.OpenForm "frmFoo"` call where `frmFoo` was
deleted or renamed.  At runtime DAO throws "The Microsoft Access
database engine cannot find the input table or query 'frmFoo'."
The button / event silently fails (or pops a confusing error).

Companion to `audit_cross_form_references.py` which catches
`Forms!<form>!<ctl>` references — `DoCmd.OpenForm` is the *opening*
side; cross-form refs poke an *already open* form.

Both should be clean for a healthy .mdb.
"""
from __future__ import annotations

import json
import re
from collections import defaultdict
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "analysis"))
from audit_lib import read_vba_lines

VBA_DIR = ROOT / "analysis" / "dump" / "vba"
CONTROL_INV = ROOT / "analysis" / "dump" / "control_inventory.json"


# Match: DoCmd.OpenForm "<form>"  with literal form name.
OPEN_FORM_RE = re.compile(
    r'\bDoCmd\.OpenForm\s+"([A-Za-z_]\w*)"',
    re.IGNORECASE,
)


def _load_form_names() -> set[str]:
    inv = json.loads(CONTROL_INV.read_text(encoding="utf-8"))
    return {fname.lower() for fname, info in inv.items()
            if isinstance(info, dict)}


def _scan(form_path: Path, known_forms: set[str]
          ) -> list[tuple[int, str]]:
    """Return (line_no, target_form)."""
    lines = read_vba_lines(form_path)
    diagnostics: list[tuple[int, str]] = []
    seen: set[str] = set()
    for ln_no, line in enumerate(lines, 1):
        stripped = line.lstrip()
        if stripped.startswith("'"):
            continue
        for m in OPEN_FORM_RE.finditer(line):
            fm = m.group(1)
            if fm.lower() in seen:
                continue
            seen.add(fm.lower())
            if fm.lower() not in known_forms:
                diagnostics.append((ln_no, fm))
    return diagnostics


def main() -> int:
    known = _load_form_names()
    forms = sorted(p for p in VBA_DIR.glob("Form_*.vb")
                    if "TMPCLP" not in p.name.upper())
    grand_total = 0
    print(f"=== Auditing {len(forms)} VBA modules for "
          f"DoCmd.OpenForm referencing non-existent forms "
          f"(known forms: {len(known)}) ===\n")
    for f in forms:
        diags = _scan(f, known)
        if not diags:
            continue
        grand_total += len(diags)
        print(f"\n[FLAG] {f.name}:")
        for ln, fm in diags[:8]:
            print(f"  line {ln}: DoCmd.OpenForm \"{fm}\" — "
                  f"form not in inventory")
    print(f"\n=== total flagged: {grand_total} ===")
    return 1 if grand_total else 0


if __name__ == "__main__":
    raise SystemExit(main())
