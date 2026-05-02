"""Static audit (16th scanner): the inverse of
`audit_event_handlers_exist.py` — find form-module Subs that look
like event handlers (`<Control>_<Event>`) where `<Control>` no
longer exists on the form.

Examples:
  - `ChkUseYears_Click` defined in Form_LookAtEntry but no
    `ChkUseYears` control on the form (deleted but the handler
    code stayed).
  - `ChkIndexYear_Click` similarly orphaned.

These aren't bugs *by themselves* — orphan handlers just sit
unused; nothing triggers them.  But they're a strong signal that
the form's design and code drifted apart and a related bug may
hide nearby (the team probably forgot something else when they
deleted the control).

`findings.md` Bug #5 mentions 4 such orphans found by
`audit_missing_controls.py` historically.  This audit makes that
finding ongoing-actionable rather than a one-off note.

Heuristic for "looks like event handler":
  Sub name matches `<Control>_(Click|AfterUpdate|BeforeUpdate|
  GotFocus|LostFocus|MouseDown|Change|Enter|Exit|DblClick|KeyDown|
  KeyPress|KeyUp)$` — the standard Access event suffixes.

Output is informational (exit 0 even if findings) since these are
code-smell signals, not bugs.  But we DO list them so a reviewer
can decide.
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


SUB_DEF_RE = re.compile(
    r"^\s*(?:Private|Public)?\s*Sub\s+([A-Za-z_]\w*)\s*\(",
    re.IGNORECASE | re.MULTILINE,
)

# Standard Access control event suffixes (via `<Control>_<Event>`).
EVENT_SUFFIXES = {
    "Click", "DblClick", "AfterUpdate", "BeforeUpdate",
    "GotFocus", "LostFocus", "MouseDown", "MouseUp", "MouseMove",
    "Change", "Enter", "Exit", "KeyDown", "KeyPress", "KeyUp",
    "OnTimer", "Timer",  # form-level
}

# Don't flag form/section-level events that don't have a per-control
# basis (Form_Open, Form_Close, etc).  These are valid even though
# their "control" name is `Form`, `Detail`, etc.
NON_CONTROL_PREFIXES = {
    "Form", "Detail", "Section", "Report", "Page",
}


def _load_inventory() -> dict[str, set[str]]:
    inv = json.loads(CONTROL_INV.read_text(encoding="utf-8"))
    out: dict[str, set[str]] = {}
    for fname, info in inv.items():
        if not isinstance(info, dict):
            continue
        ctls = {(c.get("name") or "").lower()
                for c in info.get("controls", [])
                if c.get("name")}
        out[fname.lower()] = ctls
    return out


def main() -> int:
    inv = _load_inventory()
    forms = sorted(p for p in VBA_DIR.glob("Form_*.vb")
                    if "TMPCLP" not in p.name.upper())
    grand_total = 0
    print(f"=== Auditing {len(forms)} VBA modules for orphan "
          f"event-handler Subs ===\n")
    for f in forms:
        form_name = f.stem.replace("Form_", "")
        ctls = inv.get(form_name.lower())
        if ctls is None:
            continue
        body = f.read_bytes().decode("utf-8")
        orphans: list[str] = []
        seen: set[str] = set()
        for m in SUB_DEF_RE.finditer(body):
            sub_name = m.group(1)
            if sub_name in seen:
                continue
            seen.add(sub_name)
            # Need an underscore to look like Control_Event.
            if "_" not in sub_name:
                continue
            ctl_part, _, ev_part = sub_name.rpartition("_")
            if ev_part not in EVENT_SUFFIXES:
                continue
            if ctl_part in NON_CONTROL_PREFIXES:
                continue
            if ctl_part.lower() not in ctls:
                orphans.append(sub_name)
        if orphans:
            grand_total += len(orphans)
            print(f"\n[INFO] {f.name}:")
            for o in orphans[:8]:
                print(f"  Sub {o}() — control "
                      f"{o.rsplit('_', 1)[0]!r} not on form")
            if len(orphans) > 8:
                print(f"  ... and {len(orphans) - 8} more")
    print(f"\n=== total orphan handler subs: {grand_total} "
          f"(informational — not necessarily bugs) ===")
    # Exit 0 — these are code-smell signals, not failures.
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
