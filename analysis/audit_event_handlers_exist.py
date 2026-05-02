"""Static audit (11th scanner): every form-control event handler
named in `control_inventory.json` (e.g. `CmdGIS_Click` on
LookAtPlace) must have a matching `Sub <name>()` defined in the
form's VBA module.

Bug shape: a developer renames a Sub (e.g. `CmdQuery_Click` →
`CmdQuery2_Click`) but forgets to update the form's design-time
OnClick property.  Or the property holds a stale reference to a
deleted sub.  At runtime, clicking the button does NOTHING (Access
silently swallows the missing-handler error).

Approach:
  1. Walk control_inventory.json, collecting every (form, event)
     pair where the event name is a bare identifier (NOT
     `[Event Procedure]` or `=<expr>` or `*<macro>`).
  2. For each form, read its VBA dump and grep for `Sub <event>(`.
  3. Flag any reference whose target sub is missing.

False-positive sources:
  - Macros referenced by `*<MacroName>` — skip (we don't audit macros).
  - Calculated event handlers `=Function(args)` — skip.
"""
from __future__ import annotations

import json
import re
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CONTROL_INV = ROOT / "analysis" / "dump" / "control_inventory.json"
VBA_DIR = ROOT / "analysis" / "dump" / "vba"


SUB_DEF_RE = re.compile(
    r"^\s*(?:Private|Public)?\s*(?:Sub|Function)\s+([A-Za-z_]\w*)\s*\(",
    re.IGNORECASE | re.MULTILINE,
)


def _collect_subs(vba_path: Path) -> set[str]:
    if not vba_path.exists():
        return set()
    body = vba_path.read_bytes().decode("utf-8")
    return {m.group(1) for m in SUB_DEF_RE.finditer(body)}


def main() -> int:
    inventory = json.loads(CONTROL_INV.read_text(encoding="utf-8"))
    grand_total = 0
    print(f"=== Auditing event-handler resolution across "
          f"{len(inventory)} forms ===\n")

    for form_name, info in inventory.items():
        if not isinstance(info, dict):
            continue
        vba_path = VBA_DIR / f"Form_{form_name}.vb"
        if not vba_path.exists():
            # Sub-forms with spaces in names (e.g. "STATUS_DATA_2 Subform")
            # may be exported with the underscore-replaced name; skip.
            continue
        subs = _collect_subs(vba_path)
        unknown: list[tuple[str, str]] = []
        for c in info.get("controls", []):
            ev_list = c.get("events") or []
            for ev in ev_list:
                if not ev:
                    continue
                if ev.startswith("[") or ev.startswith("=") or ev.startswith("*"):
                    continue
                # Bare identifier — must exist as a Sub.
                if not re.match(r"^[A-Za-z_]\w*$", ev):
                    continue
                if ev not in subs:
                    unknown.append((c.get("name", "?"), ev))
        if unknown:
            grand_total += len(unknown)
            print(f"\n[FLAG] {form_name}:")
            for ctl_name, ev in unknown[:8]:
                print(f"  control {ctl_name!r} -> event handler "
                      f"{ev!r} (Sub not defined in module)")
            if len(unknown) > 8:
                print(f"  ... and {len(unknown) - 8} more")
    print(f"\n=== total flagged: {grand_total} ===")
    return 1 if grand_total else 0


if __name__ == "__main__":
    raise SystemExit(main())
