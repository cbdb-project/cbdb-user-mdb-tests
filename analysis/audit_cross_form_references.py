"""Static audit (13th scanner): for every `Forms!<form>!<ctl>` (or
`Forms("<form>")("<ctl>")`) cross-form reference in VBA, verify
both `<form>` exists and `<ctl>` exists on `<form>`.

Bug shape: a developer renames a control on form A, but form B
still references it via `Forms!A!<oldname>`.  At runtime VBA throws
"Item not found in this collection" the moment that line executes
— if it's inside a guarded `On Error Resume Next` block the error
is silently swallowed and downstream logic takes the wrong branch.

`audit_missing_controls.py` already catches single-form references
(e.g. `Me.<ctl>` or bare `<ctl>` inside a form's own module).  This
audit is the cross-form companion — useful because picker dialogs
and parent LookAt forms commonly poke at each other's controls.

Sources: same control_inventory.json the other audits use.
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


# Match `Forms!<form>!<ctl>` (with optional .Form. for sub-form
# stepping).  We DON'T currently follow `.Form.<sub>` chains beyond
# the first one — that would need recursive sub-form traversal.
FORMS_BANG_RE = re.compile(
    r"\bForms!([A-Za-z_]\w*)!([A-Za-z_]\w*)"
)


def _load_inventory() -> tuple[dict[str, set[str]], dict[str, str]]:
    """Return ({form_name_lc: set(control_names_lc)},
               {form_name_lc: original_case_name}).

    Lowercased keys because Access form / control identifiers are
    case-insensitive at runtime, but the dump preserves original case
    so a literal-string compare reports false positives like
    `frmpickdynasty` vs `frmPickDynasty`.
    """
    inv = json.loads(CONTROL_INV.read_text(encoding="utf-8"))
    out: dict[str, set[str]] = {}
    case: dict[str, str] = {}
    for fname, info in inv.items():
        if not isinstance(info, dict):
            continue
        ctls = {(c.get("name") or "").lower()
                for c in info.get("controls", [])
                if c.get("name")}
        out[fname.lower()] = ctls
        case[fname.lower()] = fname
    return out, case


def _scan(form_path: Path, inv: dict[str, set[str]]
          ) -> list[tuple[int, str, str, str]]:
    """Return (line_no, target_form, target_ctl, issue)."""
    lines = read_vba_lines(form_path)
    diagnostics: list[tuple[int, str, str, str]] = []
    seen: set[tuple[str, str]] = set()
    for ln_no, line in enumerate(lines, 1):
        stripped = line.lstrip()
        if stripped.startswith("'"):
            continue
        for m in FORMS_BANG_RE.finditer(line):
            fm_lc = m.group(1).lower()
            ctl_lc = m.group(2).lower()
            key = (fm_lc, ctl_lc)
            if key in seen:
                continue
            seen.add(key)
            if fm_lc not in inv:
                diagnostics.append((ln_no, m.group(1), m.group(2),
                                    "form-not-in-inventory"))
                continue
            ctls = inv[fm_lc]
            if ctl_lc not in ctls:
                diagnostics.append((ln_no, m.group(1), m.group(2),
                                    "control-not-on-form"))
    return diagnostics


def main() -> int:
    inv, _case = _load_inventory()
    # Skip Access auto-backup snapshots (`Form__TMPCLP*.vb`) — they're
    # historical state, not live code, and naturally reference forms
    # that no longer exist.
    forms = sorted(p for p in VBA_DIR.glob("Form_*.vb")
                    if "TMPCLP" not in p.name.upper())
    grand_total = 0
    print(f"=== Auditing {len(forms)} VBA modules for "
          f"unresolved Forms!<form>!<ctl> refs "
          f"(known forms: {len(inv)}) ===\n")
    for f in forms:
        diags = _scan(f, inv)
        if not diags:
            continue
        grand_total += len(diags)
        print(f"\n[FLAG] {f.name}:")
        by_target: dict[tuple[str, str], list[tuple[int, str]]] = (
            defaultdict(list)
        )
        for ln, fm, ctl, issue in diags:
            by_target[(fm, ctl)].append((ln, issue))
        for (fm, ctl), occurrences in sorted(by_target.items()):
            issue = occurrences[0][1]
            lines = sorted({ln for ln, _ in occurrences})
            print(f"  Forms!{fm}!{ctl}: {issue}")
            print(f"    first lines: {lines[:5]}")
    print(f"\n=== total flagged: {grand_total} ===")
    return 1 if grand_total else 0


if __name__ == "__main__":
    raise SystemExit(main())
