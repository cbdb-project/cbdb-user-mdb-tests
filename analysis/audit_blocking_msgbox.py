"""Static audit (17th scanner): find every `MsgBox(...) = vb<Yes|No|
OK|Cancel>` confirmation prompt that blocks the COM thread when
running CBDB under automation.

These aren't bugs *for end users* — they're explicit yes/no
confirmations the user is supposed to answer.  But they DO block
the test driver (which has no UI) and require manual neutralization
in `tests/cbdb_driver/vba_session.py::_inject_autodetect`.  This
audit lists every one so we can keep the neutralizer list complete
when CBDB releases new code.

The driver's existing `_msgbox_replace` only neutralizes
*statement-form* `MsgBox "literal"` calls; *function-form*
`If MsgBox(...) = vbYes Then` are intentionally NOT touched (they
gate user choices that tests pre-arrange via direct table writes).
But we should still know where they are.

Output is informational (exit 0).
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

# Match `MsgBox(...)` — function-call form (parenthesized args), as
# opposed to `MsgBox "literal"` statement form.
# We look for any line that has `MsgBox(` followed by a `vbYes` /
# `vbNo` / `vbOK` / `vbCancel` comparison nearby.
MSGBOX_FN_RE = re.compile(
    r"\bMsgBox\s*\(",
    re.IGNORECASE,
)
COMPARE_RE = re.compile(
    r"\bvb(Yes|No|OK|Cancel|Ignore|Retry|Abort)\b",
    re.IGNORECASE,
)


def _scan(form_path: Path) -> list[tuple[int, str]]:
    """Return (line_no, snippet)."""
    lines = read_vba_lines(form_path)
    diagnostics: list[tuple[int, str]] = []
    for ln_no, line in enumerate(lines, 1):
        stripped = line.lstrip()
        if stripped.startswith("'"):
            continue
        if not MSGBOX_FN_RE.search(line):
            continue
        # Look for vb* comparison on the same line (most CBDB prompts
        # use one-liner `If MsgBox(...) = vbYes Then`).
        if COMPARE_RE.search(line):
            snippet = line.strip()
            if len(snippet) > 100:
                snippet = snippet[:97] + "..."
            diagnostics.append((ln_no, snippet))
    return diagnostics


def main() -> int:
    forms = sorted(p for p in VBA_DIR.glob("Form_*.vb")
                    if "TMPCLP" not in p.name.upper())
    grand_total = 0
    print(f"=== Auditing {len(forms)} VBA modules for "
          f"blocking MsgBox(...) = vb* prompts ===\n")
    for f in forms:
        diags = _scan(f)
        if not diags:
            continue
        grand_total += len(diags)
        print(f"\n[INFO] {f.name}: {len(diags)} blocking prompt(s)")
        for ln, snippet in diags[:5]:
            print(f"  line {ln}: {snippet}")
        if len(diags) > 5:
            print(f"  ... and {len(diags) - 5} more")
    print(f"\n=== total blocking prompts: {grand_total} "
          f"(informational — see _msgbox_replace docstring) ===")
    return 0  # Informational, not a failure.


if __name__ == "__main__":
    raise SystemExit(main())
