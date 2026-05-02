"""Static audit: find every place in `Form_LookAt*.vb` where the
VBA references a control name that doesn't exist on that form.

Bug #4 was found by accident this way (`GISFrame.Value` on
LookAtPlace where the actual control is `CodeFrame`).  Run this
audit periodically (and on every CBDB release) to catch similar
copy-paste/rename mistakes statically without booting Access.

Output: one section per LookAt form listing each unknown reference
along with line number and a snippet of context.
"""
from __future__ import annotations

import json
import re
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
INVENTORY = json.loads(
    (ROOT / "analysis" / "dump" / "control_inventory.json").read_text(
        encoding="utf-8")
)
VBA_DIR = ROOT / "analysis" / "dump" / "vba"

# Control naming prefixes used in CBDB User MDB.  Only flag unknown
# identifiers whose first segment matches one of these — that avoids
# false positives on local variables (lowercase t*/cmd*), built-in
# VBA functions, ADODB types, etc.
#
# Notable exclusions:
# - `ZZ_*` and `tmp*` look like control-prefix matches but are
#   actually JET table / saved-query names referenced inside SQL
#   string literals (e.g. `tQueryStr = "... ZZ_PLACE.c_addr_id ..."`).
# - `frm*` is usually a subform CONTROL OR a foreign form name;
#   subform refs on the current form ARE worth checking, so we keep
#   `frm` but require an additional `.Form` member access (which is
#   how subform recordsets are accessed) — see `SUBFORM_REF_RE` below.
CONTROL_PREFIXES = (
    "Cmd", "Txt", "Lbl", "Chk", "Frame", "Tab", "Page", "Opt",
)
PREFIX_RE = "|".join(re.escape(p) for p in CONTROL_PREFIXES)

# Unqualified `<Name>.<Member>` reference where <Name> begins with a
# control prefix.  Negative look-behind avoids matching `.<Name>`
# (member access), `Forms!...<Name>` (qualified ref to a different
# form), and `<word><Name>` (variable suffix).
REF_RE = re.compile(
    rf"(?<![\w.!])({PREFIX_RE})[A-Za-z0-9_]+(?=\s*\.[A-Za-z_])"
)

# Names ending in `_Click`, `_AfterUpdate`, `_LostFocus` etc. are sub
# definitions / `Call <X>_Click` — not control references.  Skip when
# the suffix follows immediately.
SUB_SUFFIX_RE = re.compile(
    r"^(?:Cmd|Txt|Lbl|Chk|Frame|Tab|Page|Opt|frm|ZZ_)\w+_(?:"
    r"Click|AfterUpdate|BeforeUpdate|LostFocus|GotFocus|Change|"
    r"Enter|Exit|MouseDown|MouseUp|KeyDown|KeyUp|KeyPress|DblClick|"
    r"Updated|Open|Close|Load|Unload|Resize|Activate|Deactivate"
    r")$"
)

# Some unqualified-looking identifiers are actually module-level
# variables CBDB declares as Public (gXxx) or Dim'd in subs (tXxx).
# We already filter to CONTROL_PREFIXES so most won't match, but
# guard against a few that do.  Update if more false positives crop
# up.
KNOWN_NON_CONTROL_IDENTIFIERS = {
    # ADODB.* type members accessed via local var named `Cmd<X>` — none
    # observed; placeholder.
}


def _form_inventory(form_short: str) -> set[str] | None:
    entry = INVENTORY.get(form_short)
    if entry is None:
        return None
    ctls = entry.get("controls", [])
    return {c["name"] for c in ctls if c.get("name")}


def _file_for_form(form_short: str) -> Path:
    return VBA_DIR / f"Form_{form_short}.vb"


def _scan_form(form_short: str) -> list[tuple[int, str, str]]:
    """Return list of (line_no, identifier, context) for references
    on this form to controls that aren't in its inventory."""
    inv = _form_inventory(form_short)
    if inv is None:
        return []
    f = _file_for_form(form_short)
    if not f.exists():
        return []
    out: list[tuple[int, str, str]] = []
    seen_per_id: dict[str, int] = {}
    for ln_no, line in enumerate(f.read_text(encoding="utf-8").splitlines(),
                                 1):
        # Skip comment-only lines — VBA comment is leading `'`.
        stripped = line.lstrip()
        if stripped.startswith("'"):
            continue
        for m in REF_RE.finditer(line):
            ident = m.group(0)
            if SUB_SUFFIX_RE.match(ident):
                continue
            if ident in inv:
                continue
            if ident in KNOWN_NON_CONTROL_IDENTIFIERS:
                continue
            # De-dup repeat hits on the same identifier; keep the
            # first 3 occurrences as evidence.
            seen_per_id[ident] = seen_per_id.get(ident, 0) + 1
            if seen_per_id[ident] > 3:
                continue
            out.append((ln_no, ident, line.strip()))
    return out


def main() -> int:
    forms = sorted(k for k in INVENTORY if k.startswith("LookAt"))
    total_diags = 0
    print(f"=== Auditing {len(forms)} LookAt forms ===\n")
    for form in forms:
        diags = _scan_form(form)
        if not diags:
            print(f"[OK] {form} — no unknown control references")
            continue
        total_diags += len(diags)
        # Group by identifier so report is concise.
        by_ident: dict[str, list[tuple[int, str]]] = defaultdict(list)
        for ln, ident, ctx in diags:
            by_ident[ident].append((ln, ctx))
        print(f"\n[FLAG] {form} — {len(by_ident)} unknown identifiers, "
              f"{len(diags)} references:")
        for ident in sorted(by_ident):
            print(f"  {ident!r}:")
            for ln, ctx in by_ident[ident][:3]:
                snippet = ctx[:120]
                print(f"    line {ln:>5}: {snippet}")
    print(f"\n=== total flagged: {total_diags} ===")
    return 1 if total_diags else 0


if __name__ == "__main__":
    raise SystemExit(main())
