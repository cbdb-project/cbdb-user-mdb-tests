"""Static audit (19th scanner): for every assignment of a literal
table or query name to a sub-form's RecordSource (e.g.
`Me!frmFoo.Form.RecordSource = "View_Bar"`), verify that the
target name resolves to either a real table or a saved query.

Bug shape: a saved query was renamed (View_Bar → View_Baz) but
some VBA still assigns the old name.  At runtime the sub-form
silently shows nothing.

Looks at any assignment of the form
  `<...>.RecordSource = "<name>"`
where `<name>` is a bare identifier that should be in the schema.
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
TABLES_JSON = ROOT / "analysis" / "dump" / "tables.json"
QUERIES_JSON = ROOT / "analysis" / "dump" / "queries.json"


# Match `.RecordSource = "<name>"` (bare identifier — skip SQL
# strings starting with SELECT etc.)
RECORD_SOURCE_RE = re.compile(
    r"\.RecordSource\s*=\s*\"([A-Za-z_][A-Za-z_0-9]*)\"",
    re.IGNORECASE,
)


def _load_known() -> set[str]:
    """All table + saved-query names, lowercased."""
    out: set[str] = set()
    for entry in json.loads(TABLES_JSON.read_text(encoding="utf-8")):
        out.add((entry.get("name") or "").lower())
    for q in json.loads(QUERIES_JSON.read_text(encoding="utf-8")):
        out.add((q.get("name") or "").lower())
    out.discard("")
    return out


def _scan(form_path: Path, known: set[str]) -> list[tuple[int, str]]:
    lines = read_vba_lines(form_path)
    diagnostics: list[tuple[int, str]] = []
    seen: set[str] = set()
    for ln_no, line in enumerate(lines, 1):
        stripped = line.lstrip()
        if stripped.startswith("'"):
            continue
        for m in RECORD_SOURCE_RE.finditer(line):
            name = m.group(1)
            if name.lower() in seen:
                continue
            seen.add(name.lower())
            if name.lower() not in known:
                diagnostics.append((ln_no, name))
    return diagnostics


def main() -> int:
    known = _load_known()
    forms = sorted(p for p in VBA_DIR.glob("Form_*.vb")
                    if "TMPCLP" not in p.name.upper())
    grand_total = 0
    print(f"=== Auditing {len(forms)} VBA modules for "
          f".RecordSource = \"<unknown>\" assignments ===\n")
    for f in forms:
        diags = _scan(f, known)
        if not diags:
            continue
        grand_total += len(diags)
        print(f"\n[FLAG] {f.name}:")
        for ln, name in diags[:5]:
            print(f"  line {ln}: .RecordSource = \"{name}\" — "
                  f"not a known table or saved query")
    print(f"\n=== total flagged: {grand_total} ===")
    return 1 if grand_total else 0


if __name__ == "__main__":
    raise SystemExit(main())
