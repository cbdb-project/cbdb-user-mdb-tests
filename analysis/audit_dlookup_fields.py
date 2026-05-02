"""Static audit (15th scanner): for every `DLookup("<field>",
"<table>", ...)` (or `DFirst` / `DLast`) call with literal field +
table, verify `<field>` actually exists on `<table>`.

Bug shape: developer copies `DLookup("c_X", "FOO", ...)` from
another form, but `FOO` doesn't have `c_X` (renamed or never
existed).  At runtime DLookup returns Null silently — any
downstream check for that value evaluates to its zero-value default
(0 / "" / Null) and the form takes the wrong branch.

Companion to `audit_dcount_where_columns.py` which checks the
criteria string; this one checks the lookup-field arg.

Approach is tighter than DCount's because the field arg is
typically a single column name (not an expression), so we can be
strict.
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


# Match: D{Lookup,First,Last}("<field>", "<table>", ...) — literal
# field + literal table.  We accept either a bare column name or
# `[col]` brackets in the field arg.
DLOOKUP_RE = re.compile(
    r'\bD(?:Lookup|First|Last)\s*\(\s*'
    r'"\[?([A-Za-z_]\w*)\]?"\s*,\s*'
    r'"([A-Za-z_]\w*)"\s*[,)]',
    re.IGNORECASE,
)


def _load_schema() -> dict[str, set[str]]:
    out: dict[str, set[str]] = {}
    for entry in json.loads(TABLES_JSON.read_text(encoding="utf-8")):
        n = entry.get("name", "").upper()
        cols = entry.get("columns", []) or []
        out[n] = {
            (c.get("name") if isinstance(c, dict) else str(c)).lower()
            for c in cols
        }
    return out


def _scan(form_path: Path, schema: dict[str, set[str]]
          ) -> list[tuple[int, str, str]]:
    """Return (line_no, table, missing_field)."""
    lines = read_vba_lines(form_path)
    diagnostics: list[tuple[int, str, str]] = []
    seen: set[tuple[str, str]] = set()
    for ln_no, line in enumerate(lines, 1):
        stripped = line.lstrip()
        if stripped.startswith("'"):
            continue
        for m in DLOOKUP_RE.finditer(line):
            field = m.group(1).lower()
            table = m.group(2).upper()
            tbl_cols = schema.get(table)
            if tbl_cols is None:
                continue  # unknown table — covered elsewhere
            if field not in tbl_cols:
                key = (table, field)
                if key in seen:
                    continue
                seen.add(key)
                diagnostics.append((ln_no, table, field))
    return diagnostics


def main() -> int:
    schema = _load_schema()
    forms = sorted(VBA_DIR.glob("Form_*.vb"))
    grand_total = 0
    print(f"=== Auditing {len(forms)} VBA modules for "
          f"DLookup/DFirst/DLast with unknown field ===\n")
    for f in forms:
        diags = _scan(f, schema)
        if not diags:
            continue
        grand_total += len(diags)
        print(f"\n[FLAG] {f.name}:")
        for ln, table, field in diags[:8]:
            print(f"  line {ln}: D-lookup on {table}, field "
                  f"{field!r} not in schema")
    print(f"\n=== total flagged: {grand_total} ===")
    return 1 if grand_total else 0


if __name__ == "__main__":
    raise SystemExit(main())
