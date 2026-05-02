"""Static audit (12th scanner): for every `DCount` / `DLookup` /
`DSum` / `DAvg` / `DMax` / `DMin` (the VBA "domain aggregate"
family) call with a literal table+criteria, verify that columns
referenced in the criteria string actually exist on the named
table.

Bug shape: developer copies `DCount("*", "ZZ_FOO", "c_old_field > 0")`
from another form, but `ZZ_FOO` doesn't have `c_old_field` (renamed
or never existed).  At runtime DCount returns Null silently — any
`If DCount(...) > 0 Then` check evaluates to False (because Null is
not > 0), the form takes the wrong branch, and the user sees
"unexpectedly empty" output without an error.

Approach:
  1. Find every `D<Func>(<expr>, "<table>", "<criteria>")` pattern
     in VBA, with literal table and literal criteria string.
  2. Strip surface VBA from criteria (literal '...' strings,
     # date #, leading ()s) and extract bare identifier tokens.
  3. Filter to tokens that look like CBDB column names
     (`c_<word>` / `x_<word>` / `y_<word>` etc.).
  4. Look up each in the table's schema; flag mismatches.

Caveats:
  - Criteria may reference VBA variables via `'" & x & "'` etc. —
    skip flagging tokens inside dynamic concatenations (we look for
    pure-literal criteria strings only).
  - SQL functions (`Trim`, `Nz`, `IsNull`, ...) are filtered.
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


# Match: D<Func>(<expr>, "<table>", "<criteria>") with literal table
# and criteria.  Conservative — non-literal args produce no match.
DAGG_RE = re.compile(
    r'\bD(?:Count|Lookup|Sum|Avg|Max|Min|First|Last)\s*\(\s*'
    r'(?:"[^"]*"|\*|[A-Za-z_]\w*(?:\([^)]*\))?)'
    r'\s*,\s*"([A-Za-z_]\w*)"\s*,\s*"([^"]*)"\s*\)',
    re.IGNORECASE,
)

# Common SQL / VBA tokens to skip when extracting column names from
# criteria.
SQL_KEYWORDS = {
    "AND", "OR", "NOT", "TRUE", "FALSE", "NULL", "IS",
    "BETWEEN", "LIKE", "IN", "NULL", "TRIM", "NZ", "ISNULL",
    "STR", "CSTR", "CLNG", "CDBL", "CINT", "INT", "ABS",
    "LEFT", "RIGHT", "MID", "UCASE", "LCASE", "LEN", "VAL",
    "DATE", "NOW", "YEAR", "MONTH", "DAY", "TIME",
    "ALL", "ANY", "EXISTS", "DISTINCT",
}


def _extract_criteria_columns(criteria: str) -> set[str]:
    """Pull out column-like identifiers from a SQL WHERE criteria."""
    # Strip inline string literals 'foo' (not for VBA but for SQL).
    no_strings = re.sub(r"'[^']*'", "''", criteria)
    # Strip date literals #...#.
    no_dates = re.sub(r"#[^#]*#", "##", no_strings)
    # Pull out bare identifiers.
    cols: set[str] = set()
    for m in re.finditer(r"\b([A-Za-z_][A-Za-z_0-9]*)\b", no_dates):
        tok = m.group(1)
        if tok.upper() in SQL_KEYWORDS:
            continue
        # Only consider ones that look like CBDB column names
        # (c_*, x_*, y_*, z_*) — skip standalone numbers / VBA vars.
        if re.match(r"^[cxyzn]_", tok, re.IGNORECASE):
            cols.add(tok.lower())
    return cols


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
          ) -> list[tuple[int, str, str, str]]:
    """Return (line_no, table, criteria, unknown_col)."""
    lines = read_vba_lines(form_path)
    diagnostics: list[tuple[int, str, str, str]] = []
    seen: set[tuple[str, str, str]] = set()
    for ln_no, line in enumerate(lines, 1):
        stripped = line.lstrip()
        if stripped.startswith("'"):
            continue
        for m in DAGG_RE.finditer(line):
            table = m.group(1).upper()
            criteria = m.group(2)
            tbl_cols = schema.get(table)
            if tbl_cols is None:
                continue  # table itself unknown — covered by audit_sql_table_names
            criteria_cols = _extract_criteria_columns(criteria)
            for col in criteria_cols:
                if col not in tbl_cols:
                    key = (table, col, criteria)
                    if key in seen:
                        continue
                    seen.add(key)
                    diagnostics.append((ln_no, table, criteria, col))
    return diagnostics


def main() -> int:
    schema = _load_schema()
    forms = sorted(VBA_DIR.glob("Form_*.vb"))
    grand_total = 0
    print(f"=== Auditing {len(forms)} VBA modules for "
          f"DCount/DLookup criteria with unknown columns ===\n")
    for f in forms:
        diags = _scan(f, schema)
        if not diags:
            continue
        grand_total += len(diags)
        print(f"\n[FLAG] {f.name}:")
        for ln, table, criteria, col in diags[:10]:
            print(f"  line {ln}: D-agg on {table}, criteria "
                  f"references unknown col {col!r}")
            print(f"    full criteria: {criteria!r}")
    print(f"\n=== total flagged: {grand_total} ===")
    return 1 if grand_total else 0


if __name__ == "__main__":
    raise SystemExit(main())
