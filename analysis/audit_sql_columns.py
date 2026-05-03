"""Static audit: find every `<Table>.<column>` reference inside VBA
SQL string literals where the column doesn't exist on that table.

Catches the kind of silent SQL bug that *would* corrupt query
results (rather than crash with a visible error).  Bug #1 in
reports/CBDB_Issues_Report_EN.md (View_StatusData alias swap) is exactly this shape.

Conservative on purpose:
- Only flags `<TABLE>.<column>` where TABLE matches an actual table
  in `analysis/dump/tables.json` (skips aliases / saved-query views
  / linked subforms).
- VBA string concatenation across `& _` continuations is
  reconstructed before scanning.
- VBA + JET are case-insensitive; comparisons are too.
- Doesn't try to parse SQL.  If a column name appears with the
  table prefix in a comment, dump, or non-SQL context, we still flag
  it — false positives are easier to dismiss than missing bugs.
"""
from __future__ import annotations

import json
import re
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TABLES_JSON = ROOT / "analysis" / "dump" / "tables.json"
VBA_DIR = ROOT / "analysis" / "dump" / "vba"


def _load_schema() -> dict[str, set[str]]:
    """Return {TABLE_NAME_UPPER: set(column_names_upper)}."""
    out: dict[str, set[str]] = {}
    for entry in json.loads(TABLES_JSON.read_text(encoding="utf-8")):
        name = entry.get("name", "")
        cols = entry.get("columns", []) or []
        out[name.upper()] = {
            (c.get("name") if isinstance(c, dict) else str(c)).upper()
            for c in cols
        }
    return out


def _stitch_continuations(lines: list[str]) -> list[tuple[int, str]]:
    """VBA continues a logical line over physical lines via ` _\\n`.
    Return (line_number_of_first_physical_line, joined_logical_line)."""
    out: list[tuple[int, str]] = []
    buf: str = ""
    buf_start: int = 0
    for ln_no, line in enumerate(lines, 1):
        # Strip trailing whitespace.  VBA continuation marker is
        # whitespace + underscore at end of line.
        rstripped = line.rstrip()
        if rstripped.endswith(" _"):
            if not buf:
                buf_start = ln_no
            # Drop the trailing ` _` then concat
            buf += rstripped[:-2]
            continue
        if buf:
            buf += rstripped
            out.append((buf_start, buf))
            buf = ""
        else:
            out.append((ln_no, rstripped))
    if buf:
        out.append((buf_start, buf))
    return out


ND_PROPERTIES = {
    # Several CBDB scratch tables are also subform CONTROL names on
    # the LookAt forms (e.g. `ZZ_SCRATCH_STATUS` is both a table and
    # a subform control bound to it).  Subform property access like
    # `ZZ_SCRATCH_STATUS.Form.Recordset.RecordCount` looks like a
    # `<table>.<col>` reference to our regex but isn't.  Skip the
    # well-known non-column properties.
    "FORM", "RECORDSET", "RECORDSOURCE", "RECORDCOUNT",
    "VALUE", "CAPTION", "ENABLED", "VISIBLE", "TAG",
    "SOURCEOBJECT", "ROWSOURCE", "CONTROLSOURCE",
    "TIMERINTERVAL", "ONTIMER", "ONOPEN", "ONCLOSE",
    "BOOKMARK", "ABSOLUTEPOSITION",
}


def _scan(form_path: Path, schema: dict[str, set[str]]
          ) -> list[tuple[int, str, str, str]]:
    """Return list of (line_no, table, column, snippet) for table.col
    references in `form_path` whose column isn't in the table's
    schema."""
    from audit_lib import read_vba_lines
    raw = read_vba_lines(form_path)
    diagnostics: list[tuple[int, str, str, str]] = []
    seen_per_pair: dict[tuple[str, str], int] = {}

    # Stitch continuations so a SELECT statement split across 10 lines
    # is one logical line.  Then for each table, find every
    # `\b<TABLE>\.<col>\b` reference.  Case-insensitive.
    pat_cache: dict[str, re.Pattern] = {}
    for table in schema:
        # Match TABLE.column where column is a c_* / Z_* / x_* / etc.
        # identifier (CBDB columns are snake_case).
        pat_cache[table] = re.compile(
            rf"\b({re.escape(table)})\.([A-Za-z_][A-Za-z0-9_]*)\b",
            re.IGNORECASE,
        )

    for ln_no, logical in _stitch_continuations(raw):
        # Skip pure comment lines (VBA `'`).
        stripped = logical.lstrip()
        if stripped.startswith("'"):
            continue
        for table, cols in schema.items():
            pat = pat_cache[table]
            for m in pat.finditer(logical):
                col = m.group(2).upper()
                if col in cols:
                    continue
                if col in ND_PROPERTIES:
                    continue
                key = (table, col)
                seen_per_pair[key] = seen_per_pair.get(key, 0) + 1
                if seen_per_pair[key] > 3:
                    continue
                snippet = logical.strip()
                if len(snippet) > 140:
                    snippet = snippet[:140] + " …"
                diagnostics.append((ln_no, table, col, snippet))
    return diagnostics


def main() -> int:
    schema = _load_schema()
    print(f"=== Auditing {len(schema)} tables × every Form_LookAt*.vb ===\n")

    forms = sorted(VBA_DIR.glob("Form_LookAt*.vb"))
    grand_total = 0
    for form_path in forms:
        diags = _scan(form_path, schema)
        if not diags:
            print(f"[OK] {form_path.name} — no unknown column refs")
            continue
        grand_total += len(diags)
        # Group by (table, column) for readability.
        by_pair: dict[tuple[str, str], list[tuple[int, str]]] = defaultdict(list)
        for ln, t, c, snip in diags:
            by_pair[(t, c)].append((ln, snip))
        print(f"\n[FLAG] {form_path.name} — {len(by_pair)} unknown "
              f"<Table>.<col> pairs, {len(diags)} references:")
        for (t, c), occurrences in sorted(by_pair.items()):
            print(f"  {t}.{c.lower()}:")
            for ln, snip in occurrences[:3]:
                print(f"    line {ln:>5}: {snip}")
    print(f"\n=== total flagged: {grand_total} ===")
    return 1 if grand_total else 0


if __name__ == "__main__":
    raise SystemExit(main())
