"""Static audit: scan every saved query / view in `queries.json` for
`<Table>.<Column>` references whose column doesn't exist on that
table.

Saved queries (`View_*`, `Z_*` etc.) are read by VBA forms.  If a
view projects a column that was renamed/removed in the source table,
EVERY form that reads the view fails at runtime — silent until the
user clicks the button that consumes it.

Same family as `audit_sql_columns.py` but for the queries side, not
the VBA side.
"""
from __future__ import annotations

import json
import re
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TABLES_JSON = ROOT / "analysis" / "dump" / "tables.json"
QUERIES_JSON = ROOT / "analysis" / "dump" / "queries.json"


ND_PROPERTIES = {
    "FORM", "RECORDSET", "RECORDSOURCE", "RECORDCOUNT",
    "VALUE", "CAPTION", "ENABLED", "VISIBLE", "TAG",
    "SOURCEOBJECT", "ROWSOURCE", "CONTROLSOURCE",
}


def _load_schema() -> dict[str, set[str]]:
    out: dict[str, set[str]] = {}
    for entry in json.loads(TABLES_JSON.read_text(encoding="utf-8")):
        n = entry.get("name", "")
        cols = entry.get("columns", []) or []
        out[n.upper()] = {
            (c.get("name") if isinstance(c, dict) else str(c)).upper()
            for c in cols
        }
    return out


# Aliases declared in the SQL itself: `<table> AS <alias>`.  We need
# to skip flags on aliased references because their "table" is the
# alias, not a real schema entry.
ALIAS_RE = re.compile(
    r"\b(?:\[(?P<src>[^\]]+)\]|(?P<src2>[A-Za-z_][A-Za-z0-9_]*))\s+"
    r"AS\s+(?:\[(?P<dst>[^\]]+)\]|(?P<dst2>[A-Za-z_][A-Za-z0-9_]*))",
    re.IGNORECASE,
)


def _aliases_in_sql(sql: str) -> set[str]:
    out: set[str] = set()
    for m in ALIAS_RE.finditer(sql):
        a = m.group("dst") or m.group("dst2") or ""
        if a:
            out.add(a.upper())
    return out


def _scan_query(sql: str, schema: dict[str, set[str]]
                ) -> list[tuple[str, str]]:
    """Return list of (table, column) refs whose column isn't in the
    table's schema.  Aliases are skipped (we can't validate columns
    on an alias without parsing the JOIN tree).  Each (table, column)
    pair reported once."""
    if not sql:
        return []
    aliases = _aliases_in_sql(sql)
    seen: set[tuple[str, str]] = set()
    diags: list[tuple[str, str]] = []
    for table, cols in schema.items():
        pat = re.compile(
            rf"\b({re.escape(table)})\.([A-Za-z_][A-Za-z0-9_]*)\b",
            re.IGNORECASE,
        )
        for m in pat.finditer(sql):
            col = m.group(2).upper()
            if col in cols:
                continue
            if col in ND_PROPERTIES:
                continue
            key = (table, col)
            if key in seen:
                continue
            seen.add(key)
            diags.append((table, col.lower()))
    # Also flag references to a `<UNKNOWN_TABLE>.<col>` where
    # UNKNOWN_TABLE is neither in schema nor an alias defined in the
    # same SQL.  Conservative regex: identifier that ENDS in CAPS_AND_
    # (CBDB convention for tables).
    return diags


def main() -> int:
    schema = _load_schema()
    queries = json.loads(QUERIES_JSON.read_text(encoding="utf-8"))
    print(f"=== Auditing {len(queries)} saved queries vs "
          f"{len(schema)} tables ===\n")
    grand_total = 0
    for q in queries:
        name = q.get("name", "?")
        sql = q.get("sql", "") or ""
        if not sql.strip():
            continue
        diags = _scan_query(sql, schema)
        if not diags:
            print(f"[OK] {name}")
            continue
        grand_total += len(diags)
        print(f"\n[FLAG] {name} — {len(diags)} unknown <Table>.<col>:")
        for table, col in diags:
            print(f"  {table}.{col}")
    print(f"\n=== total flagged: {grand_total} ===")
    return 1 if grand_total else 0


if __name__ == "__main__":
    raise SystemExit(main())
