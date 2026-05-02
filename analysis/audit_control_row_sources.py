"""Static audit (18th scanner): for every control with a non-empty
RowSource SQL (ListBox / ComboBox dropdowns), parse it and verify
every `<Table>.<Column>` reference in the SELECT projection (and
JOIN / WHERE clauses) names a column that exists in the schema.

Bug shape: a renamed column in BIOG_ADDR_CODES / ENTRY_CODES /
similar lookup tables breaks every dropdown bound to it.  The
ComboBox displays an empty list, the user sees no options.

Companion to:
  - audit_sql_columns.py — VBA-embedded SQL strings
  - audit_saved_queries.py — saved query SQL
This is the third leg of the SQL-column-resolution stool: the
control-design-time SQL.
"""
from __future__ import annotations

import json
import re
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CONTROL_INV = ROOT / "analysis" / "dump" / "control_inventory.json"
TABLES_JSON = ROOT / "analysis" / "dump" / "tables.json"
QUERIES_JSON = ROOT / "analysis" / "dump" / "queries.json"


# `<Table>.<Column>` reference (handles `[col]` brackets too).
TABLE_COL_RE = re.compile(
    r"\b([A-Za-z_]\w*)\.\[?([A-Za-z_]\w*)\]?"
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
    # Saved-query results — treat as `__ANY__` (we don't recurse).
    for q in json.loads(QUERIES_JSON.read_text(encoding="utf-8")):
        n = (q.get("name") or "").upper()
        out.setdefault(n, {"__any__"})
    return out


def _scan_sql(sql: str, schema: dict[str, set[str]]
              ) -> list[tuple[str, str]]:
    """Return list of (table, missing_column)."""
    seen: set[tuple[str, str]] = set()
    out: list[tuple[str, str]] = []
    # Strip string literals (single-quoted) and date literals (#...#)
    # so we don't match a literal string's content.
    no_strings = re.sub(r"'[^']*'", "''", sql)
    no_dates = re.sub(r"#[^#]*#", "##", no_strings)
    for m in TABLE_COL_RE.finditer(no_dates):
        table = m.group(1).upper()
        col = m.group(2).lower()
        if table in {"AS"}:  # weird tokens
            continue
        cols = schema.get(table)
        if cols is None:
            continue  # unknown table — covered by audit_sql_table_names
        if "__any__" in cols:
            continue
        if col in cols:
            continue
        key = (table, col)
        if key in seen:
            continue
        seen.add(key)
        out.append(key)
    return out


def main() -> int:
    schema = _load_schema()
    inventory = json.loads(CONTROL_INV.read_text(encoding="utf-8"))
    grand_total = 0
    n_with_rs = 0
    print(f"=== Auditing controls' RowSource SQL "
          f"across {len(inventory)} forms ===\n")
    for form_name, info in inventory.items():
        if not isinstance(info, dict):
            continue
        for c in info.get("controls", []):
            rs = (c.get("row_source", "") or "").strip()
            if not rs:
                continue
            n_with_rs += 1
            mismatches = _scan_sql(rs, schema)
            if mismatches:
                grand_total += len(mismatches)
                print(f"\n[FLAG] {form_name} / {c.get('name')}:")
                for table, col in mismatches[:5]:
                    print(f"  RowSource references {table}.{col} "
                          f"— column not in schema")
    print(f"\n=== {n_with_rs} controls have RowSource SQL; "
          f"{grand_total} unknown <Table>.<Col> refs ===")
    return 1 if grand_total else 0


if __name__ == "__main__":
    raise SystemExit(main())
