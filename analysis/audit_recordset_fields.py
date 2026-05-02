"""Static audit: in every VBA module, track DAO recordsets opened on
a literal table name (`Set tRst = CurrentDb.OpenRecordset("FOO", ...)`)
and then check every `tRst!field` reference against FOO's schema.
Report mismatches.

Catches the silent bug shape "VBA reads `tRst!c_old_name` after the
column was renamed to `c_new_name` in the underlying table" — JET
raises 'Item not found in this collection' the moment that line
runs.

Conservative on purpose:
- Only tracks recordsets opened on a literal `"<TABLE>"` string.
  Recordsets opened on a runtime-built `tQueryStr` aren't analysed
  (we can't know their columns statically).
- Per-sub scope: variable -> table mapping is reset at every Sub
  boundary, so reassignments across subs don't leak.
- Same-sub variable reassignment overrides the earlier mapping.

Output: one per (file, sub, var, field) where field isn't on the
mapped table.
"""
from __future__ import annotations

import json
import re
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TABLES_JSON = ROOT / "analysis" / "dump" / "tables.json"
QUERIES_JSON = ROOT / "analysis" / "dump" / "queries.json"
VBA_DIR = ROOT / "analysis" / "dump" / "vba"


def _load_schema() -> dict[str, set[str]]:
    """Tables AND saved queries — both can be passed to
    `OpenRecordset()` and both have a column projection."""
    out: dict[str, set[str]] = {}
    for entry in json.loads(TABLES_JSON.read_text(encoding="utf-8")):
        n = entry.get("name", "")
        cols = entry.get("columns", []) or []
        out[n.upper()] = {
            (c.get("name") if isinstance(c, dict) else str(c)).upper()
            for c in cols
        }
    # For saved queries, parse SELECT projection columns.  We can be
    # lazy: if the query name shows up but we can't parse cols, treat
    # column lookup as "anything goes" (skip flagging) by mapping to
    # a sentinel.
    SENTINEL_ANY = frozenset({"__ANY__"})
    for entry in json.loads(QUERIES_JSON.read_text(encoding="utf-8")):
        n = entry.get("name", "")
        out.setdefault(n.upper(), SENTINEL_ANY)  # type: ignore
    return out


SUB_RE = re.compile(
    r"^\s*(?:Private|Public)?\s*Sub\s+([A-Za-z_]\w*)\s*\("
)
END_SUB_RE = re.compile(r"^\s*End\s+Sub\b", re.IGNORECASE)
# Any `Set <var> = ...` invalidates whatever <var> was bound to,
# whether or not the new RHS is a recognisable OpenRecordset.
ANY_SET_RE = re.compile(
    r"\bSet\s+([A-Za-z_]\w*)\s*=", re.IGNORECASE
)
OPEN_RS_RE = re.compile(
    r"\bSet\s+([A-Za-z_]\w*)\s*=\s*CurrentDb\.OpenRecordset\(\s*"
    r'"(?P<table>[A-Za-z_][A-Za-z0-9_]*)"\s*[,)]',
    re.IGNORECASE,
)
# `<var>!<field>` references.  Avoid `<form>!<subform>!<col>` by
# requiring the var to be a single identifier (handled by the
# capturing group).  VBA bracket-quoted: `<var>![<field>]` also.
BANG_RE = re.compile(
    r"\b([A-Za-z_]\w*)\!\s*(?:\[(?P<f1>[^\]]+)\]"
    r"|(?P<f2>[A-Za-z_][A-Za-z0-9_]*))"
)


def _scan(form_path: Path, schema: dict[str, set[str]]
          ) -> list[tuple[int, str, str, str, str]]:
    """Per-file diagnostics: (line_no, sub, var, field, table)."""
    from audit_lib import read_vba_lines
    raw = read_vba_lines(form_path)
    diagnostics: list[tuple[int, str, str, str, str]] = []

    cur_sub = ""
    var_to_table: dict[str, str] = {}
    seen_per: dict[tuple[str, str, str, str], int] = {}

    for ln_no, line in enumerate(raw, 1):
        stripped = line.lstrip()
        if stripped.startswith("'"):
            continue
        m = SUB_RE.match(line)
        if m:
            cur_sub = m.group(1)
            var_to_table = {}
            continue
        if END_SUB_RE.match(line):
            cur_sub = ""
            var_to_table = {}
            continue
        # Any `Set <var> = ...` invalidates the old binding first.
        for m in ANY_SET_RE.finditer(line):
            var_to_table.pop(m.group(1), None)
        # Then re-establish only if the RHS is a literal-table OpenRecordset.
        for m in OPEN_RS_RE.finditer(line):
            var_to_table[m.group(1)] = m.group("table").upper()
        # Check `<var>!<field>` refs
        for m in BANG_RE.finditer(line):
            var = m.group(1)
            field = (m.group("f1") or m.group("f2") or "").upper()
            table = var_to_table.get(var)
            if table is None:
                continue
            cols = schema.get(table)
            if cols is None:
                continue
            if "__ANY__" in cols:
                continue  # saved query: skip column check
            if field in cols:
                continue
            key = (cur_sub, var, table, field)
            seen_per[key] = seen_per.get(key, 0) + 1
            if seen_per[key] > 3:
                continue
            diagnostics.append((ln_no, cur_sub, var, field.lower(), table))
    return diagnostics


def main() -> int:
    schema = _load_schema()
    forms = sorted(VBA_DIR.glob("Form_*.vb"))
    grand_total = 0
    print(f"=== Auditing {len(forms)} VBA modules ===\n")
    for form_path in forms:
        diags = _scan(form_path, schema)
        if not diags:
            continue
        grand_total += len(diags)
        print(f"\n[FLAG] {form_path.name}:")
        by_loc: dict[tuple[str, str, str], list[tuple[int, str]]] = (
            defaultdict(list)
        )
        for ln, sub, var, field, table in diags:
            by_loc[(sub, var, table)].append((ln, field))
        for (sub, var, table), occurrences in sorted(by_loc.items()):
            fields = sorted({f for _, f in occurrences})
            lines = sorted({ln for ln, _ in occurrences})
            print(f"  in Sub {sub}: {var}!<field> on {table}")
            print(f"    unknown fields: {fields}")
            print(f"    first lines: {lines[:5]}")
    print(f"\n=== total flagged: {grand_total} ===")
    return 1 if grand_total else 0


if __name__ == "__main__":
    raise SystemExit(main())
