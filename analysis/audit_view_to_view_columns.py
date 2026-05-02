"""Static audit (20th scanner): for saved queries that reference
*other* saved queries in their FROM/JOIN, check column references
against the referenced view's SELECT projection.

Bug shape: View_A is built on top of View_B via
    SELECT View_B.c_x FROM View_B
but View_B's projection no longer includes c_x (renamed / dropped).
View_A breaks at runtime, every form that reads View_A breaks
silently with it.

`audit_saved_queries.py` already checks `<RealTable>.<col>` refs
against tables.json.  This audit fills the missing leg: view → view.
"""
from __future__ import annotations

import json
import re
from collections import defaultdict
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "analysis"))

QUERIES_JSON = ROOT / "analysis" / "dump" / "queries.json"


SELECT_RE = re.compile(
    r"\bSELECT\s+(?:DISTINCT\s+|TOP\s+\d+\s+)*(.+?)\s+FROM\s",
    re.IGNORECASE | re.DOTALL,
)
TABLE_COL_RE = re.compile(
    r"\b([A-Za-z_]\w*)\.\[?([A-Za-z_]\w*)\]?"
)
ALIAS_RE = re.compile(
    r"\b(?:\[(?P<src>[^\]]+)\]|(?P<src2>[A-Za-z_][A-Za-z0-9_]*))\s+"
    r"AS\s+(?:\[(?P<dst>[^\]]+)\]|(?P<dst2>[A-Za-z_][A-Za-z0-9_]*))",
    re.IGNORECASE,
)


def _parse_projection(sql: str) -> set[str] | None:
    """Lowercase column-name / alias set, or {'__any__'} for wildcard,
    or None if SELECT can't be parsed."""
    m = SELECT_RE.search(sql)
    if not m:
        return None
    proj = m.group(1)
    cols: set[str] = set()
    depth = 0
    cur: list[str] = []
    parts: list[str] = []
    for ch in proj:
        if ch == "(":
            depth += 1; cur.append(ch)
        elif ch == ")":
            depth -= 1; cur.append(ch)
        elif ch == "," and depth == 0:
            parts.append("".join(cur).strip())
            cur = []
        else:
            cur.append(ch)
    if cur:
        parts.append("".join(cur).strip())
    for p in parts:
        if not p:
            continue
        if p == "*" or p.endswith(".*"):
            return {"__any__"}
        m_alias = re.search(r"\bAS\s+([A-Za-z_]\w*)\s*$", p, re.IGNORECASE)
        if m_alias:
            cols.add(m_alias.group(1).lower())
            continue
        if "." in p:
            tail = p.rsplit(".", 1)[1].strip()
        else:
            tail = p.strip()
        tail = tail.strip("[]")
        if not re.match(r"^[A-Za-z_]\w*$", tail):
            continue
        cols.add(tail.lower())
    return cols


def main() -> int:
    queries = json.loads(QUERIES_JSON.read_text(encoding="utf-8"))
    # name (lower) -> projection
    proj: dict[str, set[str]] = {}
    sql_by_name: dict[str, str] = {}
    for q in queries:
        name = (q.get("name") or "")
        sql = q.get("sql") or ""
        if name and sql:
            sql_by_name[name.lower()] = sql
            p = _parse_projection(sql)
            if p is not None:
                proj[name.lower()] = p

    grand_total = 0
    print(f"=== Auditing {len(sql_by_name)} saved queries for "
          f"view-to-view column refs ===\n")
    for qname, sql in sql_by_name.items():
        # Skip aliases — focus on bare `<View_X>.<col>` references.
        # Strip declared aliases from consideration.
        aliases = set()
        for m in ALIAS_RE.finditer(sql):
            a = (m.group("dst") or m.group("dst2") or "").lower()
            if a:
                aliases.add(a)
        seen: set[tuple[str, str]] = set()
        flags: list[tuple[str, str]] = []
        for m in TABLE_COL_RE.finditer(sql):
            tgt = m.group(1)
            col = m.group(2).lower()
            if tgt.lower() == qname:
                continue  # self-reference (won't happen but safe)
            if tgt.lower() in aliases:
                continue
            tgt_proj = proj.get(tgt.lower())
            if tgt_proj is None:
                continue  # not a saved query — covered elsewhere
            if "__any__" in tgt_proj or col in tgt_proj:
                continue
            key = (tgt, col)
            if key in seen:
                continue
            seen.add(key)
            flags.append((tgt, col))
        if flags:
            grand_total += len(flags)
            print(f"\n[FLAG] {qname}:")
            for tgt, col in flags[:5]:
                print(f"  refs {tgt}.{col} — col not in {tgt}'s "
                      f"SELECT projection")
    print(f"\n=== total flagged: {grand_total} ===")
    return 1 if grand_total else 0


if __name__ == "__main__":
    raise SystemExit(main())
