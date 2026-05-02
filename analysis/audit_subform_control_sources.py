"""Static audit (10th scanner): for every sub-form whose RecordSource
is a saved query (View_*), verify that every bound control's
ControlSource exists in the saved query's SELECT projection.

Catches the bug shape: saved query gets refactored (column renamed
or dropped) but the sub-form design wasn't updated.  At runtime the
control silently shows blank or "#Name?" and the user sees missing
data without an error.

Sources:
  - `analysis/dump/control_inventory.json` — every form's
    record_source + per-control control_source.
  - `analysis/dump/queries.json` — every saved query's SQL.

Approach:
  1. Build {view_name -> set(projected_columns)} from queries.json
     using the same SELECT-projection parser the recordset auditor
     uses.
  2. Walk every form.  Skip those whose record_source isn't a
     `View_*` saved query (linked tables / scratch tables aren't
     candidates — they have a static schema that audit_sql_columns
     already catches).
  3. For each control with a non-empty control_source, flag if the
     column isn't in the view's projection.

Caveats / known false-positive sources:
  - Wildcard projections (`SELECT *` or `SELECT t.*`) — treated as
    "any column allowed", skip flagging.
  - Aliased projections (`AS c_alias`) — DAO exposes the alias as
    the field name; the parser handles AS aliases.
  - Calculated controls (`=Sum([c_x])`) — control_source starts
    with `=`, skip flagging (not a bare column reference).
"""
from __future__ import annotations

import json
import re
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CONTROL_INV = ROOT / "analysis" / "dump" / "control_inventory.json"
QUERIES_JSON = ROOT / "analysis" / "dump" / "queries.json"


SELECT_RE = re.compile(
    r"\bSELECT\s+(?:DISTINCT\s+|TOP\s+\d+\s+)*(.+?)\s+FROM\s",
    re.IGNORECASE | re.DOTALL,
)


def _parse_projection(sql: str) -> set[str] | None:
    """Same projection parser as audit_recordset_sql_projection.
    Returns lowercased column-name / alias set, or {'__ANY__'} for
    wildcard, or None if unparseable."""
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
            return {"__ANY__"}
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
    view_proj: dict[str, set[str]] = {}
    for q in queries:
        name = q.get("name", "")
        sql = q.get("sql", "")
        if not name or not sql:
            continue
        proj = _parse_projection(sql)
        if proj is not None:
            view_proj[name] = proj

    inventory = json.loads(CONTROL_INV.read_text(encoding="utf-8"))
    grand_total = 0
    print(f"=== Auditing {len(inventory)} forms × "
          f"{len(view_proj)} parsed view projections ===\n")
    for form_name, info in inventory.items():
        if not isinstance(info, dict):
            continue
        rs = info.get("record_source", "") or ""
        if not rs.startswith("View_"):
            continue
        proj = view_proj.get(rs)
        if proj is None:
            print(f"  [WARN] {form_name}: cannot parse projection of {rs!r}")
            continue
        if "__ANY__" in proj:
            continue
        unknown: list[tuple[str, str]] = []
        for c in info.get("controls", []):
            cs = (c.get("control_source", "") or "").strip()
            if not cs or cs.startswith("="):
                continue
            cs_low = cs.lower().strip("[]")
            if cs_low not in proj:
                unknown.append((c.get("name", "?"), cs))
        if unknown:
            grand_total += len(unknown)
            print(f"\n[FLAG] {form_name} (rs={rs!r}):")
            for ctl_name, cs in unknown[:8]:
                print(f"  control {ctl_name!r} -> control_source={cs!r}"
                      f"  (not in projection)")
            if len(unknown) > 8:
                print(f"  ... and {len(unknown) - 8} more")
    print(f"\n=== total flagged: {grand_total} ===")
    return 1 if grand_total else 0


if __name__ == "__main__":
    raise SystemExit(main())
