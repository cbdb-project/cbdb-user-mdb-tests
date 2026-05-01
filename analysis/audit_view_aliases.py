"""
For each saved query, check whether SELECT column aliases that look like
fy_*/ly_* / first/last actually map to the correct YEAR_RANGE_CODES alias.

Heuristic:
  - In FROM clause, find which YEAR_RANGE_CODES alias was joined to which
    *_range_code field (e.g. c_fy_range vs c_ly_range vs c_by_range...).
  - In SELECT list, find every "<YEAR_RANGE_CODES alias>.c_range[_chn] AS <output>"
    and check that the output prefix matches the joined source's prefix.
"""
import json, re
from pathlib import Path

DUMP = Path(__file__).resolve().parent / "dump"
queries = json.loads((DUMP / "queries.json").read_text(encoding="utf-8"))

ALIAS = re.compile(r"YEAR_RANGE_CODES(?:\s+AS\s+(YEAR_RANGE_CODES_\d+))?\s+ON\s+\S+\.c_(\w+?)_range\s*=\s*(\S+?)\.c_range_code", re.IGNORECASE)
SELECT_USE = re.compile(r"(YEAR_RANGE_CODES_?\d?)\.c_range(_chn)?\s+AS\s+c_(\w+?)_range_(desc|chn)", re.IGNORECASE)

issues = []
for q in queries:
    sql = q["sql"]
    name = q["name"]
    # Find joins
    join_map = {}  # alias -> prefix it joined for (fy/ly/by/dy/...)
    for m in ALIAS.finditer(sql):
        # m.group(1) might be None (the unaliased YEAR_RANGE_CODES)
        joined_alias = m.group(1) or "YEAR_RANGE_CODES"
        prefix = m.group(2).lower()
        join_map[joined_alias] = prefix
    # Find SELECT alias usages
    for m in SELECT_USE.finditer(sql):
        used_alias = m.group(1)
        out_prefix = m.group(3).lower()
        out_kind = m.group(4).lower()
        joined_for = join_map.get(used_alias)
        if joined_for is None:
            issues.append((name, f"output c_{out_prefix}_range_{out_kind} pulls from {used_alias} but no join found for that alias"))
        elif joined_for != out_prefix:
            issues.append((name, f"output c_{out_prefix}_range_{out_kind} pulls from {used_alias} which was joined for c_{joined_for}_range (mismatch)"))

if not issues:
    print("No alias-mismatch issues found.")
else:
    print(f"Found {len(issues)} potential alias-mismatch issues:\n")
    by_q = {}
    for n, msg in issues:
        by_q.setdefault(n, []).append(msg)
    for n, msgs in by_q.items():
        print(f"  [{n}]")
        for m in msgs:
            print(f"     - {m}")
