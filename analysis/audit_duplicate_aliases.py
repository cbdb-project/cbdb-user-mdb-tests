"""Find SELECT lists with duplicate output aliases — those will silently shadow columns."""
import json, re
from pathlib import Path

DUMP = Path(__file__).resolve().parent / "dump"
queries = json.loads((DUMP / "queries.json").read_text(encoding="utf-8"))

# crude: grab tokens "AS alias" inside the SELECT clause (before FROM)
ALIAS_RX = re.compile(r"\bAS\s+(\w+)", re.IGNORECASE)

problems = []
for q in queries:
    sql = q["sql"]
    name = q["name"]
    # take SELECT … FROM
    m = re.search(r"SELECT\s+(.*?)\s+FROM\b", sql, re.IGNORECASE | re.DOTALL)
    if not m:
        continue
    select_clause = m.group(1)
    aliases = [a.lower() for a in ALIAS_RX.findall(select_clause)]
    # also count bare-column outputs (Foo.Bar with no alias -> output is Bar)
    for col in re.findall(r"\b\w+\.([A-Za-z_]\w*)\b", select_clause):
        # only count if not followed by AS in same item
        pass
    seen = {}
    for a in aliases:
        seen[a] = seen.get(a, 0) + 1
    dups = [(a, n) for a, n in seen.items() if n > 1]
    if dups:
        problems.append((name, dups))

if not problems:
    print("No duplicate aliases found.")
else:
    print(f"Found duplicate aliases in {len(problems)} queries:\n")
    for n, dups in problems:
        print(f"  [{n}]")
        for a, c in dups:
            print(f"     - {a!r}: {c}x")
