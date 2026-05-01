"""Categorise tables by linked vs local, and show counts."""
import json
from pathlib import Path

DUMP = Path(__file__).resolve().parent / "dump"
tables = json.loads((DUMP / "tables.json").read_text(encoding="utf-8"))

zz = []        # ZZ_SCRATCH_* working/output tables
codes = []     # *_CODES lookup tables
data = []      # *_DATA event tables
biog = []      # BIOG_* core tables
other = []
for t in tables:
    n = t["name"]
    if n.startswith("ZZ"):
        zz.append(t)
    elif n.endswith("_CODES") or n in {"DYNASTIES", "NIAN_HAO", "GANZHI_CODES"}:
        codes.append(t)
    elif n.endswith("_DATA"):
        data.append(t)
    elif n.startswith("BIOG_"):
        biog.append(t)
    else:
        other.append(t)

print(f"Total tables: {len(tables)}\n")
for label, group in (("ZZ_SCRATCH (working)", zz), ("CODES (lookups)", codes),
                     ("DATA (events)", data), ("BIOG_ (people)", biog),
                     ("OTHER", other)):
    print(f"--- {label}: {len(group)} ---")
    for t in group:
        rc = t["record_count"]
        ncol = len(t["columns"])
        print(f"  {t['name']:<40} cols={ncol:<4} rows={rc:>10}")
    print()
