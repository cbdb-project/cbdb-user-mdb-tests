"""Summarize VBA dump."""
import json, sys, re
from pathlib import Path

DUMP = Path(__file__).resolve().parent / "dump"
mods = json.loads((DUMP / "vba_modules.json").read_text(encoding="utf-8"))

print(f"Total components: {len(mods)}\n")

categories = {"Form_": [], "Standard": [], "Class": [], "Other": []}
for name, info in mods.items():
    code = info.get("code") or ""
    nl = info.get("lines", 0)
    if name.startswith("Form_"):
        categories["Form_"].append((name, nl, code))
    elif info["type"] == 1:
        categories["Standard"].append((name, nl, code))
    elif info["type"] == 2:
        categories["Class"].append((name, nl, code))
    else:
        categories["Other"].append((name, nl, code))

for cat, items in categories.items():
    print(f"--- {cat} ({len(items)}) ---")
    for name, nl, _ in sorted(items, key=lambda x: -x[1]):
        print(f"  {name:<55} {nl:>6} lines")
    print()
