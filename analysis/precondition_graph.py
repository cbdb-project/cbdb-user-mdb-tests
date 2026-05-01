"""
For each form, extract the dependency graph:
  WHICH event handler enables/disables WHICH control?

Example output line:
  LookAtEntry.CmdQuery_Click  ENABLES  CmdGIS, CmdNeo4j, CmdStoreID  IF  RecordCount > 0

The intent is to surface every (precondition, dependent-control) pair
so that the test plan can:
  1. Test the precondition is enforced (control is disabled when it should be)
  2. Test the dependent action works once enabled
  3. Test that triggering a disabled control raises an error or no-ops
"""
from __future__ import annotations
import json, re
from collections import defaultdict
from pathlib import Path

DUMP = Path(__file__).resolve().parent / "dump"
vba = json.loads((DUMP / "vba_modules.json").read_text(encoding="utf-8"))

LOOKAT = [
    "LookAtEntry", "LookAtKinship", "LookAtOffice", "LookAtPlace",
    "LookAtAssociations", "LookAtAssociationPairs", "LookAtStatus",
    "LookAtTexts", "LookAtNetworks", "LookAtGroupData",
]

ENABLE_RX = re.compile(r"(\w+)\.Enabled\s*=\s*(True|False)", re.IGNORECASE)
SUB_RX = re.compile(
    r"^(?:Private|Public)?\s*Sub\s+(\w+)\s*\(.*?$",
    re.MULTILINE | re.IGNORECASE,
)
END_SUB_RX = re.compile(r"^End Sub\s*$", re.MULTILINE | re.IGNORECASE)


def parse_subs(code: str) -> list[tuple[str, int, int, str]]:
    """Return [(sub_name, start_line, end_line, body), ...]."""
    out = []
    lines = code.splitlines()
    i = 0
    while i < len(lines):
        m = SUB_RX.match(lines[i])
        if not m:
            i += 1
            continue
        name = m.group(1)
        start = i
        j = i + 1
        while j < len(lines) and not END_SUB_RX.match(lines[j]):
            j += 1
        body = "\n".join(lines[start:j+1])
        out.append((name, start+1, j+1, body))
        i = j + 1
    return out


lines = ["# Precondition graph (.Enabled flips per event handler)", ""]

for form in LOOKAT:
    code = vba.get(f"Form_{form}", {}).get("code", "")
    if not code:
        continue
    subs = parse_subs(code)
    flips = defaultdict(lambda: defaultdict(set))  # sub -> ctrl -> {True,False}
    for name, _, _, body in subs:
        for m in ENABLE_RX.finditer(body):
            ctrl = m.group(1)
            val = m.group(2).capitalize()
            flips[name][ctrl].add(val)
    if not flips:
        continue
    lines.append(f"## {form}\n")
    lines.append("| Event handler | Enables | Disables | Toggles |")
    lines.append("|---|---|---|---|")
    for sub, controls in sorted(flips.items()):
        ena, dis, tog = [], [], []
        for c, vals in sorted(controls.items()):
            if vals == {"True"}: ena.append(c)
            elif vals == {"False"}: dis.append(c)
            else: tog.append(c)
        lines.append(f"| `{sub}` | {', '.join(ena) or '—'} | "
                     f"{', '.join(dis) or '—'} | {', '.join(tog) or '—'} |")
    lines.append("")

(DUMP / "precondition_graph.md").write_text("\n".join(lines), encoding="utf-8")
print(f"wrote precondition_graph.md")
