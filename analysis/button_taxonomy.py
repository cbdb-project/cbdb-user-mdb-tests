"""
Categorise every CommandButton + CheckBox + OptionGroup across the 10
main LookAt forms by behavioural pattern.

Buckets (by name and handler-body shape):
  query    - kicks off the main SELECT (CmdQuery / CmdRunQuery)
  export   - writes a file (CmdGIS, CmdNeo4j, writeKML, CmdSaveXxx)
  import   - reads from a file
  picker   - opens a modal picker form
  state    - toggles a public flag / wipes selection (CmdAll*, CmdClear*)
  display  - language / Fanti / Jianti
  nav      - exit / help
  unknown  - falls through

Output: analysis/dump/button_taxonomy.md
"""
from __future__ import annotations
import json, re
from collections import defaultdict
from pathlib import Path

DUMP = Path(__file__).resolve().parent / "dump"
inv = json.loads((DUMP / "control_inventory.json").read_text(encoding="utf-8"))
vba = json.loads((DUMP / "vba_modules.json").read_text(encoding="utf-8"))

LOOKAT = [
    "LookAtEntry", "LookAtKinship", "LookAtOffice", "LookAtPlace",
    "LookAtAssociations", "LookAtAssociationPairs", "LookAtStatus",
    "LookAtTexts", "LookAtNetworks", "LookAtGroupData",
]


def handler_body(form: str, handler: str) -> str:
    code = vba.get(f"Form_{form}", {}).get("code", "")
    if not code:
        return ""
    rx = re.compile(
        rf"^(Private|Public)?\s*(Sub|Function)\s+{re.escape(handler)}\s*\(.*?\n(.*?)^End \2",
        re.MULTILINE | re.DOTALL | re.IGNORECASE,
    )
    m = rx.search(code)
    return m.group(3) if m else ""


def classify(name: str, body: str) -> str:
    n = name.lower()
    b = body.lower()
    if "openform" in b and ("acdialog" in b or "frmpick" in b):
        return "picker"
    if any(x in n for x in ("gis", "kml", "neo4j", "savefile", "save_to", "export")):
        return "export"
    if "filedialog" in b and ("save" in b or "msofiledialogsaveas" in b):
        return "export"
    if "filedialog" in b and "open" in b:
        return "import"
    if any(x in n for x in ("import",)):
        return "import"
    if name.startswith("CmdQuery") or "cmdrunquery" in n:
        return "query"
    if any(x in n for x in ("cmdall", "cmdclear", "cmdreset")):
        return "state"
    if any(x in n for x in ("cmdfanti", "cmdjianti", "language")):
        return "display"
    if any(x in n for x in ("cmdexit", "cmdhelp", "cmdcancel", "cmdok")):
        return "nav"
    if name.startswith("Chk"):
        return "checkbox"
    if name.startswith("Frame"):
        return "frame"
    return "unknown"


lines = ["# Button / control taxonomy across LookAt forms", ""]
totals = defaultdict(int)

for form in LOOKAT:
    info = inv.get(form, {})
    if not info:
        continue
    lines.append(f"## {form}\n")
    lines.append("| Control | Type | Bucket | Handlers |")
    lines.append("|---|---|---|---|")
    seen = set()
    rows = []
    for c in info["controls"]:
        events = c["events"]
        # collect all click-like + change handlers
        if not events:
            continue
        # bucket by primary handler (Click/AfterUpdate/Change)
        primary = None
        for e in events:
            if e.endswith("_Click") or e.endswith("_AfterUpdate") or e.endswith("_Change"):
                primary = e
                break
        if not primary:
            primary = events[0]
        body = handler_body(form, primary)
        bucket = classify(c["name"], body)
        totals[bucket] += 1
        rows.append((c["name"], c["type"], bucket, ", ".join(events)))
    for r in sorted(rows, key=lambda x: (x[2], x[0])):
        lines.append(f"| `{r[0]}` | {r[1]} | **{r[2]}** | {r[3]} |")
    lines.append("")

lines.append("## Totals across LookAt forms\n")
for b, n in sorted(totals.items(), key=lambda x: -x[1]):
    lines.append(f"- {b}: {n}")

(DUMP / "button_taxonomy.md").write_text("\n".join(lines), encoding="utf-8")
print(f"wrote button_taxonomy.md ({len(totals)} buckets, {sum(totals.values())} controls)")
