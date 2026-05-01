"""Print a concise inventory of forms (top-level only)."""
import json
from pathlib import Path

DUMP = Path(__file__).resolve().parent / "dump"
forms = json.loads((DUMP / "forms.json").read_text(encoding="utf-8"))

print(f"Total top-level forms: {len(forms)}\n")
for f in forms:
    name = f.get("name", "?")
    rs = (f.get("properties") or {}).get("RecordSource") or ""
    nctl = len(f.get("controls", []))
    nlines = f.get("code_lines") or 0
    err = f.get("error", "") or f.get("code_error", "")
    flag = " ERR" if err else ""
    print(f"  {name:<50} controls={nctl:<4} code_lines={nlines:<5} src={rs[:60]!r}{flag}")
