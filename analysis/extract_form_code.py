"""Write each form's VBA to its own .vb file under ./dump/vba/ for easy reading."""
import json, sys, re
from pathlib import Path

DUMP = Path(__file__).resolve().parent / "dump"
OUT = DUMP / "vba"
OUT.mkdir(exist_ok=True)
mods = json.loads((DUMP / "vba_modules.json").read_text(encoding="utf-8"))

written: set[str] = set()
for name, info in mods.items():
    code = info.get("code") or ""
    if not code or info.get("lines", 0) <= 2:
        continue
    safe = re.sub(r"[^A-Za-z0-9_\-]", "_", name)
    p = OUT / f"{safe}.vb"
    p.write_text(code, encoding="utf-8")
    written.add(p.name)
# Remove stale .vb files for forms no longer in this build's dump -- otherwise
# a renamed/removed form leaves a stale file that the static audits, reverify,
# and the vba_ref line-citation audit would read as if it were current source.
removed = 0
for f in OUT.glob("*.vb"):
    if f.name not in written:
        f.unlink()
        removed += 1
print(f"wrote {len(written)} files (removed {removed} stale) to {OUT}")
