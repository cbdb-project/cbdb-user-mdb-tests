"""Write each form's VBA to its own .vb file under ./dump/vba/ for easy reading."""
import json, sys, re
from pathlib import Path

DUMP = Path(__file__).resolve().parent / "dump"
OUT = DUMP / "vba"
OUT.mkdir(exist_ok=True)
mods = json.loads((DUMP / "vba_modules.json").read_text(encoding="utf-8"))

for name, info in mods.items():
    code = info.get("code") or ""
    if not code or info.get("lines", 0) <= 2:
        continue
    safe = re.sub(r"[^A-Za-z0-9_\-]", "_", name)
    p = OUT / f"{safe}.vb"
    p.write_text(code, encoding="utf-8")
print(f"wrote {len(list(OUT.glob('*.vb')))} files to {OUT}")
