"""Quick: print the raw JSON for one form to understand what fields are populated."""
import json, sys
from pathlib import Path
DUMP = Path(__file__).resolve().parent / "dump"
forms = json.loads((DUMP / "forms.json").read_text(encoding="utf-8"))
target = sys.argv[1] if len(sys.argv) > 1 else "LookAtEntry"
for f in forms:
    if f.get("name") == target:
        print(json.dumps({k: v for k, v in f.items() if k != "controls"}, ensure_ascii=False, indent=2))
        print("---")
        print(f"controls count: {len(f.get('controls', []))}")
        print("first 5 controls:")
        for c in f.get("controls", [])[:5]:
            print(json.dumps(c, ensure_ascii=False))
        sys.exit(0)
print(f"form {target!r} not found")
