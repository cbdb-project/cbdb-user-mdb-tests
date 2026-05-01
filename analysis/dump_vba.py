"""
Extract VBA source for every Form / Report / Module / Class via VBE.
Writes ./dump/vba_modules.json with {component_name: {type, lines, code}}.

Run AFTER dump_metadata.py.
"""
from __future__ import annotations
import json, time, sys
from pathlib import Path
import win32com.client

ROOT = Path(__file__).resolve().parent.parent
USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
OUT = Path(__file__).resolve().parent / "dump" / "vba_modules.json"

# vbext_ComponentType
TYPE_NAME = {
    1: "Standard Module",
    2: "Class Module",
    3: "MSForm (UserForm)",
    11: "ActiveX Designer",
    100: "Document (Form/Report module)",  # acDocument
}


def main():
    print("opening Access...")
    app = win32com.client.Dispatch("Access.Application")
    app.Visible = False
    app.OpenCurrentDatabase(str(USER_MDB))

    out: dict = {}
    try:
        vbe = app.VBE
        for proj in vbe.VBProjects:
            print(f"  project: {proj.Name}")
            comps = proj.VBComponents
            print(f"  components: {comps.Count}")
            for i, comp in enumerate(comps, 1):
                name = comp.Name
                ctype = int(comp.Type)
                try:
                    cm = comp.CodeModule
                    n = int(cm.CountOfLines)
                    code = cm.Lines(1, n) if n > 0 else ""
                except Exception as e:
                    n = -1
                    code = f"ERROR: {e}"
                out[name] = {
                    "project": proj.Name,
                    "type": ctype,
                    "type_name": TYPE_NAME.get(ctype, f"Unknown({ctype})"),
                    "lines": n,
                    "code": code,
                }
                if i % 10 == 0 or i == comps.Count:
                    print(f"    [{i}/{comps.Count}] {name}  ({n} lines)")
    finally:
        try:
            app.CloseCurrentDatabase()
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass

    OUT.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\nwrote {OUT.name}: {len(out)} components, {OUT.stat().st_size:,} bytes")


if __name__ == "__main__":
    t0 = time.time()
    try:
        main()
    finally:
        # Clean up any lingering Access processes
        import subprocess
        subprocess.run(["taskkill", "/F", "/IM", "MSACCESS.EXE"],
                       capture_output=True)
    print(f"done in {time.time() - t0:.1f}s")
