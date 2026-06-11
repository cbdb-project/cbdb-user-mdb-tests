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


def _access_pids() -> set[int]:
    """Currently-running MSACCESS.EXE PIDs (empty set if psutil unavailable)."""
    try:
        import psutil
        return {p.pid for p in psutil.process_iter(["name"])
                if (p.info["name"] or "").upper() == "MSACCESS.EXE"}
    except Exception:
        return set()


def main():
    print("opening Access...")
    # DispatchEx (NOT Dispatch): force a FRESH out-of-process Access instance.
    # Dispatch can bind to a developer's already-running Access via the ROT,
    # and the OpenCurrentDatabase/CloseCurrentDatabase/Quit below would then
    # commandeer + close THEIR window.  DispatchEx + the scoped PID-kill in
    # __main__ keep this script's Access entirely its own (cf. landmine #9).
    app = win32com.client.DispatchEx("Access.Application")
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
    # Scope-kill ONLY Access instances that appeared DURING this run (in case
    # Quit() leaves one lingering) -- never a window already open before it.  A blanket
    # `taskkill /IM MSACCESS.EXE` (the old behaviour) would violate the
    # scoped-kill discipline the harness keeps (conftest.pytest_sessionfinish),
    # which matters now that this runs as a hard step in run_tests.ps1 (Step 5a).
    pre_pids = _access_pids()
    try:
        main()
    finally:
        import subprocess
        for pid in _access_pids() - pre_pids:
            subprocess.run(["taskkill", "/F", "/PID", str(pid)],
                           capture_output=True)
    print(f"done in {time.time() - t0:.1f}s")
