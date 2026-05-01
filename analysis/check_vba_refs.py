"""List VBA project references to find broken ones."""
import shutil, sys, io, subprocess
from pathlib import Path
import win32com.client

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "data" / "CBDB_BJ_User.mdb"
COPY = ROOT / "analysis" / "_test_copy.mdb"
if not COPY.exists():
    shutil.copy2(SRC, COPY)

app = win32com.client.Dispatch("Access.Application")
app.Visible = False
app.OpenCurrentDatabase(str(COPY))
try:
    proj = app.VBE.VBProjects(1)
    print(f"Project: {proj.Name}")
    print("\n=== before fix ===")
    for r in proj.References:
        print(f"  {r.Name:<25} broken={r.IsBroken}  fullpath={getattr(r, 'FullPath', '?')}")

    # Try to fix: drop any broken reference, then add the canonical replacement.
    fixed = []
    to_remove = []
    for r in proj.References:
        if r.IsBroken:
            to_remove.append(r)
    for r in to_remove:
        full = getattr(r, "FullPath", "") or ""
        proj.References.Remove(r)
        # add canonical DAO from current Office install
        if "dao" in full.lower():
            for cand in (
                r"C:\Program Files\Common Files\Microsoft Shared\OFFICE16\ACEDAO.DLL",
                r"C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE16\ACEDAO.DLL",
            ):
                if Path(cand).exists():
                    try:
                        proj.References.AddFromFile(cand)
                        fixed.append(("DAO", cand))
                        break
                    except Exception as e:
                        print(f"  add {cand} failed: {e}")
    # save the fix into the file
    try:
        app.DoCmd.RunCommand(126)  # acCmdCompileAllModules — will surface remaining errors
    except Exception as e:
        print(f"  compile result: {e}")

    print("\n=== after fix ===")
    for r in proj.References:
        print(f"  {r.Name:<25} broken={r.IsBroken}  fullpath={getattr(r, 'FullPath', '?')}")
    print(f"\nadded: {fixed}")
finally:
    app.CloseCurrentDatabase()
    app.Quit()
    subprocess.run(["taskkill", "/F", "/IM", "MSACCESS.EXE"], capture_output=True)
