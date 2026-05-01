"""Probe against the ORIGINAL untouched CBDB_BJ_User.mdb (no injection)
to determine which invocation path works when nothing has been
modified."""
from __future__ import annotations
import sys, time, subprocess, shutil
from pathlib import Path
import win32com.client

ROOT = Path(__file__).resolve().parent.parent
LOG = ROOT / "analysis" / "_probe.log"
LOG.unlink(missing_ok=True)

def log(*args):
    msg = " ".join(str(a) for a in args)
    LOG.open("a", encoding="utf-8").write(msg + "\n")

SRC = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK = ROOT / "analysis" / "_pristine.mdb"

subprocess.run(["taskkill", "/F", "/IM", "MSACCESS.EXE"], capture_output=True)
if WORK.exists(): WORK.unlink()
shutil.copy2(SRC, WORK)
log(f"copied to: {WORK}")

app = win32com.client.Dispatch("Access.Application")
try:
    app.AutomationSecurity = 1
    log("AutomationSecurity = 1 (low / enable all)")
except Exception as e:
    log(f"AutomationSecurity failed: {e}")
app.Visible = True
log("opening db...")
app.OpenCurrentDatabase(str(WORK))
log("db open")

# fix DAO ref
proj = app.VBE.VBProjects(1)
for r in list(proj.References):
    if r.IsBroken:
        full = getattr(r, "FullPath", "") or ""
        log(f"removing broken ref: {full}")
        proj.References.Remove(r)
        if "dao" in full.lower():
            for cand in (
                r"C:\Program Files\Microsoft Office\root\Office16\ACEDAO.DLL",
                r"C:\Program Files (x86)\Microsoft Office\root\Office16\ACEDAO.DLL",
            ):
                if Path(cand).exists():
                    proj.References.AddFromFile(cand)
                    log(f"added: {cand}")
                    break

log("opening LookAtEntry...")
app.DoCmd.OpenForm("LookAtEntry", 0, "", "", 0, 0)
log(f"  loaded: {bool(app.CurrentProject.AllForms('LookAtEntry').IsLoaded)}")

# Probe paths against the PRISTINE form (private CmdQuery_Click)
import pyodbc
conn = pyodbc.connect(
    "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    f"DBQ={WORK};", autocommit=True
)
cur = conn.cursor()
cur.execute("DELETE FROM ZZ_SCRATCH_ENTRY")
cur.execute("DELETE FROM ZZ_SCRATCH_ENTRY_CODE")
cur.execute("INSERT INTO ZZ_SCRATCH_ENTRY_CODE (c_entry_code) VALUES (118)")
cur.execute("DELETE FROM ZZ_SCRATCH_ADDR")
cur.execute("INSERT INTO ZZ_SCRATCH_ADDR (c_addr_id) VALUES (100658)")
cur.close()

def setctl(name, val):
    c = app.Forms("LookAtEntry").Controls(name)
    try: c.SetFocus()
    except Exception: pass
    c.Value = val
setctl("TxtFromYear", 900); setctl("TxtToYear", 1100); setctl("FrameYears", 2)
setctl("TxtEntryDesc", "yin privilege: general"); setctl("TxtTypeCode", "N/A")

def count():
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM ZZ_SCRATCH_ENTRY")
    n = cur.fetchone()[0]; cur.close(); return int(n)

def wipe():
    cur = conn.cursor()
    cur.execute("DELETE FROM ZZ_SCRATCH_ENTRY")
    cur.close()

paths = [
    ("Application.Run('Form_LookAtEntry.CmdQuery_Click')",
     lambda: app.Run("Form_LookAtEntry.CmdQuery_Click")),
    ("app.Forms('LookAtEntry').CmdQuery_Click()",
     lambda: app.Forms("LookAtEntry").CmdQuery_Click()),
    ("DoCmd.SelectObject + GoToControl + SendKeys ENTER",
     lambda: (
         app.DoCmd.SelectObject(2, "LookAtEntry", False),
         app.DoCmd.GoToControl("CmdQuery"),
         app.SendKeys("{ENTER}", True),
     )),
    ("RunCommand 11 (acCmdEnter)",
     lambda: (
         app.DoCmd.SelectObject(2, "LookAtEntry", False),
         app.DoCmd.GoToControl("CmdQuery"),
         app.RunCommand(11),
     )),
]

for label, fn in paths:
    wipe()
    log(f"--- {label}")
    try:
        fn()
        time.sleep(2)
        n = count()
        log(f"    OK -> rows={n}")
    except Exception as e:
        n = count()
        log(f"    FAIL ({type(e).__name__}): {e!s} ; rows={n}")

log("=== teardown ===")
try: conn.close()
except: pass
try: app.CloseCurrentDatabase()
except Exception as e: log(f" close: {e}")
try: app.Quit()
except Exception as e: log(f" quit: {e}")
subprocess.run(["taskkill", "/F", "/IM", "MSACCESS.EXE"], capture_output=True)
log("done")
