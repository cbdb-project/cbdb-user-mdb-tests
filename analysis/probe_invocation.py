"""
Minimal probe: which invocation paths work after the user enabled
trusted content?  Try simplest first, no injection at all.
"""
from __future__ import annotations
import sys, time, subprocess
from pathlib import Path
import win32com.client

ROOT = Path(__file__).resolve().parent.parent
LOG = ROOT / "analysis" / "_probe.log"
LOG.unlink(missing_ok=True)

def log(*args):
    msg = " ".join(str(a) for a in args)
    LOG.open("a", encoding="utf-8").write(msg + "\n")

WORK = ROOT / "analysis" / "_test_work.mdb"

if not WORK.exists():
    log(f"working copy {WORK} missing — rerun debug_dispatcher.py first")
    raise SystemExit(1)

subprocess.run(["taskkill", "/F", "/IM", "MSACCESS.EXE"], capture_output=True)

log(f"opening: {WORK}")
app = win32com.client.Dispatch("Access.Application")
# msoAutomationSecurityLow=1 (enable macros), 2=ByUI (default, prompts), 3=ForceDisable
try:
    app.AutomationSecurity = 1
    log(f"set AutomationSecurity=1 (low)")
except Exception as e:
    log(f"AutomationSecurity set: {e}")
app.Visible = True
app.OpenCurrentDatabase(str(WORK))
log("opened")

# Quick: are macros enabled?  This API tells us trust state.
try:
    log(f"AutomationSecurity = {app.AutomationSecurity}")
except Exception as e:
    log(f"AutomationSecurity probe: {e}")

# Open LookAtEntry
log("=== open LookAtEntry ===")
app.DoCmd.OpenForm("LookAtEntry", 0, "", "", 0, 0)
loaded = bool(app.CurrentProject.AllForms("LookAtEntry").IsLoaded)
log(f"  loaded: {loaded}")

# wipe scratch so we can detect a real run
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

# Set form controls
def setctl(name, val):
    c = app.Forms("LookAtEntry").Controls(name)
    try: c.SetFocus()
    except Exception: pass
    c.Value = val
setctl("TxtFromYear", 900)
setctl("TxtToYear", 1100)
setctl("FrameYears", 2)
setctl("TxtEntryDesc", "yin privilege: general")
setctl("TxtTypeCode", "N/A")
log("controls set")

def count_scratch():
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM ZZ_SCRATCH_ENTRY")
    n = cur.fetchone()[0]
    cur.close()
    return int(n)

def wipe():
    cur = conn.cursor()
    cur.execute("DELETE FROM ZZ_SCRATCH_ENTRY")
    cur.close()

paths = [
    ("Application.Run('Form_LookAtEntry.CmdQuery_Click')",
     lambda: app.Run("Form_LookAtEntry.CmdQuery_Click")),
    ("app.Forms('LookAtEntry').CmdQuery_Click()",
     lambda: app.Forms("LookAtEntry").CmdQuery_Click()),
    ("DoCmd.RunCommand 11 (acCmdEnter) on focused button",
     lambda: (
         app.DoCmd.SelectObject(2, "LookAtEntry", False),
         app.Forms("LookAtEntry").Controls("CmdQuery").SetFocus(),
         app.SendKeys("{ENTER}", True),
     )),
]

for label, fn in paths:
    wipe()
    log(f"--- {label}")
    try:
        fn()
        # give Access a moment to actually run
        time.sleep(2)
        n = count_scratch()
        log(f"    OK -> ZZ_SCRATCH_ENTRY rows = {n}")
    except Exception as e:
        n = count_scratch()
        log(f"    FAIL ({type(e).__name__}): {e!s} ; rows={n}")

log("=== teardown ===")
try: conn.close()
except: pass
try: app.CloseCurrentDatabase()
except Exception as e: log(f" close db: {e}")
try: app.Quit()
except Exception as e: log(f" quit: {e}")
subprocess.run(["taskkill", "/F", "/IM", "MSACCESS.EXE"], capture_output=True)
log("done")
