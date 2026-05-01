"""Just open Office, set CmdQuery.Enabled=True via COM, read back."""
import sys, time
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))
from cbdb_driver.vba_session import VbaSession

LIVE = open(r"C:\Users\how612\AppData\Local\Temp\probeoff_enabled.log",
            "w", encoding="utf-8", buffering=1)
def log(*a):
    s = " ".join(str(x) for x in a)
    LIVE.write(s + "\n"); LIVE.flush()
    print(s)

SRC = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK = ROOT / "analysis" / "_btn_enabled.mdb"

s = VbaSession(SRC, WORK).open()
s.open_form("LookAtOffice")
time.sleep(1)
ctl = s.app.Forms("LookAtOffice").Controls("CmdQuery")
log(f"CmdQuery.Enabled before: {ctl.Enabled}")
ctl.Enabled = True
log(f"CmdQuery.Enabled after  : {ctl.Enabled}")
# Try clicking the form's CmdQuery via COM (Click method on a Control)
log("Try ctl.SetFocus + Application.RunCommand acCmdCopy as no-op:")
try:
    ctl.SetFocus()
    log("  SetFocus OK")
except Exception as e:
    log(f"  SetFocus failed: {e}")
# What about invoking the underlying VBA event?
# Try via DoCmd.RunMacro... no macro defined.
# Try Eval():
try:
    log("Eval Forms!LookAtOffice.CmdQuery:")
    v = s.app.Eval("Forms!LookAtOffice!CmdQuery")
    log(f"  -> {v}")
except Exception as e:
    log(f"  Eval err: {e}")
# Try .Properties("Enabled"):
try:
    log(f"  Properties('Enabled') = {ctl.Properties('Enabled').Value}")
except Exception as e:
    log(f"  Properties err: {e}")

s.close()
LIVE.close()
