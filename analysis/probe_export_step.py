"""Step-by-step probe of LookAtEntry CmdGIS export."""
import sys, time
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))
from cbdb_driver.vba_session import VbaSession

LIVE = open(r"C:\Users\how612\AppData\Local\Temp\probe_export.log",
            "w", encoding="utf-8", buffering=1)
def log(*a):
    s = " ".join(str(x) for x in a)
    LIVE.write(s + "\n"); LIVE.flush()
    print(s)

SRC = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK = ROOT / "analysis" / "_export_probe.mdb"

s = VbaSession(SRC, WORK).open()
log("Open form, populate pickers (entry_code 36 only, no addr)")
s.open_form("LookAtEntry")
s.set_picker_codes("ZZ_SCRATCH_ENTRY_CODE", [36], column="c_entry_code")

log("Patch FileDialog")
s.patch_filedialog("LookAtEntry")

log("Read ChkKML, GISFrame defaults")
f = s.app.Forms("LookAtEntry")
try: log(f"  ChkKML.Value = {f.Controls('ChkKML').Value!r}")
except Exception as e: log(f"  ChkKML err: {e}")
try: log(f"  GISFrame.Value = {f.Controls('GISFrame').Value!r}")
except Exception as e: log(f"  GISFrame err: {e}")

log("Set export path BEFORE triggering chain")
out = ROOT / "analysis" / "_exports_tmp" / "probe_gis.tab"
out.parent.mkdir(parents=True, exist_ok=True)
if out.exists(): out.unlink()
log(f"  export_path = {out}")

log("Trigger CmdQuery + CmdGIS as a single timer chain")
s.click_chain_via_timer(
    "LookAtEntry", ["CmdQuery", "CmdGIS"],
    export_path=str(out),
)

log("Wait 30s + check file")
for i in range(30):
    time.sleep(1)
    if out.exists():
        log(f"  [{i}s] file exists, size={out.stat().st_size}")
        break
else:
    log("  file never appeared in 30s")

log("Final ZZ_TEST_DEBUG:")
try:
    df = s.read("ZZ_TEST_DEBUG", order_by="id")
    for r in df.itertuples():
        log(f"  [{r.id}] {r.msg}")
    if df.empty:
        log("  (empty)")
except Exception as e:
    log(f"  read err: {e}")

log("Check Access processes for visible MsgBox or other dialog")
import subprocess as _sp
try:
    r = _sp.run(["powershell","-NoProfile","-Command",
                 "Get-Process MSACCESS -ErrorAction SilentlyContinue | "
                 "Select-Object Id, MainWindowTitle | Format-Table -AutoSize"],
                capture_output=True, text=True, timeout=10)
    log(r.stdout)
except Exception as e:
    log(f"  ps err: {e}")

log("Dump CmdQuery_Click body (post-autodetect-inject)")
cm = s.app.VBE.VBProjects(1).VBComponents("Form_LookAtEntry").CodeModule
body = cm.Lines(1, cm.CountOfLines)
in_q = False
for i, line in enumerate(body.splitlines(), 1):
    if "Sub CmdQuery_Click(" in line:
        in_q = True
    if in_q:
        log(f"  {i}: {line}")
        if line.strip() == "End Sub":
            break

log("Dump Form_Timer body (post-injection)")
cm = s.app.VBE.VBProjects(1).VBComponents("Form_LookAtEntry").CodeModule
body = cm.Lines(1, cm.CountOfLines)
in_ft = False
for i, line in enumerate(body.splitlines(), 1):
    if "Sub Form_Timer(" in line:
        in_ft = True
    if in_ft:
        log(f"  {i}: {line}")
        if line.strip() == "End Sub":
            break

log("Check OnTimer property + TimerInterval BEFORE re-trigger")
try:
    ff = s.app.Forms("LookAtEntry")
    log(f"  OnTimer = {ff.OnTimer!r}")
    log(f"  TimerInterval = {ff.TimerInterval}")
except Exception as e:
    log(f"  read err: {e}")

log("Dump full CmdGIS_Click body (post-patch)")
cm = s.app.VBE.VBProjects(1).VBComponents("Form_LookAtEntry").CodeModule
body = cm.Lines(1, cm.CountOfLines)
in_sub = False
for i, line in enumerate(body.splitlines(), 1):
    if "Sub CmdGIS_Click(" in line:
        in_sub = True
    if in_sub:
        log(f"  {i}: {line}")
        if line.strip() == "End Sub":
            break
s.close()
LIVE.close()
