"""Trigger LookAtOffice CmdQuery_Click via Form_Timer (bypasses
disabled-button + Application.Run-unreachable issues)."""
import sys, time
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))
from cbdb_driver.vba_session import VbaSession

LIVE = open(r"C:\Users\how612\AppData\Local\Temp\probeoff_timer.log",
            "w", encoding="utf-8", buffering=1)
def log(*a):
    s = " ".join(str(x) for x in a)
    LIVE.write(s + "\n"); LIVE.flush()
    print(s)

SRC = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK = ROOT / "analysis" / "_office_timer.mdb"

s = VbaSession(SRC, WORK).open()
log("STEP 1: open form")
s.open_form("LookAtOffice")
time.sleep(1)

log("STEP 2: set TxtTypeDesc='[All]'")
s.set_control("LookAtOffice", "TxtTypeDesc", "[All]")

log("STEP 3: populate ZZ_OFFICE_CODE [80944]")
s.set_picker_codes("ZZ_OFFICE_CODE", [80944], column="c_office_id")
log(f"  ZZ_OFFICE_CODE rows = {s.row_count('ZZ_OFFICE_CODE')}")

log("STEP 4: trigger via Form_Timer")
n = s.click_via_timer("LookAtOffice", ctl="CmdQuery",
                     result_table="ZZ_SCRATCH_OFFICE", timeout=120)
log(f"  ZZ_SCRATCH_OFFICE rows = {n}")

log("STEP 5: read ZZ_TEST_DEBUG (autodetect output, if any)")
try:
    df = s.read("ZZ_TEST_DEBUG", order_by="id")
    if df.empty:
        log("  (empty)")
    else:
        for r in df.itertuples():
            log(f"  [{r.id}] {r.msg}")
except Exception as e:
    log(f"  read failed: {e}")

s.close()
LIVE.close()
