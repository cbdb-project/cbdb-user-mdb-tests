"""Minimal LookAtEntry CmdQuery via timer — matrix-style."""
import sys, time
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))
from cbdb_driver.vba_session import VbaSession

LIVE = open(r"C:\Users\how612\AppData\Local\Temp\probe_min.log",
            "w", encoding="utf-8", buffering=1)
def log(*a):
    s = " ".join(str(x) for x in a)
    LIVE.write(s + "\n"); LIVE.flush()

s = VbaSession(ROOT/"data"/"CBDB_BJ_User.mdb", ROOT/"analysis"/"_probe_min.mdb").open()
s.open_form("LookAtEntry")
# Use small fixture: yin privilege general (118) + kaifeng (100658)
# + years 900-1100 → ~103 rows.
s.set_picker_codes("ZZ_SCRATCH_ENTRY_CODE", [118], column="c_entry_code")
s.set_picker_addrs([100658])
s.set_control("LookAtEntry", "TxtFromYear", 900)
s.set_control("LookAtEntry", "TxtToYear", 1100)
s.set_control("LookAtEntry", "FrameYears", 2)
s.patch_filedialog("LookAtEntry")
log("Set Tag with chain + path; trigger CmdQuery via timer")
out = ROOT / "analysis" / "_exports_tmp" / "min_gis.tab"
out.parent.mkdir(parents=True, exist_ok=True)
if out.exists(): out.unlink()
s.set_form_tag("LookAtEntry", "CmdQuery,CmdGIS", str(out))
n = s.click_via_timer("LookAtEntry", ctl="CmdQuery",
                     result_table="ZZ_SCRATCH_ENTRY", timeout=120)
log(f"CmdQuery -> {n} rows")
df = s.read("ZZ_TEST_DEBUG", order_by="id")
log(f"ZZ_TEST_DEBUG: {df.to_dict('records')}")
log(f"file exists? {out.exists()} size={out.stat().st_size if out.exists() else 'n/a'}")
s.close()
LIVE.close()
