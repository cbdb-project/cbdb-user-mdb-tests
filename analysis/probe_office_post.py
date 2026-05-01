"""After click_via_timer returns, time various reads to localize hang."""
import sys, time
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))
from cbdb_driver.vba_session import VbaSession

LIVE = open(r"C:\Users\how612\AppData\Local\Temp\probe_post.log",
            "w", encoding="utf-8", buffering=1)
def log(*a):
    s = " ".join(str(x) for x in a)
    LIVE.write(s + "\n"); LIVE.flush()
    print(s)

SRC = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK = ROOT / "analysis" / "_post.mdb"

def t(label, fn):
    t0 = time.time()
    try:
        v = fn()
        log(f"  [{time.time()-t0:6.2f}s] {label}: {v!r}")
    except Exception as e:
        log(f"  [{time.time()-t0:6.2f}s] {label}: FAIL {e}")

s = VbaSession(SRC, WORK).open()
s.open_form("LookAtOffice")
time.sleep(1)
s.set_control("LookAtOffice", "TxtTypeDesc", "[All]")
s.set_picker_codes("ZZ_OFFICE_CODE", [80944], column="c_office_id")
log("Triggering via timer...")
n = s.click_via_timer("LookAtOffice", ctl="CmdQuery",
                     result_table="ZZ_SCRATCH_OFFICE", timeout=180)
log(f"Returned n={n}")

log("Now timing post-click ops:")
t("row_count ZZ_SCRATCH_OFFICE",
  lambda: s.row_count("ZZ_SCRATCH_OFFICE"))
t("row_count ZZ_TEST_DEBUG",
  lambda: s.row_count("ZZ_TEST_DEBUG"))
t("cursor SELECT TOP 1 c_personid",
  lambda: list(s.conn.cursor().execute("SELECT TOP 1 c_personid FROM ZZ_SCRATCH_OFFICE")))
t("cursor SELECT TOP 1 *",
  lambda: list(s.conn.cursor().execute("SELECT TOP 1 * FROM ZZ_SCRATCH_OFFICE")))
t("import pandas + read_sql TOP 1 *",
  lambda: __import__("pandas").read_sql("SELECT TOP 1 * FROM [ZZ_SCRATCH_OFFICE]", s.conn).columns.tolist())
log("Done.")
s.close()
LIVE.close()
