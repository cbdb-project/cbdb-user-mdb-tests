"""For each LookAt form, open it and check whether CmdQuery starts
enabled or disabled."""
import sys, time
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))
from cbdb_driver.vba_session import VbaSession

SRC = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK = ROOT / "analysis" / "_btn_state.mdb"

FORMS = ["LookAtEntry", "LookAtStatus", "LookAtTexts",
         "LookAtAssociations", "LookAtOffice"]

def main():
    s = VbaSession(SRC, WORK).open()
    for f in FORMS:
        try:
            s.open_form(f)
            time.sleep(1)
            ctl = s.app.Forms(f).Controls("CmdQuery")
            print(f"  {f}: CmdQuery.Enabled = {ctl.Enabled}")
            try:
                s.close_form(f)
            except Exception:
                pass
        except Exception as e:
            print(f"  {f}: ERROR {e}")
    s.close()

if __name__ == "__main__":
    main()
