"""Step-by-step probe of LookAtOffice flow to find why we get 0 rows.

Mirrors test_cross_form_matrix exactly, but prints state at every step.
"""
from __future__ import annotations
import sys, time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))

from cbdb_driver.vba_session import VbaSession

SRC = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK = ROOT / "analysis" / "_probe_office_step.mdb"

# Live log to file so we can tail without waiting for buffer flush
LOG = open(r"C:\Users\how612\AppData\Local\Temp\probeoff_live.log",
           "w", encoding="utf-8", buffering=1)
_orig_print = print
def print(*args, **kw):  # type: ignore
    s = " ".join(str(a) for a in args)
    LOG.write(s + "\n")
    LOG.flush()
    _orig_print(s, **kw)


def main():
    print("=" * 60)
    print("STEP 1: Open VbaSession (fresh copy + Access launched)")
    s = VbaSession(SRC, WORK).open()
    print(f"  Access app pid={s.app.hWndAccessApp if s.app else 'n/a'}")

    print("\nSTEP 2: Open Form_LookAtOffice")
    s.open_form("LookAtOffice")
    time.sleep(2)

    print("\nSTEP 3: Read ZZ_OFFICE_CODE rows (should be 0 — Form_Open wipes)")
    n = s.row_count("ZZ_OFFICE_CODE")
    print(f"  ZZ_OFFICE_CODE rows: {n}")

    print("\nSTEP 4: set_control TxtTypeDesc='[All]'")
    try:
        s.set_control("LookAtOffice", "TxtTypeDesc", "[All]")
        print("  set_control OK")
    except Exception as e:
        print(f"  set_control FAILED: {e}")

    print("\nSTEP 5: set_control FrameFilterYears=1")
    try:
        s.set_control("LookAtOffice", "FrameFilterYears", 1)
        print("  set_control OK")
    except Exception as e:
        print(f"  set_control FAILED: {e}")

    print("\nSTEP 6: Read controls back via COM")
    try:
        f = s.app.Forms("LookAtOffice")
        print(f"  TxtTypeDesc: {f.Controls('TxtTypeDesc').Value!r}")
        print(f"  FrameFilterYears: {f.Controls('FrameFilterYears').Value!r}")
        print(f"  TxtOfficeID: {f.Controls('TxtOfficeID').Value!r}")
    except Exception as e:
        print(f"  read FAILED: {e}")

    print("\nSTEP 7: set_picker_codes(ZZ_OFFICE_CODE, [80944])")
    s.set_picker_codes("ZZ_OFFICE_CODE", [80944], column="c_office_id")
    n = s.row_count("ZZ_OFFICE_CODE")
    print(f"  ZZ_OFFICE_CODE rows after insert: {n}")

    print("\nSTEP 8: Verify autodetect injection in Form_LookAtOffice.CmdQuery_Click")
    cm = s.app.VBE.VBProjects(1).VBComponents("Form_LookAtOffice").CodeModule
    body = cm.Lines(1, cm.CountOfLines)
    if "AUTO-DETECT PICKER STATE v2" in body:
        print("  ✓ marker v2 present")
    else:
        print("  ✗ marker v2 MISSING — injection didn't happen!")
    if "gUseOfficeID = (tdOC > 0)" in body:
        print("  ✓ gUseOfficeID setter present")
    else:
        print("  ✗ gUseOfficeID setter MISSING")

    print("\nSTEP 9: Click Run Query (force-enable + pywinauto click)")
    n = s.click_button_and_wait_table(
        "Run Query", form="LookAtOffice",
        result_table="ZZ_SCRATCH_OFFICE",
        force_enable_ctl="CmdQuery", timeout=90,
    )
    print(f"  ZZ_SCRATCH_OFFICE rows: {n}")

    print("\nSTEP 9b: list all top-level Access windows (look for MsgBox)")
    try:
        from pywinauto import Application as PWA
        pwa = PWA(backend="uia").connect(path="MSACCESS.EXE")
        for w in pwa.windows():
            try:
                title = w.window_text()
                ct = w.element_info.control_type
                print(f"  window: {ct!r} title={title!r}")
            except Exception:
                pass
    except Exception as e:
        print(f"  enum failed: {e}")

    print("\nSTEP 9c: try the picker SQL directly via pyodbc on same work copy")
    try:
        s.exec_sql("DELETE FROM ZZ_SCRATCH_OFFICE")
        rc = s.exec_sql(
            "INSERT INTO ZZ_SCRATCH_OFFICE ("
            "c_posting_id, c_personid, c_index_year, c_female, c_person_dy, "
            "c_office_id, c_sequence, c_firstyear, c_fy_nh_code, c_fy_nh_year, "
            "c_fy_range, c_lastyear, c_ly_nh_code, c_ly_nh_year, c_ly_range, "
            "c_appt_code, c_assume_office_code, c_inst_code, c_inst_name_code, "
            "c_source, c_pages, c_notes, c_fy_intercalary, c_fy_month, "
            "c_ly_intercalary, c_ly_month, c_fy_day, c_ly_day, c_fy_day_gz, "
            "c_ly_day_gz, c_dy, c_office_category_id, c_addr_id, c_addr_type, "
            "c_office_addr_id, c_index_year_type_code) "
            "SELECT POD.c_posting_id, POD.c_personid, BIOG_MAIN.c_index_year, "
            "BIOG_MAIN.c_female, BIOG_MAIN.c_dy, POD.c_office_id, POD.c_sequence, "
            "POD.c_firstyear, POD.c_fy_nh_code, POD.c_fy_nh_year, POD.c_fy_range, "
            "POD.c_lastyear, POD.c_ly_nh_code, POD.c_ly_nh_year, POD.c_ly_range, "
            "POD.c_appt_code, POD.c_assume_office_code, POD.c_inst_code, "
            "POD.c_inst_name_code, POD.c_source, POD.c_pages, POD.c_notes, "
            "POD.c_fy_intercalary, POD.c_fy_month, POD.c_ly_intercalary, "
            "POD.c_ly_month, POD.c_fy_day, POD.c_ly_day, POD.c_fy_day_gz, "
            "POD.c_ly_day_gz, BIOG_MAIN.c_dy, POD.c_office_category_id, "
            "BIOG_MAIN.c_index_addr_id, BIOG_MAIN.c_index_addr_type_code, "
            "PAD.c_addr_id, BIOG_MAIN.c_index_year_type_code "
            "FROM BIOG_MAIN INNER JOIN ((POSTED_TO_OFFICE_DATA AS POD "
            "INNER JOIN POSTED_TO_ADDR_DATA AS PAD "
            "ON (POD.c_posting_id = PAD.c_posting_id) "
            "AND (POD.c_office_id = PAD.c_office_id)) "
            "INNER JOIN ZZ_OFFICE_CODE ON POD.c_office_id = ZZ_OFFICE_CODE.c_office_id) "
            "ON BIOG_MAIN.c_personid = POD.c_personid"
        )
        n2 = s.row_count("ZZ_SCRATCH_OFFICE")
        print(f"  pyodbc INSERT rowcount={rc}, ZZ_SCRATCH_OFFICE final={n2}")
    except Exception as e:
        print(f"  pyodbc INSERT FAILED: {e}")

    print("\nSTEP 10: Read back gUseOfficeID + TxtOfficeID + TxtTypeDesc")
    try:
        f = s.app.Forms("LookAtOffice")
        print(f"  TxtOfficeID: {f.Controls('TxtOfficeID').Value!r}")
        print(f"  TxtTypeDesc: {f.Controls('TxtTypeDesc').Value!r}")
    except Exception as e:
        print(f"  read FAILED: {e}")

    print("\nSTEP 11: Read ZZ_TEST_DEBUG (autodetect output)")
    try:
        df = s.read("ZZ_TEST_DEBUG", order_by="id")
        for row in df.itertuples():
            print(f"  [{row.id}] {row.msg}")
        if df.empty:
            print("  (empty — autodetect didn't fire / didn't reach INSERT)")
    except Exception as e:
        print(f"  read failed: {e}")

    s.close()


if __name__ == "__main__":
    main()
