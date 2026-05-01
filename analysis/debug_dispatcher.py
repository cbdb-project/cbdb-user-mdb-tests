"""Debug script: run the LookAtEntry smoke flow inline so we can
inspect what goes wrong with the trigger/timer dispatcher."""
from __future__ import annotations
import sys, io, time, subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
LOG = ROOT / "analysis" / "_debug.log"
LOG.unlink(missing_ok=True)

def _print(*args, **kwargs):
    msg = " ".join(str(a) for a in args)
    LOG.open("a", encoding="utf-8").write(msg + "\n")

print = _print  # override

sys.path.insert(0, str(ROOT / "tests"))

# pre-cleanup
subprocess.run(["taskkill", "/F", "/IM", "MSACCESS.EXE"], capture_output=True)

from cbdb_driver import AccessApp, VbaInjector, FormDriver
from cbdb_driver.access_app import make_working_copy, kill_orphan_access

WORK = ROOT / "analysis" / "_test_work.mdb"
SRC = ROOT / "data" / "CBDB_BJ_User.mdb"

print("=== copy ===")
work = make_working_copy(SRC, WORK)
print(f"  {work}")

print("=== open (VISIBLE for timer) ===")
app = AccessApp(work, hidden=False).open()
print(f"  broken refs removed: {app.broken_refs_removed}")
print(f"  DAO added:           {app.dao_added}")

print("=== inject ===")
rep = VbaInjector(app).run_all()
print(f"  handlers publicised: {rep.handlers_publicised}")
print(f"  msgbox replaced:     {rep.msgbox_replacements}")
print(f"  forms modified:      {len(rep.forms_modified)}")
print(f"  timer injected:      {rep.timer_injected_into}")
print(f"  errs={rep.error_table_created} trig={rep.trigger_table_created} state={rep.state_table_created}")

print("=== compile ===")
print(f"  compiled: {app.compile_vba()}")

print("=== open LookAtEntry NORMAL (window-visible) ===")
app.app.DoCmd.OpenForm("LookAtEntry", 0, "", "", 0, 0)  # WindowMode=0 normal
time.sleep(0.5)
loaded = bool(app.app.CurrentProject.AllForms("LookAtEntry").IsLoaded)
print(f"  IsLoaded: {loaded}")

driver = FormDriver(app)

try:
    print("=== state round-trip ===")
    driver.set_global("export_path", "PROBE")
    print(f"  get_global('export_path'): {driver.get_global('export_path')!r}")

    print("=== seed picker tables ===")
    app.exec_sql("DELETE FROM ZZ_SCRATCH_ENTRY_CODE")
    app.exec_sql("INSERT INTO ZZ_SCRATCH_ENTRY_CODE (c_entry_code) VALUES (118)")
    app.exec_sql("DELETE FROM ZZ_SCRATCH_ADDR")
    app.exec_sql("INSERT INTO ZZ_SCRATCH_ADDR (c_addr_id) VALUES (100658)")
    print(f"  ZZ_SCRATCH_ENTRY_CODE rows: {app.row_count('ZZ_SCRATCH_ENTRY_CODE')}")
    print(f"  ZZ_SCRATCH_ADDR rows:       {app.row_count('ZZ_SCRATCH_ADDR')}")

    print("=== set form controls ===")
    driver.set_control("LookAtEntry", "TxtFromYear", 900)
    driver.set_control("LookAtEntry", "TxtToYear", 1100)
    driver.set_control("LookAtEntry", "FrameYears", 2)
    driver.set_control("LookAtEntry", "TxtEntryDesc", "yin privilege: general")
    driver.set_control("LookAtEntry", "TxtTypeCode", "N/A")

    print("=== suppress msgbox ===")
    driver.set_global("suppress_msgbox", True)

    print("=== probe ALL invocation paths for CmdQuery_Click ===")
    # First make sure scratch is empty so we know if it worked
    app.exec_sql("DELETE FROM ZZ_SCRATCH_ENTRY")
    print(f"  scratch wiped, rows: {app.row_count('ZZ_SCRATCH_ENTRY')}")

    paths = [
        ("Application.Run('Form_LookAtEntry.CmdQuery_Click')",
         lambda: app.app.Run("Form_LookAtEntry.CmdQuery_Click")),
        ("Application.Run('[Form_LookAtEntry].CmdQuery_Click')",
         lambda: app.app.Run("[Form_LookAtEntry].CmdQuery_Click")),
        ("app.Forms('LookAtEntry').CmdQuery_Click()",
         lambda: app.app.Forms("LookAtEntry").CmdQuery_Click()),
        ("app.Eval('[Forms]![LookAtEntry].[CmdQuery_Click]()')",
         lambda: app.app.Eval("[Forms]![LookAtEntry].[CmdQuery_Click]()")),
        ("DoCmd.SelectObject + GoToControl + RunCommand acCmdSpaceBar(99)",
         lambda: (
             app.app.DoCmd.SelectObject(2, "LookAtEntry", False),
             app.app.DoCmd.GoToControl("CmdQuery"),
             app.app.SendKeys("{ENTER}", True),
         )),
    ]
    for label, fn in paths:
        try:
            fn()
            n = app.row_count("ZZ_SCRATCH_ENTRY")
            print(f"  [OK]   {label}  -> rows={n}")
            if n > 0:
                # cleanup for next attempt
                app.exec_sql("DELETE FROM ZZ_SCRATCH_ENTRY")
        except Exception as e:
            print(f"  [FAIL] {label}  -> {type(e).__name__}: {e}")

    print("=== probe TimerInterval ===")
    f = app.app.Forms("LookAtEntry")
    print(f"  initial TimerInterval: {f.TimerInterval}")
    print(f"  initial OnTimer:       {f.OnTimer!r}")
    f.OnTimer = "[Event Procedure]"
    print(f"  after set OnTimer:     {f.OnTimer!r}")
    f.TimerInterval = 100
    print(f"  after set Interval:    {f.TimerInterval}")
    print(f"  Visible:               {f.Visible}")

    print("=== insert trigger row ===")
    app.exec_sql(
        "INSERT INTO ZZ_TEST_TRIGGER ([form], [action], args) "
        "VALUES ('LookAtEntry', 'CmdQuery_Click', '')"
    )
    print(f"  trigger rows now: {app.row_count('ZZ_TEST_TRIGGER')}")

    print("=== wait 5s, pumping msgs ===")
    import pythoncom
    deadline = time.time() + 5
    while time.time() < deadline:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.05)
    print(f"  trigger rows after wait: {app.row_count('ZZ_TEST_TRIGGER')}")
    print(f"  ZZ_TEST_ERRORS rows:     {app.row_count('ZZ_TEST_ERRORS')}")
    print(f"  ZZ_SCRATCH_ENTRY rows:   {app.row_count('ZZ_SCRATCH_ENTRY')}")
    print(f"  Visible:                 {f.Visible}")
    print(f"  TimerInterval:           {f.TimerInterval}")

    print("=== inspect outcome ===")
    n = app.row_count("ZZ_SCRATCH_ENTRY")
    print(f"  ZZ_SCRATCH_ENTRY rows: {n}")
    n_errs = app.row_count("ZZ_TEST_ERRORS")
    print(f"  ZZ_TEST_ERRORS rows:   {n_errs}")
    if n_errs:
        for r in app.fetch_all("SELECT form_name, event_name, err_desc FROM ZZ_TEST_ERRORS ORDER BY id"):
            print(f"    [{r.form_name}.{r.event_name}] {r.err_desc}")
    n_trig = app.row_count("ZZ_TEST_TRIGGER")
    print(f"  ZZ_TEST_TRIGGER rows:  {n_trig}")
    if n_trig:
        for r in app.fetch_all("SELECT form, action FROM ZZ_TEST_TRIGGER ORDER BY id"):
            print(f"    pending: form={r.form!r} action={r.action!r}")
    if n > 0:
        n_distinct = app.fetch_one(
            "SELECT COUNT(*) FROM (SELECT DISTINCT c_personid FROM ZZ_SCRATCH_ENTRY) AS sub"
        )[0]
        print(f"  distinct persons: {n_distinct}")
        # enable states
        for c in ("CmdGIS", "CmdNeo4j", "CmdStoreID"):
            ena = app.app.Forms("LookAtEntry").Controls(c).Enabled
            print(f"  {c}.Enabled = {ena}")
finally:
    print("=== teardown ===")
    try: app.app.DoCmd.Close(2, "LookAtEntry", 2)
    except Exception as e: print(f"  close form: {e}")
    try: app.close()
    except Exception as e: print(f"  close app: {e}")
    subprocess.run(["taskkill", "/F", "/IM", "MSACCESS.EXE"], capture_output=True)
    print("done")
