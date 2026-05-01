"""
Phase-1 infrastructure smoke test (pywinauto + COM hybrid).

Validates:
  1. AccessApp opens a working copy + repairs broken DAO ref
  2. LinkListInit pre-patch prevents NAVIGATION_PANE.Form_Open hang
  3. VbaInjector creates ZZ_TEST_ERRORS / ZZ_TEST_STATE tables
     and replaces MsgBox calls
  4. FormDriver.click_button via pywinauto fires CmdQuery_Click
     (with force-enable to bypass picker prerequisites)
  5. ZZ_SCRATCH_ENTRY gets populated
  6. Post-query: CmdGIS / CmdNeo4j / CmdStoreID become Enabled
  7. ZZ_TEST_ERRORS stays empty
"""
from __future__ import annotations

import pytest


def test_infra_helpers_injected(com_app):
    """Test tables exist; helpers added to host form module."""
    cur = com_app.conn.cursor()
    for t in ("ZZ_TEST_ERRORS", "ZZ_TEST_STATE"):
        cur.execute(f"SELECT TOP 1 * FROM [{t}]")
        cur.fetchone()
    cur.close()
    proj = com_app.app.VBE.VBProjects(1)
    host = proj.VBComponents("Form_LookAtEntry")
    code = host.CodeModule.Lines(1, host.CodeModule.CountOfLines)
    assert "AUTO-INJECTED CBDB TEST HELPERS v3" in code
    assert "Public Function TestMsgBox" in code


def test_state_round_trip(com_app, driver, clean_state):
    driver.set_global("export_path", r"C:\tmp\probe.tab")
    assert driver.get_global("export_path") == r"C:\tmp\probe.tab"
    driver.set_global("suppress_msgbox", True)
    assert driver.get_global("suppress_msgbox") == "1"


@pytest.mark.skip(
    reason="pywinauto+COM end-to-end works in standalone probe but not "
           "under pytest session-state; tracked in PHASE1_BREAKTHROUGH.md. "
           "Use the SQL-replay tests in test_lookatentry.py for now."
)
def test_lookatentry_full_workflow(com_app, fresh_form):
    """End-to-end: set inputs, click Run Query (via pywinauto),
    verify ZZ_SCRATCH_ENTRY filled, verify post-query enable states."""

    # 1. Open LookAtEntry fresh (each test gets a clean form instance)
    driver = fresh_form("LookAtEntry")

    # 2. Suppress dialogs so unattended test doesn't block
    driver.set_global("suppress_msgbox", True)

    # 3. Bypass the entry/place pickers by inserting their scratch rows
    com_app.exec_sql("DELETE FROM ZZ_SCRATCH_ENTRY_CODE")
    com_app.exec_sql("INSERT INTO ZZ_SCRATCH_ENTRY_CODE (c_entry_code) VALUES (118)")
    com_app.exec_sql("DELETE FROM ZZ_SCRATCH_ADDR")
    com_app.exec_sql("INSERT INTO ZZ_SCRATCH_ADDR (c_addr_id) VALUES (100658)")

    # 4. Set form input controls
    driver.set_control("LookAtEntry", "TxtFromYear", 900)
    driver.set_control("LookAtEntry", "TxtToYear", 1100)
    driver.set_control("LookAtEntry", "FrameYears", 2)         # Index Years
    driver.set_control("LookAtEntry", "TxtEntryDesc", "yin privilege: general")
    driver.set_control("LookAtEntry", "TxtTypeCode", "N/A")

    # 5. Click Run Query (force-enable since picker buttons would
    #    normally have toggled Enabled=True for us)
    n_rows = driver.run_query(
        "LookAtEntry",
        result_table="ZZ_SCRATCH_ENTRY",
        cmd_caption="Run Query",
        cmd_name="CmdQuery",
        timeout=30,
    )
    assert n_rows > 0, f"ZZ_SCRATCH_ENTRY = {n_rows}"

    # 6. Distinct person count: HelpFile says ~104 (today's data ≈103)
    n_distinct = com_app.fetch_one(
        "SELECT COUNT(*) FROM "
        "(SELECT DISTINCT c_personid FROM ZZ_SCRATCH_ENTRY) AS sub"
    )[0]
    assert 50 <= n_distinct <= 1500, (
        f"distinct persons = {n_distinct}; sanity range failed"
    )

    # 7. Post-query enable states
    assert driver.get_control_property("LookAtEntry", "CmdGIS", "Enabled") is True
    assert driver.get_control_property("LookAtEntry", "CmdNeo4j", "Enabled") is True
    assert driver.get_control_property("LookAtEntry", "CmdStoreID", "Enabled") is True

    # 8. No VBA errors
    n_errs = com_app.row_count("ZZ_TEST_ERRORS")
    assert n_errs == 0, f"unexpected VBA errors: {n_errs}"
