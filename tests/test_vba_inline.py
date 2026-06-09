"""
Inline-probe pytest test.  Verbatim port of analysis/probe_pywinauto.py
into a single pytest function.  No conftest fixtures, no FormDriver
abstraction — exactly the code path that DOES work standalone.

If this passes, fixtures are the issue (fix the abstractions).
If this fails, something about pytest's process model differs from a
plain Python script.

This test is INTENTIONALLY ugly. It's a proof of life, not production.
"""
from __future__ import annotations

import os
import shutil
import time
from pathlib import Path

import pyodbc
import pytest
import win32com.client

from cbdb_driver.access_app import (
    _pid_for_access_app, register_spawned_pid, kill_access_pid,
    _kill_file_holder,
)


ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK = ROOT / "analysis" / "_inline_test_copy.mdb"


@pytest.fixture(scope="function")
def inline_session():
    """Self-contained per-test session: kill orphans, copy mdb, patch
    LinkListInit, open Access visible, fix DAO ref, open LookAtEntry.

    Yields a dict of handles {app, conn}. Tears down by closing
    everything.
    """
    # B10: do NOT global-kill MSACCESS (that nukes the developer's unrelated
    # Access windows).  Only free WORK if a stale process is holding it —
    # _kill_file_holder scopes to the single process that has WORK open.
    if WORK.exists():
        try:
            WORK.unlink()
        except PermissionError:
            _kill_file_holder(WORK)
            time.sleep(1)
            WORK.unlink()
    shutil.copy2(SRC, WORK)

    # pre-patch LinkListInit so NAVIGATION_PANE.Form_Open exits early
    conn = pyodbc.connect(
        "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={WORK};", autocommit=True
    )
    conn.cursor().execute(
        f"UPDATE LinkListInit SET c_path = '{str(WORK).replace(chr(39), chr(39)*2)}'"
    )
    # don't close yet — pyodbc needs to remain open for the test

    # open Access VISIBLE so pywinauto can click
    app = win32com.client.DispatchEx("Access.Application")
    try:
        app.AutomationSecurity = 1
    except Exception:
        pass
    app.Visible = True
    app.OpenCurrentDatabase(str(WORK))
    # Capture + register the PID now (HWND alive) so teardown can scope-kill
    # only THIS process and the sessionfinish safety-net knows about it.
    _pid = _pid_for_access_app(app)
    register_spawned_pid(_pid)

    # fix DAO ref
    proj = app.VBE.VBProjects(1)
    for r in list(proj.References):
        if r.IsBroken:
            full = getattr(r, "FullPath", "") or ""
            proj.References.Remove(r)
            if "dao" in full.lower():
                for cand in (
                    r"C:\Program Files\Microsoft Office\root\Office16\ACEDAO.DLL",
                    r"C:\Program Files (x86)\Microsoft Office\root\Office16\ACEDAO.DLL",
                ):
                    if Path(cand).exists():
                        proj.References.AddFromFile(cand)
                        break

    # open LookAtEntry visible-normal
    app.DoCmd.OpenForm("LookAtEntry", 0, "", "", 0, 0)
    time.sleep(1)

    yield {"app": app, "conn": conn, "work": WORK}

    # teardown
    try: conn.close()
    except Exception: pass
    try: app.DoCmd.Close(2, "LookAtEntry", 2)   # acForm, acSaveNo
    except Exception: pass
    try: app.CloseCurrentDatabase()
    except Exception: pass
    try: app.Quit()
    except Exception: pass
    # B10: scoped kill of ONLY this test's Access process (was global /IM).
    if _pid is not None:
        kill_access_pid(_pid)


def test_vba_lookatentry_run_query_yin_general_kaifeng(inline_session):
    """Drive the actual VBA CmdQuery_Click via pywinauto and verify
    ZZ_SCRATCH_ENTRY gets populated. This mirrors probe_pywinauto.py
    exactly — no fixture abstractions in the click path."""
    app = inline_session["app"]
    conn = inline_session["conn"]

    # seed picker tables (bypass picker UI)
    cur = conn.cursor()
    cur.execute("DELETE FROM ZZ_SCRATCH_ENTRY")
    cur.execute("DELETE FROM ZZ_SCRATCH_ENTRY_CODE")
    cur.execute("INSERT INTO ZZ_SCRATCH_ENTRY_CODE (c_entry_code) VALUES (118)")
    cur.execute("DELETE FROM ZZ_SCRATCH_ADDR")
    cur.execute("INSERT INTO ZZ_SCRATCH_ADDR (c_addr_id) VALUES (100658)")
    cur.close()

    # set form input controls via COM
    def setctl(name, val):
        c = app.Forms("LookAtEntry").Controls(name)
        try: c.SetFocus()
        except Exception: pass
        c.Value = val

    setctl("TxtFromYear", 900)
    setctl("TxtToYear", 1100)
    setctl("FrameYears", 2)              # Index Years
    setctl("TxtEntryDesc", "yin privilege: general")
    setctl("TxtTypeCode", "N/A")

    # Force-enable Run Query so the click takes effect
    # (in real UI, picker handlers would have done this)
    app.Forms("LookAtEntry").Controls("CmdQuery").Enabled = True

    # Now click via pywinauto
    from pywinauto import Application as PWA
    pwa = PWA(backend="uia").connect(path="MSACCESS.EXE", timeout=10)
    main = pwa.window(title="Welcome to CBDB!")
    main.wait("ready", timeout=10).set_focus()    # bring Access to foreground
    time.sleep(0.5)

    target = None
    for d in main.descendants():
        try:
            if (d.element_info.control_type == "Button"
                    and (d.window_text() or "").strip().lower() == "run query"):
                target = d
                break
        except Exception:
            continue
    assert target is not None, "Run Query button not found via pywinauto"
    assert target.is_enabled(), "Run Query button is not enabled"

    target.click_input()

    # poll for ZZ_SCRATCH_ENTRY to fill
    rows = 0
    deadline = time.time() + 30
    while time.time() < deadline:
        time.sleep(0.5)
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM ZZ_SCRATCH_ENTRY")
        rows = int(cur.fetchone()[0])
        cur.close()
        if rows > 0:
            break

    assert rows > 0, (
        f"ZZ_SCRATCH_ENTRY = {rows} after pywinauto click; the VBA "
        "CmdQuery_Click did not fire (or fired with no results)."
    )

    # Re-acquire COM (pywinauto may have invalidated the proxy)
    try:
        _ = app.Forms("LookAtEntry").Controls("CmdGIS").Enabled
        app2 = app
    except Exception:
        app2 = win32com.client.GetActiveObject("Access.Application")

    # Post-query: export buttons should be Enabled
    assert app2.Forms("LookAtEntry").Controls("CmdGIS").Enabled, \
        "CmdGIS not enabled after query"
    assert app2.Forms("LookAtEntry").Controls("CmdNeo4j").Enabled, \
        "CmdNeo4j not enabled after query"
    assert app2.Forms("LookAtEntry").Controls("CmdStoreID").Enabled, \
        "CmdStoreID not enabled after query"

    # Distinct person count sanity check
    cur = conn.cursor()
    cur.execute(
        "SELECT COUNT(*) FROM "
        "(SELECT DISTINCT c_personid FROM ZZ_SCRATCH_ENTRY) AS sub"
    )
    n_distinct = int(cur.fetchone()[0])
    cur.close()
    # HelpFile says ~104 for this exact query; current data ~103
    assert 50 <= n_distinct <= 1500, (
        f"distinct persons = {n_distinct}, expected ~104 (HelpFile)"
    )
