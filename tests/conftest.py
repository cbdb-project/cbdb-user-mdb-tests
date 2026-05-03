"""
pytest fixtures for the CBDB user-mdb test suite.

Three scopes:
  - session-scoped read-only conn to the original CBDB_BJ_User.mdb
    (used by tests that only SELECT)
  - session-scoped COM driver against an INJECTED working copy
    (used by COM-driven tests in test_*com*.py)
  - per-test cleanup of ZZ_SCRATCH_* / ZZ_TEST_ERRORS between cases
"""
from __future__ import annotations

import os
import sys
import shutil
import tempfile
from pathlib import Path

import pyodbc
import pytest

# Make tests/ importable
TESTS_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(TESTS_DIR))

from cbdb_replay.driver import open_connection, working_copy
from cbdb_driver import AccessApp, VbaInjector, FormDriver
from cbdb_driver.access_app import make_working_copy, kill_orphan_access

ROOT = TESTS_DIR.parent
USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK_MDB = ROOT / "analysis" / "_test_work.mdb"


def pytest_addoption(parser):
    parser.addoption(
        "--regenerate-goldens", action="store_true", default=False,
        help="Re-write golden CSVs from current results instead of comparing.",
    )
    parser.addoption(
        "--user-mdb", action="store", default=str(USER_MDB),
        help="Path to the user mdb to test. Defaults to ../CBDB_BJ_User.mdb.",
    )
    parser.addoption(
        "--no-discover-inputs", action="store_true", default=False,
        help="Skip auto-running discover_test_inputs.py at session start.",
    )
    parser.addoption(
        "--include-vba", action="store_true", default=False,
        help="Include the Access-COM ('access' marker) test files.  "
             "Defaults to OFF — the fast suite skips them so headless / "
             "non-Windows runs don't error-out.",
    )


def pytest_collection_modifyitems(config, items):
    """Skip Access-COM-dependent tests by default.

    Every test file that uses `cbdb_driver.vba_session.VbaSession` /
    win32com / pywinauto must spawn an Access process and is slow,
    Windows-only, and prone to environment errors (orphan MSACCESS
    processes, RPC unavailable, ROT collisions).

    The fast (non-Access) suite is what CI / quick-check runs want;
    pass `--include-vba` to opt in to the COM suite.
    """
    if config.getoption("--include-vba"):
        return
    skip_access = pytest.mark.skip(
        reason="needs Access COM — run with `--include-vba` to enable"
    )
    for item in items:
        path = str(item.fspath).replace("\\", "/")
        # Files that drive Access via COM:
        if ("/test_vba_" in path
                or path.endswith("/test_infra_smoke.py")):
            item.add_marker(skip_access)


@pytest.fixture(scope="session")
def user_mdb_path(request) -> Path:
    p = Path(request.config.getoption("--user-mdb")).resolve()
    if not p.exists():
        pytest.exit(f"user mdb not found: {p}")
    return p


def pytest_configure(config):
    """Register markers AND refresh test_inputs.json if stale.

    Discovery failures hard-exit pytest — running matrix tests against
    a stale fixture file silently masks data-version drift, which
    looks like 'pass' but is actually 'tested with the wrong data'.

    Disable refresh with: pytest --no-discover-inputs
    """
    config.addinivalue_line(
        "markers",
        "access: requires a running Access COM session "
        "(Windows + Office + a working data/CBDB_BJ_User.mdb).",
    )
    if config.getoption("--no-discover-inputs"):
        return
    inputs_json = ROOT / "analysis" / "dump" / "test_inputs.json"
    mdb = ROOT / "data" / "CBDB_BJ_User.mdb"
    if not mdb.exists():
        return
    needs_refresh = (
        not inputs_json.exists()
        or inputs_json.stat().st_mtime < mdb.stat().st_mtime
    )
    if needs_refresh:
        print(f"\n[conftest] refreshing {inputs_json.name} "
              f"(stale or missing) ...")
        import subprocess
        rc = subprocess.run(
            [sys.executable, str(ROOT / "analysis" / "discover_test_inputs.py")],
            capture_output=True, text=True,
        )
        if rc.returncode != 0:
            # Hard-exit rather than continue with stale fixtures.  The
            # original behaviour (warn + continue) would let matrix
            # tests pass against an outdated test_inputs.json — which
            # silently masks data-version drift and produces misleading
            # green CI runs.
            pytest.exit(
                f"[conftest] discover_test_inputs.py FAILED (rc="
                f"{rc.returncode}).  Tests would otherwise run against "
                f"a stale fixture file ({inputs_json.name}).  Fix the "
                f"discovery error or pass `--no-discover-inputs` to "
                f"skip refresh.\n\n  stderr tail:\n{rc.stderr[-1000:]}"
            )
        else:
            print(f"[conftest] discovery refreshed.")


@pytest.fixture(scope="session")
def ro_conn(user_mdb_path) -> pyodbc.Connection:
    """Read-only connection to the production user-mdb."""
    conn = open_connection(user_mdb_path, readonly=True)
    yield conn
    conn.close()


@pytest.fixture(scope="session")
def golden_dir() -> Path:
    p = TESTS_DIR / "golden"
    p.mkdir(exist_ok=True)
    return p


@pytest.fixture(scope="session")
def regenerate_goldens(request) -> bool:
    return bool(request.config.getoption("--regenerate-goldens"))


# ---------- COM driver fixtures (Phase 1 infra) ----------

@pytest.fixture(scope="session")
def com_app(user_mdb_path):
    """Session-wide AccessApp running against an injected working copy.

    Access is opened VISIBLE because pywinauto-based clicks need the
    form's window to receive focus.  This fixture is session-scoped so
    one Access process serves all tests."""
    kill_orphan_access()
    work = make_working_copy(user_mdb_path, WORK_MDB)
    # Pre-patch LinkListInit BEFORE Access opens, otherwise
    # NAVIGATION_PANE.Form_Open hangs trying to relink to a non-existent
    # CBDB_<ver>_DATA.mdb at our working path.
    import pyodbc as _pyodbc
    _conn = _pyodbc.connect(
        "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={work};", autocommit=True
    )
    _cur = _conn.cursor()
    _cur.execute(
        f"UPDATE LinkListInit SET c_path = '{str(work).replace(chr(39), chr(39)*2)}'"
    )
    _cur.close(); _conn.close()
    # Open Access VISIBLE — pywinauto needs a visible window to click on.
    app = AccessApp(work, hidden=False).open()
    rep = VbaInjector(app).run_all()
    print(
        f"\n[com_app] working copy at {work.name}\n"
        f"  broken refs removed: {app.broken_refs_removed}\n"
        f"  DAO added:           {app.dao_added}\n"
        f"  msgbox replaced:     {rep.msgbox_replacements}\n"
        f"  forms modified:      {len(rep.forms_modified)}\n"
        f"  helpers host:        {rep.helpers_host}\n"
        f"  tables created:      "
        f"errs={rep.error_table_created} "
        f"state={rep.state_table_created}\n"
        f"  linklist patched:    {rep.linklist_patched}"
    )
    compiled = app.compile_vba()
    print(f"  vba compiled:        {compiled}")
    # Do NOT pre-open any form here.  Each test fixture (fresh_form)
    # opens its target form afresh so per-test state is clean.
    yield app
    app.close()


@pytest.fixture(scope="session")
def driver(com_app) -> FormDriver:
    drv = FormDriver(com_app)
    yield drv
    drv.close_all()


@pytest.fixture(scope="function")
def clean_state(com_app):
    """Reset per-form scratch tables + ZZ_TEST_ERRORS + ZZ_TEST_STATE
    between tests."""
    for tbl in (
        "ZZ_TEST_ERRORS", "ZZ_TEST_STATE",
        "ZZ_SCRATCH_ENTRY", "ZZ_SCRATCH_ENTRY_CODE",
        "ZZ_SCRATCH_ADDR", "ZZ_SCRATCH_ADDR_LIST",
    ):
        try:
            com_app.exec_sql(f"DELETE FROM [{tbl}]")
        except Exception:
            pass
    yield


@pytest.fixture(scope="function")
def fresh_form(com_app, driver, clean_state):
    """Open the target form fresh for each test (closes any prior open
    forms first), and reset the pywinauto cache so we get clean window
    handles. Yield (driver, opener) where opener(form_name) loads the
    given form."""
    def _open(form_name: str):
        # Close any LookAt* form that's still loaded from a prior test
        for ao in com_app.app.CurrentProject.AllForms:
            if ao.IsLoaded:
                try:
                    com_app.app.DoCmd.Close(2, ao.Name, 2)  # acForm, acSaveNo
                except Exception:
                    pass
        # Drop pywinauto cache (window handles may be stale after closes)
        driver.reset_pywinauto()
        # Open the requested form, visible normal so pywinauto can click it
        com_app.app.DoCmd.OpenForm(form_name, 0, "", "", 0, 0)
        return driver
    yield _open
    # post-test: close everything, clear pwa cache for next test
    for ao in com_app.app.CurrentProject.AllForms:
        try:
            if ao.IsLoaded:
                com_app.app.DoCmd.Close(2, ao.Name, 2)
        except Exception:
            pass
    driver.reset_pywinauto()
