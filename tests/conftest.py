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

import pytest

# Make tests/ importable
TESTS_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(TESTS_DIR))

# IMPORTANT: do NOT import pyodbc / cbdb_driver / win32com / pywinauto
# at module top-level.  The fast suite (no `--include-vba`) must collect
# cleanly on Linux / headless / fresh machines that don't have Office or
# pywin32 installed.  Each fixture below imports its dependencies
# locally — that way pytest collection only needs `pytest` itself.

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
        help="Force-include the Access-COM ('access' marker) test files.  "
             "On a capable Windows box (pywin32 present) the COM suite is "
             "ON by default already; use --fast to opt OUT.",
    )
    parser.addoption(
        "--fast", action="store_true", default=False,
        help="Explicitly run the FAST subset (skip the Access-COM test "
             "files).  Use this to opt out of the auto-on COM suite on "
             "Windows — so 'running the tests' is always an explicit choice, "
             "never a silent smaller set.",
    )


def _is_access_com_test_file(name: str) -> bool:
    return name.startswith("test_vba_") or name == "test_infra_smoke.py"


def _com_capable() -> bool:
    """True if this box can actually drive Access COM (Windows + pywin32)."""
    if sys.platform != "win32":
        return False
    try:
        import win32com  # noqa: F401
        return True
    except ImportError:
        return False


def _should_include_vba(config) -> bool:
    """Whether to collect the Access-COM test files this run.

    B3 (no silent shrinkage): on a COM-capable Windows box the full suite
    is ON BY DEFAULT — "running the tests" must not silently drop the whole
    COM behavioural suite.  Opt out explicitly with --fast.  --include-vba
    forces ON anywhere.  Non-Windows / no-pywin32 stays OFF so the fast
    suite still collects cleanly there.
    """
    if config.getoption("--fast"):
        return False
    if config.getoption("--include-vba"):
        return True
    return _com_capable()


# Set True only once an Access-COM test actually ENTERS setup this session.
# pytest_sessionfinish's MSACCESS taskkill keys off THIS, not off collection
# eligibility — otherwise --collect-only or a pure-unit targeted run on a
# COM-capable Windows box would kill the developer's unrelated Access windows.
_ACCESS_TESTS_EXECUTED = False


def pytest_runtest_setup(item):
    global _ACCESS_TESTS_EXECUTED
    try:
        name = Path(str(item.fspath)).name
    except Exception:
        return
    if _is_access_com_test_file(name):
        _ACCESS_TESTS_EXECUTED = True


def pytest_ignore_collect(collection_path, config):
    """Skip *importing* Access-COM-dependent test files by default.

    Every test file that uses `cbdb_driver.vba_session.VbaSession` /
    win32com / pywinauto imports those Windows-only modules at top
    level.  Just adding a skip marker (the previous approach) doesn't
    help — pytest still imports the file during collection, which
    fails on Linux / headless / fresh machines that lack pywin32.

    Skipping the file at collect-time is the only way to keep
    `pytest tests/ -W ignore` green on a non-Windows box.  On a
    COM-capable Windows box the COM suite is auto-included (B3); use
    `--fast` to skip it, or `--include-vba` to force it anywhere.
    """
    if _should_include_vba(config):
        return False
    return _is_access_com_test_file(collection_path.name)


@pytest.fixture(scope="session")
def user_mdb_path(request) -> Path:
    p = Path(request.config.getoption("--user-mdb")).resolve()
    if not p.exists():
        pytest.exit(f"user mdb not found: {p}")
    return p


def _resolve_data_mdb(root: Path) -> Path | None:
    """Resolve the DATA mdb in data/.

    Delegates to analysis/discover_test_inputs._find_data_mdb() so both
    the conftest gate and the discovery script use identical selection
    logic (pick newest by YYYYMMDD when multiple exist).

    Returns None only when data/ has no CBDB_*_DATA.mdb files at all
    (FileNotFoundError from _find_data_mdb is caught and converted).

    NOTE: DATA mdb is always resolved from ROOT/data regardless of the
    --user-mdb option.  The standard workflow (AGENTS.md) keeps both
    files in the same repo data/ directory.  If --user-mdb points to an
    external path with a different DATA mdb, update data/ accordingly.
    """
    # Import _find_data_mdb from the repo's analysis/ directory (always the
    # same location regardless of the `root` parameter, which can be a
    # tmp_path in tests).
    # Delegates to analysis/_data_mdb_finder.find_data_mdb — a
    # side-effect-free module (no pyodbc, no module-level file access)
    # that is the single source of truth shared with discover_test_inputs.
    _analysis = TESTS_DIR.parent / "analysis"
    if str(_analysis) not in sys.path:
        sys.path.insert(0, str(_analysis))
    try:
        from _data_mdb_finder import find_data_mdb  # type: ignore[import]
        return find_data_mdb(root)
    except FileNotFoundError:
        return None
    except ImportError:
        # Module not found — shouldn't happen in a normal repo checkout.
        # Fall back to single-match-only glob so collection stays clean.
        matches = list((root / "data").glob("CBDB_*_DATA.mdb"))
        return matches[0] if len(matches) == 1 else None


def _refresh_decision(
    inputs_json: Path,
    user_mdb: Path,
    data_mdb: Path | None,
) -> tuple[str, str]:
    """Decide whether to refresh `analysis/dump/test_inputs.json`.

    Returns ``(action, reason)`` where ``action`` is ``"skip"``
    or ``"refresh"`` and ``reason`` is one of:
      - ``"no_user_mdb"``  — user mdb missing; can't gate
      - ``"missing"``       — test_inputs.json doesn't exist
      - ``"stale_user_mdb"`` — older than user mdb
      - ``"stale_data_mdb"`` — older than linked DATA mdb
      - ``"fresh"``         — newer than both gates

    Pure file-mtime logic; no COM, no pyodbc.  Tested in
    `tests/test_infra_refresh_decision.py`.
    """
    if not user_mdb.exists():
        return ("skip", "no_user_mdb")
    if not inputs_json.exists():
        return ("refresh", "missing")
    inputs_mtime = inputs_json.stat().st_mtime
    if inputs_mtime < user_mdb.stat().st_mtime:
        return ("refresh", "stale_user_mdb")
    if data_mdb is not None and data_mdb.exists():
        if inputs_mtime < data_mdb.stat().st_mtime:
            return ("refresh", "stale_data_mdb")
    return ("skip", "fresh")


def _data_mdb_relink_needed(user_mdb: Path, data_mdb: Path | None) -> bool:
    """Return True if LinkedTables in user_mdb are stale vs. data_mdb.

    Reads LinkListInit.c_dataset from the user mdb via pyodbc (a local
    table — readable even if the linked DATA tables are broken) and
    compares it against the YYYYMMDD embedded in data_mdb's filename.

    Returns False on any error (no pyodbc, no user mdb, etc.) so that
    the fast path never blocks headless / Linux collection.
    """
    if data_mdb is None or not user_mdb.exists():
        return False
    try:
        import pyodbc as _pyodbc
    except ImportError:
        return False   # pyodbc not installed (headless / Linux) — safe no-op
    try:
        conn = _pyodbc.connect(
            "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
            f"DBQ={user_mdb};ReadOnly=True;",
            autocommit=True,
        )
        cur = conn.cursor()
        cur.execute("SELECT c_dataset FROM LinkListInit")
        row = cur.fetchone()
        conn.close()
        if row is None:
            return True
        stored = (row[0] or "").strip()
        parts = data_mdb.stem.split("_")  # ["CBDB", "YYYYMMDD", "DATA"]
        current = parts[1] if len(parts) >= 2 else ""
        return stored != current
    except Exception as e:
        # On a real Windows+Access box, ODBC failures here are genuine
        # problems (locked MDB, broken driver, unreadable LinkListInit).
        # Warn so the user knows the relink check was skipped, rather
        # than silently proceeding with potentially stale links.
        print(
            f"\n[conftest] WARNING: could not check LinkListInit.c_dataset "
            f"({type(e).__name__}: {e}). "
            f"If linked tables are stale, run: "
            f"python analysis/relink_data_mdb.py"
        )
        return False


def pytest_sessionfinish(session, exitstatus):
    """Kill ONLY the MSACCESS.EXE processes THIS session spawned (B10).

    Scoped per-PID: the AccessApp/VbaSession fixtures register every PID
    they open in the cbdb_driver spawned-PID registry, and this hook kills
    only those — NEVER the developer's unrelated Access windows.  Gated on
    _ACCESS_TESTS_EXECUTED so --collect-only / --fast / pure-unit runs do
    nothing.  Fixture close()s normally kill their own PID already; this is
    the safety net for when a dialog block or interrupt skipped teardown.
    """
    if sys.platform != "win32":
        return
    if not _ACCESS_TESTS_EXECUTED:
        return
    try:
        from cbdb_driver.access_app import spawned_pids, kill_access_pid
    except Exception:
        return  # COM driver never imported => nothing we spawned to clean up
    pids = spawned_pids()
    if not pids:
        return
    for pid in pids:
        try:
            kill_access_pid(pid)  # taskkill /F /PID + wait; harmless if gone
        except Exception:
            pass
    # Belt-and-suspenders for any of OUR pids that survived taskkill (modal
    # dialog / pending I/O).  Re-check the name so a RECYCLED pid that is no
    # longer Access is never killed.
    try:
        import psutil
        import time
        time.sleep(0.3)
        for pid in pids:
            if not psutil.pid_exists(pid):
                continue
            try:
                proc = psutil.Process(pid)
                if (proc.name() or "").upper().startswith("MSACCESS"):
                    proc.kill()
                    proc.wait(timeout=3)
            except (psutil.NoSuchProcess, psutil.AccessDenied,
                    psutil.TimeoutExpired):
                pass
    except ImportError:
        pass


def pytest_configure(config):
    """Register markers, relink DATA mdb if stale, refresh test_inputs.json.

    Step 1 — DATA mdb relink: if LinkListInit.c_dataset in the User mdb
    doesn't match the YYYYMMDD of the DATA mdb found in data/, the linked
    tables are stale.  We call analysis/relink_data_mdb.py to fix them
    before any test can open the User mdb.  Stale links silently corrupt
    every pyodbc-based test (SQL replay, saved-view checks, etc.).

    Step 2 — test_inputs.json refresh (existing logic).

    Disable both steps with: pytest --no-discover-inputs
    """
    config.addinivalue_line(
        "markers",
        "access: requires a running Access COM session "
        "(Windows + Office + a working data/CBDB_BJ_User.mdb).",
    )

    # --- Build pin (B7 part 2): fail loudly if a pinned expected build
    # doesn't match the DATA mdb actually in data/.  No pin set => no-op
    # (record-only).  Runs even with --no-discover-inputs.
    _ana = str(ROOT / "analysis")
    if _ana not in sys.path:
        sys.path.insert(0, _ana)
    try:
        from build_stamp import build_pin_error  # type: ignore[import]
        _pin_err = build_pin_error(ROOT)
    except ImportError:
        _pin_err = None
    if _pin_err:
        pytest.exit(f"[conftest] build pin: {_pin_err}")

    # B3: tell the operator when the heavy COM suite is auto-included, so a
    # plain `pytest tests/` on Windows is never a silent surprise (or a silent
    # smaller set).  --fast opts out; --include-vba is the explicit force.
    if _should_include_vba(config) and not config.getoption("--include-vba"):
        print("\n[conftest] Access-COM/VBA suite AUTO-INCLUDED (capable Windows "
              "box).  This is the full standardized suite (slow).  Pass --fast "
              "to run only the fast subset.")

    if config.getoption("--no-discover-inputs"):
        return

    import subprocess

    # Respect --user-mdb if provided (mirrors user_mdb_path fixture logic).
    raw_user_mdb = config.getoption("--user-mdb")
    mdb = Path(raw_user_mdb).resolve() if raw_user_mdb else ROOT / "data" / "CBDB_BJ_User.mdb"
    data_mdb = _resolve_data_mdb(ROOT)

    # --- Step 1: relink if c_dataset is stale ---
    if _data_mdb_relink_needed(mdb, data_mdb):
        print(f"\n[conftest] LinkListInit.c_dataset stale; "
              f"relinking to {data_mdb.name} ...")
        try:
            rc = subprocess.run(
                [sys.executable, str(ROOT / "analysis" / "relink_data_mdb.py"),
                 "--user-mdb", str(mdb)],  # pass resolved mdb path
                capture_output=True, text=True,
                timeout=120,  # DAO/ACE can hang; don't block pytest forever
            )
        except subprocess.TimeoutExpired:
            pytest.exit(
                "[conftest] relink_data_mdb.py timed out after 120s."
                "\nDAO/ACE may be hung (another Access process holding the MDB?)."
                "\nKill any MSACCESS.EXE processes, then re-run."
                "\nPass `--no-discover-inputs` to skip the relink gate."
            )
        if rc.returncode != 0:
            pytest.exit(
                f"[conftest] relink_data_mdb.py FAILED (rc={rc.returncode})."
                f"\nStale linked tables would silently corrupt all SQL-replay"
                f" tests.  Fix the relink error or pass `--no-discover-inputs`"
                f" to skip.\n\n  stderr:\n{rc.stderr[-1000:]}"
            )
        print(f"[conftest] relink complete.")

    # --- Step 2: refresh test_inputs.json if stale ---
    inputs_json = ROOT / "analysis" / "dump" / "test_inputs.json"
    action, reason = _refresh_decision(inputs_json, mdb, data_mdb)
    if action == "skip":
        if reason == "fresh":
            print(f"\n[conftest] {inputs_json.name} fresh; "
                  f"skipping discovery")
        return
    # action == "refresh"
    print(f"\n[conftest] refreshing {inputs_json.name} "
          f"(reason: {reason}) ...")
    rc = subprocess.run(
        [sys.executable, str(ROOT / "analysis" / "discover_test_inputs.py")],
        capture_output=True, text=True,
    )
    if rc.returncode != 0:
        pytest.exit(
            f"[conftest] discover_test_inputs.py FAILED (rc="
            f"{rc.returncode}).  Tests would otherwise run against "
            f"a stale fixture file ({inputs_json.name}).  Fix the "
            f"discovery error or pass `--no-discover-inputs` to "
            f"skip refresh.\n\n  stderr tail:\n{rc.stderr[-1000:]}"
        )
    print(f"[conftest] discovery refreshed.")


@pytest.hookimpl(optionalhook=True)
def pytest_json_modifyreport(json_report):
    """Stamp the DATA build into the pytest JSON report (B7 part 2).

    Fires only when pytest-json-report is active (--json-report).  Marked
    optionalhook so pluggy doesn't error when that plugin isn't installed.
    Lets the generated report + any consumer know which build this run tested.
    """
    try:
        _ana = str(ROOT / "analysis")
        if _ana not in sys.path:
            sys.path.insert(0, _ana)
        from build_stamp import build_stamp as _stamp  # type: ignore[import]
        json_report.setdefault("environment", {}).update(_stamp(ROOT))
    except Exception:
        # Stamping is best-effort metadata; never fail the report over it.
        pass


@pytest.fixture(scope="session")
def ro_conn(user_mdb_path):
    """Read-only connection to the production user-mdb."""
    from cbdb_replay.driver import open_connection
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
    # Lazy imports — only the COM-driven test fixtures pull in pywin32
    # and pyodbc, so the fast suite stays collectable on machines
    # without Office or even Windows.
    import pyodbc as _pyodbc
    from cbdb_driver import AccessApp, VbaInjector
    from cbdb_driver.access_app import make_working_copy

    # Used to call `kill_orphan_access()` here unconditionally, which
    # also killed any Access DB the developer was editing manually.
    # Each AccessApp now does a scoped per-PID kill in close(), so a
    # clean shutdown leaves nothing to clean up.  If a previous run
    # crashed and left orphan MSACCESS.EXE processes blocking the
    # working copy, run `CBDB_KILL_ALL_ACCESS=1 python -c "from
    # cbdb_driver.access_app import kill_orphan_access; kill_orphan_access()"`
    # once, then re-run the suite.
    work = make_working_copy(user_mdb_path, WORK_MDB)
    # Pre-patch LinkListInit BEFORE Access opens, otherwise
    # NAVIGATION_PANE.Form_Open hangs trying to relink to a non-existent
    # CBDB_<ver>_DATA.mdb at our working path.
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
def driver(com_app):
    from cbdb_driver import FormDriver
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
