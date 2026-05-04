"""Probe LookAtNetworks Form_Open hang — diagnose, don't fix.

Background
----------
`tests/test_vba_matrix_hard_forms.py` skips `LookAtNetworks` because
opening the form via the test driver hangs.  README roadmap item 7
notes that an earlier attempt to rewrite `Forms!LookAtNetworks!<sub>`
→ `Me!<sub>` self-references in `Form_Open` did not unblock it (the
hang is somewhere deeper).

Static read of `analysis/dump/vba/Form_LookAtNetworks.vb` (Form_Open
at line 6132) shows the suspect surface:

  - 3 subform-recordset swaps that touch `Forms!LookAtNetworks!
    <sub>.Form.Recordset` AND/OR `<sub>.Form.Recordset` directly
    (lines 6173, 6177, 6189, 6195, 6200, 6204, 6210).
  - 6 `cmdDel.Execute` DELETE statements against scratch tables that
    those same subforms have as their RecordSource (lines 6181, 6182,
    6192, 6193, 6207, 6208, 6227, 6228, 6230, 6231).
  - One `DCount("*", "ZZ_STORE_PERSON_ID")` (line 6163) — should be
    fast.

Hypothesis
----------
Subform RecordSource binding loads scratch tables (ZZ_SOCIAL_NETWORK
/ ZZ_SCRATCH_PEOPLE / ZZ_SOCIAL_NETWORK_AGGREGATE) at form load.  If
those scratch tables have many rows from prior use (the test driver's
`reset_pickers` does NOT include them — see PICKER_TABLES), Access
spends time materialising them; Form_Open then immediately tries to
swap the recordset out via DAO and DELETE the rows, which can
deadlock against the still-loading subform.

This probe verifies that hypothesis (or rules it out) without
shipping a fix.

Probe design
------------
Each variation runs in its **own Access process** so a hang in one
variation cannot block the next.  Each variation is wrapped in a
watchdog thread with a hard timeout; on timeout we taskkill the
Access PID and record `hung_at_<last_marker>`.

Variations:

  V1  default                — fresh source-copy mdb, OpenForm
                              ("LookAtNetworks", acFormView, ...,
                              WindowMode=acWindowNormal=0).
                              Reproduces the known hang.

  V2  hidden                 — WindowMode=acHidden=1.

  V3  pre-clear scratch      — DELETE FROM ZZ_SOCIAL_NETWORK,
                              ZZ_SCRATCH_PEOPLE,
                              ZZ_SOCIAL_NETWORK_AGGREGATE,
                              ZZ_SCRATCH_ADDR_LIST,
                              ZZ_SCRATCH_IMPORT_PEOPLE
                              **before** OpenForm.  Tests whether
                              the hang comes from subform
                              materialising stale rows.

  V4  pre-clear + hidden     — combined V2 + V3.

Pre-flight (no Access required): query the row counts in the
suspect scratch tables on the source mdb so we know the
materialisation cost.

Outputs
-------
  - reports/lookatnetworks_form_open_hang_probe.json
  - analysis/lookatnetworks_form_open_hang.md (companion note,
    written by the human reading the JSON; this script does NOT
    overwrite it).

Usage
-----
  python analysis/probe_lookatnetworks_form_open.py

Requires Access COM (it spawns Access).  Each variation has a 30 s
hard timeout.  Total runtime ~2–4 min.
"""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
import threading
import time
import traceback
from pathlib import Path

import pyodbc

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))

USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK_BASE = ROOT / "analysis" / "_probe_lan_copy"   # per-variation suffix
OUT_JSON = ROOT / "reports" / "lookatnetworks_form_open_hang_probe.json"

# Per-variation hard timeout for the OpenForm call.
OPENFORM_TIMEOUT_SEC = 30

# Subform/scratch tables Form_Open touches.
SUSPECT_SCRATCH_TABLES = [
    "ZZ_SOCIAL_NETWORK",
    "ZZ_SCRATCH_PEOPLE",
    "ZZ_SOCIAL_NETWORK_AGGREGATE",
    "ZZ_SCRATCH_ADDR_LIST",
    "ZZ_SCRATCH_IMPORT_PEOPLE",
]


def _open_user(mdb: Path, read_only: bool = True) -> pyodbc.Connection:
    return pyodbc.connect(
        "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={mdb};", autocommit=True, readonly=read_only)


# --------------------------------------------------------------------
# Pre-flight: row counts of suspect scratch tables on source mdb
# (no Access COM)
# --------------------------------------------------------------------
def preflight_row_counts() -> dict[str, int]:
    out = {}
    conn = _open_user(USER_MDB)
    cur = conn.cursor()
    for t in SUSPECT_SCRATCH_TABLES:
        try:
            cur.execute(f"SELECT COUNT(*) FROM [{t}]")
            out[t] = int(cur.fetchone()[0])
        except pyodbc.Error as e:
            out[t] = f"ERROR: {e}"
    conn.close()
    return out


# --------------------------------------------------------------------
# Variation runner — each spawns one Access process, attempts an
# OpenForm under specified conditions, with a watchdog timeout.
# --------------------------------------------------------------------
def _kill_access_pid(pid: int) -> None:
    if pid is None:
        return
    try:
        subprocess.run(
            ["taskkill", "/F", "/PID", str(pid)],
            capture_output=True, timeout=10)
    except Exception:
        pass


def run_variation(name: str, *, hidden: bool, pre_clear: bool,
                   work_mdb: Path) -> dict:
    """Run a single variation.  Returns a result dict."""
    from cbdb_driver.vba_session import VbaSession  # noqa: F401
    from cbdb_driver.access_app import (
        kill_access_pid as _kap,  # type: ignore
    )

    result: dict = {
        "variation": name,
        "hidden": hidden,
        "pre_clear": pre_clear,
        "markers": [],
        "outcome": None,
        "elapsed_sec": None,
        "exception": None,
    }
    t0 = time.time()
    sess = None
    pid_for_kill = None

    def mark(s: str) -> None:
        result["markers"].append({"t": round(time.time() - t0, 2), "marker": s})

    try:
        # Fresh copy + fresh Access process per variation.
        if work_mdb.exists():
            try:
                work_mdb.unlink()
            except Exception:
                _kill_access_pid(pid_for_kill)
                time.sleep(1)
                work_mdb.unlink()
        shutil.copy2(USER_MDB, work_mdb)
        mark("copied_mdb")

        # CRITICAL — patch LinkListInit so NAVIGATION_PANE.
        # SetLinkTables fast-paths through its `If c_path =
        # CurrentProject.FullName Then Exit Sub` check
        # (Form_NAVIGATION_PANE.vb:265) and never tries to
        # re-relink to a non-existent
        # `<work_mdb_basename - 12 chars>_<date>_DATA.mdb`
        # (which would pop a blocking modal dialog).  Set
        # c_path = work_mdb path, NOT USER_MDB — the relink
        # logic at line 282 uses CurrentProject.FullName, which
        # will be the work mdb at runtime, so c_path must match
        # it for Exit Sub to fire.
        conn_link = _open_user(work_mdb, read_only=False)
        try:
            link_value = str(work_mdb).replace("'", "''")
            conn_link.cursor().execute(
                f"UPDATE LinkListInit SET c_path = '{link_value}'"
            )
            mark("patched_LinkListInit_to_work_mdb")
        except pyodbc.Error as e:
            mark(f"LinkListInit_warn: {e}")
        finally:
            conn_link.close()

        # Pre-clear: DELETE the suspect tables BEFORE we ever
        # invoke Access COM.
        if pre_clear:
            conn = _open_user(work_mdb, read_only=False)
            cur = conn.cursor()
            for t in SUSPECT_SCRATCH_TABLES:
                try:
                    cur.execute(f"DELETE FROM [{t}]")
                except pyodbc.Error as e:
                    mark(f"pre_clear_warn_{t}: {e}")
            conn.close()
            mark("pre_cleared")

        # Spin up Access via VbaSession (lighter than re-implementing
        # the full open() routine).  We bypass `reset_pickers` and
        # `_inject_autodetect` — they're irrelevant to the hang and
        # add side-effects.
        from cbdb_driver.vba_session import VbaSession
        # Construct directly, avoid VbaSession.open() side effects.
        import win32com.client
        app = win32com.client.DispatchEx("Access.Application")
        try:
            app.AutomationSecurity = 1
        except Exception:
            pass
        app.Visible = not hidden
        app.OpenCurrentDatabase(str(work_mdb))
        mark("access_opened")

        # Capture PID for kill-on-hang.
        try:
            from cbdb_driver.access_app import _pid_for_access_app  # type: ignore
            pid_for_kill = _pid_for_access_app(app)
        except Exception:
            pid_for_kill = None
        mark(f"pid={pid_for_kill}")

        # OpenForm in a worker thread with a hard timeout.
        # acFormView=0; FilterName=""; WhereCondition="";
        # DataMode=acFormEdit=2; WindowMode=acHidden=1 OR
        # acWindowNormal=0; OpenArgs=""
        window_mode = 1 if hidden else 0
        worker_done = threading.Event()
        worker_exc: list[BaseException] = []

        def _worker():
            try:
                app.DoCmd.OpenForm(
                    "LookAtNetworks", 0, "", "",
                    2, window_mode,
                )
                worker_done.set()
            except BaseException as e:  # noqa: BLE001
                worker_exc.append(e)
                worker_done.set()

        mark("worker_starting_OpenForm")
        worker = threading.Thread(target=_worker, daemon=True)
        worker.start()
        finished = worker_done.wait(timeout=OPENFORM_TIMEOUT_SEC)
        mark(f"worker_finished={finished}")

        if not finished:
            result["outcome"] = "hung_at_OpenForm"
            # Hard kill — the worker thread is wedged in COM RPC
            # and won't return.  Killing the Access process
            # unwedges the RPC, releasing the worker thread to die.
            _kill_access_pid(pid_for_kill)
        elif worker_exc:
            result["outcome"] = "exception_at_OpenForm"
            result["exception"] = repr(worker_exc[0])
            try:
                app.DoCmd.Close(2, "LookAtNetworks", 2)
            except Exception:
                pass
            try:
                app.Quit()
            except Exception:
                pass
            _kill_access_pid(pid_for_kill)
        else:
            result["outcome"] = "succeeded"
            try:
                app.DoCmd.Close(2, "LookAtNetworks", 2)
            except Exception:
                pass
            try:
                app.Quit()
            except Exception:
                pass
            _kill_access_pid(pid_for_kill)
        mark("cleaned_up")
    except BaseException as e:  # noqa: BLE001
        result["outcome"] = "exception_before_OpenForm"
        result["exception"] = repr(e) + "\n" + traceback.format_exc()
        _kill_access_pid(pid_for_kill)
    finally:
        result["elapsed_sec"] = round(time.time() - t0, 2)
    return result


def main() -> int:
    print(f"=== LookAtNetworks Form_Open hang probe ===")
    print()
    print("Pre-flight: row counts on source mdb for suspect "
          "scratch tables...")
    preflight = preflight_row_counts()
    for t, n in preflight.items():
        print(f"  {t}: {n}")
    print()

    variations = [
        ("V1_default",                 dict(hidden=False, pre_clear=False)),
        ("V2_hidden",                  dict(hidden=True,  pre_clear=False)),
        ("V3_pre_clear_visible",       dict(hidden=False, pre_clear=True)),
        ("V4_pre_clear_hidden",        dict(hidden=True,  pre_clear=True)),
    ]
    results = []
    for name, kw in variations:
        print(f"--- {name} (hidden={kw['hidden']}, "
              f"pre_clear={kw['pre_clear']}) ---")
        work = WORK_BASE.with_suffix(f".{name}.mdb")
        r = run_variation(name, work_mdb=work, **kw)
        results.append(r)
        print(f"  outcome: {r['outcome']}, "
              f"elapsed: {r['elapsed_sec']}s")
        for m in r["markers"]:
            print(f"    +{m['t']}s {m['marker']}")
        if r.get("exception"):
            print(f"    exception: {r['exception'][:200]}")
        print()
        # Small pause between variations to let any zombies settle.
        time.sleep(2)

    out = {
        "preflight_row_counts": preflight,
        "openform_timeout_sec": OPENFORM_TIMEOUT_SEC,
        "variations": results,
        "summary_by_outcome": {
            r["variation"]: r["outcome"] for r in results
        },
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"wrote {OUT_JSON}")
    print()
    print("=== summary ===")
    for v in results:
        print(f"  {v['variation']:30s} → {v['outcome']} "
              f"(in {v['elapsed_sec']}s)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
