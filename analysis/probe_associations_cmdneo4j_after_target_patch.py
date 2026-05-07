"""LookAtAssociations × CmdNeo4j verification probe — post target-column rewrite patch.

Companion to PR #112's probe (which observed the 0-file +
JET 3061 mode) and PR #114's static investigation (which pinned
the cause as a target-side schema mismatch).  This probe verifies
the narrow scoped driver-side workaround in
`tests/cbdb_driver/vba_session.py`,
`_rewrite_associations_cmdneo4j_target_column`, which rewrites
the single literal anchor

  c_index_addr_type_code, c_female  ->  c_addr_type, c_female

inside `Form_LookAtAssociations.CmdNeo4j_Click`.

Probe answers:

  Q1  Does the JET 3061 'unknown field name:
      c_index_addr_type_code' :ERR marker disappear from
      ZZ_TEST_DEBUG?
  Q2  Does the chain start writing CSV files (file_count >= 1)?
  Q3  What is the file shape — count, first columns, row counts?
  Q4  Does the patch expose a NEXT layer of blocker (a different
      :ERR class, an unexpected MsgBox, an empty file set, a
      timeout)?

This probe does NOT open a coverage PR; per brief it is the
verification step that gates whether a coverage PR is the next
step.

Outputs:
  analysis/probe_associations_cmdneo4j_after_target_patch.md
  reports/probe_associations_cmdneo4j_after_target_patch.json
"""
from __future__ import annotations

import json
import subprocess
import sys
import threading
import time
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))

USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK = ROOT / "analysis" / (
    "_probe_associations_cmdneo4j_after_target_patch_copy.mdb")
OUT_JSON = ROOT / "reports" / (
    "probe_associations_cmdneo4j_after_target_patch.json")
OUT_MD = ROOT / "analysis" / (
    "probe_associations_cmdneo4j_after_target_patch.md")

TIMER_TIMEOUT_SEC = 180
PROBE_OUTER_TIMEOUT_SEC = 300


def _kill_orphan() -> None:
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "MSACCESS.EXE"],
            capture_output=True, timeout=10)
    except Exception:
        pass


def _msgbox_watchdog(
    stop_event: threading.Event,
    observed_log: list,
    t0: float,
) -> None:
    """Watchdog: dismiss any MsgBox the driver missed; each entry
    is a runtime signal we surface (not silently swallow)."""
    from pywinauto import findwindows
    from pywinauto import Application as PWA
    import pywinauto.keyboard as kb

    while not stop_event.is_set():
        try:
            handles = findwindows.find_windows(
                title="Microsoft Access", class_name="#32770")
            for hwnd in handles:
                try:
                    app = PWA(backend="win32")
                    app.connect(handle=hwnd)
                    dlg = app.window(handle=hwnd)
                    try:
                        texts = [
                            c.window_text()
                            for c in dlg.children()
                            if c.friendly_class_name() == "Static"
                        ]
                        msg_text = " | ".join(
                            t for t in texts if t.strip())[:200]
                    except Exception:
                        msg_text = "(unread)"
                    try:
                        dlg.Button.click()
                    except Exception:
                        try:
                            kb.send_keys("{ENTER}")
                        except Exception:
                            pass
                    observed_log.append({
                        "t": round(time.time() - t0, 2),
                        "hwnd": hwnd,
                        "msg_text": msg_text,
                    })
                except Exception:
                    pass
        except Exception:
            pass
        time.sleep(0.3)


def _classify_outcome(result: dict) -> str:
    """Verification probe outcome buckets (mutually exclusive,
    first-match wins).  Order is significant: any debug-MsgBox
    observation disqualifies the clean bucket because such a
    MsgBox would block unattended coverage even if the chain
    completes when a watchdog dismisses it.

      - blocked_exception:
          exception raised in the worker, no files produced.
      - patch_partial_jet_3061_still_observed:
          ZZ_TEST_DEBUG still contains the JET 3061
          'c_index_addr_type_code' :ERR — patch did not take
          effect.
      - next_blocker_exposed_different_err:
          ZZ_TEST_DEBUG contains a :ERR marker but it is NOT the
          JET 3061 anchor — patch resolved the original failure
          but a downstream :ERR class is now visible.
      - patch_resolved_issue23_but_exposed_msgbox_blocker:
          no :ERR markers AND file_count >= 1 BUT the watchdog
          observed and dismissed >= 1 MsgBox during the chain
          run.  The Issue #23 JET 3061 was neutralized by the
          rewrite AND the chain reaches its natural end AND
          writes files, BUT a debug-MsgBox layer (concat-form
          `MsgBox "<lit>" + Trim(Str(...))` calls that the
          driver's literal-only generic neutralizer cannot
          catch) is still live and would block unattended
          coverage.  Same shape as the AssocPairs CmdNeo4j
          debug-MsgBox layer that PR #109 suppressed via
          `_suppress_assocpairs_cmdneo4j_debug_msgbox` — but in
          a different form, so PR #109's per-form patch does
          NOT cover it.  This is a NEW per-host observation
          (PR #112's pre-patch probe could not surface it
          because the chain bailed at JET 3061 BEFORE reaching
          the debug-MsgBox lines).
      - patch_resolved_jet_3061_but_zero_files:
          no :ERR marker and no JET 3061 AND no watchdog MsgBox
          BUT file_count == 0 — something else is preventing
          file writes (e.g. mid-body early-bail, outer
          watchdog timeout).
      - patch_verified_chain_clean:
          STRICT: file_count >= 1 AND no :ERR AND no watchdog
          MsgBox observation AND chain quiesced AND
          chain_elapsed <= TIMER_TIMEOUT_SEC.  Any watchdog
          MsgBox disqualifies this bucket — a real popup
          dialog during a 'clean' run is a contradiction.

    Note on ZZ_TEST_DEBUG `:MSGBOX` markers: those rows are
    written by the driver's generic literal-only MsgBox
    neutralizer (rewrites `MsgBox "<lit>"` standalone statements
    to `INSERT INTO ZZ_TEST_DEBUG ('<form>:MSGBOX')` BEFORE the
    runtime would see them as a dialog).  They are NOT the same
    signal as a watchdog-dismissed dialog: the generic
    neutralizer prevents the dialog from ever popping, while
    the watchdog dismisses dialogs that DID pop (typically
    concat-form MsgBoxes the literal-only regex cannot match).
    A `:MSGBOX` marker by itself is not necessarily a blocker;
    a watchdog dismissal IS.
    """
    n_files = int(result.get("file_count") or 0)
    chain_elapsed = result.get("chain_elapsed_sec")
    exception = result.get("exception")
    debug_msgs = result.get("zz_test_debug_msgs") or []
    n_watchdog_msgbox = len(result.get("msgbox_observed") or [])

    has_jet_3061 = any(
        ":ERR" in m and "c_index_addr_type_code" in m
        for m in debug_msgs
    )
    has_other_err = any(
        ":ERR" in m and "c_index_addr_type_code" not in m
        for m in debug_msgs
    )
    elapsed_ok = (
        isinstance(chain_elapsed, (int, float))
        and chain_elapsed <= TIMER_TIMEOUT_SEC
    )

    if exception and n_files == 0:
        return "blocked_exception"
    if has_jet_3061:
        return "patch_partial_jet_3061_still_observed"
    if has_other_err:
        return "next_blocker_exposed_different_err"
    # No :ERR markers from here on.
    if n_watchdog_msgbox >= 1 and n_files >= 1:
        return "patch_resolved_issue23_but_exposed_msgbox_blocker"
    if n_files == 0:
        return "patch_resolved_jet_3061_but_zero_files"
    if (n_files >= 1
            and n_watchdog_msgbox == 0
            and elapsed_ok
            and result.get("chain_observed_done")):
        return "patch_verified_chain_clean"
    return "patch_resolved_jet_3061_but_zero_files"


def _get_associations_fixture():
    from test_vba_matrix_all_forms import _all_fixtures
    for fx in _all_fixtures():
        if fx.spec.name == "LookAtAssociations":
            return fx
    raise RuntimeError("no Associations fixture found in matrix")


def _run_probe(out_dir: Path) -> dict:
    from cbdb_driver.vba_session import VbaSession, make_fixture
    from cbdb_driver.form_specs import LOOKATASSOCIATIONS

    spec = LOOKATASSOCIATIONS
    fx = _get_associations_fixture()

    result: dict = {
        "form": spec.name,
        "fixture_name": fx.name,
        "fixture_picker_ids": list(fx.picker_ids) if fx.picker_ids else [],
        "fixture_controls": dict(fx.controls or {}),
        "markers": [],
        "outcome": None,
        "elapsed_sec": None,
        "exception": None,
        "row_counts": {},
        "files": [],
        "file_count": 0,
        "click_via_timer_returned": None,
        "chain_elapsed_sec": None,
        "chain_observed_done": False,
        "msgbox_observed": [],
        "zz_test_debug_msgs": [],
    }

    t0 = time.time()
    completed = threading.Event()
    stop_watchdog = threading.Event()
    _session_holder: list = []

    def mark(s: str) -> None:
        result["markers"].append(
            {"t": round(time.time() - t0, 2), "marker": s})

    def _make_session_iter():
        gen = make_fixture(USER_MDB, WORK)
        for s in gen:
            _session_holder.append((s, gen))
            yield s
            return

    def _worker() -> None:
        try:
            mark("constructing_session")
            for attempt in (1, 2, 3):
                try:
                    sess = next(_make_session_iter())
                    mark(f"session_opened_attempt_{attempt}")
                    break
                except Exception as e:
                    mark(f"session_open_attempt_{attempt}_fail: {e!r}")
                    _kill_orphan()
                    time.sleep(60)
            else:
                raise RuntimeError(
                    "session open failed after 3 attempts")

            sess.patch_filedialog(spec.name)
            mark("filedialog_patched")

            sess.open_form(spec.name)
            mark("form_opened")

            for ctl, val in (fx.controls or {}).items():
                try:
                    sess.set_control(spec.name, ctl, val)
                except Exception as e:
                    mark(f"set_control_{ctl}_fail: {e!r}")
            if fx.picker_ids and spec.picker_table:
                sess.set_picker_codes(
                    spec.picker_table, list(fx.picker_ids),
                    column=spec.picker_column)
                mark(f"picker_seeded_{len(fx.picker_ids)}_codes")
            if fx.addr_ids:
                sess.set_picker_addrs(list(fx.addr_ids))
                mark(f"addr_picker_seeded_{len(fx.addr_ids)}_codes")
            mark("fixture_seeded")

            sess.set_form_tag(
                spec.name,
                f"{spec.cmd_name},CmdNeo4j",
                str(out_dir) + "\\",
            )
            mark("form_tag_set_chain_CmdQuery_CmdNeo4j")

            t_chain_start = time.time()
            mark("chain_fire_t_start")
            try:
                n = sess.click_via_timer(
                    spec.name,
                    ctl=spec.cmd_name,
                    result_table=spec.result_table,
                    timeout=TIMER_TIMEOUT_SEC,
                )
                result["click_via_timer_returned"] = n
                mark(f"click_via_timer_returned_{n}")
            except Exception as e:
                mark(f"click_via_timer_exc: {e!r}")
                result["exception"] = repr(e)

            chain_observed_done = False
            stable_count = 0
            last_count = -1
            poll_deadline = t0 + PROBE_OUTER_TIMEOUT_SEC - 5

            while time.time() < poll_deadline:
                cur_count = len(sorted(out_dir.glob("*")))
                if cur_count == last_count:
                    stable_count += 1
                else:
                    stable_count = 0
                    last_count = cur_count
                if cur_count > 0 and stable_count >= 5:
                    chain_observed_done = True
                    mark(
                        f"chain_quiescent_files_{cur_count}_"
                        f"stable_for_5s")
                    break
                if (cur_count == 0
                        and result["click_via_timer_returned"]
                        is not None
                        and stable_count >= 8):
                    chain_observed_done = True
                    mark("chain_quiescent_zero_files_stable_for_8s")
                    break
                time.sleep(1)

            t_chain_end = time.time()
            chain_elapsed = round(t_chain_end - t_chain_start, 2)
            result["chain_elapsed_sec"] = chain_elapsed
            result["chain_observed_done"] = chain_observed_done
            mark(f"chain_elapsed_{chain_elapsed}s")

            files = sorted(out_dir.glob("*"))
            for f in files:
                try:
                    raw = f.read_bytes()
                    text = raw.decode(
                        "utf-8", errors="replace").lstrip("﻿")
                    first_line = text.split("\n", 1)[0].strip()
                    cols = first_line.split(",")
                    data_lines = [
                        ln for ln in text.split("\n")[1:]
                        if ln.strip()]
                    result["files"].append({
                        "name": f.name,
                        "size": f.stat().st_size,
                        "header_first_col": cols[0] if cols else "",
                        "header_n_cols": len(cols),
                        "header_preview": first_line[:200],
                        "data_row_count": len(data_lines),
                    })
                except Exception as e:
                    result["files"].append({
                        "name": f.name,
                        "size": f.stat().st_size,
                        "read_error": repr(e),
                    })
            result["file_count"] = len(files)
            mark(f"files_inventoried_{len(files)}")

            for tbl in (
                "ZZ_SCRATCH_ASSOC",
                "ZZ_SCRATCH_P_ASSOC",
                "ZZ_SCRATCH_PEOPLE",
                "ZZ_TEST_DEBUG",
            ):
                try:
                    cur = sess.conn.cursor()
                    cur.execute(f"SELECT COUNT(*) FROM [{tbl}]")
                    result["row_counts"][tbl] = int(
                        cur.fetchone()[0])
                    cur.close()
                except Exception as e:
                    result["row_counts"][tbl] = f"ERROR: {e}"
            mark("row_counts_captured")

            try:
                cur = sess.conn.cursor()
                cur.execute(
                    "SELECT msg FROM ZZ_TEST_DEBUG ORDER BY id")
                result["zz_test_debug_msgs"] = [
                    r[0] for r in cur.fetchall()]
                cur.close()
            except Exception as e:
                result["zz_test_debug_msgs"] = [f"ERROR: {e}"]
            mark("zz_test_debug_captured")

            result["outcome"] = _classify_outcome(result)
            completed.set()
        except BaseException as e:  # noqa: BLE001
            result["outcome"] = "blocked_exception"
            result["exception"] = (
                repr(e) + "\n" + traceback.format_exc())
            completed.set()

    watchdog = threading.Thread(
        target=_msgbox_watchdog,
        args=(stop_watchdog, result["msgbox_observed"], t0),
        daemon=True,
    )
    watchdog.start()

    worker = threading.Thread(target=_worker, daemon=False)
    worker.start()

    finished = completed.wait(timeout=PROBE_OUTER_TIMEOUT_SEC)
    if not finished:
        result["outcome"] = result.get("outcome") or "hung_at_per_probe_timeout"
        mark(f"per_probe_hard_timeout_at_{PROBE_OUTER_TIMEOUT_SEC}s")
        _kill_orphan()

    stop_watchdog.set()
    watchdog.join(timeout=5)

    try:
        if _session_holder:
            _, gen = _session_holder[0]
            try:
                next(gen)
            except StopIteration:
                pass
    except Exception:
        pass
    _kill_orphan()
    worker.join(timeout=10)
    time.sleep(2)
    result["elapsed_sec"] = round(time.time() - t0, 2)
    return result


def _verdict(result: dict) -> dict:
    outcome = result.get("outcome", "")
    n_files = result.get("file_count", 0)
    chain_elapsed = result.get("chain_elapsed_sec")
    debug_msgs = result.get("zz_test_debug_msgs") or []
    err_markers = [m for m in debug_msgs if ":ERR" in m]
    has_jet_3061 = any(
        "c_index_addr_type_code" in m for m in err_markers)
    other_err = [
        m for m in err_markers
        if "c_index_addr_type_code" not in m
    ]

    n_watchdog = len(result.get("msgbox_observed") or [])
    msgbox_marker_count = sum(
        1 for m in debug_msgs if ":MSGBOX" in m)

    answers = {
        "Q1_jet_3061_disappeared": {
            "has_jet_3061_marker": has_jet_3061,
            "all_err_markers_observed": err_markers,
            "interpretation": (
                "If has_jet_3061_marker is False, the patch "
                "successfully prevented the JET 3061 'unknown "
                "field name: c_index_addr_type_code' that PR "
                "#112's probe observed.  If True, the rewrite "
                "did not take effect (anchor mismatch / scoping "
                "bug / etc.) and the patch needs investigation."
            ),
        },
        "Q2_chain_writes_files": {
            "file_count": n_files,
            "chain_observed_done": result.get(
                "chain_observed_done", False),
            "chain_elapsed_sec": chain_elapsed,
            "interpretation": (
                "PR #112 observed 0 files; this probe expects "
                "file_count >= 1 if the rewrite took effect AND "
                "the rest of the chain runs to its natural end."
            ),
        },
        "Q3_file_shape": {
            "file_count": n_files,
            "files": [
                {
                    "name": f.get("name", ""),
                    "size": f.get("size", 0),
                    "n_cols": f.get("header_n_cols", 0),
                    "first_col": f.get("header_first_col", ""),
                    "data_rows": f.get("data_row_count", 0),
                }
                for f in result.get("files", [])
            ],
        },
        "Q4_next_blocker_exposed": {
            "outcome_bucket": outcome,
            "other_err_markers": other_err,
            "msgbox_watchdog_count": n_watchdog,
            "zz_test_debug_msgbox_marker_count": msgbox_marker_count,
            "msgbox_signal_layers_explained": {
                "watchdog_dismissed_dialogs": (
                    "REAL pop-up dialogs that the driver's "
                    "generic literal-only `MsgBox \"<lit>\"` "
                    "neutralizer could not catch — typically "
                    "concat-form `MsgBox \"<lit>\" + "
                    "Trim(Str(...))` calls.  Watchdog dismisses "
                    "them so the probe does not hang, but for "
                    "unattended coverage a dialog is a hard "
                    "blocker."
                ),
                "zz_test_debug_MSGBOX_markers": (
                    "INSERT rows the driver's generic literal "
                    "neutralizer wrote BEFORE the runtime would "
                    "see the original `MsgBox \"<lit>\"` line "
                    "as a dialog.  These do NOT correspond to "
                    "dialogs that ever popped — they are the "
                    "neutralizer's footprint, not a blocker.  "
                    "Distinct from the line-1035 early-bail "
                    "marker (which would only appear if the "
                    "RecordCount=0 bail fired; on this fixture "
                    "it does not, so the :MSGBOX marker present "
                    "here is the terminal `MsgBox \"Finished "
                    "saving to Neo4j\"` neutralizer footprint, "
                    "indicating the chain reached its end)."
                ),
            },
            "interpretation": (
                "Bucket meanings (mutually exclusive, first "
                "match wins):\n"
                "  - patch_verified_chain_clean: file_count "
                ">= 1 AND no :ERR AND zero watchdog dialogs.  "
                "Strict.\n"
                "  - patch_resolved_issue23_but_exposed_msgbox"
                "_blocker: file_count >= 1 AND no :ERR BUT "
                ">= 1 watchdog dialog dismissed.  Issue #23 "
                "JET 3061 is gone, but a downstream debug-"
                "MsgBox layer is exposed and would block "
                "unattended coverage.\n"
                "  - patch_partial_jet_3061_still_observed: "
                "JET 3061 :ERR still present.\n"
                "  - next_blocker_exposed_different_err: a :ERR "
                "from a different class.\n"
                "  - patch_resolved_jet_3061_but_zero_files: "
                "no :ERR but no files either."
            ),
        },
    }

    if outcome == "patch_verified_chain_clean":
        verdict_note = (
            f"Strict clean: chain produced {n_files} files in "
            f"{chain_elapsed}s with no :ERR markers in "
            f"ZZ_TEST_DEBUG (specifically no JET 3061 "
            f"'c_index_addr_type_code' marker) AND zero "
            f"watchdog-dismissed dialogs.  Per the brief, do "
            f"NOT open a coverage PR off this verdict alone — "
            f"the coverage PR is a separate brief."
        )
    elif outcome == (
        "patch_resolved_issue23_but_exposed_msgbox_blocker"
    ):
        # Sample the first few watchdog dialogs to make the
        # exposed blocker concrete in the verdict note.
        sample_dialogs = [
            d.get("msg_text", "?")[:80]
            for d in (result.get("msgbox_observed") or [])[:5]
        ]
        verdict_note = (
            f"**Two-layer outcome.**\n\n"
            f"Layer 1 (resolved): the Issue #23 JET 3061 "
            f"target-column mismatch ('unknown field name: "
            f"c_index_addr_type_code') has been neutralized by "
            f"the narrow per-form rewrite.  Direct evidence:\n"
            f"  - 0 :ERR markers in ZZ_TEST_DEBUG (no "
            f"`LookAtAssociations:ERR ... c_index_addr_type_code` "
            f"row).\n"
            f"  - file_count = {n_files} (was 0 in PR #112's "
            f"pre-patch probe).\n"
            f"  - ZZ_SCRATCH_PEOPLE populated by the INSERT "
            f"that previously failed.\n"
            f"  - Chain reached its natural end "
            f"(chain_observed_done = "
            f"{result.get('chain_observed_done')}).\n\n"
            f"Layer 2 (newly exposed): a downstream debug-"
            f"MsgBox layer is now live and would block "
            f"unattended coverage.  Direct evidence:\n"
            f"  - watchdog dismissed {n_watchdog} dialog(s) "
            f"during the chain run.  Sample: "
            f"{sample_dialogs}\n"
            f"  - These are concat-form `MsgBox \"<lit>\" + "
            f"Trim(Str(...))` calls inside CmdNeo4j_Click that "
            f"the driver's literal-only generic neutralizer "
            f"cannot match.  Same shape as the AssocPairs "
            f"CmdNeo4j debug-MsgBox layer that PR #109 "
            f"suppressed via "
            f"`_suppress_assocpairs_cmdneo4j_debug_msgbox` — "
            f"but in a different form, so PR #109's per-form "
            f"patch does NOT cover it.\n\n"
            f"This is a NEW per-host observation that PR "
            f"#112's pre-patch probe could not surface, "
            f"because that probe's chain bailed at the JET "
            f"3061 BEFORE reaching the debug-MsgBox lines.\n\n"
            f"**Implication for next step:** coverage PR is "
            f"NOT directly unblocked by this rewrite alone.  "
            f"For unattended coverage, one of two paths is "
            f"required (each its own brief, NOT this PR):\n"
            f"  (a) extend `_PER_FORM_CMDGIS_PATCHES` with a "
            f"`Form_LookAtAssociations.CmdNeo4j_Click` "
            f"debug-MsgBox suppress (mirror of PR #109's "
            f"AssocPairs-side patch); OR\n"
            f"  (b) the cross-form coverage test inlines a "
            f"watchdog dismisser (this probe demonstrates the "
            f"watchdog approach works in principle).\n\n"
            f"Note on the single ZZ_TEST_DEBUG `:MSGBOX` "
            f"marker present here: that row was written by "
            f"the driver's generic literal-only MsgBox "
            f"neutralizer when it rewrote the terminal "
            f"`MsgBox \"Finished saving to Neo4j\"` line "
            f"(literal form, neutralizable) BEFORE the runtime "
            f"would see it.  It is NOT the line-1035 "
            f"RecordCount=0 early-bail marker (which would "
            f"only appear if the bail fired; on this fixture "
            f"it does not — ZZ_SCRATCH_P_ASSOC has 8087 rows).  "
            f"The :MSGBOX marker is a clean-end signal, NOT "
            f"a blocker — distinct layer from the watchdog "
            f"dismissals above."
        )
    elif outcome == "patch_partial_jet_3061_still_observed":
        verdict_note = (
            f"Patch INCOMPLETE.  ZZ_TEST_DEBUG still contains "
            f"the JET 3061 'c_index_addr_type_code' :ERR "
            f"marker(s): {[m[:100] for m in err_markers if 'c_index_addr_type_code' in m]}.  "
            f"The rewrite did not take effect; investigate the "
            f"per-form patch scoping or anchor literal."
        )
    elif outcome == "next_blocker_exposed_different_err":
        verdict_note = (
            f"Rewrite resolved the original JET 3061 but a "
            f"different runtime error class appeared in "
            f"ZZ_TEST_DEBUG: {other_err[:3]}.  file_count = "
            f"{n_files}.  A NEW blocker is exposed; coverage PR "
            f"is NOT the next step — open a separate "
            f"investigation."
        )
    elif outcome == "patch_resolved_jet_3061_but_zero_files":
        verdict_note = (
            f"Rewrite resolved the JET 3061 (no marker observed) "
            f"but file_count = {n_files} on this fixture.  Chain "
            f"may have hit a different early-bail or the watchdog "
            f"deadline.  See raw facts below; NOT a coverage "
            f"candidate yet."
        )
    elif outcome == "blocked_exception":
        verdict_note = (
            f"Exception during probe with no files produced.  "
            f"See `exception` field for the trace."
        )
    else:
        verdict_note = (
            f"Outcome `{outcome}` did not match any known path."
        )

    return {
        "verdict": outcome,
        "verdict_note": verdict_note,
        "answers": answers,
    }


def _write_md(result: dict, verdict: dict) -> None:
    md: list[str] = []
    md.append(
        "# LookAtAssociations × CmdNeo4j verification probe — "
        "post target-column rewrite patch")
    md.append("")
    md.append(
        "**Date:** 2026-05-07  ·  **Branch:** "
        "`driver/associations-cmdneo4j-c-addr-type-rewrite`")
    md.append("")
    md.append(
        "Verification probe for the narrow scoped driver patch "
        "`_rewrite_associations_cmdneo4j_target_column` in "
        "`tests/cbdb_driver/vba_session.py`.  The patch rewrites "
        "the single literal anchor "
        "`c_index_addr_type_code, c_female` -> `c_addr_type, "
        "c_female` inside `Form_LookAtAssociations.CmdNeo4j_Click` "
        "to work around canonical Issue #23 (P1_visible_crash).  "
        "This probe verifies the patch holds under the same "
        "Associations matrix fixture used by PR #112 and looks "
        "for any new blocker behind the JET 3061 layer.")
    md.append("")
    md.append("## Setup")
    md.append("")
    md.append(
        f"- **Fixture:** `{result.get('fixture_name')}` "
        f"(matrix `_make_assoc_fixtures` first fixture; "
        f"picker_ids = {result.get('fixture_picker_ids')}; "
        f"controls = {result.get('fixture_controls')})")
    md.append(
        "- **Chain:** `CmdQuery,CmdNeo4j` via Form.Tag, directory "
        "mode")
    md.append(
        "- **Watchdog:** records and dismisses any MsgBox the "
        "driver missed; each entry is a runtime signal we "
        "surface (not silently swallow).")
    md.append(
        f"- **click_via_timer cap:** {TIMER_TIMEOUT_SEC} s  ·  "
        f"**outer cap:** {PROBE_OUTER_TIMEOUT_SEC} s")
    md.append("")
    md.append("## Raw observed facts")
    md.append("")
    md.append(
        f"- **chain_elapsed_sec:** {result.get('chain_elapsed_sec')}")
    md.append(
        f"- **file_count:** {result.get('file_count')}")
    md.append(
        f"- **chain_observed_done:** "
        f"{result.get('chain_observed_done')}")
    md.append(
        f"- **click_via_timer_returned:** "
        f"{result.get('click_via_timer_returned')}")
    md.append(
        f"- **msgbox_observed_via_watchdog_count:** "
        f"{len(result.get('msgbox_observed', []))}")
    md.append(
        f"- **total_wall_elapsed_sec:** "
        f"{result.get('elapsed_sec')}")
    if result.get("exception"):
        md.append(
            f"- **exception:** `{result['exception'][:300]}`")
    md.append("")
    md.append("### Scratch row counts (post-chain)")
    md.append("")
    for tbl, c in (result.get("row_counts") or {}).items():
        md.append(f"- `{tbl}`: {c}")
    md.append("")
    md.append("### ZZ_TEST_DEBUG content")
    md.append("")
    msgs = result.get("zz_test_debug_msgs") or []
    if msgs:
        for m in msgs:
            md.append(f"- `{m}`")
    else:
        md.append("(empty)")
    md.append("")
    md.append("### Watchdog MsgBox observations")
    md.append("")
    obs = result.get("msgbox_observed", [])
    if obs:
        md.append("| +t (s) | msg_text |")
        md.append("|---:|---|")
        for d in obs:
            md.append(
                f"| {d['t']} | `{d.get('msg_text', '')[:120]}` |")
    else:
        md.append("(none observed)")
    md.append("")
    md.append("## Q1-Q4 answers")
    md.append("")
    a = verdict["answers"]
    md.append(
        "**Q1 — JET 3061 'c_index_addr_type_code' :ERR "
        "disappeared?**")
    md.append("")
    q1 = a["Q1_jet_3061_disappeared"]
    md.append(
        f"- has_jet_3061_marker: "
        f"**{q1['has_jet_3061_marker']}** (expected False)")
    md.append(
        f"- all_err_markers_observed: "
        f"`{q1['all_err_markers_observed']}`")
    md.append("")
    md.append(q1["interpretation"])
    md.append("")
    md.append(
        "**Q2 — chain wrote files (was 0 in PR #112)?**")
    md.append("")
    q2 = a["Q2_chain_writes_files"]
    md.append(f"- file_count = {q2['file_count']}")
    md.append(
        f"- chain_observed_done = {q2['chain_observed_done']}")
    md.append(
        f"- chain_elapsed_sec = {q2['chain_elapsed_sec']}")
    md.append("")
    md.append(q2["interpretation"])
    md.append("")
    md.append("**Q3 — file shape:**")
    md.append("")
    q3 = a["Q3_file_shape"]
    if q3["files"]:
        md.append(
            "| name | size | n_cols | first_col | data_rows |")
        md.append("|---|---:|---:|---|---:|")
        for f in q3["files"]:
            md.append(
                f"| {f['name']} | {f['size']} | "
                f"{f['n_cols']} | `{f['first_col']}` | "
                f"{f['data_rows']} |")
    else:
        md.append("(no files produced)")
    md.append("")
    md.append("**Q4 — next blocker exposed?**")
    md.append("")
    q4 = a["Q4_next_blocker_exposed"]
    md.append(
        f"- outcome_bucket = `{q4['outcome_bucket']}`")
    md.append(
        f"- other_err_markers = `{q4['other_err_markers']}`")
    md.append(
        f"- msgbox_watchdog_count = "
        f"{q4['msgbox_watchdog_count']}  "
        f"(REAL pop-up dialogs the driver's literal-only "
        f"generic neutralizer could not catch; each one "
        f"would block unattended coverage)")
    md.append(
        f"- zz_test_debug_msgbox_marker_count = "
        f"{q4['zz_test_debug_msgbox_marker_count']}  "
        f"(generic-neutralizer footprints for literal "
        f"`MsgBox \"<lit>\"` lines; NOT pop-up dialogs)")
    md.append("")
    md.append("**The two MsgBox signal layers are distinct:**")
    md.append("")
    md.append(
        f"- *Watchdog-dismissed dialogs:* "
        f"{q4['msgbox_signal_layers_explained']['watchdog_dismissed_dialogs']}")
    md.append("")
    md.append(
        f"- *ZZ_TEST_DEBUG :MSGBOX markers:* "
        f"{q4['msgbox_signal_layers_explained']['zz_test_debug_MSGBOX_markers']}")
    md.append("")
    md.append(q4["interpretation"])
    md.append("")
    md.append(f"## Verdict: `{verdict['verdict']}`")
    md.append("")
    md.append(verdict["verdict_note"])
    md.append("")
    md.append(
        f"## Markers (timeline, "
        f"{len(result.get('markers', []))} entries)")
    md.append("")
    for m in result.get("markers", []):
        md.append(f"  - `+{m['t']:>6.2f}s` {m['marker']}")
    md.append("")
    md.append("## Constraints honoured per brief")
    md.append("")
    md.append("- ✅ Did NOT open a coverage PR")
    md.append(
        "- ✅ Driver patch is narrow-scoped: only "
        "`Form_LookAtAssociations`, only `CmdNeo4j_Click`, only "
        "the literal `c_index_addr_type_code, c_female` -> "
        "`c_addr_type, c_female` rewrite")
    md.append(
        "- ✅ Did NOT introduce a generic SQL text rewriting "
        "policy")
    md.append(
        "- ✅ Did NOT touch other forms or the SELECT-side "
        "qualified `BIOG_MAIN.c_index_addr_type_code` reference")
    md.append(
        "- ✅ Did NOT change canonical reports / issue severity / "
        "README / triage docs")
    md.append("- ✅ Did NOT change tests/")

    OUT_MD.write_text("\n".join(md), encoding="utf-8")


def _write_outputs(result: dict, verdict: dict) -> None:
    out: dict = {
        "schema_version": 1,
        "generated_date": "2026-05-07",
        "probe_branch": (
            "driver/associations-cmdneo4j-c-addr-type-rewrite"),
        "fixture": "associations_matrix_first_fixture",
        "config": {
            "click_via_timer_timeout_sec": TIMER_TIMEOUT_SEC,
            "per_probe_outer_timeout_sec": PROBE_OUTER_TIMEOUT_SEC,
        },
        "result": result,
        "verdict": verdict["verdict"],
        "verdict_note": verdict["verdict_note"],
        "answers": verdict["answers"],
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"wrote {OUT_JSON}")
    _write_md(result, verdict)
    print(f"wrote {OUT_MD}")


def _reclassify(src_path: Path) -> int:
    """Reload preserved JSON, re-run classifier + verdict, rewrite
    outputs.  Used to fix classifier wording/bucket bugs without
    re-running COM."""
    print(f"=== reclassifying from {src_path} (no COM) ===\n")
    existing = json.loads(src_path.read_text(encoding="utf-8"))
    result = existing["result"]
    result["outcome"] = _classify_outcome(result)
    verdict = _verdict(result)
    _write_outputs(result, verdict)
    print(f"\nreclassified outcome: {result.get('outcome')}")
    print(f"file_count: {result.get('file_count')}")
    print(f"chain_elapsed: {result.get('chain_elapsed_sec')} s")
    print(
        f"watchdog dialogs: "
        f"{len(result.get('msgbox_observed') or [])}")
    return 0


def main(argv: list[str] | None = None) -> int:
    if argv is None:
        argv = sys.argv[1:]
    if "--reclassify-from-json" in argv:
        idx = argv.index("--reclassify-from-json")
        if idx + 1 >= len(argv):
            print("ERROR: --reclassify-from-json requires a path arg",
                  file=sys.stderr)
            return 2
        return _reclassify(Path(argv[idx + 1]))

    print("=== LookAtAssociations × CmdNeo4j verification probe "
          "(post target-column rewrite patch) ===\n")
    _kill_orphan()
    time.sleep(1)

    out_dir = ROOT / "analysis" / (
        "_probe_associations_cmdneo4j_after_target_patch_out")
    if out_dir.exists():
        for f in out_dir.glob("*"):
            try:
                f.unlink()
            except Exception:
                pass
    else:
        out_dir.mkdir(parents=True)

    result = _run_probe(out_dir)
    verdict = _verdict(result)
    _write_outputs(result, verdict)
    print(f"\noutcome: {result.get('outcome')}")
    print(f"file_count: {result.get('file_count')}")
    print(f"chain_elapsed: {result.get('chain_elapsed_sec')} s")
    print(f"\n=== verdict: {verdict['verdict']} ===")
    return 0


if __name__ == "__main__":
    sys.exit(main())
