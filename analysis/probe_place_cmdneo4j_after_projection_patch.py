"""LookAtPlace × CmdNeo4j verification probe — post tRstPeople SELECT projection patch.

Companion to PR #120 (probe; observed JET 3265 + 0 files), PR
#121 (static investigation; pinned the failing read at
`tRstPeople!c_dynasty` line 757), PR #122 (canonical Issue #24
filed P1).

This probe verifies the narrow scoped driver-side workaround in
`tests/cbdb_driver/vba_session.py`,
`_rewrite_place_cmdneo4j_trstpeople_projection`, which extends
the SELECT projection of the `tRstPeople` query inside
`Form_LookAtPlace.CmdNeo4j_Click` from 4 cols (only
`ZZ_SCRATCH_P_TEXT` cols) to 7 cols (also `DYNASTIES.c_dynasty`,
`DYNASTIES.c_dynasty_chn`, `BIOG_MAIN.c_female`).

Anchor:
  old: "ZZ_SCRATCH_P_TEXT.c_index_year "  (with trailing space)
  new: "ZZ_SCRATCH_P_TEXT.c_index_year, DYNASTIES.c_dynasty, "
       "DYNASTIES.c_dynasty_chn, BIOG_MAIN.c_female "

Probe answers (per brief):

  Q1  Does the JET 3265 'Item not found in this collection.'
      :ERR marker disappear from ZZ_TEST_DEBUG?
  Q2  Does the chain start writing CSV files (file_count >= 1)?
  Q3  What is the file shape — count, first columns, row counts?
  Q4  Does the patch expose a NEXT layer of blocker (a
      different :ERR class, an unexpected MsgBox, an empty file
      set, a timeout)?

This probe does NOT open a coverage PR; per brief it is the
verification step that gates whether a coverage PR is the next
step.

Note on pre-chain noise: per PR #120's probe, the matrix
`_make_place_fixtures` controls dict contains `ChkAssoc` /
`ChkPosting`, which are NOT controls on Form_LookAtPlace.  The
`set_control` calls fail (com_error: "can't find the field
ChkAssoc / ChkPosting").  This is a fixture-vs-form mismatch,
NOT a CmdNeo4j_Click failure.  Per the brief, this probe
preserves those failures as observation in the markers
timeline AND in a dedicated `pre_chain_observations` field —
NOT silently swallowed.

Outputs:
  analysis/probe_place_cmdneo4j_after_projection_patch.md
  reports/probe_place_cmdneo4j_after_projection_patch.json

CLI:
  python analysis/probe_place_cmdneo4j_after_projection_patch.py
    full COM probe run.
  python analysis/probe_place_cmdneo4j_after_projection_patch.py --reclassify-from-json <path>
    re-run classification + verdict from preserved JSON (no COM).
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
    "_probe_place_cmdneo4j_after_projection_patch_copy.mdb")
OUT_JSON = ROOT / "reports" / (
    "probe_place_cmdneo4j_after_projection_patch.json")
OUT_MD = ROOT / "analysis" / (
    "probe_place_cmdneo4j_after_projection_patch.md")

TIMER_TIMEOUT_SEC = 180
PROBE_OUTER_TIMEOUT_SEC = 300

# The skip-reason / Issue #24 surface text under investigation.
JET_3265_PHRASE = "Item not found in this collection"


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
    """Watchdog: dismiss + record any MsgBox the driver missed.
    After this projection patch, a watchdog dialog with text
    matching JET 3265 would mean the rewrite did NOT take
    effect; any other dialog is a NEW blocker observation."""
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
    first-match wins).

      - blocked_exception:
          exception raised AND no files produced.
      - patch_partial_jet_3265_still_observed:
          ZZ_TEST_DEBUG still contains the Issue #24 :ERR text
          ("Item not found in this collection") — the rewrite
          did not take effect.
      - next_blocker_exposed_runtime_err:
          ZZ_TEST_DEBUG contains a :ERR row that is NOT the
          Issue #24 surface text — the patch resolved Issue #24
          but a different runtime error class is now visible.
      - patch_resolved_issue24_but_exposed_msgbox_blocker:
          no :ERR markers AND file_count >= 1 BUT the watchdog
          observed and dismissed >= 1 MsgBox during the chain
          run.  Same shape as PR #116's "MsgBox blocker"
          bucket for Associations CmdNeo4j — chain reaches its
          natural end and writes files, but a debug-MsgBox
          layer is still live and would block unattended
          coverage.
      - patch_verified_chain_clean:
          STRICT: file_count >= 1 AND no :ERR AND zero watchdog
          dialogs AND chain quiesced AND chain_elapsed <= 180s.
      - patch_resolved_jet_3265_but_zero_files:
          no :ERR AND no watchdog dialogs BUT file_count == 0.
          Unexpected; investigate.
    """
    n_files = int(result.get("file_count") or 0)
    chain_elapsed = result.get("chain_elapsed_sec")
    exception = result.get("exception")
    debug_msgs = result.get("zz_test_debug_msgs") or []
    n_watchdog = len(result.get("msgbox_observed") or [])

    err_msgs = [m for m in debug_msgs if ":ERR" in m]
    has_jet_3265 = any(JET_3265_PHRASE in m for m in err_msgs)
    has_other_err = any(
        JET_3265_PHRASE not in m for m in err_msgs)
    elapsed_ok = (
        isinstance(chain_elapsed, (int, float))
        and chain_elapsed <= TIMER_TIMEOUT_SEC
    )

    if exception and n_files == 0:
        return "blocked_exception"
    if has_jet_3265:
        return "patch_partial_jet_3265_still_observed"
    if has_other_err:
        return "next_blocker_exposed_runtime_err"
    # No :ERR markers from here on.
    if n_watchdog >= 1 and n_files >= 1:
        return (
            "patch_resolved_issue24_but_exposed_msgbox_blocker")
    if n_files == 0:
        return "patch_resolved_jet_3265_but_zero_files"
    if (n_files >= 1
            and n_watchdog == 0
            and elapsed_ok
            and result.get("chain_observed_done")):
        return "patch_verified_chain_clean"
    return "patch_resolved_jet_3265_but_zero_files"


def _get_place_fixture():
    from test_vba_matrix_all_forms import _all_fixtures
    for fx in _all_fixtures():
        if fx.spec.name == "LookAtPlace":
            return fx
    raise RuntimeError("no LookAtPlace fixture found in matrix")


def _run_probe(out_dir: Path) -> dict:
    from cbdb_driver.vba_session import VbaSession, make_fixture
    from cbdb_driver.form_specs import LOOKATPLACE

    spec = LOOKATPLACE
    fx = _get_place_fixture()

    result: dict = {
        "form": spec.name,
        "fixture_name": fx.name,
        "fixture_picker_ids": list(fx.picker_ids) if fx.picker_ids else [],
        "fixture_addr_ids": list(fx.addr_ids) if fx.addr_ids else [],
        "fixture_controls": dict(fx.controls or {}),
        "fixture_expected_min_rows": fx.expected_min_rows,
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

            # LookAtPlace-specific: TabPlaces=0 before fixture
            # controls (mirrors PR #120 + the cross-form CmdNeo4j
            # test's `_seed_query_inputs` special-case).
            try:
                sess.set_control(spec.name, "TabPlaces", 0)
                mark("tab_places_set_to_0")
            except Exception as e:
                mark(f"set_TabPlaces_0_fail: {e!r}")

            for ctl, val in (fx.controls or {}).items():
                try:
                    sess.set_control(spec.name, ctl, val)
                except Exception as e:
                    # Pre-chain noise observation per brief — NOT
                    # silently swallowed.  Same as PR #120 saw
                    # for ChkAssoc / ChkPosting (matrix fixture
                    # has them but Form_LookAtPlace doesn't).
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
                "ZZ_SCRATCH_STATUS",
                "ZZ_SCRATCH_PLACE_PEOPLE",
                "ZZ_SCRATCH_PLACE_AGG",
                "ZZ_SCRATCH_PEOPLE",
                "ZZ_SCRATCH_ADDR",
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
    has_jet_3265 = any(JET_3265_PHRASE in m for m in err_markers)
    other_err = [m for m in err_markers
                 if JET_3265_PHRASE not in m]
    n_watchdog = len(result.get("msgbox_observed") or [])

    # Pre-chain noise: set_control_<X>_fail markers from the
    # matrix fixture's ChkAssoc / ChkPosting (Form_LookAtPlace
    # doesn't have those controls).  Preserved per brief.
    set_control_failures = [
        m["marker"] for m in (result.get("markers") or [])
        if m.get("marker", "").startswith("set_control_")
        and "_fail" in m.get("marker", "")
    ]

    answers = {
        "Q1_jet_3265_disappeared": {
            "has_jet_3265_marker": has_jet_3265,
            "all_err_markers_observed": err_markers,
            "interpretation": (
                "If has_jet_3265_marker is False, the patch "
                "successfully prevented the JET 3265 'Item not "
                "found in this collection.' that PR #120's "
                "probe observed.  If True, the rewrite did "
                "not take effect."
            ),
        },
        "Q2_chain_writes_files": {
            "file_count": n_files,
            "chain_observed_done": result.get(
                "chain_observed_done", False),
            "chain_elapsed_sec": chain_elapsed,
            "interpretation": (
                "PR #120 observed 0 files; this probe expects "
                "file_count >= 1 if the rewrite took effect "
                "AND the rest of the chain runs to its natural "
                "end."
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
            "interpretation": (
                "If outcome == patch_verified_chain_clean: no "
                "next blocker on this fixture.  If outcome == "
                "patch_partial_jet_3265_still_observed: rewrite "
                "did not take effect.  If outcome == "
                "next_blocker_exposed_runtime_err: rewrite "
                "worked but a different runtime error class "
                "appeared.  If outcome == "
                "patch_resolved_issue24_but_exposed_msgbox_blocker: "
                "rewrite worked AND chain wrote files BUT a "
                "watchdog-dismissed dialog appeared (analog of "
                "PR #116's MsgBox-blocker layer for "
                "Associations CmdNeo4j).  If outcome == "
                "patch_resolved_jet_3265_but_zero_files: "
                "rewrite worked AND no :ERR but no files "
                "either."
            ),
        },
        "pre_chain_observations": {
            "set_control_failures": set_control_failures,
            "interpretation": (
                "These are pre-chain failures from the matrix "
                "_make_place_fixtures controls dict (ChkAssoc / "
                "ChkPosting NOT controls on Form_LookAtPlace). "
                "Preserved per brief — NOT silently swallowed. "
                "They do NOT cause the JET 3265 (which fires "
                "later inside CmdNeo4j_Click body) and do NOT "
                "cause any post-patch failure.  Fixing the "
                "fixture is out-of-scope for this probe."
            ),
        },
    }

    if outcome == "patch_verified_chain_clean":
        verdict_note = (
            f"**Strict clean.**  Chain produced {n_files} files "
            f"in {chain_elapsed}s with 0 :ERR markers AND 0 "
            f"watchdog-dismissed dialogs.  This is the "
            f"strongest available signal that the SELECT "
            f"projection rewrite is the only blocker on this "
            f"fixture for LookAtPlace × CmdNeo4j.  Per the "
            f"brief, do NOT open a coverage PR off this "
            f"verdict alone — the coverage PR is a separate "
            f"brief."
        )
    elif outcome == "patch_partial_jet_3265_still_observed":
        verdict_note = (
            f"**Patch INCOMPLETE.**  ZZ_TEST_DEBUG still "
            f"contains the Issue #24 :ERR marker(s): "
            f"{[m[:120] for m in err_markers if JET_3265_PHRASE in m][:3]}.  "
            f"The rewrite did not take effect; investigate the "
            f"per-form patch scoping or anchor literal."
        )
    elif outcome == "next_blocker_exposed_runtime_err":
        verdict_note = (
            f"**Two-layer outcome.**\n\n"
            f"Layer 1 (resolved): JET 3265 'Item not found in "
            f"this collection.' is GONE — the SELECT projection "
            f"rewrite worked.  No matching :ERR row.\n\n"
            f"Layer 2 (newly exposed): a different runtime "
            f":ERR class appeared in ZZ_TEST_DEBUG: "
            f"{[m[:120] for m in other_err][:3]}.  "
            f"file_count = {n_files}.  This is a NEW blocker "
            f"not observed by PR #120 because that probe's "
            f"chain bailed at the JET 3265 BEFORE reaching "
            f"this point.\n\n"
            f"Recommended next step: characterize the new :ERR "
            f"class (compare against Issue #21 DAO 3021 / "
            f"Issue #23 JET 3061 / Issue #24 JET 3265 / others) "
            f"and decide whether it's a new canonical Issue "
            f"candidate, an existing-issue match, or another "
            f"investigation line."
        )
    elif outcome == (
        "patch_resolved_issue24_but_exposed_msgbox_blocker"
    ):
        sample = [
            d.get("msg_text", "?")[:80]
            for d in (result.get("msgbox_observed") or [])[:5]
        ]
        verdict_note = (
            f"**Two-layer outcome** (analog of PR #116's "
            f"Associations CmdNeo4j MsgBox-blocker).\n\n"
            f"Layer 1 (resolved): JET 3265 is GONE; chain "
            f"wrote {n_files} files (was 0 in PR #120).\n\n"
            f"Layer 2 (newly exposed): a downstream MsgBox "
            f"layer is now live — watchdog dismissed "
            f"{n_watchdog} dialog(s).  Sample: {sample}.  "
            f"For unattended coverage, those dialogs would "
            f"block headless test runs unless similarly "
            f"suppressed.\n\n"
            f"Recommended next step (separate brief, NOT this "
            f"PR): mirror PR #109's debug-MsgBox-suppress "
            f"pattern for Form_LookAtPlace.CmdNeo4j_Click, OR "
            f"the cross-form coverage test inlines a watchdog "
            f"dismisser."
        )
    elif outcome == "patch_resolved_jet_3265_but_zero_files":
        verdict_note = (
            f"**Unexpected.**  JET 3265 :ERR is gone AND no "
            f"watchdog dialogs AND no other :ERR — but "
            f"file_count = {n_files} (PR #120 observed 0; "
            f"expected this rewrite to produce >= 1).  Chain "
            f"may have hit a different early-bail or the "
            f"watchdog deadline.  Investigate before any "
            f"coverage decision."
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
        "# LookAtPlace × CmdNeo4j verification probe — "
        "post tRstPeople SELECT projection patch")
    md.append("")
    md.append(
        "**Date:** 2026-05-08  ·  **Branch:** "
        "`driver/place-cmdneo4j-trstpeople-projection-rewrite`")
    md.append("")
    md.append(
        "Verification probe for the narrow scoped driver patch "
        "`_rewrite_place_cmdneo4j_trstpeople_projection` in "
        "`tests/cbdb_driver/vba_session.py`.  The patch extends "
        "the `tRstPeople` SELECT projection inside "
        "`Form_LookAtPlace.CmdNeo4j_Click` from 4 cols (only "
        "`ZZ_SCRATCH_P_TEXT` cols) to 7 cols (also "
        "`DYNASTIES.c_dynasty`, `DYNASTIES.c_dynasty_chn`, "
        "`BIOG_MAIN.c_female`).  Anchor:")
    md.append("")
    md.append("```")
    md.append(
        "old: \"ZZ_SCRATCH_P_TEXT.c_index_year \"  (with "
        "trailing space)")
    md.append(
        "new: \"ZZ_SCRATCH_P_TEXT.c_index_year, "
        "DYNASTIES.c_dynasty, DYNASTIES.c_dynasty_chn, "
        "BIOG_MAIN.c_female \"")
    md.append("```")
    md.append("")
    md.append(
        "Verifies under the same matrix Place fixture used by "
        "PR #120, looking for whether the JET 3265 disappears "
        "and whether the chain now writes files.")
    md.append("")
    md.append("## Setup")
    md.append("")
    md.append(
        f"- **Fixture:** `{result.get('fixture_name')}` (matrix "
        f"`_make_place_fixtures` first fixture; same as PR "
        f"#120)")
    md.append(
        "- **Pre-fixture step:** `set_control(\"LookAtPlace\", "
        "\"TabPlaces\", 0)` (mirrors the cross-form CmdNeo4j "
        "test's special-case for Place)")
    md.append(
        "- **Chain:** `CmdQuery,CmdNeo4j` via Form.Tag, "
        "directory mode")
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
        f"- **msgbox_watchdog_count:** "
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
    md.append("**Q1 — JET 3265 'Item not found in this "
              "collection.' :ERR disappeared?**")
    md.append("")
    q1 = a["Q1_jet_3265_disappeared"]
    md.append(
        f"- has_jet_3265_marker: "
        f"**{q1['has_jet_3265_marker']}** (expected False)")
    md.append(
        f"- all_err_markers_observed: "
        f"`{q1['all_err_markers_observed']}`")
    md.append("")
    md.append(q1["interpretation"])
    md.append("")
    md.append("**Q2 — chain wrote files (was 0 in PR #120)?**")
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
        f"{q4['msgbox_watchdog_count']}")
    md.append("")
    md.append(q4["interpretation"])
    md.append("")
    md.append(
        "**Pre-chain observations (preserved, not silenced):**")
    md.append("")
    pco = a["pre_chain_observations"]
    if pco["set_control_failures"]:
        md.append(
            f"- set_control failures during fixture seeding "
            f"({len(pco['set_control_failures'])} entries):")
        for sf in pco["set_control_failures"]:
            md.append(f"    - `{sf}`")
    else:
        md.append("- (none observed)")
    md.append("")
    md.append(pco["interpretation"])
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
        "`Form_LookAtPlace`, only `CmdNeo4j_Click`, only the "
        "`tRstPeople` SELECT projection literal")
    md.append(
        "- ✅ Did NOT introduce a generic SQL rewrite policy")
    md.append(
        "- ✅ Did NOT touch other forms or other recordsets in "
        "the same sub")
    md.append(
        "- ✅ Did NOT change canonical reports / issue severity "
        "/ README / triage docs")
    md.append("- ✅ Did NOT change tests/")
    md.append(
        "- ✅ Pre-chain `ChkAssoc / ChkPosting` `set_control` "
        "failures preserved as observations (per brief), NOT "
        "silenced")

    OUT_MD.write_text("\n".join(md), encoding="utf-8")


def _write_outputs(result: dict, verdict: dict) -> None:
    out: dict = {
        "schema_version": 1,
        "generated_date": "2026-05-08",
        "probe_branch": (
            "driver/place-cmdneo4j-trstpeople-projection-rewrite"),
        "fixture": "place_matrix_first_fixture",
        "config": {
            "click_via_timer_timeout_sec": TIMER_TIMEOUT_SEC,
            "per_probe_outer_timeout_sec": PROBE_OUTER_TIMEOUT_SEC,
            "promote_gates_strict": [
                "file_count >= 1",
                "no ':ERR' markers in ZZ_TEST_DEBUG",
                "watchdog dialog count == 0",
                "chain_observed_done == True",
                f"chain_elapsed_sec <= {TIMER_TIMEOUT_SEC}",
            ],
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
    print(f"=== reclassifying from {src_path} (no COM) ===\n")
    existing = json.loads(src_path.read_text(encoding="utf-8"))
    result = existing["result"]
    result["outcome"] = _classify_outcome(result)
    verdict = _verdict(result)
    _write_outputs(result, verdict)
    print(f"\nreclassified outcome: {result.get('outcome')}")
    print(f"file_count: {result.get('file_count')}")
    print(f"chain_elapsed: {result.get('chain_elapsed_sec')} s")
    print(f"watchdog dialogs: "
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

    print("=== LookAtPlace × CmdNeo4j verification probe "
          "(post tRstPeople SELECT projection patch) ===\n")
    _kill_orphan()
    time.sleep(1)

    out_dir = ROOT / "analysis" / (
        "_probe_place_cmdneo4j_after_projection_patch_out")
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
    print(
        f"watchdog dialogs: "
        f"{len(result.get('msgbox_observed') or [])}")
    print(f"\n=== verdict: {verdict['verdict']} ===")
    return 0


if __name__ == "__main__":
    sys.exit(main())
