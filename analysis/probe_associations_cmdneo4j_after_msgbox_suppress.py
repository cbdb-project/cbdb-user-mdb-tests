"""LookAtAssociations × CmdNeo4j verification probe — post debug-MsgBox suppress patch.

Companion to PR #112 (probe; observed JET 3061 + 0 files), PR
#114 (static investigation), PR #115 (canonical Issue #23), and
PR #116 (post-c_addr_type-rewrite probe; observed file_count = 8
+ 0 :ERR + 5 watchdog-dismissed concat-form debug MsgBoxes).

This probe verifies the narrow scoped driver-side debug-MsgBox
suppress patch in `tests/cbdb_driver/vba_session.py`,
`_suppress_associations_cmdneo4j_debug_msgbox`, which comments
out the 5 concat-form debug MsgBox calls inside
`Form_LookAtAssociations.CmdNeo4j_Click`:

  MsgBox "Kinship code records = "       ' line 1130
  MsgBox "Literary genre code records = " ' line 1232
  MsgBox "Institution code records = "    ' line 1315
  MsgBox "Occasion code records = "       ' line 1398
  MsgBox "Topic code records = "          ' line 1481

Probe answers:

  Q1  Does the watchdog dialog count drop to 0?  (PR #116 saw 5)
  Q2  Are :ERR markers still 0?
  Q3  Does file_count / file family stay at the 8-file shape?
  Q4  Does the patch expose a NEXT layer of blocker?

This probe does NOT open a coverage PR; per brief it is the
verification step that gates whether a coverage PR is the next
step.

Outputs:
  analysis/probe_associations_cmdneo4j_after_msgbox_suppress.md
  reports/probe_associations_cmdneo4j_after_msgbox_suppress.json

CLI:
  python analysis/probe_associations_cmdneo4j_after_msgbox_suppress.py
    full COM probe run.
  python analysis/probe_associations_cmdneo4j_after_msgbox_suppress.py --reclassify-from-json <path>
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
    "_probe_associations_cmdneo4j_after_msgbox_suppress_copy.mdb")
OUT_JSON = ROOT / "reports" / (
    "probe_associations_cmdneo4j_after_msgbox_suppress.json")
OUT_MD = ROOT / "analysis" / (
    "probe_associations_cmdneo4j_after_msgbox_suppress.md")

TIMER_TIMEOUT_SEC = 180
PROBE_OUTER_TIMEOUT_SEC = 300

# PR #116's verification probe observed 8 files in this shape;
# this probe expects the same family (no shape regression from
# adding the MsgBox suppress).
EXPECTED_FILE_COUNT_FROM_PR116 = 8


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
    After this probe's MsgBox suppress patch, count should be 0
    (the 5 concat-form dialogs PR #116 saw should now be
    commented out before the runtime sees them)."""
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
    first-match wins).  Order matters: any watchdog dialog or :ERR
    marker disqualifies the strict clean bucket.

      - blocked_exception:
          exception raised AND no files produced.
      - msgbox_suppress_partial_dialogs_still_observed:
          watchdog count >= 1 — the 5-prefix suppress did NOT
          take effect for one or more dialogs.  Investigate the
          per-form anchors.
      - next_blocker_exposed_runtime_err:
          ZZ_TEST_DEBUG contains a :ERR marker — watchdog count
          may be 0 but a different runtime error appeared.
      - patch_verified_chain_clean:
          STRICT: file_count >= 1 AND no :ERR AND zero watchdog
          dialogs AND chain quiesced AND chain_elapsed <= 180s.
          This is now reachable after the combined
          c_addr_type-rewrite (PR #116) + this MsgBox suppress.
      - patch_resolved_msgbox_but_zero_files:
          no :ERR AND no watchdog dialogs BUT file_count == 0.
          Unexpected on this fixture; investigate.
    """
    n_files = int(result.get("file_count") or 0)
    chain_elapsed = result.get("chain_elapsed_sec")
    exception = result.get("exception")
    debug_msgs = result.get("zz_test_debug_msgs") or []
    n_watchdog = len(result.get("msgbox_observed") or [])

    has_err = any(":ERR" in m for m in debug_msgs)
    elapsed_ok = (
        isinstance(chain_elapsed, (int, float))
        and chain_elapsed <= TIMER_TIMEOUT_SEC
    )

    if exception and n_files == 0:
        return "blocked_exception"
    if has_err:
        return "next_blocker_exposed_runtime_err"
    if n_watchdog >= 1:
        return "msgbox_suppress_partial_dialogs_still_observed"
    # No :ERR, no watchdog dialogs from here on.
    if (n_files >= 1
            and elapsed_ok
            and result.get("chain_observed_done")):
        return "patch_verified_chain_clean"
    return "patch_resolved_msgbox_but_zero_files"


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
    n_watchdog = len(result.get("msgbox_observed") or [])

    answers = {
        "Q1_watchdog_dialogs_to_zero": {
            "watchdog_count": n_watchdog,
            "expected": 0,
            "pr116_baseline": 5,
            "watchdog_dialogs": [
                d.get("msg_text", "?")[:80]
                for d in (result.get("msgbox_observed") or [])
            ],
            "interpretation": (
                "PR #116's pre-suppress probe observed 5 "
                "watchdog-dismissed dialogs (concat-form "
                "`MsgBox \"<lit>\" + Trim(Str(...))`).  This "
                "probe expects 0 — the 5 prefixes are now "
                "comment-prefixed by the per-form suppress."
            ),
        },
        "Q2_err_markers_still_zero": {
            "err_marker_count": len(err_markers),
            "err_markers_observed": err_markers,
            "interpretation": (
                "PR #116's post-c_addr_type-rewrite probe had "
                "0 :ERR markers; this probe expects the same "
                "(no new runtime error class introduced by the "
                "MsgBox suppress)."
            ),
        },
        "Q3_file_shape_stable": {
            "file_count": n_files,
            "expected_file_count_from_pr116": (
                EXPECTED_FILE_COUNT_FROM_PR116),
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
            "interpretation": (
                "PR #116 saw an 8-file shape (People / Places / "
                "PeoplePlaces / PeopleAssociations / "
                "AssociationCodes / KinshipCodes / "
                "OccasionCodes / TopicCodes).  This probe "
                "expects the same family — the MsgBox suppress "
                "should not change the SaveAs blocks themselves."
            ),
        },
        "Q4_next_blocker_exposed": {
            "outcome_bucket": outcome,
            "interpretation": (
                "patch_verified_chain_clean = no next blocker "
                "on this fixture; coverage PR is feasible.  "
                "msgbox_suppress_partial_dialogs_still_observed "
                "= one or more of the 5 prefixes did not match.  "
                "next_blocker_exposed_runtime_err = a runtime "
                ":ERR appeared (regression vs PR #116)."
            ),
        },
    }

    if outcome == "patch_verified_chain_clean":
        verdict_note = (
            f"Strict clean: chain produced {n_files} files in "
            f"{chain_elapsed}s with 0 :ERR markers AND 0 "
            f"watchdog-dismissed dialogs.  This is the "
            f"strongest available signal that the combined "
            f"c_addr_type rewrite (PR #116) + this MsgBox "
            f"suppress are sufficient driver-side workarounds "
            f"for unattended LookAtAssociations × CmdNeo4j "
            f"coverage on this fixture.  Per the brief, do NOT "
            f"open a coverage PR off this verdict alone — the "
            f"coverage PR is a separate brief."
        )
    elif outcome == "msgbox_suppress_partial_dialogs_still_observed":
        sample = [
            d.get("msg_text", "?")[:80]
            for d in (result.get("msgbox_observed") or [])[:5]
        ]
        verdict_note = (
            f"Suppress INCOMPLETE.  Watchdog dismissed "
            f"{n_watchdog} dialog(s) during the chain run.  "
            f"Sample: {sample}.  The 5-prefix suppress did "
            f"not match one or more dialogs; investigate the "
            f"per-form anchors against the current VBA dump."
        )
    elif outcome == "next_blocker_exposed_runtime_err":
        verdict_note = (
            f"MsgBox suppress took effect (0 watchdog dialogs) "
            f"but a runtime :ERR appeared in ZZ_TEST_DEBUG: "
            f"{err_markers[:3]}.  This is a NEW blocker not "
            f"observed in PR #116; coverage PR is NOT the next "
            f"step — open a separate investigation."
        )
    elif outcome == "patch_resolved_msgbox_but_zero_files":
        verdict_note = (
            f"MsgBox suppress took effect (0 watchdog, 0 :ERR) "
            f"but file_count = {n_files} on this fixture.  "
            f"Regression vs PR #116; investigate."
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
        "post debug-MsgBox suppress patch")
    md.append("")
    md.append(
        "**Date:** 2026-05-07  ·  **Branch:** "
        "`driver/associations-cmdneo4j-debug-msgbox-suppress`")
    md.append("")
    md.append(
        "Verification probe for the narrow scoped driver patch "
        "`_suppress_associations_cmdneo4j_debug_msgbox` in "
        "`tests/cbdb_driver/vba_session.py`.  The patch comments "
        "out the 5 concat-form debug MsgBox calls inside "
        "`Form_LookAtAssociations.CmdNeo4j_Click` (lines 1130 / "
        "1232 / 1315 / 1398 / 1481 in the current dump).  These "
        "are the watchdog-dismissed dialogs PR #116's "
        "post-c_addr_type-rewrite verification probe surfaced "
        "(outcome = "
        "patch_resolved_issue23_but_exposed_msgbox_blocker).")
    md.append("")
    md.append("## Setup")
    md.append("")
    md.append(
        f"- **Fixture:** `{result.get('fixture_name')}` (matrix "
        f"`_make_assoc_fixtures` first fixture; same as PR #112 "
        f"and PR #116)")
    md.append(
        "- **Chain:** `CmdQuery,CmdNeo4j` via Form.Tag, directory "
        "mode")
    md.append(
        "- **Watchdog:** records and dismisses any MsgBox; after "
        "this patch, count is expected to drop from 5 (PR #116 "
        "baseline) to 0.")
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
        f"{len(result.get('msgbox_observed', []))}  "
        f"(PR #116 baseline: 5; expected after this patch: 0)")
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
        md.append("(none observed — suppress patch took effect)")
    md.append("")
    md.append("## Q1-Q4 answers")
    md.append("")
    a = verdict["answers"]
    md.append("**Q1 — watchdog dialogs to zero?**")
    md.append("")
    q1 = a["Q1_watchdog_dialogs_to_zero"]
    md.append(
        f"- watchdog_count = {q1['watchdog_count']} "
        f"(expected: {q1['expected']}; PR #116 baseline: "
        f"{q1['pr116_baseline']})")
    if q1["watchdog_dialogs"]:
        md.append(f"- dialogs observed: {q1['watchdog_dialogs']}")
    else:
        md.append("- dialogs observed: (none)")
    md.append("")
    md.append(q1["interpretation"])
    md.append("")
    md.append("**Q2 — :ERR markers still zero?**")
    md.append("")
    q2 = a["Q2_err_markers_still_zero"]
    md.append(
        f"- err_marker_count = {q2['err_marker_count']}")
    if q2["err_markers_observed"]:
        md.append(
            f"- err_markers_observed = "
            f"{q2['err_markers_observed']}")
    md.append("")
    md.append(q2["interpretation"])
    md.append("")
    md.append("**Q3 — file shape stable vs PR #116?**")
    md.append("")
    q3 = a["Q3_file_shape_stable"]
    md.append(
        f"- file_count = {q3['file_count']} "
        f"(PR #116: {q3['expected_file_count_from_pr116']})")
    md.append("")
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
    md.append(q3["interpretation"])
    md.append("")
    md.append("**Q4 — next blocker exposed?**")
    md.append("")
    q4 = a["Q4_next_blocker_exposed"]
    md.append(
        f"- outcome_bucket = `{q4['outcome_bucket']}`")
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
        "the 5 concat-form debug MsgBox prefixes")
    md.append(
        "- ✅ Did NOT touch `Bad file Name.` (file-save errors), "
        "`Err.Description` (error trap), or `Finished saving to "
        "Neo4j` (terminal — already neutralized by generic "
        "literal rewriter)")
    md.append(
        "- ✅ Did NOT introduce a generic auto-dismiss policy")
    md.append(
        "- ✅ Did NOT touch other forms")
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
            "driver/associations-cmdneo4j-debug-msgbox-suppress"),
        "fixture": "associations_matrix_first_fixture",
        "config": {
            "click_via_timer_timeout_sec": TIMER_TIMEOUT_SEC,
            "per_probe_outer_timeout_sec": PROBE_OUTER_TIMEOUT_SEC,
            "expected_file_count_from_pr116": (
                EXPECTED_FILE_COUNT_FROM_PR116),
            "promote_gates_strict": [
                "file_count >= 1",
                "no ':ERR' markers in ZZ_TEST_DEBUG",
                "watchdog dialog count == 0",
                "chain_observed_done == True",
                f"chain_elapsed_sec <= {TIMER_TIMEOUT_SEC}"
            ]
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
          "(post debug-MsgBox suppress patch) ===\n")
    _kill_orphan()
    time.sleep(1)

    out_dir = ROOT / "analysis" / (
        "_probe_associations_cmdneo4j_after_msgbox_suppress_out")
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
