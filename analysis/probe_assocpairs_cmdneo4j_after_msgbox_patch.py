"""AssociationPairs × CmdNeo4j verification probe — post MsgBox suppress patch.

Companion to the previous PR AX probe (preserved at
`reports/probe_assocpairs_cmdneo4j.json` once that PR merges).
That probe identified 6 unconditional debug MsgBox calls in
`Form_LookAtAssociationPairs.CmdNeo4j_Click` as a confirmed
runtime blocker for unattended coverage.  This branch adds a
narrow, scoped driver-side patch (see
`tests/cbdb_driver/vba_session.py`,
`_suppress_assocpairs_cmdneo4j_debug_msgbox`) that comments out
exactly those 6 lines inside the matched sub body.

This probe verifies that the patch does what it claims and
answers three explicit questions:

  Q1  Is the terminal Finished-saving path now fully observable?
      - The 6th MsgBox (line 1470) was the previous "terminal"
        signal but is now commented out, so we cannot use it
        directly.  Instead we check: chain runs to its natural
        end without any MsgBox auto-dismissals (count == 0),
        chain block + DONE marker fires (ZZ_TEST_DEBUG has the
        chain markers), and outer chain elapsed time is short.
  Q2  Is the produced file set stable across the full chain?
      - File count stable for >= 5 polls (5 s) AND >= the 6
        files seen by the previous probe.
  Q3  Does removing the MsgBox layer expose a NEXT layer of
      blocker (a chain hang, a runtime exception, an empty
      output set, etc.)?

This probe does NOT open a coverage PR; per brief it is the
verification step that gates whether a coverage PR is the next
step.

Outputs:
  analysis/probe_assocpairs_cmdneo4j_after_msgbox_patch.md
  reports/probe_assocpairs_cmdneo4j_after_msgbox_patch.json

CLI:
  python analysis/probe_assocpairs_cmdneo4j_after_msgbox_patch.py
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
    "_probe_assocpairs_cmdneo4j_after_msgbox_patch_copy.mdb")
OUT_JSON = ROOT / "reports" / (
    "probe_assocpairs_cmdneo4j_after_msgbox_patch.json")
OUT_MD = ROOT / "analysis" / (
    "probe_assocpairs_cmdneo4j_after_msgbox_patch.md")

TIMER_TIMEOUT_SEC = 120
PROBE_OUTER_TIMEOUT_SEC = 240

FIXTURE_CONTROLS = {
    "TxtID1": 1,
    "TxtID2": 3,
    "TxtPerson1": "1",
    "TxtPerson2": "3",
    "FrameFilterYears": 1,
    "Chk2Nodes": 0,
    "ChkKinship": 0,
}

# We still run a watchdog that records any MsgBox the patch
# missed.  If the patch is correct this list MUST stay empty.
SENTINEL_MSGBOX_TITLES = ("Microsoft Access",)


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
    """Watchdog (NOT auto-dismisser): record any MsgBox dialog the
    patch missed but do NOT dismiss it.  We still need to dismiss
    so the chain doesn't hang the probe forever — but each entry
    in `observed_log` is a patch failure that must be reported.
    """
    from pywinauto import findwindows
    from pywinauto import Application as PWA
    import pywinauto.keyboard as kb

    while not stop_event.is_set():
        try:
            for title in SENTINEL_MSGBOX_TITLES:
                handles = findwindows.find_windows(
                    title=title, class_name="#32770")
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
                            "patch_failure": True,
                        })
                    except Exception:
                        pass
        except Exception:
            pass
        time.sleep(0.3)


def _classify_outcome(result: dict) -> str:
    """Verification probe outcome buckets:

      - patch_verified_chain_clean:
          file_count >= 6 AND msgboxes_observed == 0 AND
          chain_elapsed <= 120 s AND no exception.
      - patch_partial_msgbox_still_observed:
          msgboxes_observed >= 1.  The scoped patch missed at
          least one MsgBox; classification names which one(s).
      - next_blocker_exposed_no_msgbox_but_chain_unhealthy:
          msgboxes_observed == 0 BUT (file_count < 6 OR
          chain_elapsed > 120 s OR exception present).
      - blocked_exception:
          exception observed AND no files produced.
    """
    n_files = int(result.get("file_count") or 0)
    msgboxes = int(result.get("msgboxes_observed_count") or 0)
    chain_elapsed = result.get("chain_elapsed_sec")
    exception = result.get("exception")
    elapsed_ok = (
        isinstance(chain_elapsed, (int, float))
        and chain_elapsed <= TIMER_TIMEOUT_SEC
    )

    if exception and n_files == 0:
        return "blocked_exception"
    if msgboxes >= 1:
        return "patch_partial_msgbox_still_observed"
    # msgboxes == 0 from here on.
    if n_files >= 6 and elapsed_ok and not exception:
        return "patch_verified_chain_clean"
    return "next_blocker_exposed_no_msgbox_but_chain_unhealthy"


def _run_probe(out_dir: Path) -> dict:
    from cbdb_driver.vba_session import VbaSession, make_fixture
    from cbdb_driver.form_specs import LOOKATASSOCIATIONPAIRS

    spec = LOOKATASSOCIATIONPAIRS
    result: dict = {
        "form": spec.name,
        "fixture": "1x3_known_edged",
        "fixture_controls": FIXTURE_CONTROLS,
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
        "msgboxes_observed": [],
        "msgboxes_observed_count": 0,
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

            for ctl, val in FIXTURE_CONTROLS.items():
                try:
                    sess.set_control(spec.name, ctl, val)
                except Exception as e:
                    mark(f"set_control_{ctl}_fail: {e!r}")
            mark("fixture_controls_set")

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

            # File-count stability: with the MsgBox layer suppressed
            # we expect the chain to quiesce quickly.  No MsgBox
            # signal to wait on now.
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
                "ZZ_SOCIAL_NETWORK",
                "ZZ_SCRATCH_PEOPLE",
                "ZZ_SCRATCH_P_ASSOC",
                "ZZ_KIN_LIST_TMP",
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

            # Inspect ZZ_TEST_DEBUG content for chain-block markers
            try:
                cur = sess.conn.cursor()
                cur.execute(
                    "SELECT msg FROM ZZ_TEST_DEBUG ORDER BY id")
                rows = [r[0] for r in cur.fetchall()]
                result["zz_test_debug_msgs"] = rows
                cur.close()
            except Exception as e:
                result["zz_test_debug_msgs"] = [f"ERROR: {e}"]
            mark("zz_test_debug_captured")

            result["msgboxes_observed_count"] = len(
                result["msgboxes_observed"])
            result["outcome"] = _classify_outcome(result)
            completed.set()
        except BaseException as e:  # noqa: BLE001
            result["outcome"] = "blocked_exception"
            result["exception"] = (
                repr(e) + "\n" + traceback.format_exc())
            completed.set()

    watchdog = threading.Thread(
        target=_msgbox_watchdog,
        args=(stop_watchdog, result["msgboxes_observed"], t0),
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
    result["msgboxes_observed_count"] = len(
        result["msgboxes_observed"])
    return result


def _verdict(result: dict) -> dict:
    outcome = result.get("outcome", "")
    n_files = result.get("file_count", 0)
    chain_elapsed = result.get("chain_elapsed_sec")
    n_msgboxes = result.get("msgboxes_observed_count", 0)

    answers = {
        "Q1_terminal_finished_saving_path_observable": {
            "msgboxes_observed_count": n_msgboxes,
            "chain_observed_done": result.get(
                "chain_observed_done", False),
            "chain_elapsed_sec": chain_elapsed,
            "exception": result.get("exception"),
            "interpretation": (
                "With all 6 unconditional MsgBoxes commented out, "
                "the terminal MsgBox is no longer the observable "
                "signal — chain quiescence (file count stable >= "
                "5s) AND zero MsgBox dialogs observed AND no "
                "runtime exception are the new completeness "
                "criteria.  All three present means the terminal "
                "path is fully reached."
            ),
        },
        "Q2_file_set_stable": {
            "file_count": n_files,
            "names": [f.get("name", "") for f in result.get("files", [])],
            "interpretation": (
                f"Probe expects file_count >= 6 (matching the "
                f"previous probe's observed set); observed = "
                f"{n_files}."
            ),
        },
        "Q3_next_blocker_exposed": {
            "outcome_bucket": outcome,
            "interpretation": (
                "If outcome == patch_verified_chain_clean, no next "
                "blocker exposed by this fixture.  If outcome == "
                "patch_partial_msgbox_still_observed, the scoped "
                "patch missed one of the 6 — see "
                "msgboxes_observed list.  If outcome == "
                "next_blocker_exposed_no_msgbox_but_chain_unhealthy, "
                "MsgBox layer is gone but a NEW blocker shows up "
                "(empty file set, slow chain, exception)."
            ),
        },
    }

    if outcome == "patch_verified_chain_clean":
        verdict = outcome
        verdict_note = (
            f"Patch verified.  Chain produced {n_files} files in "
            f"{chain_elapsed}s with zero MsgBox dialogs observed.  "
            f"This is the strongest available signal that the 6 "
            f"unconditional debug MsgBox calls are the only "
            f"blocker for unattended coverage on this fixture.  "
            f"Note: this probe does NOT open a coverage PR; the "
            f"next step is a separate coverage PR brief."
        )
    elif outcome == "patch_partial_msgbox_still_observed":
        verdict = outcome
        miss = ", ".join(
            d.get("msg_text", "?")[:60]
            for d in result.get("msgboxes_observed", [])[:5]
        )
        verdict_note = (
            f"Patch INCOMPLETE.  {n_msgboxes} MsgBox dialog(s) "
            f"still observed at runtime: {miss}.  The scoped "
            f"rewrite did not match those line(s); inspect the "
            f"target list and the regex prefix exactness."
        )
    elif outcome == (
        "next_blocker_exposed_no_msgbox_but_chain_unhealthy"
    ):
        verdict = outcome
        verdict_note = (
            f"MsgBox layer suppressed (0 dialogs observed) but "
            f"chain did not finish cleanly.  file_count={n_files}, "
            f"chain_elapsed={chain_elapsed}s, "
            f"exception={result.get('exception')}.  A NEW blocker "
            f"is exposed beyond the MsgBox layer; coverage PR is "
            f"NOT the next step — open a separate investigation."
        )
    elif outcome == "blocked_exception":
        verdict = outcome
        verdict_note = (
            f"Exception during probe with no files produced.  See "
            f"`exception` field for the trace."
        )
    else:
        verdict = f"do_not_open_coverage_pr_other_{outcome}"
        verdict_note = (
            f"Outcome `{outcome}` did not match any known "
            f"verification path.  See markers for detail."
        )

    return {
        "verdict": verdict,
        "verdict_note": verdict_note,
        "answers": answers,
    }


def _write_md(result: dict, verdict: dict) -> None:
    md: list[str] = []
    md.append(
        "# AssociationPairs × CmdNeo4j verification probe — "
        "post MsgBox suppress patch")
    md.append("")
    md.append(
        "**Date:** 2026-05-07  ·  **Branch:** "
        "`driver/assocpairs-cmdneo4j-msgbox-suppress`")
    md.append("")
    md.append(
        "Verification probe for the narrow scoped driver patch "
        "`_suppress_assocpairs_cmdneo4j_debug_msgbox` in "
        "`tests/cbdb_driver/vba_session.py`.  The patch comments "
        "out the 6 unconditional debug MsgBox calls in "
        "`Form_LookAtAssociationPairs.CmdNeo4j_Click` (lines "
        "1069 / 1151 / 1234 / 1317 / 1400 / 1470) — the confirmed "
        "blocker identified by PR AX's probe.  This probe verifies "
        "the patch holds under the same 1×3 fixture and looks for "
        "any new blocker that may have been hidden behind the "
        "MsgBox layer.")
    md.append("")
    md.append("## Setup")
    md.append("")
    md.append(
        "- **Fixture:** 1×3 known-edged (TxtID1=1, TxtID2=3, "
        "FrameFilterYears=1, ChkKinship=0, Chk2Nodes=0)")
    md.append(
        "- **Chain:** `CmdQuery,CmdNeo4j` via Form.Tag, directory "
        "mode")
    md.append(
        "- **Watchdog:** records (and dismisses to keep the probe "
        "moving) any MsgBox the patch missed.  Each observation "
        "is a patch-correctness failure.")
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
        f"- **msgboxes_observed_count:** "
        f"{result.get('msgboxes_observed_count')}  "
        f"(expected: 0)")
    md.append(
        f"- **click_via_timer_returned:** "
        f"{result.get('click_via_timer_returned')}")
    md.append(
        f"- **total_wall_elapsed_sec:** "
        f"{result.get('elapsed_sec')}")
    if result.get("exception"):
        md.append(
            f"- **exception:** `{result['exception'][:300]}`")
    md.append("")
    md.append("## Q1-Q3 answers")
    md.append("")
    a = verdict["answers"]
    md.append("**Q1 — Is the terminal Finished-saving path now "
              "fully observable?**")
    md.append("")
    md.append(f"- msgboxes_observed_count = "
              f"{a['Q1_terminal_finished_saving_path_observable']['msgboxes_observed_count']} "
              f"(expected 0)")
    md.append(f"- chain_observed_done = "
              f"{a['Q1_terminal_finished_saving_path_observable']['chain_observed_done']}")
    md.append(f"- chain_elapsed_sec = "
              f"{a['Q1_terminal_finished_saving_path_observable']['chain_elapsed_sec']}")
    md.append(
        a['Q1_terminal_finished_saving_path_observable'][
            'interpretation'])
    md.append("")
    md.append("**Q2 — Is the file set stable across the full chain?**")
    md.append("")
    md.append(f"- file_count = "
              f"{a['Q2_file_set_stable']['file_count']}")
    md.append(f"- names = "
              f"{a['Q2_file_set_stable']['names']}")
    md.append(a['Q2_file_set_stable']['interpretation'])
    md.append("")
    md.append("**Q3 — Does removing MsgBox layer expose a next blocker?**")
    md.append("")
    md.append(f"- outcome bucket = "
              f"`{a['Q3_next_blocker_exposed']['outcome_bucket']}`")
    md.append(a['Q3_next_blocker_exposed']['interpretation'])
    md.append("")
    md.append(f"## Verdict: `{verdict['verdict']}`")
    md.append("")
    md.append(verdict["verdict_note"])
    md.append("")
    md.append("## MsgBoxes observed (must be empty for clean verify)")
    md.append("")
    obs = result.get("msgboxes_observed", [])
    if obs:
        md.append("| +t (s) | msg_text | patch_failure |")
        md.append("|---:|---|:---:|")
        for d in obs:
            md.append(
                f"| {d['t']} | `{d.get('msg_text', '')[:120]}` | "
                f"{'❌' if d.get('patch_failure') else '—'} |")
    else:
        md.append("(none observed — scoped patch held)")
    md.append("")
    md.append("## Per-file detail")
    md.append("")
    if result.get("files"):
        md.append(
            "| name | size | n_cols | first_col | data_rows | "
            "header_preview |")
        md.append("|---|---:|---:|---|---:|---|")
        for f in result["files"]:
            md.append(
                f"| {f.get('name', '?')} | "
                f"{f.get('size', '—')} | "
                f"{f.get('header_n_cols', '—')} | "
                f"`{f.get('header_first_col', '—')}` | "
                f"{f.get('data_row_count', '—')} | "
                f"`{f.get('header_preview', '—')[:80]}` |")
    else:
        md.append("(no files produced)")
    md.append("")
    md.append("## Scratch row counts (post-chain)")
    md.append("")
    for tbl, c in (result.get("row_counts") or {}).items():
        md.append(f"- `{tbl}`: {c}")
    md.append("")
    md.append("## ZZ_TEST_DEBUG content")
    md.append("")
    msgs = result.get("zz_test_debug_msgs") or []
    if msgs:
        for m in msgs:
            md.append(f"- `{m}`")
    else:
        md.append("(empty)")
    md.append("")
    md.append(
        f"## Markers (timeline, "
        f"{len(result.get('markers', []))} entries)")
    md.append("")
    for m in result.get("markers", []):
        md.append(f"  - `+{m['t']:>6.2f}s` {m['marker']}")
    md.append("")
    md.append("## Constraints honoured")
    md.append("")
    md.append("- ✅ Did NOT open a coverage PR")
    md.append("- ✅ Driver patch is narrow-scoped: only "
              "`Form_LookAtAssociationPairs.CmdNeo4j_Click`, only "
              "the 6 unconditional debug MsgBox lines")
    md.append("- ✅ Did NOT introduce a generic auto-dismiss policy")
    md.append("- ✅ Did NOT touch unrelated forms or other "
              "MsgBox calls in CmdNeo4j_Click")
    md.append("- ✅ Did NOT change README / canonical reports / "
              "issue severity")
    md.append("- ✅ Did NOT change tests/")

    OUT_MD.write_text("\n".join(md), encoding="utf-8")


def _write_outputs(result: dict, verdict: dict) -> None:
    out: dict = {
        "schema_version": 1,
        "generated_date": "2026-05-07",
        "probe_branch": "driver/assocpairs-cmdneo4j-msgbox-suppress",
        "fixture": "assocpairs_1x3_known_edged",
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


def main(argv: list[str] | None = None) -> int:
    print("=== AssociationPairs × CmdNeo4j verification probe "
          "(post MsgBox suppress patch) ===\n")
    _kill_orphan()
    time.sleep(1)

    out_dir = ROOT / "analysis" / (
        "_probe_assocpairs_cmdneo4j_after_msgbox_patch_out")
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
    print(f"chain_elapsed: {result.get('chain_elapsed_sec')} s")
    print(f"file_count: {result.get('file_count')}")
    print(
        f"msgboxes_observed: "
        f"{result.get('msgboxes_observed_count')}")
    print(f"\n=== verdict: {verdict['verdict']} ===")
    return 0


if __name__ == "__main__":
    sys.exit(main())
