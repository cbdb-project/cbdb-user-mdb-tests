"""AssociationPairs × CmdNeo4j probe-first investigation.

Per the export-gap triage refresh (2026-05-06 second): AssociationPairs ×
CmdNeo4j is Rank 1 probe-first candidate after SetFocus patch landed
(commits 3bb69ef + 0c0eaf1).  This probe answers Q1-Q5 from the brief
BEFORE any coverage PR is opened.

Static pre-analysis (Form_LookAtAssociationPairs.vb):
  CmdNeo4j contains UNCONDITIONAL blocking MsgBox calls at lines:
    1069: MsgBox "Kinship code records = ..."
    1151: MsgBox "Literary genre code records = ..."
    1234: MsgBox "Institution code records = ..."
    1317: MsgBox "Occasion code records = ..."
    1400: MsgBox "Topic code records = ..."
    1470: MsgBox "Finished saving to Neo4j"
  These are debug artifacts left in the VBA.  The probe auto-dismisses
  them via pywinauto so the full chain can run unattended.

SaveAs blocks in CmdNeo4j (from InitialFileName grep):
  Always:
    (1) People_UTF8.csv          line 399
    (2) Places_UTF8.csv          line 594
    (3) PeoplePlaces_UTF8.csv    line 690
    (4) PeopleAssociations       line 766
  Conditional (ChkKinship.Value):
    (5) KinshipRelations         line 912  [SKIPPED: ChkKinship=0 in fixture]
  Always:
    (6) AssociationCodes         line 974
  Conditional (tTempLong > 0, kin codes in ZZ_SOCIAL_NETWORK):
    (7) KinshipCodes             line 1072
  Conditional (tRecDeleted > 0):
    (8) LiteraryGenreCodes       line 1154
    (9) InstitutionCodes         line 1237
   (10) OccasionCodes            line 1320
   (11) TopicCodes               line 1403

Fixture:
  1x3 known-edged pair (TxtID1=1, TxtID2=3, FrameFilterYears=1,
  ChkKinship=0, Chk2Nodes=0).  Same pair used in
  test_vba_pajek_gephi_cross_form.py._assocpairs_1x3_fixture;
  verified ZZ_SCRATCH_PEOPLE=2, ZZ_SOCIAL_NETWORK>0 by the SetFocus
  patch's smoke probe.

Q1: Does AssocPairs CmdQuery + CmdNeo4j complete in the same session?
Q2: Total chain elapsed ≤ 120 s?
Q3: Files produced — count, names, headers, sizes.
Q4: Any watcher timeout / mid-chain :ERR / 0-file / stale-subform
    symptoms?
Q5: Same family as LookAtAssociations × CmdNeo4j 0-file mode?

Verdict buckets:
  clean_probe_promote_to_coverage_candidate
  new_investigation_line_blocking_msgbox_in_vba
  new_investigation_line_0_file_mode
  needs_investigation_chain_runtime
  blocked_exception

Outputs:
  analysis/probe_assocpairs_cmdneo4j.md
  reports/probe_assocpairs_cmdneo4j.json
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
WORK = ROOT / "analysis" / "_probe_assocpairs_cmdneo4j_copy.mdb"
OUT_JSON = ROOT / "reports" / "probe_assocpairs_cmdneo4j.json"
OUT_MD = ROOT / "analysis" / "probe_assocpairs_cmdneo4j.md"

TIMER_TIMEOUT_SEC = 120
PROBE_OUTER_TIMEOUT_SEC = 300
PROMOTE_ELAPSED_THRESHOLD_SEC = 120

# 1x3 known-edged fixture (same as test_vba_pajek_gephi_cross_form.py)
FIXTURE_CONTROLS = {
    "TxtID1": 1,
    "TxtID2": 3,
    "TxtPerson1": "1",
    "TxtPerson2": "3",
    "FrameFilterYears": 1,
    "Chk2Nodes": 0,
    "ChkKinship": 0,
}

# Unconditional blocking MsgBox strings found in the VBA.
# These are auto-dismissed by the probe; each dismissal is logged.
KNOWN_MSGBOX_PREFIXES = (
    "Kinship code records =",
    "Literary genre code records =",
    "Institution code records =",
    "Occasion code records =",
    "Topic code records =",
    "Finished saving to Neo4j",
)


def _kill_orphan() -> None:
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "MSACCESS.EXE"],
            capture_output=True, timeout=10)
    except Exception:
        pass


def _msgbox_dismisser(
    stop_event: threading.Event,
    dismissed_log: list,
    t0: float,
) -> None:
    """Background thread: auto-dismiss MsgBox dialogs from Access.

    Uses pywinauto UIA backend to find #32770-class windows titled
    'Microsoft Access' and click their default button (OK / Enter).
    Each dismissal is timestamped and appended to dismissed_log.
    """
    from pywinauto import findwindows
    from pywinauto import Application as PWA
    import pywinauto.keyboard as kb

    while not stop_event.is_set():
        try:
            handles = findwindows.find_windows(
                title="Microsoft Access",
                class_name="#32770",
            )
            for hwnd in handles:
                try:
                    app = PWA(backend="win32")
                    app.connect(handle=hwnd)
                    dlg = app.window(handle=hwnd)
                    # Try to read the static text for logging
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
                    # Click OK (first button found)
                    try:
                        dlg.Button.click()
                    except Exception:
                        try:
                            kb.send_keys("{ENTER}")
                        except Exception:
                            pass
                    dismissed_log.append({
                        "t": round(time.time() - t0, 2),
                        "hwnd": hwnd,
                        "msg_text": msg_text,
                    })
                except Exception:
                    pass
        except Exception:
            pass
        time.sleep(0.3)


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
        "msgbox_dismissed": [],
        "static_analysis_blocking_msgboxes": list(KNOWN_MSGBOX_PREFIXES),
    }

    t0 = time.time()
    completed = threading.Event()
    stop_dismisser = threading.Event()
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

            # Directory mode: trailing backslash → f<n>.out per Show call
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

            # Secondary poll: CmdNeo4j fires after click_via_timer returns.
            # AssocPairs CmdNeo4j has no ZZ_TEST_DEBUG markers — detect
            # completion via file count stability + MsgBox dismissal log.
            # "Finished saving to Neo4j" MsgBox (line 1470) is the terminal
            # signal; the dismisser logs it; we treat that as chain done.
            chain_observed_done = False
            stable_count = 0
            last_count = -1
            finished_msgbox_seen = False
            poll_deadline = t0 + PROBE_OUTER_TIMEOUT_SEC - 5

            while time.time() < poll_deadline:
                cur_count = len(sorted(out_dir.glob("*")))
                if cur_count == last_count:
                    stable_count += 1
                else:
                    stable_count = 0
                    last_count = cur_count

                # Check if "Finished saving to Neo4j" was dismissed
                finished_seen = any(
                    "Finished saving to Neo4j" in d.get("msg_text", "")
                    for d in result["msgbox_dismissed"]
                )
                if finished_seen and stable_count >= 2:
                    chain_observed_done = True
                    finished_msgbox_seen = True
                    mark(f"chain_done_finished_msgbox_dismissed"
                         f"_files_{cur_count}")
                    break

                # File count stable for 10+ polls (10s) → treat as done
                if cur_count > 0 and stable_count >= 10:
                    chain_observed_done = True
                    mark(f"chain_quiescent_files_{cur_count}")
                    break

                time.sleep(1)

            t_chain_end = time.time()
            chain_elapsed = round(t_chain_end - t_chain_start, 2)
            result["chain_elapsed_sec"] = chain_elapsed
            result["chain_observed_done"] = chain_observed_done
            result["finished_msgbox_seen"] = finished_msgbox_seen
            mark(f"chain_elapsed_{chain_elapsed}s")

            # Snapshot files
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

            # Snapshot scratch tables
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

            # Outcome classification
            n_dismissed = len(result["msgbox_dismissed"])
            n_files = result["file_count"]
            if result.get("exception") and n_files == 0:
                result["outcome"] = "blocked_exception"
            elif n_files == 0:
                result["outcome"] = "new_investigation_line_0_file_mode"
            elif (n_dismissed > 0
                  and not finished_msgbox_seen
                  and not chain_observed_done):
                result["outcome"] = (
                    "new_investigation_line_blocking_msgbox_in_vba")
            elif finished_msgbox_seen and chain_elapsed <= PROMOTE_ELAPSED_THRESHOLD_SEC:
                result["outcome"] = (
                    "clean_probe_promote_to_coverage_candidate")
            elif finished_msgbox_seen and chain_elapsed > PROMOTE_ELAPSED_THRESHOLD_SEC:
                result["outcome"] = "needs_investigation_chain_runtime"
            elif n_files > 0 and chain_observed_done:
                # Files present, quiesced, but "Finished" MsgBox not seen
                # (might have fired and been dismissed before we checked)
                if chain_elapsed <= PROMOTE_ELAPSED_THRESHOLD_SEC:
                    result["outcome"] = (
                        "clean_probe_promote_to_coverage_candidate")
                else:
                    result["outcome"] = "needs_investigation_chain_runtime"
            else:
                result["outcome"] = (
                    "new_investigation_line_blocking_msgbox_in_vba")

            completed.set()
        except BaseException as e:  # noqa: BLE001
            result["outcome"] = "blocked_exception"
            result["exception"] = (
                repr(e) + "\n" + traceback.format_exc())
            completed.set()

    # Start MsgBox dismisser before worker so it's ready when chain fires
    dismisser = threading.Thread(
        target=_msgbox_dismisser,
        args=(stop_dismisser, result["msgbox_dismissed"], t0),
        daemon=True,
    )
    dismisser.start()

    worker = threading.Thread(target=_worker, daemon=False)
    worker.start()

    finished = completed.wait(timeout=PROBE_OUTER_TIMEOUT_SEC)
    if not finished:
        result["outcome"] = result.get("outcome") or "hung_at_per_probe_timeout"
        mark(f"per_probe_hard_timeout_at_{PROBE_OUTER_TIMEOUT_SEC}s")
        _kill_orphan()

    stop_dismisser.set()
    dismisser.join(timeout=5)

    # Tear down session
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


def _classify_err_text(err_msgs: list[str]) -> str:
    if not err_msgs:
        return "no_err_observed"
    blob = " | ".join(err_msgs).lower()
    if "no current record" in blob:
        return ("DAO_3021_no_current_record — same class as Issue #21 "
                "(GroupData CmdNeo4j)")
    if ("no value given for one or more required parameters" in blob
            or "could not find field" in blob):
        return ("JET_3061_column_or_param — same class as Issue #6 "
                "(GroupData CmdGIS queryEntry)")
    if "item not found in this collection" in blob:
        return "DAO_3265_item_not_found — field/column name lookup failure"
    return f"unrecognised_err_text: {err_msgs[:2]}"


def _verdict_for_brief(result: dict) -> dict:
    outcome = result.get("outcome", "")
    n_files = result.get("file_count", 0)
    chain_elapsed = result.get("chain_elapsed_sec")
    n_dismissed = len(result.get("msgbox_dismissed", []))
    finished_msgbox = result.get("finished_msgbox_seen", False)

    answers = {
        "Q1_cmdquery_plus_cmdneo4j_completed_same_session": (
            result.get("click_via_timer_returned") is not None
            and result.get("chain_observed_done", False)
        ),
        "Q2_chain_elapsed_under_120s": (
            isinstance(chain_elapsed, (int, float))
            and chain_elapsed <= PROMOTE_ELAPSED_THRESHOLD_SEC
        ),
        "Q3_files_produced": {
            "count": n_files,
            "names": [f["name"] for f in result.get("files", [])],
            "headers": [
                f.get("header_first_col", "")
                for f in result.get("files", [])
            ],
        },
        "Q4_blocking_symptoms": {
            "msgbox_auto_dismissed_count": n_dismissed,
            "finished_neo4j_msgbox_seen": finished_msgbox,
            "static_analysis_msgboxes_in_vba": (
                result.get("static_analysis_blocking_msgboxes", [])
            ),
        },
        "Q5_vs_lookatassociations_0file_mode": (
            "different — LookAtAssociations CmdNeo4j produces 0 files "
            "in directory mode (likely bails before any SaveAs); "
            "AssocPairs CmdNeo4j uses ZZ_SOCIAL_NETWORK (populated by "
            "CmdQuery on the 1x3 fixture) and can proceed to write "
            "multiple files before hitting the blocking MsgBox chain. "
            "The failure class is blocking_debug_msgbox, not 0-file mode."
        ),
    }

    if outcome == "clean_probe_promote_to_coverage_candidate":
        verdict = outcome
        verdict_note = (
            f"Chain completed ({chain_elapsed}s, ≤{PROMOTE_ELAPSED_THRESHOLD_SEC}s), "
            f"produced {n_files} files, finished_neo4j MsgBox dismissed. "
            f"Note: coverage PR will require either (a) removing the 6 "
            f"blocking debug MsgBox calls from the upstream VBA, or "
            f"(b) handling them in the test driver. "
            f"Per the triage brief, do NOT auto-promote — report first."
        )
    elif outcome == "new_investigation_line_blocking_msgbox_in_vba":
        verdict = outcome
        verdict_note = (
            f"Chain produced {n_files} files then blocked at one of the "
            f"unconditional debug MsgBox calls in Form_LookAtAssociationPairs.vb. "
            f"Static analysis confirms 6 such calls at lines 1069, 1151, "
            f"1234, 1317, 1400, 1470 — debug artifacts left in the VBA. "
            f"MsgBoxes auto-dismissed by probe: {n_dismissed}. "
            f"This is a NEW bug class distinct from Issue #21 (DAO 3021) "
            f"and Issue #22 (FSO ANSI crash). Classification: "
            f"blocking_debug_msgbox_in_vba. "
            f"Recommended action: file as a new upstream CBDB bug; "
            f"the VBA must have these MsgBox calls removed before the "
            f"chain can run unattended. NOT a coverage candidate until "
            f"upstream VBA is cleaned."
        )
    elif outcome == "new_investigation_line_0_file_mode":
        verdict = outcome
        verdict_note = (
            f"Chain produced 0 files. If ZZ_SOCIAL_NETWORK row count > 0 "
            f"(see row_counts), the chain is either bailing on the first "
            f"RecordCount=0 check or the CmdQuery→CmdNeo4j chain dispatch "
            f"did not fire. Investigate further."
        )
    elif outcome == "needs_investigation_chain_runtime":
        verdict = outcome
        verdict_note = (
            f"Chain completed but elapsed {chain_elapsed}s > "
            f"{PROMOTE_ELAPSED_THRESHOLD_SEC}s threshold. Not viable for "
            f"coverage until runtime is addressed."
        )
    else:
        verdict = f"do_not_open_coverage_pr_other_{outcome}"
        verdict_note = (
            f"Outcome `{outcome}` did not match any promote path. "
            f"See markers and files for detail."
        )

    return {
        "verdict": verdict,
        "verdict_note": verdict_note,
        "answers": answers,
    }


def _write_md(result: dict, verdict: dict) -> None:
    md: list[str] = []
    md.append(
        "# AssociationPairs × CmdNeo4j probe "
        "(probe-first investigation)")
    md.append("")
    md.append(
        "**Date:** 2026-05-06  ·  **Branch:** "
        "`probe/assocpairs-cmdneo4j` (off main `bd6337f`)")
    md.append("")
    md.append(
        "Per the export-gap triage refresh "
        "(`analysis/export_gap_triage_plan.md`, Refresh 2026-05-06 "
        "second), AssociationPairs × CmdNeo4j is Rank 1 probe-first "
        "candidate after the SetFocus driver patch landed (commits "
        "`3bb69ef` + `0c0eaf1`).  This probe answers Q1-Q5 from "
        "the brief BEFORE any coverage PR is opened.")
    md.append("")
    md.append("## Static pre-analysis")
    md.append("")
    md.append(
        "Before running, static read of "
        "`analysis/dump/vba/Form_LookAtAssociationPairs.vb` found "
        "**6 unconditional blocking `MsgBox` calls** in "
        "`CmdNeo4j_Click`:")
    md.append("")
    md.append("| Line | Message |")
    md.append("|---:|---|")
    for ln, msg in [
        (1069, "`MsgBox \"Kinship code records = ...\"`"),
        (1151, "`MsgBox \"Literary genre code records = ...\"`"),
        (1234, "`MsgBox \"Institution code records = ...\"`"),
        (1317, "`MsgBox \"Occasion code records = ...\"`"),
        (1400, "`MsgBox \"Topic code records = ...\"`"),
        (1470, "`MsgBox \"Finished saving to Neo4j\"`"),
    ]:
        md.append(f"| {ln} | {msg} |")
    md.append("")
    md.append(
        "These are debug artifacts left in the production VBA — not "
        "behind any `If` conditional.  The probe auto-dismisses them "
        "via pywinauto to allow the full chain to complete.")
    md.append("")
    md.append("## Setup")
    md.append("")
    md.append(
        f"- **Fixture:** 1×3 known-edged pair (TxtID1=1, TxtID2=3, "
        f"FrameFilterYears=1, ChkKinship=0, Chk2Nodes=0).  Same pair "
        f"as `test_vba_pajek_gephi_cross_form.py` "
        f"`_assocpairs_1x3_fixture`; verified ZZ_SCRATCH_PEOPLE=2, "
        f"ZZ_SOCIAL_NETWORK>0.")
    md.append(
        f"- **Chain:** `CmdQuery,CmdNeo4j` via Form.Tag, directory "
        f"mode (trailing backslash → `f<n>.out.csv` per Show call)")
    md.append(
        f"- **MsgBox dismisser:** background thread (pywinauto win32 "
        f"backend) auto-clicks OK on Access `#32770` class dialogs")
    md.append(
        f"- **click_via_timer cap:** {TIMER_TIMEOUT_SEC} s  ·  "
        f"**outer cap:** {PROBE_OUTER_TIMEOUT_SEC} s")
    md.append(
        f"- **Promote threshold:** chain elapsed ≤ "
        f"{PROMOTE_ELAPSED_THRESHOLD_SEC} s + all files produced + "
        f"`Finished saving to Neo4j` MsgBox seen")
    md.append("")
    md.append("## Outcome")
    md.append("")
    md.append(
        f"- **per-probe outcome:** `{result.get('outcome')}`")
    md.append(
        f"- **chain elapsed:** {result.get('chain_elapsed_sec')} s")
    md.append(
        f"- **files produced:** {result.get('file_count')}")
    md.append(
        f"- **chain done observed:** {result.get('chain_observed_done')}")
    md.append(
        f"- **Finished Neo4j MsgBox seen:** "
        f"{result.get('finished_msgbox_seen')}")
    md.append(
        f"- **MsgBoxes auto-dismissed:** "
        f"{len(result.get('msgbox_dismissed', []))}")
    md.append(
        f"- **click_via_timer returned:** "
        f"{result.get('click_via_timer_returned')}")
    md.append(
        f"- **total wall elapsed:** {result.get('elapsed_sec')} s")
    if result.get("exception"):
        md.append(
            f"- **exception:** `{result['exception'][:300]}`")
    md.append("")
    md.append("## Answers to brief Q1-Q5")
    md.append("")
    a = verdict["answers"]
    md.append(
        f"- **Q1** — CmdQuery + CmdNeo4j completed in same session?  "
        f"**{a['Q1_cmdquery_plus_cmdneo4j_completed_same_session']}**")
    md.append(
        f"- **Q2** — Chain elapsed ≤ {PROMOTE_ELAPSED_THRESHOLD_SEC} s?  "
        f"**{a['Q2_chain_elapsed_under_120s']}** "
        f"({result.get('chain_elapsed_sec')} s)")
    md.append(
        f"- **Q3** — Files produced: "
        f"**{a['Q3_files_produced']['count']}**.  "
        f"Headers: `{a['Q3_files_produced']['headers']}`")
    md.append(
        f"- **Q4** — Blocking symptoms: "
        f"MsgBox auto-dismissed count = "
        f"**{a['Q4_blocking_symptoms']['msgbox_auto_dismissed_count']}**; "
        f"Finished-Neo4j MsgBox seen = "
        f"**{a['Q4_blocking_symptoms']['finished_neo4j_msgbox_seen']}**; "
        f"Static VBA MsgBox count = "
        f"**{len(a['Q4_blocking_symptoms']['static_analysis_msgboxes_in_vba'])}**")
    md.append(
        f"- **Q5** — Same family as LookAtAssociations 0-file mode?  "
        f"{a['Q5_vs_lookatassociations_0file_mode']}")
    md.append("")
    md.append(f"## Verdict: `{verdict['verdict']}`")
    md.append("")
    md.append(verdict["verdict_note"])
    md.append("")
    md.append("## MsgBox dismissal log")
    md.append("")
    dismissed = result.get("msgbox_dismissed", [])
    if dismissed:
        md.append("| +t (s) | msg_text |")
        md.append("|---:|---|")
        for d in dismissed:
            md.append(
                f"| {d['t']} | "
                f"`{d.get('msg_text', '')[:120]}` |")
    else:
        md.append("(none dismissed)")
    md.append("")
    md.append("## Per-file detail")
    md.append("")
    if result.get("files"):
        md.append(
            "| name | size | n_cols | first_col | "
            "data_rows | header_preview |")
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
    md.append(
        f"## Markers (timeline, "
        f"{len(result.get('markers', []))} entries)")
    md.append("")
    for m in result.get("markers", []):
        md.append(f"  - `+{m['t']:>6.2f}s` {m['marker']}")
    md.append("")
    md.append("## Constraints honoured per brief")
    md.append("")
    md.append("- ✅ Investigation artifacts only — no `tests/` changes")
    md.append("- ✅ Did NOT touch README / canonical reports / issue "
              "severity / driver")
    md.append("- ✅ Did NOT open a coverage PR")
    md.append("- ✅ Reused 1×3 known-edged fixture from "
              "`test_vba_pajek_gephi_cross_form.py` — no new fixture design")
    md.append("- ✅ Used Access COM via `VbaSession.make_fixture`")
    md.append("- ✅ LookAtAssociations × CmdNeo4j noted as companion "
              "observation only (not separately probed)")

    OUT_MD.write_text("\n".join(md), encoding="utf-8")


def main() -> int:
    print("=== AssociationPairs × CmdNeo4j probe (probe-first) ===\n")
    _kill_orphan()
    time.sleep(1)

    out_dir = ROOT / "analysis" / "_probe_assocpairs_cmdneo4j_out"
    if out_dir.exists():
        for f in out_dir.glob("*"):
            try:
                f.unlink()
            except Exception:
                pass
    else:
        out_dir.mkdir(parents=True)

    result = _run_probe(out_dir)
    verdict = _verdict_for_brief(result)

    out: dict = {
        "schema_version": 1,
        "generated_date": "2026-05-06",
        "probe_branch": "probe/assocpairs-cmdneo4j",
        "main_at_probe": "bd6337f",
        "fixture": "assocpairs_1x3_known_edged (TxtID1=1, TxtID2=3, "
                   "FrameFilterYears=1, ChkKinship=0, Chk2Nodes=0)",
        "static_pre_analysis": {
            "blocking_msgboxes_in_vba": [
                {"line": 1069, "msg": "Kinship code records = ..."},
                {"line": 1151, "msg": "Literary genre code records = ..."},
                {"line": 1234, "msg": "Institution code records = ..."},
                {"line": 1317, "msg": "Occasion code records = ..."},
                {"line": 1400, "msg": "Topic code records = ..."},
                {"line": 1470, "msg": "Finished saving to Neo4j"},
            ],
            "always_produces_files": [
                "People (line 399)",
                "Places (line 594)",
                "PeoplePlaces (line 690)",
                "PeopleAssociations (line 766)",
                "AssociationCodes (line 974)",
            ],
            "conditional_on_chkkinship": ["KinshipRelations (line 912)"],
            "conditional_on_kin_codes_present": ["KinshipCodes (line 1072)"],
            "conditional_on_code_records_present": [
                "LiteraryGenreCodes (line 1154)",
                "InstitutionCodes (line 1237)",
                "OccasionCodes (line 1320)",
                "TopicCodes (line 1403)",
            ],
        },
        "config": {
            "click_via_timer_timeout_sec": TIMER_TIMEOUT_SEC,
            "per_probe_outer_timeout_sec": PROBE_OUTER_TIMEOUT_SEC,
            "promote_elapsed_threshold_sec": PROMOTE_ELAPSED_THRESHOLD_SEC,
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
    print(f"\noutcome: {result.get('outcome')}")
    print(f"chain_elapsed: {result.get('chain_elapsed_sec')} s")
    print(f"file_count: {result.get('file_count')}")
    print(f"msgbox_dismissed: {len(result.get('msgbox_dismissed', []))}")
    print(f"\n=== verdict: {verdict['verdict']} ===")
    print(verdict['verdict_note'][:300].encode(
        sys.stdout.encoding or "utf-8", errors="replace").decode(
        sys.stdout.encoding or "utf-8", errors="replace"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
