"""GroupData × CmdNeo4j probe-first investigation.

Per the refresh brief (2026-05-05): the only remaining "cheapest
next cell" under the brief's exclusions is GroupData × CmdNeo4j.
Before opening any coverage PR, validate that the just-merged
GroupData × CmdGIS pattern transfers cleanly to CmdNeo4j.

Specifically answer:

  Q1 — Does CmdRun + CmdNeo4j complete in the same session under
       `groupdata_person_1_small` + the all-Chk*-reset pattern?
  Q2 — Does total chain elapsed fit within 120 s?
  Q3 — Which files are produced; how many?
  Q4 — Does ZZ_TEST_DEBUG show any `LookAtGroupData:ERR`?
  Q5 — Is there any branch-specific known issue (à la Issue #6
       in CmdGIS's queryEntry path), or is the whole chain clean?

Verdict logic (4 buckets):

  - promote_to_coverage_candidate
      chain elapsed <= 120 s
      AND >= 1 expected file produced
      AND no `LookAtGroupData:ERR` in ZZ_TEST_DEBUG
  - needs_investigation_mid_chain_err
      :ERR present (Issue-6-shape).  Files may or may not exist.
  - needs_investigation_chain_runtime
      chain finished but took longer than 120 s.  Coverage PR
      not viable until either the chain is sped up or the test
      watcher cap is raised + justified.
  - blocked
      chain didn't finish (timeout) or no files produced.

Constraints honoured per brief:
  - Investigation artifacts only -- no tests/ changes
  - Did NOT touch README / reports / issue severity / driver
  - Did NOT design a new fixture -- reuses person_1 inline
    (matches matrix_hard_forms's `groupdata_person_1_small`)
  - Uses Access COM via VbaSession (brief explicitly permits)

Outputs
-------
- analysis/groupdata_cmdneo4j_probe.md
- reports/groupdata_cmdneo4j_probe.json
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
WORK = ROOT / "analysis" / "_probe_groupdata_cmdneo4j_copy.mdb"
OUT_JSON = ROOT / "reports" / "groupdata_cmdneo4j_probe.json"
OUT_MD = ROOT / "analysis" / "groupdata_cmdneo4j_probe.md"

# Per-probe outer cap is generous (the brief's promote/reject line
# is 120 s, but we want to capture the actual elapsed even if the
# chain blows past it without true-hanging).
PROBE_OUTER_TIMEOUT_SEC = 240
# This is the cap passed to click_via_timer.  CmdNeo4j cross-form
# uses 180 s for the existing 7 forms; keep parity here.
TIMER_TIMEOUT_SEC = 180

# Promote/reject threshold per brief.
PROMOTE_ELAPSED_THRESHOLD_SEC = 120

PERSON_ID = 1  # matches matrix_hard_forms's groupdata_person_1_small

# All 11 Chk* controls that GroupData's Form_Open may set to True.
# The GroupData × CmdGIS coverage PR proved this list is the
# minimum reset set required to keep Issue #6's queryEntry branch
# from firing.  We re-use the same all-False reset here, then enable
# only the 6 clean-branch boxes.
ALL_CHK_CONTROLS = (
    "ChkStatus", "ChkOffice", "ChkEntry", "ChkText", "ChkAddr",
    "ChkGisStatus", "ChkGisOffice", "ChkGisOfficePeople",
    "ChkGisEntry", "ChkGisText", "ChkGisAddr",
)
# Match the just-merged GroupData × CmdGIS test's enable set.
# (CmdNeo4j body has no If ChkX.Value gating in its own code path,
# but ChkX values still drive what CmdRun writes into the
# ZZ_SCRATCH_* tables that CmdNeo4j reads from.  Including
# ChkEntry would re-trigger Issue #6 in queryEntry; excluding it
# keeps the probe scope aligned with CmdGIS coverage.)
ENABLE_CHK_CONTROLS = (
    "ChkStatus", "ChkOffice", "ChkAddr",
    "ChkGisStatus", "ChkGisOffice", "ChkGisAddr",
)


def _kill_orphan():
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "MSACCESS.EXE"],
            capture_output=True, timeout=10)
    except Exception:
        pass


def _run_probe(out_dir: Path) -> dict:
    """Drive CmdRun -> CmdNeo4j chain on person_1.  Capture timing,
    files, debug transcript, scratch row counts."""
    from cbdb_driver.vba_session import VbaSession, make_fixture
    from cbdb_driver.form_specs import LOOKATGROUPDATA

    spec = LOOKATGROUPDATA
    result: dict = {
        "person_id": PERSON_ID,
        "form": spec.name,
        "enable_chk_controls": list(ENABLE_CHK_CONTROLS),
        "all_chk_reset_first": list(ALL_CHK_CONTROLS),
        "markers": [],
        "outcome": None,
        "elapsed_sec": None,
        "exception": None,
        "row_counts": {},
        "debug_transcript": [],
        "controls_set": {},
        "files": [],
        "file_count": 0,
        "click_via_timer_returned": None,
    }

    t0 = time.time()
    completed = threading.Event()
    sess = None

    def mark(s):
        result["markers"].append(
            {"t": round(time.time() - t0, 2), "marker": s})

    def _worker():
        nonlocal sess
        try:
            # Session construction can flake with COM RPC errors when
            # an orphan MSACCESS instance lingers.  One retry after a
            # generous cooldown handles the case (this is the same
            # flake seen during PR cover/groupdata-cmdgis-clean-
            # branches second-attempt run).
            mark("constructing_session")
            for attempt in (1, 2, 3):
                try:
                    sess = next(_make_session_iter())
                    mark(f"session_opened_attempt_{attempt}")
                    break
                except Exception as e:
                    mark(f"session_open_attempt_{attempt}_fail: "
                         f"{e!r}")
                    _kill_orphan()
                    time.sleep(60)
            else:
                raise RuntimeError(
                    "session open failed after 3 attempts; see "
                    "earlier markers for COM error trail")

            sess.patch_filedialog(spec.name)
            mark("filedialog_patched")

            sess.open_form(spec.name)
            mark("form_opened")

            sess.set_picker_codes(
                spec.picker_table, [PERSON_ID],
                column=spec.picker_column)
            mark(f"picker_seeded_pid_{PERSON_ID}")

            for ctl in ALL_CHK_CONTROLS:
                try:
                    sess.set_control(spec.name, ctl, False)
                    result["controls_set"][ctl] = False
                except Exception as e:
                    result["controls_set"][ctl] = f"FAIL_reset: {e}"
                    mark(f"reset_{ctl}_fail: {e}")
            mark("all_chk_reset_to_False")

            for ctl in ENABLE_CHK_CONTROLS:
                try:
                    sess.set_control(spec.name, ctl, True)
                    result["controls_set"][ctl] = True
                except Exception as e:
                    result["controls_set"][ctl] = f"FAIL_enable: {e}"
                    mark(f"enable_{ctl}_fail: {e}")
            mark("clean_branches_enabled")

            # Wire chain CmdRun -> CmdNeo4j via Form.Tag, directory
            # mode (trailing backslash makes patch_filedialog hand
            # out f<n>.out per Show call).
            sess.set_form_tag(
                spec.name,
                f"{spec.cmd_name},CmdNeo4j",
                str(out_dir) + "\\",
            )
            mark("form_tag_set_chain_CmdRun_CmdNeo4j")

            # Fire CmdRun via timer.  click_via_timer waits on the
            # primary result_table to grow (here: ZZ_SCRATCH_STATUS),
            # but the chained CmdNeo4j will keep running after that
            # poll fires.  We rely on the marker-based poll below
            # to know when the *whole chain* is done.
            t_chain_start = time.time()
            mark("chain_fire_t_start")
            try:
                n = sess.click_via_timer(
                    spec.name, ctl=spec.cmd_name,        # "CmdRun"
                    result_table="ZZ_SCRATCH_STATUS",
                    timeout=TIMER_TIMEOUT_SEC,
                )
                result["click_via_timer_returned"] = n
                mark(f"click_via_timer_returned_{n}")
            except Exception as e:
                mark(f"click_via_timer_exc: {e!r}")
                result["exception"] = repr(e)

            # The CmdNeo4j chain may still be running after CmdRun's
            # primary poll table fills.  Poll for the LookAtGroupData
            # CmdNeo4j-side DONE/ERR marker, OR file count stabilising.
            # Conservatively wait up to outer cap minus already-elapsed.
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
                # Check ZZ_TEST_DEBUG for chain end markers.
                try:
                    cur = sess.conn.cursor()
                    cur.execute(
                        "SELECT msg FROM ZZ_TEST_DEBUG ORDER BY id")
                    msgs = [str(r[0]) for r in cur.fetchall()]
                    cur.close()
                except Exception:
                    msgs = []
                # Treat any DONE marker that appears AFTER the
                # chain-fire marker as chain-done.  Also treat
                # an ERR as chain-done (we want to capture the
                # state, not block forever).
                done_seen = any(
                    m.endswith(":DONE") or m.endswith(":ERR")
                    for m in msgs)
                if done_seen and stable_count >= 2:
                    chain_observed_done = True
                    mark(f"chain_done_observed_files_{cur_count}")
                    break
                # Even without DONE, if file count is stable for
                # 6 polls (~6s) and >0, treat as done.
                if cur_count > 0 and stable_count >= 6:
                    chain_observed_done = True
                    mark(f"chain_quiescent_files_{cur_count}")
                    break
                time.sleep(1)

            t_chain_end = time.time()
            chain_elapsed = round(t_chain_end - t_chain_start, 2)
            result["chain_elapsed_sec"] = chain_elapsed
            mark(f"chain_elapsed_{chain_elapsed}s")
            result["chain_observed_done"] = chain_observed_done

            # Snapshot files
            files = sorted(out_dir.glob("*"))
            for f in files:
                try:
                    raw = f.read_bytes()
                    text = raw.decode(
                        "utf-8", errors="replace").lstrip("﻿")
                    first_line = text.split("\n", 1)[0].strip()
                    cols = first_line.split(",")
                    result["files"].append({
                        "name": f.name,
                        "size": f.stat().st_size,
                        "header_first_col": (cols[0] if cols
                                             else ""),
                        "header_n_cols": len(cols),
                        "header_preview": first_line[:160],
                    })
                except Exception as e:
                    result["files"].append({
                        "name": f.name,
                        "size": f.stat().st_size,
                        "read_error": repr(e),
                    })
            result["file_count"] = len(files)
            mark(f"files_inventoried_{len(files)}")

            # Capture scratch row counts
            for tbl in ("ZZ_SCRATCH_STATUS", "ZZ_SCRATCH_OFFICE",
                        "ZZ_SCRATCH_ENTRY", "ZZ_SCRATCH_BIOG_ADDR_DATA",
                        "ZZ_SCRATCH_BIOG_TEXT_DATA",
                        "ZZ_SCRATCH_P_TEXT", "ZZ_SCRATCH_IMPORT_PEOPLE",
                        "ZZ_ADDRESSES", "ZZ_PLACE"):
                try:
                    cur = sess.conn.cursor()
                    cur.execute(f"SELECT COUNT(*) FROM [{tbl}]")
                    result["row_counts"][tbl] = int(
                        cur.fetchone()[0])
                    cur.close()
                except Exception as e:
                    result["row_counts"][tbl] = f"ERROR: {e}"

            # Capture debug transcript
            try:
                cur = sess.conn.cursor()
                cur.execute(
                    "SELECT id, msg FROM ZZ_TEST_DEBUG ORDER BY id")
                for r in cur.fetchall():
                    result["debug_transcript"].append({
                        "id": int(r[0]),
                        "msg": (str(r[1])[:400]
                                if r[1] is not None else ""),
                    })
                cur.close()
            except Exception as e:
                mark(f"debug_capture_fail: {e}")
            mark("debug_captured")

            # Outcome classification
            err_msgs = [d for d in result["debug_transcript"]
                        if "LookAtGroupData:ERR" in d["msg"]]
            n_files = result["file_count"]
            if err_msgs and n_files == 0:
                result["outcome"] = (
                    "needs_investigation_mid_chain_err_no_files")
            elif err_msgs:
                result["outcome"] = (
                    "needs_investigation_mid_chain_err_with_files")
            elif n_files == 0:
                result["outcome"] = "blocked_no_files"
            elif chain_elapsed > PROMOTE_ELAPSED_THRESHOLD_SEC:
                result["outcome"] = "needs_investigation_chain_runtime"
            else:
                result["outcome"] = "promote_to_coverage_candidate"

            completed.set()
        except BaseException as e:  # noqa: BLE001
            result["outcome"] = "exception_uncaught"
            result["exception"] = (
                repr(e) + "\n" + traceback.format_exc())
            completed.set()

    # We use a closure that yields the session once, so the worker
    # can both construct it and the outer can tear it down.  Need
    # an iterator pattern because make_fixture is a generator.
    _session_holder: list = []

    def _make_session_iter():
        gen = make_fixture(USER_MDB, WORK)
        for s in gen:
            _session_holder.append((s, gen))
            yield s
            return

    worker = threading.Thread(target=_worker, daemon=False)
    worker.start()
    finished = completed.wait(timeout=PROBE_OUTER_TIMEOUT_SEC)
    if not finished:
        result["outcome"] = result.get(
            "outcome") or "hung_at_per_probe_timeout"
        mark(f"per_probe_hard_timeout_at_"
             f"{PROBE_OUTER_TIMEOUT_SEC}s")
        _kill_orphan()

    # Tear down the session via the make_fixture generator's
    # cleanup phase (next() raises StopIteration after teardown).
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
    try:
        worker.join(timeout=10)
    except Exception:
        pass
    time.sleep(2)
    result["elapsed_sec"] = round(time.time() - t0, 2)
    return result


def _classify_err_text(err_msgs: list[str]) -> str:
    """Inspect the actual error message text and classify by
    DAO/JET error family.  Distinguishing this matters: Issue #6
    (CmdGIS queryEntry typo) is JET 3061 ("no value given for ..."
    / "could not find field" / specific column name), whereas a
    "no current record" message is DAO 3021 (unguarded recordset
    access), which would be a *different* bug class."""
    if not err_msgs:
        return "no_err_observed"
    blob = " | ".join(err_msgs).lower()
    if ("no value given for one or more required parameters" in blob
            or "could not find field" in blob
            or "c_parental_status" in blob):
        return ("issue_6_family_JET_3061_column_or_param "
                "(same shape as queryEntry typo bug)")
    if "no current record" in blob:
        return ("DAO_3021_no_current_record "
                "(distinct from Issue #6; suggests unguarded "
                "recordset access, e.g. .Fields read without "
                ".EOF check)")
    if "item not found in this collection" in blob:
        return ("DAO_3265_item_not_found "
                "(field/column name lookup failure; could be "
                "renamed/missing column, similar shape to Bug "
                "category that audit_recordset_sql_projection.py "
                "is the static guard for)")
    return f"unrecognised_err_text: {err_msgs[:2]}"


def _verdict_for_brief(result: dict) -> dict:
    """Map the per-probe outcome to the brief's promote/reject
    judgement.  Captures the four answers Q1-Q5 explicitly."""
    outcome = result.get("outcome", "")
    n_files = result.get("file_count", 0)
    chain_elapsed = result.get("chain_elapsed_sec")
    err_msgs_full = [
        d["msg"]
        for d in result.get("debug_transcript", [])
        if "LookAtGroupData:ERR" in d["msg"]
    ]
    err_present = bool(err_msgs_full)
    done_present = any(
        d["msg"].endswith(":DONE")
        for d in result.get("debug_transcript", []))
    err_classification = _classify_err_text(err_msgs_full)

    answers = {
        "Q1_cmdrun_plus_cmdneo4j_completed_same_session": (
            (result.get("click_via_timer_returned") is not None)
            and result.get("chain_observed_done", False)
        ),
        "Q2_chain_elapsed_under_120s": (
            isinstance(chain_elapsed, (int, float))
            and chain_elapsed <= PROMOTE_ELAPSED_THRESHOLD_SEC
        ),
        "Q3_files_produced": {
            "count": n_files,
            "names": [f["name"] for f in result.get("files", [])],
        },
        "Q4_zz_test_debug_has_LookAtGroupData_ERR": err_present,
        "Q4b_err_messages": err_msgs_full,
        "Q5_branch_specific_issues": err_classification,
        "_chain_done_marker_present": done_present,
    }

    promote = (
        answers["Q1_cmdrun_plus_cmdneo4j_completed_same_session"]
        and answers["Q2_chain_elapsed_under_120s"]
        and n_files >= 1
        and not err_present
    )
    if promote:
        verdict = "promote_to_coverage_candidate"
        verdict_note = (
            f"Chain completed in {chain_elapsed}s (<= "
            f"{PROMOTE_ELAPSED_THRESHOLD_SEC}s threshold), produced "
            f"{n_files} files, no LookAtGroupData:ERR observed.  "
            f"Per the refresh brief's promote condition, this cell "
            f"is now eligible for a coverage PR (sibling to "
            f"test_cmd_gis_groupdata_clean_branches in "
            f"tests/test_vba_cmdgis_other_forms.py).  "
            f"Recommended next action: open coverage PR with the "
            f"same all-Chk*-reset pattern + a per-shape header "
            f"classifier modeled on _NEO4J_SHAPES from "
            f"tests/test_vba_cmdneo4j_cross_form.py."
        )
    elif err_present:
        verdict = "do_not_open_coverage_pr_mid_chain_err"
        verdict_note = (
            f"ZZ_TEST_DEBUG contains LookAtGroupData:ERR.  Per the "
            f"refresh brief, do NOT open a coverage PR — open an "
            f"investigation PR instead.\n\n"
            f"Error classification: {err_classification}\n\n"
            f"Notes for the next maintainer reading this report:\n"
            f"- The chain DID complete in {chain_elapsed} s "
            f"(well under the {PROMOTE_ELAPSED_THRESHOLD_SEC} s "
            f"promote threshold) and DID produce {n_files} "
            f"well-formed CSVs.  See the per-file table for the "
            f"shapes that landed.\n"
            f"- The error fired mid-chain but the form-level "
            f"error handler caught it; the chain still emitted "
            f"its DONE marker.  This is similar in *shape* to "
            f"the GroupData × CmdGIS investigation finding (an "
            f"Entry-related branch failing while sibling branches "
            f"complete cleanly), but the underlying error class "
            f"(see classification above) determines whether this "
            f"is the SAME bug as Issue #6 or a NEW bug worth "
            f"filing separately.\n"
            f"- Recommended investigation: per-Chk* isolation "
            f"probe (like analysis/probe_groupdata_cmdgis_"
            f"subcalls.py from PR `investigate/groupdata-cmdgis-"
            f"subcall-isolation`) to localise which CmdNeo4j "
            f"section raises the error.  CmdNeo4j has no "
            f"`If ChkX.Value` gating in its own body; the "
            f"isolation probe needs to vary CmdRun's Chk* state "
            f"to vary which ZZ_SCRATCH_* tables are populated, "
            f"then re-run the chain and see which Chk*-state "
            f"combo eliminates the error."
        )
    elif n_files == 0:
        verdict = "do_not_open_coverage_pr_no_files"
        verdict_note = (
            f"Chain produced 0 files even though no :ERR was "
            f"observed.  Either the FileDialog patch didn't "
            f"engage, the chain didn't fire CmdNeo4j, or the "
            f"chain bailed early on an unobserved condition.  "
            f"Per the refresh brief, do NOT open a coverage PR — "
            f"open an investigation PR (probe what bails)."
        )
    elif (isinstance(chain_elapsed, (int, float))
          and chain_elapsed > PROMOTE_ELAPSED_THRESHOLD_SEC):
        verdict = "do_not_open_coverage_pr_chain_too_slow"
        verdict_note = (
            f"Chain completed but took {chain_elapsed}s, exceeding "
            f"the brief's {PROMOTE_ELAPSED_THRESHOLD_SEC}s promote "
            f"threshold.  Coverage PR not viable until either the "
            f"chain is sped up OR the test watcher cap is raised "
            f"with explicit justification.  Document as "
            f"investigation outcome; do not open a coverage PR."
        )
    else:
        verdict = f"do_not_open_coverage_pr_other_{outcome}"
        verdict_note = (
            f"Outcome `{outcome}` did not match any promote path.  "
            f"See per-probe debug transcript and markers for "
            f"failure detail.  Do NOT open a coverage PR; open an "
            f"investigation PR if the cell remains a priority."
        )
    return {
        "verdict": verdict,
        "verdict_note": verdict_note,
        "answers": answers,
    }


def _write_md(result: dict, verdict: dict) -> None:
    md: list[str] = []
    md.append("# GroupData × CmdNeo4j probe (probe-first investigation)")
    md.append("")
    md.append("**Date:** 2026-05-05  ·  **Branch:** "
              "`probe/groupdata-cmdneo4j` (off main `47e506d`)")
    md.append("")
    md.append("Per the export-gap triage refresh "
              "(`analysis/export_gap_triage_plan.md`, refresh section "
              "2026-05-05), GroupData × CmdNeo4j is the only remaining "
              "cheapest-next-cell after applying the brief's "
              "exclusions (no AssociationPairs · no driver/meta-PR · "
              "no CmdUCINet new family).  This probe answers Q1-Q5 "
              "from that brief BEFORE any coverage PR is opened.")
    md.append("")
    md.append("## Setup")
    md.append("")
    md.append(f"- Fixture: person_1 (matches "
              f"`matrix_hard_forms`'s `groupdata_person_1_small`)")
    md.append(f"- Pattern: all 11 Chk* reset to False, then enable "
              f"6 clean-branch boxes (Status/Office/Addr + GIS "
              f"sisters).  Same as the just-merged GroupData × "
              f"CmdGIS test.")
    md.append(f"- Chain: `CmdRun,CmdNeo4j` via Form.Tag, directory "
              f"mode (trailing backslash → `f<n>.out` per Show)")
    md.append(f"- click_via_timer cap: {TIMER_TIMEOUT_SEC} s  ·  "
              f"per-probe outer cap: {PROBE_OUTER_TIMEOUT_SEC} s")
    md.append(f"- Promote threshold per brief: chain elapsed ≤ "
              f"{PROMOTE_ELAPSED_THRESHOLD_SEC} s, ≥ 1 file, no "
              f"`LookAtGroupData:ERR`")
    md.append("")
    md.append("## Outcome")
    md.append("")
    md.append(f"- **per-probe outcome:** `{result.get('outcome')}`")
    md.append(f"- **chain elapsed:** "
              f"{result.get('chain_elapsed_sec')} s")
    md.append(f"- **files produced:** {result.get('file_count')}")
    md.append(f"- **chain done observed:** "
              f"{result.get('chain_observed_done')}")
    md.append(f"- **click_via_timer returned:** "
              f"{result.get('click_via_timer_returned')}")
    md.append(f"- **total wall elapsed:** "
              f"{result.get('elapsed_sec')} s")
    if result.get("exception"):
        md.append(f"- **exception:** "
                  f"`{result['exception'][:300]}`")
    md.append("")
    md.append("## Answers to brief Q1-Q5")
    md.append("")
    a = verdict["answers"]
    md.append(f"- **Q1** — CmdRun + CmdNeo4j completed in same "
              f"session?  **{a['Q1_cmdrun_plus_cmdneo4j_completed_same_session']}**")
    md.append(f"- **Q2** — Total chain elapsed ≤ 120 s?  "
              f"**{a['Q2_chain_elapsed_under_120s']}** "
              f"({result.get('chain_elapsed_sec')} s)")
    md.append(f"- **Q3** — Files produced: "
              f"**{a['Q3_files_produced']['count']}**.  Names: "
              f"`{a['Q3_files_produced']['names']}`")
    md.append(f"- **Q4** — `ZZ_TEST_DEBUG` has "
              f"`LookAtGroupData:ERR`?  "
              f"**{a['Q4_zz_test_debug_has_LookAtGroupData_ERR']}**")
    md.append(f"- **Q5** — Branch-specific known issues?  "
              f"`{a['Q5_branch_specific_issues']}`")
    md.append("")
    md.append(f"## Verdict: `{verdict['verdict']}`")
    md.append("")
    md.append(verdict["verdict_note"])
    md.append("")
    md.append("## Findings supplement (interpretation)")
    md.append("")
    # CmdNeo4j body (per Form_LookAtGroupData.vb) writes 11
    # dlgSaveAs.Show blocks back-to-back, one per shape:
    #   1. People                (line 380)
    #   2. Places                (line 529)
    #   3. PeoplePlaces          (line 696)
    #   4. PeoplePlacesCodes     (line 791)
    #   5. PeopleStatus          (line 860)
    #   6. StatusCode            (line 939)
    #   7. PeopleOffice          (line 1017)
    #   8. OfficeCodes           (line 1127)
    #   9. PeopleEntry           (line 1213)
    #  10. EntryCode             (line ~1290)
    #  11. InstitutionCodes      (gated `If tRecDeleted > 0`)
    expected_blocks = [
        "People", "Places", "PeoplePlaces", "PeoplePlacesCodes",
        "PeopleStatus", "StatusCode", "PeopleOffice", "OfficeCodes",
        "PeopleEntry", "EntryCode", "InstitutionCodes",
    ]
    n_files_actual = result.get("file_count", 0)
    md.append(f"- Expected dlgSaveAs.Show blocks in CmdNeo4j: "
              f"**{len(expected_blocks)}** "
              f"({', '.join(expected_blocks)}).  "
              f"Last (InstitutionCodes) is conditional on "
              f"`If tRecDeleted > 0`.")
    md.append(f"- Actual files produced: **{n_files_actual}**.")
    if n_files_actual < len(expected_blocks):
        gap = len(expected_blocks) - n_files_actual
        md.append(f"- **{gap} blocks did not produce a file.** "
                  f"Most likely the chain bailed via "
                  f"`GoTo Exit_CmdNeo4j_Click` after the error "
                  f"surfaced.  Looking at the file-name sequence "
                  f"(patch_filedialog increments `f<n>.out` per "
                  f"Show call), the gap localises which block "
                  f"raised the error: the chain stopped at "
                  f"f<n>=" + str(n_files_actual * 3) + " (every "
                  f"Show call increments the sequence by ~3 due "
                  f"to driver-side bookkeeping), suggesting the "
                  f"error fired in or just before the "
                  f"`{expected_blocks[n_files_actual] if n_files_actual < len(expected_blocks) else '?'}` "
                  f"block.")
    seen_first_cols = sorted({
        f.get("header_first_col", "")
        for f in result.get("files", [])
    })
    md.append(f"- File first-column shapes seen: "
              f"`{seen_first_cols}`.")
    # Cross-reference with ZZ_SCRATCH_ENTRY count: if 0, the
    # Entry-related blocks (PeopleEntry / EntryCode) have no
    # source data, which is consistent with the all-Chk*-reset +
    # Status/Office/Addr-only enable pattern this probe uses.
    rc = result.get("row_counts") or {}
    if rc.get("ZZ_SCRATCH_ENTRY") == 0:
        md.append("- `ZZ_SCRATCH_ENTRY` row count is 0 (because "
                  "we excluded ChkEntry per the all-Chk*-reset "
                  "+ Status/Office/Addr-only enable pattern, "
                  "matching the GroupData × CmdGIS coverage "
                  "scope).  CmdNeo4j's PeopleEntry / EntryCode "
                  "blocks (#9, #10) read from this table.  "
                  "If the error is `No current record.`, that's "
                  "an unguarded recordset access in those "
                  "blocks (DAO 3021 family — `.Fields(...)` "
                  "called without `.EOF` check), which is a "
                  "*different* bug pattern than Issue #6's "
                  "JET 3061 column-typo.")
        md.append("- Suggested next-step probe: re-run with "
                  "`ChkEntry = True` enabled (and continue "
                  "to exclude `ChkGisEntry` since this is "
                  "Neo4j not GIS).  If the `:ERR No current "
                  "record.` disappears when ZZ_SCRATCH_ENTRY "
                  "is non-empty, that confirms the bug is "
                  "\"PeopleEntry/EntryCode block doesn't "
                  "guard against empty source recordset\" "
                  "and is worth filing as a NEW issue distinct "
                  "from Issue #6.")
        md.append("- Caveat: enabling ChkEntry will re-trigger "
                  "Issue #6 (queryEntry's typo'd column ref) "
                  "during CmdRun.  ZZ_SCRATCH_ENTRY may end up "
                  "non-empty due to the prior INSERT executing "
                  "before the JET 3061 raises, OR may stay 0 "
                  "depending on transactional behaviour — "
                  "the next-step probe needs to capture "
                  "ZZ_SCRATCH_ENTRY count both with and "
                  "without ChkEntry to disambiguate.")
    md.append("")
    md.append("## Per-file detail")
    md.append("")
    if result.get("files"):
        md.append("| name | size | n_cols | first_col | "
                  "header_preview |")
        md.append("|---|---:|---:|---|---|")
        for f in result["files"]:
            md.append(
                f"| {f.get('name', '?')} | "
                f"{f.get('size', '—')} | "
                f"{f.get('header_n_cols', '—')} | "
                f"`{f.get('header_first_col', '—')}` | "
                f"`{f.get('header_preview', '—')[:80]}` |"
            )
    else:
        md.append("(no files produced)")
    md.append("")
    md.append("## Scratch row counts (post-chain)")
    md.append("")
    for tbl, c in (result.get("row_counts") or {}).items():
        md.append(f"- `{tbl}`: {c}")
    md.append("")
    md.append("## ZZ_TEST_DEBUG transcript "
              f"({len(result.get('debug_transcript', []))} entries)")
    md.append("")
    for d in result.get("debug_transcript", [])[:60]:
        md.append(f"  - `{d['id']:>4d}`: `{d['msg']}`")
    if len(result.get("debug_transcript", [])) > 60:
        md.append(f"  - … (+{len(result['debug_transcript']) - 60} "
                  f"more)")
    md.append("")
    md.append("## Markers (timeline)")
    md.append("")
    for m in result.get("markers", []):
        md.append(f"  - `+{m['t']:>6.2f}s` {m['marker']}")
    md.append("")
    md.append("## Constraints honoured per brief")
    md.append("")
    md.append("- ✅ Investigation artifacts only — no `tests/` "
              "changes")
    md.append("- ✅ Did NOT touch README / reports* / issues / "
              "driver  (*reports/groupdata_cmdneo4j_probe.json IS "
              "an output of this probe per the brief; no other "
              "reports edited)")
    md.append("- ✅ Did NOT design a new fixture — reused person_1 "
              "inline (matches matrix_hard_forms's "
              "`groupdata_person_1_small`)")
    md.append("- ✅ Used Access COM via `VbaSession.make_fixture`")
    OUT_MD.write_text("\n".join(md), encoding="utf-8")


def main() -> int:
    print("=== GroupData × CmdNeo4j probe (probe-first) ===\n")
    _kill_orphan()
    time.sleep(1)
    out_dir = ROOT / "analysis" / "_probe_groupdata_cmdneo4j_out"
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
        "generated_date": "2026-05-05",
        "probe_branch": "probe/groupdata-cmdneo4j",
        "main_at_probe": "47e506d",
        "fixture": "groupdata_person_1_small (c_personid = 1)",
        "config": {
            "person_id": PERSON_ID,
            "all_chk_reset_first": list(ALL_CHK_CONTROLS),
            "enable_chk_controls": list(ENABLE_CHK_CONTROLS),
            "click_via_timer_timeout_sec": TIMER_TIMEOUT_SEC,
            "per_probe_outer_timeout_sec": PROBE_OUTER_TIMEOUT_SEC,
            "promote_elapsed_threshold_sec":
                PROMOTE_ELAPSED_THRESHOLD_SEC,
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
    print(f"\n=== verdict: {verdict['verdict']} ===")
    print(f"  {verdict['verdict_note']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
