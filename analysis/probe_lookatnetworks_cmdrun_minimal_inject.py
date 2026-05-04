"""LookAtNetworks CmdRun probe via minimal injection (PR AX).

Bisection arc:
  PR AR  — CmdRun probe via VbaSession default → hung at open_form
  PR AS  — DataMode bisect → not the trigger
  PR AT  — autodetect bisect → trigger is `_inject_autodetect`
  PR AU  — per-other-form bisect → ANY sibling injection is enough
  PR AV  — compile-after-inject mitigation → falsified
  PR AW  — warm-open variants → W3 (keep loaded) works; close+reopen
           still hangs.  PR AU V12 (inject ONLY Networks) opens fine.
  PR AX  — THIS — try CmdRun on the 3 PR AQ anchor candidates using
           the minimal-injection path proven by PR AU V12.

Setup per candidate:
  skip = all 9 sibling Form_LookAt* (Networks injection ON)
  sess = VbaSession(USER_MDB, WORK,
                     skip_inject_autodetect_forms=skip)
  sess.open()                                # only Networks injected
  sess.open_form("LookAtNetworks")           # ~2 s, expected to work
  sess.set_picker_codes(
      "ZZ_SCRATCH_IMPORT_PEOPLE", [pid],
      column="c_person_id")
  set minimal kin-only control state:
      TxtNodeDist=1, TxtMaxLoop=0,
      ChkKin=-1, ChkNonKin=0,
      ChkMale=-1, ChkFemale=-1
  sess.click_via_timer("LookAtNetworks", ctl="CmdRun",
                        result_table="ZZ_SOCIAL_NETWORK",
                        timeout=120)

Per candidate: fresh Access process; per-candidate hard cap.

Outputs
-------
- analysis/dump/lookatnetworks_cmdrun_minimal_inject.json
- analysis/lookatnetworks_cmdrun_minimal_inject.md
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
WORK_BASE = ROOT / "analysis" / "_probe_lan_ax_copy"
OUT_JSON = ROOT / "analysis" / "dump" / "lookatnetworks_cmdrun_minimal_inject.json"
OUT_MD = ROOT / "analysis" / "lookatnetworks_cmdrun_minimal_inject.md"

PER_CANDIDATE_TIMEOUT_SEC = 240
CMDRUN_TIMER_TIMEOUT_SEC = 120

CANDIDATES = [
    {"c_personid": 30270, "name_chn": "曹植", "name_py": "Cao Zhi",
     "assocs": 10, "kin": 1, "est_1hop": 10},
    {"c_personid":     4, "name_chn": "查道", "name_py": "Zha Dao",
     "assocs": 5,  "kin": 9, "est_1hop": 99},
    {"c_personid":  3135, "name_chn": "張君平",
     "name_py": "Zhang Junping",
     "assocs": 5,  "kin": 2, "est_1hop": 31},
]

SKIP_SIBLINGS = {
    "Form_LookAtEntry", "Form_LookAtOffice", "Form_LookAtStatus",
    "Form_LookAtTexts", "Form_LookAtAssociations",
    "Form_LookAtPlace", "Form_LookAtKinship",
    "Form_LookAtAssociationPairs", "Form_LookAtGroupData",
}

# Minimal kin-only run.  Avoids gUseFilter machinery.
MIN_CONTROL_STATE = {
    "TxtNodeDist": 1,
    "TxtMaxLoop":  0,
    "ChkKin":      -1,
    "ChkNonKin":    0,
    "ChkMale":     -1,
    "ChkFemale":   -1,
}

RESULT_TABLES = ["ZZ_SOCIAL_NETWORK", "ZZ_SCRATCH_PEOPLE",
                 "ZZ_SOCIAL_NETWORK_AGGREGATE"]


def _kill_orphan():
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "MSACCESS.EXE"],
            capture_output=True, timeout=10)
    except Exception:
        pass


def run_candidate(candidate: dict) -> dict:
    from cbdb_driver.vba_session import VbaSession

    pid = candidate["c_personid"]
    work = WORK_BASE.with_suffix(f".pid{pid}.mdb")

    result: dict = {
        "candidate": candidate,
        "markers": [],
        "outcome": None,
        "elapsed_sec": None,
        "exception": None,
        "row_counts": {},
        "debug_transcript": [],
        "controls_set": {},
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
            mark("constructing_session_minimal_inject")
            sess = VbaSession(
                USER_MDB, work,
                skip_inject_autodetect_forms=SKIP_SIBLINGS,
            )
            mark("opening_session")
            sess.open()
            mark("session_opened_only_Networks_injected")

            sess.open_form("LookAtNetworks")
            mark("form_opened")

            sess.set_picker_codes(
                "ZZ_SCRATCH_IMPORT_PEOPLE", [pid],
                column="c_person_id")
            mark("picker_seeded")

            for ctl, val in MIN_CONTROL_STATE.items():
                try:
                    sess.set_control("LookAtNetworks", ctl, val)
                    result["controls_set"][ctl] = val
                except Exception as e:
                    result["controls_set"][ctl] = f"FAIL: {e}"
                    mark(f"set_{ctl}_fail: {e}")
            mark("controls_set")

            try:
                n = sess.click_via_timer(
                    "LookAtNetworks", ctl="CmdRun",
                    result_table="ZZ_SOCIAL_NETWORK",
                    timeout=CMDRUN_TIMER_TIMEOUT_SEC,
                )
                mark(f"cmdrun_returned_{n}_rows_via_timer")
                result["row_counts"][
                    "ZZ_SOCIAL_NETWORK_via_click_via_timer"] = n
            except Exception as e:
                mark(f"click_via_timer_exc: {e!r}")
                result["exception"] = repr(e)

            for tbl in RESULT_TABLES:
                try:
                    cur = sess.conn.cursor()
                    cur.execute(f"SELECT COUNT(*) FROM [{tbl}]")
                    result["row_counts"][tbl] = int(cur.fetchone()[0])
                except Exception as e:
                    result["row_counts"][tbl] = f"ERROR: {e}"
            mark("row_counts_captured")

            try:
                cur = sess.conn.cursor()
                cur.execute(
                    "SELECT id, msg FROM ZZ_TEST_DEBUG ORDER BY id")
                for r in cur.fetchall():
                    result["debug_transcript"].append({
                        "id": int(r[0]),
                        "msg": str(r[1])[:300] if r[1] else "",
                    })
            except Exception as e:
                mark(f"debug_capture_fail: {e}")
            mark("debug_captured")

            # Determine outcome based on row counts.
            sn = result["row_counts"].get("ZZ_SOCIAL_NETWORK")
            sp = result["row_counts"].get("ZZ_SCRATCH_PEOPLE")
            saw_done = any(
                d["msg"].endswith(":DONE")
                for d in result["debug_transcript"]
            )
            if (isinstance(sn, int) and sn > 0
                    and isinstance(sp, int) and sp > 0
                    and saw_done):
                result["outcome"] = "succeeded_with_output"
            elif saw_done:
                result["outcome"] = "succeeded_but_empty_output"
            elif result["exception"]:
                result["outcome"] = "exception"
            else:
                result["outcome"] = "no_DONE_marker_likely_timeout"
            completed.set()
        except BaseException as e:  # noqa: BLE001
            result["outcome"] = "exception_uncaught"
            result["exception"] = repr(e) + "\n" + traceback.format_exc()
            completed.set()

    # NOTE: daemon=False so we can join() cleanly before tearing
    # down the session.  daemon=True caused a Python-internal
    # SystemError on RLock during thread cleanup when the parent
    # closed sess at the same time the worker was being GC'd.
    worker = threading.Thread(target=_worker, daemon=False)
    worker.start()
    finished = completed.wait(timeout=PER_CANDIDATE_TIMEOUT_SEC)
    if not finished:
        result["outcome"] = result["outcome"] or "hung_at_per_candidate_timeout"
        mark(f"per_candidate_hard_timeout_at_{PER_CANDIDATE_TIMEOUT_SEC}s")
        _kill_orphan()

    # Tear down session BEFORE waiting on worker join — sess.close()
    # kills MSACCESS, which unwedges any worker stuck in COM RPC.
    try:
        if sess is not None:
            sess.close()
    except Exception:
        pass
    _kill_orphan()
    # Now wait for the worker to actually finish so its thread
    # state is gone before we move on to the next candidate.
    try:
        worker.join(timeout=10)
    except Exception:
        pass
    time.sleep(2)
    result["elapsed_sec"] = round(time.time() - t0, 2)
    return result


def _classify_failure(r: dict) -> str | None:
    """If the candidate didn't succeed, classify the failure."""
    if r["outcome"] in ("succeeded_with_output",
                         "succeeded_but_empty_output"):
        return None
    msgs = [d["msg"].lower() for d in r["debug_transcript"]]
    # Look for autodetect ENTER (means CmdRun fired)
    saw_enter = any("networks:enter" in m for m in msgs)
    saw_msgbox = any("networks:msgbox" in m for m in msgs)
    saw_err = any("networks:err" in m for m in msgs)
    if not saw_enter:
        return "cmdrun_never_fired (no Networks:ENTER marker)"
    if saw_msgbox:
        return "blocked_by_MsgBox (gating check failed; gUsePersonID etc.)"
    if saw_err:
        err_msgs = [m for m in msgs if "networks:err" in m]
        return f"vba_error_in_cmdrun: {err_msgs[:2]}"
    return "cmdrun_fired_but_no_DONE (likely recursive expansion timeout)"


def main() -> int:
    print("=== LookAtNetworks CmdRun probe — minimal injection ===\n")
    _kill_orphan()
    time.sleep(1)
    results = []
    for c in CANDIDATES:
        print(f"--- candidate pid={c['c_personid']} "
              f"({c['name_chn']} / {c['name_py']}) ---")
        r = run_candidate(c)
        results.append(r)
        print(f"  outcome: {r['outcome']}, elapsed: {r['elapsed_sec']}s")
        for m in r["markers"]:
            print(f"    +{m['t']}s {m['marker']}")
        print(f"  row_counts: {r['row_counts']}")
        if r["debug_transcript"]:
            print(f"  debug ({len(r['debug_transcript'])} rows):")
            for d in r["debug_transcript"][:6]:
                print(f"    {d['id']}: {d['msg']!r}")
        if r.get("exception"):
            print(f"  exception: {r['exception'][:200]}")
        if r["outcome"] != "succeeded_with_output":
            cls = _classify_failure(r)
            if cls:
                print(f"  failure_classification: {cls}")
                r["failure_classification"] = cls
        print()

    succeeded = [
        r for r in results
        if r["outcome"] in ("succeeded_with_output",
                             "succeeded_but_empty_output")
    ]
    succeeded_with_output = [
        r for r in results
        if r["outcome"] == "succeeded_with_output"
    ]
    if succeeded_with_output:
        verdict = "SUCCESS_at_least_one_anchor_completed"
        verdict_note = (
            f"{len(succeeded_with_output)} of {len(CANDIDATES)} anchors "
            f"completed CmdRun with non-empty bounded output under "
            f"{CMDRUN_TIMER_TIMEOUT_SEC} s.  Recommend PR AY to add "
            f"a small-fixture LookAtNetworks test using one of "
            f"these anchors + the minimal-injection helper."
        )
    elif succeeded:
        verdict = "PARTIAL_completed_but_empty_output"
        verdict_note = (
            f"{len(succeeded)} anchor(s) completed (DONE marker "
            f"observed) but produced empty output.  Could be "
            f"control-state issue (kin-only with no kin in scope) "
            f"or anchor with no qualifying neighbours.  Investigate "
            f"per-candidate row counts."
        )
    else:
        # All failed — check if they share the same failure mode
        cls_set = {r.get("failure_classification") for r in results
                   if r.get("failure_classification")}
        verdict = "FAILED_no_anchor_completed"
        verdict_note = (
            f"No anchor completed CmdRun.  Failure classifications: "
            f"{cls_set}.  See per-candidate detail for next-step "
            f"diagnosis."
        )

    out = {
        "skip_inject_autodetect_forms": sorted(SKIP_SIBLINGS),
        "min_control_state": MIN_CONTROL_STATE,
        "per_candidate_timeout_sec": PER_CANDIDATE_TIMEOUT_SEC,
        "cmdrun_timer_timeout_sec": CMDRUN_TIMER_TIMEOUT_SEC,
        "results": results,
        "summary_by_outcome": {
            r["candidate"]["c_personid"]: r["outcome"] for r in results
        },
        "headline_verdict": verdict,
        "verdict_note": verdict_note,
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"wrote {OUT_JSON}")
    print(f"\n=== verdict: {verdict} ===")
    print(f"  {verdict_note}")

    md = []
    md.append("# LookAtNetworks CmdRun probe — minimal injection (PR AX)")
    md.append("")
    md.append("Bisection arc complete: PR AR diagnosis → PR AS-AV "
              "ruling out triggers → PR AW finding W3 (keep loaded) "
              "and PR AU V12 (minimal injection) as the two viable "
              "paths.  This PR uses the minimal-injection path on "
              "PR AQ's 3 anchor candidates.")
    md.append("")
    md.append("## Setup (constant across candidates)")
    md.append("")
    md.append("- skip_inject_autodetect_forms = all 9 sibling "
              "Form_LookAt* (Networks autodetect ON only)")
    md.append("- minimal kin-only control state:")
    for k, v in MIN_CONTROL_STATE.items():
        md.append(f"  - `{k}` = `{v}`")
    md.append("- result_table = ZZ_SOCIAL_NETWORK")
    md.append("- CmdRun timer cap = 120 s; per-candidate hard cap = 240 s")
    md.append("")
    md.append("## Outcomes")
    md.append("")
    md.append("| pid | name | outcome | elapsed | ZZ_SOCIAL_NETWORK | ZZ_SCRATCH_PEOPLE |")
    md.append("|---:|---|---|---:|---:|---:|")
    for r in results:
        c = r["candidate"]
        rc = r["row_counts"]
        md.append(
            f"| {c['c_personid']} | {c['name_chn']} ({c['name_py']}) "
            f"| `{r['outcome']}` | {r['elapsed_sec']}s "
            f"| {rc.get('ZZ_SOCIAL_NETWORK', '—')} "
            f"| {rc.get('ZZ_SCRATCH_PEOPLE', '—')} |"
        )
    md.append("")
    md.append(f"## Headline verdict: `{verdict}`")
    md.append("")
    md.append(verdict_note)
    md.append("")
    md.append("## Per-candidate detail")
    md.append("")
    for r in results:
        c = r["candidate"]
        md.append(f"### `c_personid = {c['c_personid']}` "
                  f"({c['name_chn']} / {c['name_py']})")
        md.append("")
        md.append(f"- est_1hop_assoc_total (PR AQ): {c['est_1hop']}")
        md.append(f"- assocs={c['assocs']}, kin={c['kin']}")
        md.append(f"- elapsed: {r['elapsed_sec']}s")
        md.append(f"- outcome: **`{r['outcome']}`**")
        if r.get("failure_classification"):
            md.append(f"- failure_classification: `{r['failure_classification']}`")
        md.append(f"- row counts:")
        for k, v in r["row_counts"].items():
            md.append(f"  - `{k}`: {v}")
        md.append(f"- controls_set: `{r['controls_set']}`")
        if r["debug_transcript"]:
            md.append(f"- ZZ_TEST_DEBUG ({len(r['debug_transcript'])} entries):")
            for d in r["debug_transcript"][:25]:
                md.append(f"  - {d['id']}: `{d['msg']}`")
            if len(r["debug_transcript"]) > 25:
                md.append(f"  - … (+{len(r['debug_transcript']) - 25} more)")
        if r.get("exception"):
            md.append(f"- exception: `{r['exception'][:300]}`")
        md.append(f"- markers:")
        for m in r["markers"]:
            md.append(f"  - +{m['t']}s {m['marker']}")
        md.append("")

    md.append("## Implications")
    md.append("")
    if verdict == "SUCCESS_at_least_one_anchor_completed":
        md.append("**The minimal-injection path is viable** for "
                  "Networks tests.  Recommended follow-up "
                  "(separate PR AY): add a small-fixture "
                  "`tests/test_vba_networks_small_fixture.py` "
                  "that uses one of the successful anchors and "
                  "the same minimal-injection setup, gated by "
                  "`--include-vba`.  Use the existing "
                  "`test_vba_matrix_hard_forms.py` patterns for "
                  "fixture conventions; do NOT unskip the matrix "
                  "test in the same PR.")
        md.append("")
        md.append("Operating recipe to encode in AY:")
        md.append("1. Construct VbaSession with the 9-sibling "
                  "skip set (use the same SKIP_SIBLINGS constant "
                  "as this probe).")
        md.append("2. Open LookAtNetworks (~2 s).")
        md.append("3. Seed picker (`ZZ_SCRATCH_IMPORT_PEOPLE`).")
        md.append("4. Set the 6 minimal-control values above.")
        md.append("5. `click_via_timer` for CmdRun, 120 s cap.")
        md.append("6. Assert ZZ_SOCIAL_NETWORK row count > 0 (or "
                  "another sanity bound from the per-anchor "
                  "outcomes captured here).")
    elif verdict == "PARTIAL_completed_but_empty_output":
        md.append("CmdRun completes but emits no rows.  Likely a "
                  "control-state issue (kin-only with no kin in "
                  "scope?) or an anchor that has no qualifying "
                  "1-hop neighbours under the minimal settings.  "
                  "Try adding ChkNonKin=-1 + at least one assoc "
                  "filter checkbox to broaden the walk.")
    else:
        md.append("CmdRun didn't complete on any anchor.  See "
                  "per-candidate failure_classification for the "
                  "specific blocker.  Common categories tested:")
        md.append("- `cmdrun_never_fired`: autodetect didn't write "
                  "ENTER marker.  Possible Form_Timer trigger "
                  "issue.")
        md.append("- `blocked_by_MsgBox`: a gating check failed "
                  "(gUsePersonID, ChkKin/ChkNonKin, ChkMale/"
                  "ChkFemale, gUseFilter).  Adjust control state.")
        md.append("- `vba_error_in_cmdrun`: VBA raised; check the "
                  "Networks:ERR transcript line for the message.")
        md.append("- `cmdrun_fired_but_no_DONE`: recursive "
                  "expansion timeout — try a smaller anchor.")
    md.append("")
    md.append("## Constraints respected per AX brief")
    md.append("- Probe-only.  No matrix unskips.")
    md.append("- Minimal injection (only Form_LookAtNetworks).")
    md.append("- Did NOT use W3 keep-loaded path.")
    md.append("- Per-candidate fresh Access process.")
    md.append("- 120 s CmdRun timer cap + 240 s per-candidate "
              "outer cap.")
    md.append("- Killed orphan MSACCESS.EXE between candidates.")
    md.append("- Default driver behaviour preserved (probe uses "
              "PR AT `skip_inject_autodetect_forms` kwarg only).")
    md.append("- Fast suite still 111 passed, 9 skipped (pre-probe).")
    OUT_MD.write_text("\n".join(md), encoding="utf-8")
    print(f"wrote {OUT_MD}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
