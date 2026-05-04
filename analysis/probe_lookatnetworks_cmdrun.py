"""LookAtNetworks CmdRun focused COM probe (PR AR).

Picks 3 candidate anchors from PR AQ
(`analysis/lookatnetworks_anchor_candidates.md`):

  1. c_personid = 30270  曹植  Cao Zhi  (assocs=10, est_1hop=10)
  2. c_personid =     4  查道  Zha Dao  (assocs=5,  est_1hop=99)
  3. c_personid =  3135  張君平         (assocs=5,  est_1hop=31)

For each candidate:
  - fresh `VbaSession` (per-process Access; LinkListInit fix +
    autodetect injection handled by VbaSession.open())
  - open LookAtNetworks
  - seed `ZZ_SCRATCH_IMPORT_PEOPLE` with the candidate pid
  - set minimal-expansion control state:
      TxtNodeDist = 1, TxtMaxLoop = 0
      ChkKin = -1 (true), ChkNonKin = 0
      ChkMale = -1, ChkFemale = -1
  - fire CmdRun via Form_Timer, poll for `Networks:DONE` in
    ZZ_TEST_DEBUG with a 120 s budget
  - record:
      elapsed time
      result row counts (ZZ_SOCIAL_NETWORK, ZZ_SCRATCH_PEOPLE)
      ZZ_TEST_DEBUG transcript
      whether Access had to be killed
  - close session, kill any orphan MSACCESS, sleep 2 s, next

Outputs:
  - analysis/dump/lookatnetworks_cmdrun_probe.json
  - analysis/lookatnetworks_cmdrun_probe.md (companion note)

Requires Access COM.  Per-candidate ~30-180 s.  Safe to run
unattended IF cold-cache OpenCurrentDatabase doesn't wedge —
the script kills MSACCESS on per-candidate timeout, but the
classic PR AA "popup on cold start" risk is now addressed by
the LinkListInit patch in VbaSession.open().

Usage:
  python analysis/probe_lookatnetworks_cmdrun.py
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
WORK_MDB = ROOT / "analysis" / "_probe_lan_cmdrun_copy.mdb"
OUT_JSON = ROOT / "analysis" / "dump" / "lookatnetworks_cmdrun_probe.json"
OUT_MD = ROOT / "analysis" / "lookatnetworks_cmdrun_probe.md"

# Per-candidate hard cap.  The Form_Timer-based click_via_timer in
# vba_session uses its own 120 s timeout argument; we add a safety
# margin for OpenCurrentDatabase warmup + setup overhead.
PER_CANDIDATE_TIMEOUT_SEC = 240
CMDRUN_TIMER_TIMEOUT_SEC = 120

CANDIDATES = [
    {"c_personid": 30270, "name_chn": "曹植",   "name_py": "Cao Zhi",   "assocs": 10, "kin": 1, "est_1hop": 10},
    {"c_personid":     4, "name_chn": "查道",   "name_py": "Zha Dao",   "assocs": 5,  "kin": 9, "est_1hop": 99},
    {"c_personid":  3135, "name_chn": "張君平", "name_py": "Zhang Junping", "assocs": 5, "kin": 2, "est_1hop": 31},
]

# Minimal kin-only run: kin yes, non-kin no, both sexes,
# distance=1, no recursion.  Avoids the gUseFilter machinery.
MIN_CONTROL_STATE = {
    "TxtNodeDist": 1,
    "TxtMaxLoop":  0,
    "ChkKin":      -1,
    "ChkNonKin":    0,
    "ChkMale":     -1,
    "ChkFemale":   -1,
}


def _kill_orphan_access():
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "MSACCESS.EXE"],
            capture_output=True, timeout=10)
    except Exception:
        pass


def run_candidate(candidate: dict) -> dict:
    """Run one CmdRun probe with a hard timeout via threading."""
    from cbdb_driver.vba_session import VbaSession, make_fixture

    pid = candidate["c_personid"]
    result: dict = {
        "candidate": candidate,
        "markers": [],
        "outcome": None,
        "elapsed_sec": None,
        "exception": None,
        "row_counts": {},
        "debug_transcript": [],
    }
    t0 = time.time()

    def mark(s: str) -> None:
        result["markers"].append(
            {"t": round(time.time() - t0, 2), "marker": s})

    sess = None
    sess_iter = None
    completed = threading.Event()
    worker_exc: list[BaseException] = []

    def _worker():
        nonlocal sess, sess_iter
        try:
            mark("opening_session")
            sess_iter = make_fixture(USER_MDB, WORK_MDB)
            sess = next(sess_iter)
            mark("session_opened")

            sess.open_form("LookAtNetworks")
            mark("form_opened")

            # Seed picker.
            sess.set_picker_codes(
                "ZZ_SCRATCH_IMPORT_PEOPLE", [pid],
                column="c_person_id")
            mark("picker_seeded")

            # Apply minimal-expansion control state.
            for ctl, val in MIN_CONTROL_STATE.items():
                try:
                    sess.set_control("LookAtNetworks", ctl, val)
                except Exception as e:
                    mark(f"set_{ctl}_fail: {e}")
            mark("controls_set")

            # Fire CmdRun via Form_Timer; poll for DONE up to
            # CMDRUN_TIMER_TIMEOUT_SEC.
            try:
                n = sess.click_via_timer(
                    "LookAtNetworks", ctl="CmdRun",
                    result_table="ZZ_SOCIAL_NETWORK",
                    timeout=CMDRUN_TIMER_TIMEOUT_SEC,
                )
                mark(f"cmdrun_returned_{n}_rows")
                result["row_counts"]["ZZ_SOCIAL_NETWORK_via_click_via_timer"] = n
            except Exception as e:
                mark(f"click_via_timer_exc: {e}")
                result["exception"] = repr(e)

            # Capture additional row counts.
            for tbl in ("ZZ_SOCIAL_NETWORK", "ZZ_SCRATCH_PEOPLE",
                         "ZZ_SOCIAL_NETWORK_AGGREGATE"):
                try:
                    cur = sess.conn.cursor()
                    cur.execute(f"SELECT COUNT(*) FROM [{tbl}]")
                    result["row_counts"][tbl] = int(cur.fetchone()[0])
                except Exception as e:
                    result["row_counts"][tbl] = f"ERROR: {e}"
            mark("row_counts_captured")

            # Capture ZZ_TEST_DEBUG transcript.
            try:
                cur = sess.conn.cursor()
                cur.execute("SELECT id, msg FROM ZZ_TEST_DEBUG ORDER BY id")
                for r in cur.fetchall():
                    result["debug_transcript"].append({
                        "id": int(r[0]),
                        "msg": str(r[1])[:300] if r[1] else "",
                    })
            except Exception as e:
                mark(f"debug_capture_fail: {e}")
            mark("debug_captured")

            result["outcome"] = "succeeded"
            completed.set()
        except BaseException as e:  # noqa: BLE001
            worker_exc.append(e)
            result["outcome"] = "exception"
            result["exception"] = repr(e) + "\n" + traceback.format_exc()
            completed.set()

    worker = threading.Thread(target=_worker, daemon=True)
    worker.start()
    finished = completed.wait(timeout=PER_CANDIDATE_TIMEOUT_SEC)
    if not finished:
        result["outcome"] = "hung_at_per_candidate_timeout"
        mark(f"hard_timeout_at_{PER_CANDIDATE_TIMEOUT_SEC}s")
        _kill_orphan_access()

    # Best-effort cleanup.
    try:
        if sess_iter is not None:
            try:
                next(sess_iter)
            except StopIteration:
                pass
    except Exception:
        pass
    _kill_orphan_access()
    # Let any kill-induced cleanup settle.
    time.sleep(2)

    result["elapsed_sec"] = round(time.time() - t0, 2)
    return result


def main() -> int:
    print(f"=== LookAtNetworks CmdRun focused probe ===\n")
    _kill_orphan_access()
    time.sleep(1)
    results = []
    for c in CANDIDATES:
        print(f"--- candidate pid={c['c_personid']} ({c['name_chn']}) ---")
        r = run_candidate(c)
        results.append(r)
        print(f"  outcome: {r['outcome']}, elapsed: {r['elapsed_sec']}s")
        for m in r["markers"]:
            print(f"    +{m['t']}s {m['marker']}")
        print(f"  row_counts: {r['row_counts']}")
        if r.get("exception"):
            print(f"  exception: {r['exception'][:200]}")
        print()

    out = {
        "candidates_probed": len(results),
        "outcomes_summary": {
            r["candidate"]["c_personid"]: r["outcome"] for r in results
        },
        "per_candidate_timeout_sec": PER_CANDIDATE_TIMEOUT_SEC,
        "cmdrun_timer_timeout_sec": CMDRUN_TIMER_TIMEOUT_SEC,
        "min_control_state": MIN_CONTROL_STATE,
        "results": results,
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"wrote {OUT_JSON}")

    # ----- markdown -----
    md = []
    md.append("# LookAtNetworks CmdRun focused probe (PR AR)")
    md.append("")
    md.append("Picks 3 candidate anchors from PR AQ and attempts a "
              "minimal-expansion CmdRun on each, with per-candidate "
              "watchdog (240 s hard cap, 120 s CmdRun timer cap).")
    md.append("")
    md.append("Minimal control state:")
    for k, v in MIN_CONTROL_STATE.items():
        md.append(f"  - `{k}` = `{v}`")
    md.append("")
    md.append("## Outcomes")
    md.append("")
    md.append("| pid | name | outcome | elapsed | "
              "ZZ_SOCIAL_NETWORK rows | ZZ_SCRATCH_PEOPLE rows |")
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
    md.append("## Per-candidate detail")
    md.append("")
    for r in results:
        c = r["candidate"]
        md.append(f"### `c_personid = {c['c_personid']}` "
                  f"({c['name_chn']} / {c['name_py']})")
        md.append("")
        md.append(f"- est_1hop_assoc_total (PR AQ): {c['est_1hop']}")
        md.append(f"- elapsed: {r['elapsed_sec']}s")
        md.append(f"- outcome: **`{r['outcome']}`**")
        md.append(f"- row counts:")
        for k, v in r["row_counts"].items():
            md.append(f"  - `{k}`: {v}")
        if r["debug_transcript"]:
            md.append(f"- ZZ_TEST_DEBUG transcript "
                      f"({len(r['debug_transcript'])} entries):")
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
    succeeded = [r for r in results if r["outcome"] == "succeeded"]
    if succeeded:
        md.append(f"**{len(succeeded)} of {len(results)} candidates "
                  f"completed CmdRun under {CMDRUN_TIMER_TIMEOUT_SEC} s**.  "
                  f"Recommended next step (separate PR): use the "
                  f"smallest-elapsed candidate as the LookAtNetworks "
                  f"matrix fixture, with the same minimal control "
                  f"state.  Do NOT unskip in this PR — the focused "
                  f"probe doesn't exercise the full matrix path "
                  f"(e.g. picker UI, CmdGIS chain).")
    else:
        md.append(f"**No candidate completed** under "
                  f"{CMDRUN_TIMER_TIMEOUT_SEC} s.  The focused probe "
                  f"did NOT reproduce a working LookAtNetworks "
                  f"CmdRun.  See per-candidate `markers` field for "
                  f"the exact failure point and `exception` field "
                  f"for the COM-side error if any.  Do NOT broaden "
                  f"to driver refactor — the per-candidate Access "
                  f"COM behaviour first needs more diagnostic data "
                  f"(probably a smaller fixture from PR AQ's "
                  f"recommended list).")
    md.append("")
    md.append("## Constraints respected per AR brief")
    md.append("")
    md.append("- Per-candidate fresh Access process (via fresh "
              "`VbaSession`).")
    md.append("- 240 s per-candidate hard timeout + 120 s CmdRun "
              "timer cap.")
    md.append("- Tiny driver helper added: Form_LookAtNetworks "
              "autodetect entry now sets `gUsePersonID = (DCount "
              "ZZ_SCRATCH_IMPORT_PEOPLE > 0)` so CmdRun's gating "
              "check passes when the test seeds the picker via "
              "pyodbc.  Mirrors the existing `gUseADDRID` autodetect "
              "in Status / Texts / etc.  See "
              "`tests/cbdb_driver/vba_session.py` _AUTODETECT.")
    md.append("- No matrix test unskips.  No production fixture "
              "changes.  No reports / ISSUES touched.")
    OUT_MD.write_text("\n".join(md), encoding="utf-8")
    print(f"wrote {OUT_MD}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
