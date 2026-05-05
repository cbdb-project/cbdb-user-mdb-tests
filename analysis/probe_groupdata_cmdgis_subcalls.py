"""GroupData × CmdGIS sub-call ERR localisation probe.

Follow-up to PR investigate/groupdata-cmdgis-probe (45dca39's
sibling, branch investigate/groupdata-cmdgis-probe), which found
that GroupData's chain reaches DONE and produces files BUT emits
one non-fatal ERR mid-chain:

    LookAtGroupData:ERR No value given for one or more
    required parameters.

The previous probe couldn't tell WHICH sub raised the ERR
(form-level error handler catches all sub-call errs and logs
one entry).  This probe isolates each sub by toggling individual
Chk* checkboxes per iteration:

  CmdRun's per-checkbox dispatches (Form_LookAtGroupData.vb:1589-1607):
    Call queryStatus     gated by ChkStatus
    Call queryOffice     gated by ChkOffice
    Call queryEntry      gated by ChkEntry
    Call queryText       gated by ChkText
    Call queryAddr       gated by ChkAddr

  CmdGIS's per-checkbox dispatches (Form_LookAtGroupData.vb:114-135):
    Call WriteGIS_Status         gated by ChkGisStatus
    Call WriteGIS_OfficeOffice   gated by ChkGisOffice
    Call WriteGIS_OfficePeople   gated by ChkGisOfficePeople
    Call WriteGIS_Entry          gated by ChkGisEntry
    Call WriteGIS_Text           gated by ChkGisText
    Call WriteGIS_Addr           gated by ChkGisAddr

Per-iteration plan: reset all Chk* to False, set just the target
Chk*, run CmdRun → CmdGIS chain, snapshot ZZ_TEST_DEBUG delta +
scratch-table counts + produced files.

Run:
    python analysis/probe_groupdata_cmdgis_subcalls.py

Outputs:
    reports/groupdata_cmdgis_subcall_trace.json

The companion markdown summary lives at
    analysis/groupdata_cmdgis_subcall_trace.md
and is hand-written from this script's JSON.
"""
from __future__ import annotations

import json
import os
import sys
import time
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK_MDB = ROOT / "analysis" / "_groupdata_subcall_probe_test_copy.mdb"
OUT_DIR = ROOT / "analysis" / "_groupdata_subcall_probe_out"
OUT_JSON = ROOT / "reports" / "groupdata_cmdgis_subcall_trace.json"

PERSON_ID = 1   # matrix_hard_forms's groupdata_person_1_small

# All Chk* controls relevant to the chain (CmdRun side + CmdGIS side).
# Need to reset every one to False each iteration.
ALL_CHK_CONTROLS: tuple[str, ...] = (
    # CmdRun side
    "ChkStatus", "ChkOffice", "ChkEntry", "ChkText", "ChkAddr",
    # CmdGIS side
    "ChkGisStatus", "ChkGisOffice", "ChkGisOfficePeople",
    "ChkGisEntry", "ChkGisText", "ChkGisAddr",
)

# Per-iteration plan.  `chk_set` lists the ONLY checkboxes that
# should be True for that iteration (everything else False).
# `tests_sub` names the sub-call(s) we expect this iteration to
# exercise.  `result_table` is the scratch table to count after.
ITERATIONS: tuple[dict, ...] = (
    # ---- CmdRun's 5 query* sub-calls in isolation (no GIS) ----
    {"name": "queryStatus_alone",
     "chk_set": ("ChkStatus",),
     "tests_subs": ["queryStatus"],
     "result_tables": ["ZZ_SCRATCH_STATUS"]},
    {"name": "queryOffice_alone",
     "chk_set": ("ChkOffice",),
     "tests_subs": ["queryOffice"],
     "result_tables": ["ZZ_SCRATCH_OFFICE"]},
    {"name": "queryEntry_alone",
     "chk_set": ("ChkEntry",),
     "tests_subs": ["queryEntry"],
     "result_tables": ["ZZ_SCRATCH_ENTRY"]},
    {"name": "queryText_alone",
     "chk_set": ("ChkText",),
     "tests_subs": ["queryText"],
     "result_tables": ["ZZ_SCRATCH_TEXT"]},
    {"name": "queryAddr_alone",
     "chk_set": ("ChkAddr",),
     "tests_subs": ["queryAddr"],
     "result_tables": ["ZZ_SCRATCH_ADDR"]},
    # ---- queryX + WriteGIS_X pairs ----
    {"name": "Status_full_chain",
     "chk_set": ("ChkStatus", "ChkGisStatus"),
     "tests_subs": ["queryStatus", "WriteGIS_Status"],
     "result_tables": ["ZZ_SCRATCH_STATUS"]},
    {"name": "Office_OfficeOffice",
     "chk_set": ("ChkOffice", "ChkGisOffice"),
     "tests_subs": ["queryOffice", "WriteGIS_OfficeOffice"],
     "result_tables": ["ZZ_SCRATCH_OFFICE"]},
    {"name": "Office_OfficePeople",
     "chk_set": ("ChkOffice", "ChkGisOfficePeople"),
     "tests_subs": ["queryOffice", "WriteGIS_OfficePeople"],
     "result_tables": ["ZZ_SCRATCH_OFFICE"]},
    {"name": "Entry_full_chain",
     "chk_set": ("ChkEntry", "ChkGisEntry"),
     "tests_subs": ["queryEntry", "WriteGIS_Entry"],
     "result_tables": ["ZZ_SCRATCH_ENTRY"]},
    {"name": "Text_full_chain",
     "chk_set": ("ChkText", "ChkGisText"),
     "tests_subs": ["queryText", "WriteGIS_Text"],
     "result_tables": ["ZZ_SCRATCH_TEXT"]},
    {"name": "Addr_full_chain",
     "chk_set": ("ChkAddr", "ChkGisAddr"),
     "tests_subs": ["queryAddr", "WriteGIS_Addr"],
     "result_tables": ["ZZ_SCRATCH_ADDR"]},
)


def main() -> int:
    sys.path.insert(0, str(ROOT / "tests"))
    os.environ["CBDB_KILL_ALL_ACCESS"] = "1"
    from cbdb_driver.vba_session import VbaSession  # noqa: E402
    from cbdb_driver.access_app import kill_orphan_access  # noqa: E402
    from cbdb_driver.form_specs import LOOKATGROUPDATA  # noqa: E402

    if not USER_MDB.exists():
        print(f"USER_MDB not found: {USER_MDB}", file=sys.stderr)
        return 2

    out: dict = {
        "schema_version": 1,
        "purpose": (
            "Localise the LookAtGroupData:ERR ('No value given for "
            "one or more required parameters') from the previous "
            "GroupData × CmdGIS probe to a specific sub-call."
        ),
        "fixture": {
            "form": "LookAtGroupData",
            "person_id": PERSON_ID,
        },
        "iterations": [],
        "errors": [],
    }

    # ---- Long cooldown before COM (the previous GroupData probe
    # showed COM is flaky after recent activity) ----
    kill_orphan_access(); time.sleep(120)
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    sess = None
    try:
        sess = VbaSession(USER_MDB, WORK_MDB).open()
        sess.open_form("LookAtGroupData")
        sess.set_picker_codes(
            LOOKATGROUPDATA.picker_table,
            [PERSON_ID],
            column=LOOKATGROUPDATA.picker_column)
        sess.patch_filedialog("LookAtGroupData")

        for it in ITERATIONS:
            iter_rec: dict = {
                "name": it["name"],
                "chk_set": list(it["chk_set"]),
                "tests_subs": list(it["tests_subs"]),
                "errors": [],
            }
            print(f"\n[iter] === {it['name']} ===")

            # Snapshot ZZ_TEST_DEBUG max id + ZZ_SCRATCH_*
            # cutoffs so we can compute deltas afterwards.
            cur = sess.conn.cursor()
            cur.execute("SELECT MAX(id) FROM ZZ_TEST_DEBUG")
            cutoff_id = (cur.fetchone()[0] or 0)
            iter_rec["zz_test_debug_cutoff_id_before"] = int(cutoff_id)

            # Reset all Chk* to False, then set the target subset.
            for ctl in ALL_CHK_CONTROLS:
                try:
                    sess.set_control("LookAtGroupData", ctl, False)
                except Exception as e:
                    iter_rec["errors"].append(
                        f"reset {ctl}=False: {e}")
            for ctl in it["chk_set"]:
                try:
                    sess.set_control("LookAtGroupData", ctl, True)
                except Exception as e:
                    iter_rec["errors"].append(
                        f"set {ctl}=True: {e}")

            # Re-seed picker (each CmdRun consumes
            # ZZ_SCRATCH_IMPORT_PEOPLE; re-seed each iteration so
            # the chain has fresh input).
            sess.set_picker_codes(
                LOOKATGROUPDATA.picker_table,
                [PERSON_ID],
                column=LOOKATGROUPDATA.picker_column)

            # Per-iteration output dir (directory mode so each
            # WriteGIS_X gets a distinct f<n>.out per call).
            iter_dir = OUT_DIR / it["name"]
            if iter_dir.exists():
                for f in iter_dir.iterdir():
                    try: f.unlink()
                    except OSError: pass
            else:
                iter_dir.mkdir(parents=True)

            sess.set_form_tag(
                "LookAtGroupData",
                "CmdRun,CmdGIS",
                str(iter_dir) + "\\")

            # Fire chain.  Watch the first declared result_table
            # for row-count change; fall through on timeout.
            watch_table = it["result_tables"][0]
            try:
                n = sess.click_via_timer(
                    "LookAtGroupData",
                    ctl=LOOKATGROUPDATA.cmd_name,    # "CmdRun"
                    result_table=watch_table,
                    timeout=180)
            except Exception as e:
                iter_rec["errors"].append(
                    f"click_via_timer raised: {e!r}")
                n = None
            iter_rec["click_via_timer_returned_n"] = n

            # Snapshot scratch-table counts.
            scratch_counts = {}
            for tbl in (
                "ZZ_SCRATCH_IMPORT_PEOPLE",
                "ZZ_SCRATCH_STATUS", "ZZ_SCRATCH_OFFICE",
                "ZZ_SCRATCH_ENTRY", "ZZ_SCRATCH_TEXT",
                "ZZ_SCRATCH_ADDR", "ZZ_SCRATCH_PEOPLE",
            ):
                try:
                    cur.execute(f"SELECT COUNT(*) FROM {tbl}")
                    scratch_counts[tbl] = int(cur.fetchone()[0] or 0)
                except Exception:
                    scratch_counts[tbl] = None
            iter_rec["scratch_counts_after"] = scratch_counts

            # ZZ_TEST_DEBUG delta — the per-iteration transcript.
            cur.execute(
                "SELECT id, msg FROM ZZ_TEST_DEBUG "
                "WHERE id > ? ORDER BY id",
                cutoff_id)
            new_msgs = [
                {"id": int(r[0]), "msg": (r[1] or "")[:300]}
                for r in cur.fetchall()
            ]
            iter_rec["zz_test_debug_delta"] = new_msgs
            iter_rec["enter_marker_seen"] = any(
                "LookAtGroupData:ENTER" in m["msg"] for m in new_msgs)
            iter_rec["done_marker_seen"] = any(
                "LookAtGroupData:DONE" in m["msg"] for m in new_msgs)
            err_msgs = [
                m["msg"] for m in new_msgs
                if "LookAtGroupData:ERR" in m["msg"]
            ]
            iter_rec["err_marker_seen"] = bool(err_msgs)
            iter_rec["err_marker_messages"] = err_msgs
            iter_rec["msgbox_marker_count"] = sum(
                1 for m in new_msgs
                if "LookAtGroupData:MSGBOX" in m["msg"])

            # File inventory.
            files = sorted(iter_dir.glob("*"))
            file_recs = []
            for f in files:
                try:
                    raw = f.read_bytes()
                    text = raw.decode("utf-8",
                                      errors="replace").lstrip("﻿")
                    first = text.split("\n", 1)[0].strip()
                except Exception:
                    first = "<read failed>"
                file_recs.append({
                    "name": f.name,
                    "size_bytes": f.stat().st_size,
                    "header": first[:200],
                })
            iter_rec["output_files"] = file_recs
            iter_rec["output_file_count"] = len(file_recs)

            # Console summary.
            print(f"  ENTER:{iter_rec['enter_marker_seen']} "
                  f"DONE:{iter_rec['done_marker_seen']} "
                  f"ERR:{iter_rec['err_marker_seen']} "
                  f"MSGBOX:{iter_rec['msgbox_marker_count']} "
                  f"files:{iter_rec['output_file_count']}")
            if err_msgs:
                for m in err_msgs:
                    print(f"  err: {m}")

            out["iterations"].append(iter_rec)

    except Exception as e:
        out["errors"].append(f"COM probe raised: {e!r}")
        print(f"[err] {e!r}", file=sys.stderr)
    finally:
        if sess is not None:
            try: sess.close()
            except Exception: pass

    # ---- Verdict synthesis ----------------------------------------
    # For each (sub, iteration) pair, attribute the ERR.
    err_localisation = {}
    for it_rec in out["iterations"]:
        if it_rec.get("err_marker_seen"):
            for sub in it_rec["tests_subs"]:
                err_localisation.setdefault(sub, []).append(
                    it_rec["name"])
    out["err_localisation"] = err_localisation

    # Compute "first ERR-firing iteration" for each sub.
    # If a sub's "_alone" iteration ERRs, sub is responsible.
    # If only the chain iteration ERRs (after _alone is clean),
    # the OTHER sub in the chain is responsible.
    sub_verdicts: dict[str, str] = {}
    for it_rec in out["iterations"]:
        if not it_rec.get("err_marker_seen"):
            continue
        if it_rec["name"].endswith("_alone"):
            # The single sub in this iteration is the culprit.
            sub_verdicts[it_rec["tests_subs"][0]] = (
                "ERR_in_alone_iteration")
        else:
            # Two subs ran; identify the one whose _alone was
            # clean (it must be the chain partner that ERRs).
            for sub in it_rec["tests_subs"]:
                alone_name = (sub.replace("WriteGIS_", "")
                                 .replace("query", "")
                              + "_alone")
                # Find the _alone record for this sub's underlying
                # query (only query subs have _alone iterations).
                if not sub.startswith("query"):
                    continue
                alone = next(
                    (r for r in out["iterations"]
                     if r["name"].lower().startswith(
                         sub.replace("query", "").lower()
                         + "_alone")),
                    None)
                if alone and not alone.get("err_marker_seen"):
                    # query side clean -> WriteGIS side is suspect
                    other = [s for s in it_rec["tests_subs"]
                             if s != sub]
                    for o in other:
                        sub_verdicts[o] = (
                            "ERR_in_chain_iteration_clean_alone")
    out["per_sub_verdicts"] = sub_verdicts

    # Final classification.  Three buckets per the brief:
    #   A. benign_probe_gap
    #   B. real_cbdb_bug_candidate
    #   C. still_ambiguous_but_localized
    n_iters = len(out["iterations"])
    n_err = sum(1 for r in out["iterations"]
                if r.get("err_marker_seen"))
    if n_err == 0:
        verdict = "benign_probe_gap"
        reason = (
            "Re-running the chain with isolated checkboxes never "
            "reproduced the ERR — the previous probe's ERR was "
            "specific to the simultaneous default-checkbox state, "
            "not to any individual sub.  Once the probe seeds "
            "checkboxes explicitly, the chain is clean."
        )
    elif sub_verdicts and all(
            v == "ERR_in_alone_iteration" for v in sub_verdicts.values()):
        verdict = "real_cbdb_bug_candidate"
        reason = (
            f"ERR fired even when each suspect sub was isolated "
            f"(no other Chk* set).  The sub itself raises the "
            f"JET parameter error: {sorted(sub_verdicts)}.  "
            f"Same bug-class as Issues #7-9.  Investigate per "
            f"the issue-report-maintainer skill."
        )
    elif sub_verdicts:
        verdict = "real_cbdb_bug_candidate"
        reason = (
            f"ERR localised to: {sub_verdicts}.  The chain-only "
            f"failures suggest a WriteGIS_X sub has a JET "
            f"parameter error in its SELECT projection — same "
            f"bug-class as Issues #7-9.  Investigate per the "
            f"issue-report-maintainer skill."
        )
    else:
        verdict = "still_ambiguous_but_localized"
        reason = (
            f"ERR fired in {n_err}/{n_iters} iterations but the "
            f"per-sub attribution heuristic couldn't single out a "
            f"culprit.  Manual transcript review needed."
        )

    out["verdict"] = verdict
    out["verdict_reason"] = reason
    print(f"\nverdict: {verdict}")
    print(f"reason:  {reason}")

    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, indent=2, ensure_ascii=False),
        encoding="utf-8")
    print(f"\nwrote {OUT_JSON}")
    return 0 if verdict == "benign_probe_gap" else 1


if __name__ == "__main__":
    raise SystemExit(main())
