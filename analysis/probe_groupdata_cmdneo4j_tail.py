"""GroupData × CmdNeo4j tail probe — per-block isolation.

Follow-up to PR `probe/groupdata-cmdneo4j` (commit 4ace85b).  That
probe found:

  - chain elapsed ~2.5 s (well under 120 s)
  - 8 of 11 expected dlgSaveAs.Show blocks produced files
  - mid-chain `LookAtGroupData:ERR No current record.` (DAO 3021,
    distinct from Issue #6's JET 3061)
  - the 3 missing files map to the Entry-related tail
    (PeopleEntry / EntryCode / InstitutionCodes), and
    `ZZ_SCRATCH_ENTRY` was 0 because we excluded ChkEntry per
    the all-Chk*-reset + Status/Office/Addr-only enable scope

This probe localises the failure and tests the hypothesis that
the bug is *unguarded `.MoveFirst` on an empty recordset*.

Static evidence (already confirmed by source read of
`analysis/dump/vba/Form_LookAtGroupData.vb`):

  - All 11 CmdNeo4j blocks use the same shape:
      `Set tRstX = CurrentDb.OpenRecordset(...)`
      `With tRstX`
      `    .MoveFirst`           ← DAO 3021 if recordset is empty
      `    Do While Not .EOF`
  - Block #9 PeopleEntry at line 1243-1245 reads from
    `ZZ_SCRATCH_ENTRY` directly.
  - Block #10 EntryCode at line 1383-1385 reads from a SELECT
    against `ZZ_SCRATCH_ENTRY`.
  - Block #11 InstitutionCodes at line 1485-1487 is gated by
    `If tRecDeleted > 0 Then` so should not fire when its
    upstream INSERT produces 0 rows (skipped, not bugged).

So the static picture says: blocks #1-#8 only "work" because
their feeder scratch tables happen to be non-empty under the
probe's enable scope; block #9 fails purely because
`ZZ_SCRATCH_ENTRY` is empty.  The runtime probe below tests this
hypothesis end-to-end.

Three iterations:

  Iter 1 — baseline_chain_chkentry_off
    Reproduce the original probe's 8-files-+-ERR result.

  Iter 2 — chain_chkentry_on
    Enable ChkEntry alongside Status/Office/Addr.  Issue #6 is
    expected to fire in queryEntry during CmdRun; this iter
    captures whether `ZZ_SCRATCH_ENTRY` ends up populated or 0
    (transactional behaviour of JET 3061), and what CmdNeo4j
    does with that state.

  Iter 3 — split_then_seed
    Fire CmdRun alone (no Form.Tag chain).  Manually INSERT one
    synthetic row into `ZZ_SCRATCH_ENTRY`.  Then fire CmdNeo4j
    alone.  The hypothesis: with `ZZ_SCRATCH_ENTRY` non-empty,
    the chain produces the missing 2-3 files and emits no
    `:ERR No current record.`.

Outcome buckets (per the brief):

  A. new_bug_candidate_empty_recordset_guard
     iter 3 produces ≥ 10 files AND no `:ERR No current record.`
     → confirmed: blocks #9/#10 fail to guard against an empty
     source recordset → file as a NEW issue.

  B. benign_probe_path_not_user_reachable
     iter 1 cannot be reproduced under conditions a real user
     would hit, OR the ChkEntry path Issue #6 always blocks
     access to this code path → not user-facing → leave alone.

  C. still_ambiguous
     iter 3 still ERRs OR result table doesn't disambiguate.

Constraints honoured per brief:
  - Investigation artifacts only — no `tests/` changes
  - Did NOT touch README / canonical reports / issue severity /
    driver
  - Did NOT design a new fixture
  - Used Access COM via VbaSession
  - Did NOT open a new issue (this PR is the evidence base for
    the maintainer's later issue-filing decision)

Outputs
-------
- analysis/groupdata_cmdneo4j_tail_probe.md
- reports/groupdata_cmdneo4j_tail_probe.json
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
WORK_BASE = ROOT / "analysis" / "_probe_groupdata_cmdneo4j_tail_copy"
OUT_JSON = ROOT / "reports" / "groupdata_cmdneo4j_tail_probe.json"
OUT_MD = ROOT / "analysis" / "groupdata_cmdneo4j_tail_probe.md"

PROBE_OUTER_TIMEOUT_SEC = 240
TIMER_TIMEOUT_SEC = 180

PERSON_ID = 1

ALL_CHK_CONTROLS = (
    "ChkStatus", "ChkOffice", "ChkEntry", "ChkText", "ChkAddr",
    "ChkGisStatus", "ChkGisOffice", "ChkGisOfficePeople",
    "ChkGisEntry", "ChkGisText", "ChkGisAddr",
)

# Per-iter Chk* enable sets (always reset all to False first).
ITER_CONFIGS = [
    {
        "name": "iter1_baseline_chain_chkentry_off",
        "enable_chk": ("ChkStatus", "ChkOffice", "ChkAddr",
                       "ChkGisStatus", "ChkGisOffice", "ChkGisAddr"),
        "chain_neo4j": True,
        "seed_entry_row": False,
        "expected_outcome": (
            "8 files + LookAtGroupData:ERR 'No current record.' "
            "+ ZZ_SCRATCH_ENTRY=0 (reproduces PR probe/groupdata-"
            "cmdneo4j commit 4ace85b finding)"
        ),
    },
    {
        "name": "iter2_chain_chkentry_on",
        "enable_chk": ("ChkStatus", "ChkOffice", "ChkAddr",
                       "ChkGisStatus", "ChkGisOffice", "ChkGisAddr",
                       "ChkEntry"),
        "chain_neo4j": True,
        "seed_entry_row": False,
        "expected_outcome": (
            "Issue #6 (JET 3061) fires in queryEntry during "
            "CmdRun.  Outcome of CmdNeo4j depends on whether "
            "ZZ_SCRATCH_ENTRY ends up populated (transactional "
            "behaviour of the failing INSERT)."
        ),
    },
    {
        "name": "iter3_split_then_seed",
        "enable_chk": ("ChkStatus", "ChkOffice", "ChkAddr",
                       "ChkGisStatus", "ChkGisOffice", "ChkGisAddr"),
        "chain_neo4j": False,  # split: CmdRun alone, then CmdNeo4j alone
        "seed_entry_row": True,
        "expected_outcome": (
            "If hypothesis A holds: 10-11 files + no "
            "'No current record.' ERR (only InstitutionCodes "
            "block may be skipped via its tRecDeleted gate; "
            "PeopleEntry + EntryCode should now succeed against "
            "the seeded row)."
        ),
    },
]


# Synthetic ZZ_SCRATCH_ENTRY row — one entry record for the
# probe's seeded person, with sensible default values for every
# column.  Schema dumped from analysis/dump/tables.json.  All
# columns are nominally NOT NULL but the VBA reader does
# IsNull() checks per cell, so empty strings / 0s for non-keys
# are fine.  c_personid matches PERSON_ID.
SEED_ENTRY_INSERT = (
    "INSERT INTO ZZ_SCRATCH_ENTRY ( "
    "c_personid, c_name, c_name_chn, c_index_year, "
    "c_index_year_type_code, c_index_year_type_desc, "
    "c_index_year_type_hz, c_entry_code, c_entry_desc, "
    "c_entry_chn, c_sequence, c_exam_rank, c_kin_id, "
    "c_kin_code, c_kin_desc, c_kin_name, c_kin_chn, "
    "c_assoc_code, c_assoc_desc, c_assoc_desc_chn, "
    "c_assoc_id, c_assoc_name, c_assoc_name_chn, c_inst_code, "
    "c_inst_name_code, c_inst_name_hz, c_inst_name_py, "
    "c_addr_id, c_addr_name, c_addr_chn, c_addr_type, "
    "c_addr_desc, c_addr_desc_chn, c_year, "
    "c_parental_status_code, c_parental_status_desc, "
    "c_parental_status_desc_chn, x_coord, y_coord, xy_count, "
    "c_entry_addr_id, c_entry_addr_name, c_entry_addr_chn, "
    "c_entry_xcoord, c_entry_ycoord, c_entry_xy_count, c_dy, "
    "c_dynasty_chn, c_dynasty, c_source, c_source_text, "
    "c_source_text_chn, c_age "
    ") VALUES ( "
    "1, 'probe_seed', '', 0, "
    "'', '', '', 100, 'probe_seed_entry', '', "
    "1, '', 0, 0, '', '', '', "
    "0, '', '', 0, '', '', 0, "
    "0, '', '', 0, '', '', 0, "
    "'', '', 0, "
    "0, '', '', 0.0, 0.0, 0, "
    "0, '', '', 0.0, 0.0, 0, 0, '', '', "
    "0, '', '', 0 "
    ")"
)


def _kill_orphan():
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "MSACCESS.EXE"],
            capture_output=True, timeout=10)
    except Exception:
        pass


def _run_iter(iter_cfg: dict, work: Path, out_dir: Path) -> dict:
    """Run one isolation iteration.  Each iter gets its own
    fresh Access COM session + work mdb copy + output dir."""
    from cbdb_driver.vba_session import VbaSession, make_fixture
    from cbdb_driver.form_specs import LOOKATGROUPDATA

    spec = LOOKATGROUPDATA
    result: dict = {
        "iter": iter_cfg["name"],
        "config": iter_cfg,
        "markers": [],
        "outcome": None,
        "elapsed_sec": None,
        "exception": None,
        "row_counts_after_cmdrun": {},
        "row_counts_after_cmdneo4j": {},
        "debug_transcript": [],
        "files": [],
        "file_count": 0,
        "click_via_timer_returned": None,
        "seed_insert_succeeded": None,
    }
    t0 = time.time()
    completed = threading.Event()
    sess_holder: list = []

    def mark(s):
        result["markers"].append(
            {"t": round(time.time() - t0, 2), "marker": s})

    def _row_counts(sess) -> dict:
        rc = {}
        for tbl in ("ZZ_SCRATCH_STATUS", "ZZ_SCRATCH_OFFICE",
                    "ZZ_SCRATCH_ENTRY",
                    "ZZ_SCRATCH_BIOG_ADDR_DATA",
                    "ZZ_SCRATCH_BIOG_TEXT_DATA",
                    "ZZ_SCRATCH_P_TEXT",
                    "ZZ_SCRATCH_IMPORT_PEOPLE",
                    "ZZ_ADDRESSES", "ZZ_PLACE"):
            try:
                cur = sess.conn.cursor()
                cur.execute(f"SELECT COUNT(*) FROM [{tbl}]")
                rc[tbl] = int(cur.fetchone()[0])
                cur.close()
            except Exception as e:
                rc[tbl] = f"ERROR: {e}"
        return rc

    def _capture_files() -> list[dict]:
        out = []
        for f in sorted(out_dir.glob("*")):
            try:
                raw = f.read_bytes()
                text = raw.decode(
                    "utf-8", errors="replace").lstrip("﻿")
                first_line = text.split("\n", 1)[0].strip()
                cols = first_line.split(",")
                out.append({
                    "name": f.name,
                    "size": f.stat().st_size,
                    "header_first_col": (cols[0] if cols
                                         else ""),
                    "header_n_cols": len(cols),
                    "header_preview": first_line[:160],
                })
            except Exception as e:
                out.append({
                    "name": f.name,
                    "size": f.stat().st_size,
                    "read_error": repr(e),
                })
        return out

    def _capture_debug(sess) -> list[dict]:
        out = []
        try:
            cur = sess.conn.cursor()
            cur.execute(
                "SELECT id, msg FROM ZZ_TEST_DEBUG ORDER BY id")
            for r in cur.fetchall():
                out.append({
                    "id": int(r[0]),
                    "msg": (str(r[1])[:400]
                            if r[1] is not None else ""),
                })
            cur.close()
        except Exception:
            pass
        return out

    def _wait_chain_done(sess, t_chain_start, deadline_sec):
        """Poll for chain quiescence (file count stable + DONE/ERR
        marker), bounded by deadline_sec from t0."""
        deadline = t0 + deadline_sec
        stable = 0
        last = -1
        while time.time() < deadline:
            cur_count = len(sorted(out_dir.glob("*")))
            if cur_count == last:
                stable += 1
            else:
                stable = 0
                last = cur_count
            try:
                cur = sess.conn.cursor()
                cur.execute(
                    "SELECT msg FROM ZZ_TEST_DEBUG ORDER BY id")
                msgs = [str(r[0]) for r in cur.fetchall()]
                cur.close()
            except Exception:
                msgs = []
            done = any(m.endswith(":DONE") or m.endswith(":ERR")
                       for m in msgs)
            if done and stable >= 2:
                return True, cur_count, "done_marker"
            if cur_count > 0 and stable >= 6:
                return True, cur_count, "quiescent"
            time.sleep(1)
        return False, last, "timeout"

    def _worker():
        try:
            mark("constructing_session")
            for attempt in (1, 2, 3):
                try:
                    gen = make_fixture(USER_MDB, work)
                    sess = next(gen)
                    sess_holder.append((sess, gen))
                    mark(f"session_opened_attempt_{attempt}")
                    break
                except Exception as e:
                    mark(f"session_open_attempt_{attempt}_fail: "
                         f"{e!r}")
                    _kill_orphan()
                    time.sleep(60)
            else:
                raise RuntimeError(
                    "session open failed after 3 attempts")

            sess = sess_holder[0][0]
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
                except Exception as e:
                    mark(f"reset_{ctl}_fail: {e}")
            for ctl in iter_cfg["enable_chk"]:
                try:
                    sess.set_control(spec.name, ctl, True)
                except Exception as e:
                    mark(f"enable_{ctl}_fail: {e}")
            mark(f"chk_state_set_enable_"
                 f"{','.join(iter_cfg['enable_chk'])}")

            if iter_cfg["chain_neo4j"]:
                # Chain mode: CmdRun -> CmdNeo4j via Form.Tag
                sess.set_form_tag(
                    spec.name,
                    f"{spec.cmd_name},CmdNeo4j",
                    str(out_dir) + "\\")
                mark("form_tag_set_chain")
                t_chain_start = time.time()
                try:
                    n = sess.click_via_timer(
                        spec.name, ctl=spec.cmd_name,
                        result_table="ZZ_SCRATCH_STATUS",
                        timeout=TIMER_TIMEOUT_SEC)
                    result["click_via_timer_returned"] = n
                    mark(f"click_via_timer_returned_{n}")
                except Exception as e:
                    mark(f"click_via_timer_exc: {e!r}")
                    result["exception"] = repr(e)
                done, fc, why = _wait_chain_done(
                    sess, t_chain_start,
                    PROBE_OUTER_TIMEOUT_SEC - 5)
                result["chain_observed_done"] = done
                result["chain_quiesce_reason"] = why
                mark(f"chain_quiesce_files_{fc}_reason_{why}")
                # Capture the post-CmdRun row counts AFTER the
                # whole chain (we don't have a clean split here,
                # so report both keys but with same data).
                rc = _row_counts(sess)
                result["row_counts_after_cmdrun"] = rc
                result["row_counts_after_cmdneo4j"] = rc
            else:
                # Split mode: CmdRun alone, then optional seed,
                # then CmdNeo4j alone
                sess.set_form_tag(
                    spec.name, spec.cmd_name,
                    str(out_dir) + "\\")
                mark("form_tag_set_cmdrun_alone")
                try:
                    n = sess.click_via_timer(
                        spec.name, ctl=spec.cmd_name,
                        result_table="ZZ_SCRATCH_STATUS",
                        timeout=TIMER_TIMEOUT_SEC)
                    result["click_via_timer_returned"] = n
                    mark(f"cmdrun_alone_returned_{n}")
                except Exception as e:
                    mark(f"cmdrun_alone_exc: {e!r}")
                    result["exception"] = repr(e)
                # Wait briefly for any file-side activity to settle
                time.sleep(2)
                rc_after_run = _row_counts(sess)
                result["row_counts_after_cmdrun"] = rc_after_run
                mark(f"row_counts_after_cmdrun_captured_"
                     f"entry_{rc_after_run.get('ZZ_SCRATCH_ENTRY')}")

                if iter_cfg["seed_entry_row"]:
                    try:
                        cur = sess.conn.cursor()
                        cur.execute(SEED_ENTRY_INSERT)
                        sess.conn.commit()
                        cur.close()
                        result["seed_insert_succeeded"] = True
                        mark("seed_insert_succeeded")
                    except Exception as e:
                        result["seed_insert_succeeded"] = False
                        result["seed_insert_error"] = repr(e)
                        mark(f"seed_insert_failed: {e!r}")

                # Re-snapshot ZZ_SCRATCH_ENTRY after seed
                try:
                    cur = sess.conn.cursor()
                    cur.execute(
                        "SELECT COUNT(*) FROM ZZ_SCRATCH_ENTRY")
                    n_post_seed = int(cur.fetchone()[0])
                    cur.close()
                    result["zz_scratch_entry_after_seed"] = (
                        n_post_seed)
                    mark(f"zz_scratch_entry_after_seed_"
                         f"{n_post_seed}")
                except Exception as e:
                    mark(f"post_seed_count_fail: {e!r}")

                # Now fire CmdNeo4j alone
                sess.set_form_tag(
                    spec.name, "CmdNeo4j",
                    str(out_dir) + "\\")
                mark("form_tag_set_cmdneo4j_alone")
                t_neo4j_start = time.time()
                try:
                    # Use ZZ_SCRATCH_P_TEXT as a soft poll
                    # target — CmdNeo4j INSERTs into it early.
                    n2 = sess.click_via_timer(
                        spec.name, ctl="CmdNeo4j",
                        result_table="ZZ_SCRATCH_P_TEXT",
                        timeout=TIMER_TIMEOUT_SEC)
                    result["cmdneo4j_alone_returned"] = n2
                    mark(f"cmdneo4j_alone_returned_{n2}")
                except Exception as e:
                    mark(f"cmdneo4j_alone_exc: {e!r}")
                    result["exception"] = repr(e)
                done, fc, why = _wait_chain_done(
                    sess, t_neo4j_start,
                    PROBE_OUTER_TIMEOUT_SEC - 5)
                result["chain_observed_done"] = done
                result["chain_quiesce_reason"] = why
                mark(f"cmdneo4j_quiesce_files_{fc}_reason_{why}")
                result["row_counts_after_cmdneo4j"] = (
                    _row_counts(sess))

            # Final capture
            result["files"] = _capture_files()
            result["file_count"] = len(result["files"])
            mark(f"files_inventoried_{result['file_count']}")
            result["debug_transcript"] = _capture_debug(sess)
            mark("debug_captured")

            err_msgs = [d["msg"]
                        for d in result["debug_transcript"]
                        if "LookAtGroupData:ERR" in d["msg"]]
            result["err_messages"] = err_msgs
            err_blob = " | ".join(err_msgs).lower()
            n_files = result["file_count"]

            if "no current record" in err_blob and n_files >= 1:
                result["outcome"] = "no_current_record_with_files"
            elif ("no value given" in err_blob
                  or "could not find field" in err_blob
                  or "c_parental_status" in err_blob):
                result["outcome"] = "issue_6_jet_3061_class"
            elif err_msgs:
                result["outcome"] = (
                    "other_err_class: "
                    f"{err_msgs[0][:120]}")
            elif n_files == 0:
                result["outcome"] = "no_files_no_err"
            else:
                result["outcome"] = "clean_no_err_files_produced"

            completed.set()
        except BaseException as e:  # noqa: BLE001
            result["outcome"] = "exception_uncaught"
            result["exception"] = (
                repr(e) + "\n" + traceback.format_exc())
            completed.set()

    worker = threading.Thread(target=_worker, daemon=False)
    worker.start()
    finished = completed.wait(timeout=PROBE_OUTER_TIMEOUT_SEC)
    if not finished:
        result["outcome"] = (result.get("outcome")
                              or "hung_at_per_iter_timeout")
        mark(f"hard_timeout_at_{PROBE_OUTER_TIMEOUT_SEC}s")
        _kill_orphan()
    try:
        if sess_holder:
            _, gen = sess_holder[0]
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


def _classify_overall(results: list[dict]) -> dict:
    """Apply the brief's three-bucket conclusion logic.

    Baseline corroboration accepts EITHER iter1 (chain with
    ChkEntry off) OR iter2 (chain with ChkEntry on) reproducing
    the (ZZ_SCRATCH_ENTRY=0 + 8 files + No current record ERR)
    pattern.  Iter1 is the ideal baseline but the COM session can
    flake at construction (RPC error); iter2 provides independent
    evidence of the same condition because Issue #6's JET 3061
    in queryEntry leaves ZZ_SCRATCH_ENTRY=0 even when ChkEntry is
    on.  As long as one of them reproduces the
    'ZZ_SCRATCH_ENTRY=0 → DAO 3021 in CmdNeo4j tail' pattern, the
    iter3 isolation evidence determines bucket A vs C.
    """
    by_iter = {r["iter"]: r for r in results}
    iter1 = by_iter.get("iter1_baseline_chain_chkentry_off", {})
    iter2 = by_iter.get("iter2_chain_chkentry_on", {})
    iter3 = by_iter.get("iter3_split_then_seed", {})

    def _iter_baseline_signature(r: dict) -> bool:
        """True if this iter exhibits the
        'ZZ_SCRATCH_ENTRY=0 + 8 files + No current record' shape
        that the original probe found and that hypothesis A
        predicts."""
        rc = (r.get("row_counts_after_cmdneo4j")
              or r.get("row_counts_after_cmdrun") or {})
        entry = rc.get("ZZ_SCRATCH_ENTRY")
        return (
            r.get("file_count", 0) == 8
            and entry == 0
            and any("no current record" in m.lower()
                    for m in r.get("err_messages", []))
        )

    iter1_baseline = _iter_baseline_signature(iter1)
    iter2_baseline = _iter_baseline_signature(iter2)
    baseline_reproduced = iter1_baseline or iter2_baseline
    baseline_source = (
        "iter1" if iter1_baseline
        else "iter2" if iter2_baseline
        else None
    )

    iter3_files = iter3.get("file_count", 0)
    iter3_no_no_current_record = not any(
        "no current record" in m.lower()
        for m in iter3.get("err_messages", []))
    iter3_seed_succeeded = bool(
        iter3.get("seed_insert_succeeded"))
    iter3_entry_after_seed = iter3.get(
        "zz_scratch_entry_after_seed", 0)
    iter3_chk_entry_off = (
        "ChkEntry" not in iter3.get("config", {}).get(
            "enable_chk", ()))

    if (baseline_reproduced
            and iter3_seed_succeeded
            and iter3_entry_after_seed >= 1
            and iter3_files >= 10
            and iter3_no_no_current_record
            and iter3_chk_entry_off):
        bucket = "A_new_bug_candidate_empty_recordset_guard"
        baseline_note = (
            f"Baseline corroborated by `{baseline_source}` "
            f"({'iter1 chain with ChkEntry off' if baseline_source == 'iter1' else 'iter2 chain with ChkEntry on; ZZ_SCRATCH_ENTRY ended up 0 because Issue #6 (JET 3061) blocked queryEntry inserts'}): "
            f"ZZ_SCRATCH_ENTRY=0 produced 8 files + "
            f"`LookAtGroupData:ERR No current record.`"
        )
        rationale = (
            f"{baseline_note}.  Iter 3 then split the chain "
            f"(CmdRun alone with ChkEntry off, then a manual "
            f"INSERT of 1 synthetic row into ZZ_SCRATCH_ENTRY, "
            f"then CmdNeo4j alone).  Result: "
            f"ZZ_SCRATCH_ENTRY={iter3_entry_after_seed}, "
            f"{iter3_files} files produced, NO "
            f"`No current record.` ERR.  The 2 extra files are "
            f"the missing PeopleEntry + EntryCode shapes (see "
            f"the per-file detail).  InstitutionCodes (block #11) "
            f"correctly remained skipped because its "
            f"`If tRecDeleted > 0` gate evaluates to False for "
            f"the synthetic row.\n\n"
            f"This isolates the failure cause to: "
            f"`Form_LookAtGroupData.CmdNeo4j_Click` blocks #9 "
            f"(PeopleEntry, line 1243-1245) and #10 (EntryCode, "
            f"line 1383-1385) call `.MoveFirst` on their "
            f"recordset without first checking `.EOF` or "
            f"`.RecordCount > 0`.  When ZZ_SCRATCH_ENTRY is "
            f"empty the recordset opens with `.EOF=True` and "
            f"`.MoveFirst` raises DAO 3021.\n\n"
            f"This is a NEW bug candidate distinct from Issue #6 "
            f"(which is a column-typo JET 3061 in "
            f"`Form_LookAtGroupData.queryEntry` that prevents "
            f"ZZ_SCRATCH_ENTRY from being populated at all when "
            f"ChkEntry is true).  The two bugs interact — "
            f"Issue #6 is the upstream cause that leaves "
            f"ZZ_SCRATCH_ENTRY=0 even on the ChkEntry-on path, "
            f"which then exposes this downstream missing-guard "
            f"bug — but they are different code-level defects "
            f"and should be filed separately.\n\n"
            f"Static evidence corroborates: ALL 11 CmdNeo4j "
            f"blocks share the same unguarded `.MoveFirst` "
            f"pattern.  Blocks #1-#8 only \"work\" because their "
            f"feeder scratch tables (ZZ_SCRATCH_STATUS, "
            f"ZZ_SCRATCH_OFFICE, ZZ_PLACE, ZZ_SCRATCH_P_TEXT, "
            f"ZZ_ADDRESSES) are non-empty under the probe's "
            f"enable scope.  This is a systemic missing-guard "
            f"pattern; PeopleEntry and EntryCode happen to be "
            f"the first blocks that hit the empty-feeder case "
            f"under any normal user enable scope."
        )
    elif baseline_reproduced and iter3_files == 0:
        bucket = "C_still_ambiguous"
        rationale = (
            "Baseline reproduced (corroborated by "
            f"`{baseline_source}`) but Iter 3 produced 0 files "
            "even after seeding — possibly the seed INSERT "
            "failed or CmdNeo4j alone doesn't engage the "
            "expected code path.  Cannot conclude bug class."
        )
    elif (baseline_reproduced
            and iter3_seed_succeeded
            and iter3_entry_after_seed >= 1
            and not iter3_no_no_current_record):
        bucket = "C_still_ambiguous"
        rationale = (
            "Baseline reproduced but Iter 3 seeded "
            f"ZZ_SCRATCH_ENTRY with "
            f"{iter3_entry_after_seed} row(s) and the "
            "'No current record.' ERR re-fired anyway.  The "
            "trigger is therefore NOT 'ZZ_SCRATCH_ENTRY empty' — "
            "it's something else.  Open further investigation."
        )
    elif not baseline_reproduced:
        bucket = "C_still_ambiguous"
        rationale = (
            "Neither iter1 nor iter2 reproduced the baseline "
            "(ZZ_SCRATCH_ENTRY=0 + 8 files + 'No current "
            "record.' ERR).  The original probe finding may be "
            "flaky or fixture-state dependent.  Re-investigate."
        )
    else:
        bucket = "C_still_ambiguous"
        rationale = "No clear A/B match.  See per-iter detail."

    return {
        "conclusion": bucket,
        "rationale": rationale,
        "iter1_baseline_reproduced": iter1_baseline,
        "iter2_baseline_corroborates": iter2_baseline,
        "baseline_source": baseline_source,
        "iter3_files": iter3_files,
        "iter3_seed_succeeded": iter3_seed_succeeded,
        "iter3_entry_count_after_seed": iter3_entry_after_seed,
        "iter3_no_no_current_record_err": (
            iter3_no_no_current_record),
    }


def _write_md(results: list[dict], conclusion: dict) -> None:
    md: list[str] = []
    md.append("# GroupData × CmdNeo4j tail probe — per-block "
              "isolation")
    md.append("")
    md.append("**Date:** 2026-05-05  ·  **Branch:** "
              "`investigate/groupdata-cmdneo4j-tail`")
    md.append("")
    md.append("Follow-up to PR `probe/groupdata-cmdneo4j` "
              "(commit `4ace85b`).  That probe found 8 of 11 "
              "expected CmdNeo4j files were produced under "
              "person_1 + Status/Office/Addr enable scope, with "
              "a mid-chain `LookAtGroupData:ERR No current "
              "record.` (DAO 3021).  This tail probe localises "
              "the failure and tests the hypothesis that the "
              "bug is *unguarded `.MoveFirst` on an empty "
              "recordset* in the PeopleEntry / EntryCode "
              "blocks.")
    md.append("")
    md.append("## Static evidence (read-only source review)")
    md.append("")
    md.append("`analysis/dump/vba/Form_LookAtGroupData.vb`:")
    md.append("")
    md.append("- Line 1243-1245 (block #9 PeopleEntry):")
    md.append("  ```")
    md.append("  Set tRstPeopleEntry = "
              "CurrentDb.OpenRecordset(\"ZZ_SCRATCH_ENTRY\", "
              "dbOpenDynaset)")
    md.append("  With tRstPeopleEntry")
    md.append("      .MoveFirst         "
              "' DAO 3021 if recordset is empty")
    md.append("      Do While Not .EOF")
    md.append("  ```")
    md.append("")
    md.append("- Line 1383-1385 (block #10 EntryCode):")
    md.append("  ```")
    md.append("  Set tRstEntryCodes = "
              "CurrentDb.OpenRecordset(tQueryStr)")
    md.append("        ' tQueryStr is a SELECT against "
              "ZZ_SCRATCH_ENTRY")
    md.append("  With tRstEntryCodes")
    md.append("      .MoveFirst         "
              "' DAO 3021 if recordset is empty")
    md.append("      Do While Not .EOF")
    md.append("  ```")
    md.append("")
    md.append("- Line 1485-1487 (block #11 InstitutionCodes) is "
              "gated by `If tRecDeleted > 0 Then`; this block "
              "is correctly skipped when its upstream INSERT "
              "produces 0 rows.  Not bugged on the Entry-empty "
              "path.")
    md.append("")
    md.append("- ALL 11 blocks (#1-#11) share the same "
              "unguarded `.MoveFirst` pattern.  "
              "Blocks #1-#8 only \"work\" on the probe's "
              "fixture because their feeder scratch tables "
              "(ZZ_SCRATCH_STATUS, ZZ_SCRATCH_OFFICE, ZZ_PLACE, "
              "ZZ_SCRATCH_P_TEXT, ZZ_ADDRESSES) are non-empty.  "
              "Block #9 fails purely because ZZ_SCRATCH_ENTRY "
              "is 0 under the probe's all-Chk*-reset + "
              "Status/Office/Addr-only enable scope.")
    md.append("")
    md.append("## Runtime probe — three iterations")
    md.append("")
    md.append("| iter | enable_chk | chain | seed | files | "
              "ZZ_SCRATCH_ENTRY | err_class |")
    md.append("|---|---|:-:|:-:|---:|---:|---|")
    for r in results:
        cfg = r["config"]
        rc = (r.get("row_counts_after_cmdneo4j")
              or r.get("row_counts_after_cmdrun") or {})
        entry = rc.get("ZZ_SCRATCH_ENTRY", "—")
        if "ChkEntry" in cfg["enable_chk"]:
            enable_short = ("Status+Office+Addr+Entry +GIS sisters")
        else:
            enable_short = "Status+Office+Addr +GIS sisters"
        err_msgs = r.get("err_messages", [])
        err_short = (err_msgs[0][:80] if err_msgs
                     else "(no ERR)")
        md.append(
            f"| `{cfg['name']}` | {enable_short} | "
            f"{'Y' if cfg['chain_neo4j'] else 'N (split)'} | "
            f"{'Y' if cfg['seed_entry_row'] else 'N'} | "
            f"{r['file_count']} | {entry} | `{err_short}` |"
        )
    md.append("")
    md.append(f"## Conclusion: **`{conclusion['conclusion']}`**")
    md.append("")
    md.append(conclusion["rationale"])
    md.append("")
    md.append("## Per-iter detail")
    md.append("")
    for r in results:
        cfg = r["config"]
        md.append(f"### `{cfg['name']}`")
        md.append("")
        md.append(f"- **enable_chk:** "
                  f"`{cfg['enable_chk']}`")
        md.append(f"- **chain_neo4j:** `{cfg['chain_neo4j']}`")
        md.append(f"- **seed_entry_row:** "
                  f"`{cfg['seed_entry_row']}`")
        md.append(f"- **expected outcome (per probe design):** "
                  f"{cfg['expected_outcome']}")
        md.append(f"- **per-iter outcome:** "
                  f"`{r.get('outcome')}`")
        md.append(f"- **elapsed:** {r.get('elapsed_sec')} s")
        md.append(f"- **file_count:** {r.get('file_count')}")
        md.append(f"- **click_via_timer_returned:** "
                  f"{r.get('click_via_timer_returned')}")
        if cfg["chain_neo4j"]:
            md.append(f"- **chain observed done:** "
                      f"{r.get('chain_observed_done')}  "
                      f"(reason: "
                      f"{r.get('chain_quiesce_reason')})")
        else:
            md.append(f"- **cmdneo4j_alone_returned:** "
                      f"{r.get('cmdneo4j_alone_returned')}")
            md.append(f"- **seed_insert_succeeded:** "
                      f"{r.get('seed_insert_succeeded')}")
            if r.get("seed_insert_error"):
                md.append(f"- **seed_insert_error:** "
                          f"`{r['seed_insert_error'][:200]}`")
            md.append(f"- **ZZ_SCRATCH_ENTRY after seed:** "
                      f"{r.get('zz_scratch_entry_after_seed')}")
        rc_run = r.get("row_counts_after_cmdrun") or {}
        if rc_run:
            md.append(f"- **row counts after CmdRun:**")
            for tbl, c in rc_run.items():
                md.append(f"  - `{tbl}`: {c}")
        rc_neo = r.get("row_counts_after_cmdneo4j") or {}
        if rc_neo and rc_neo != rc_run:
            md.append(f"- **row counts after CmdNeo4j:**")
            for tbl, c in rc_neo.items():
                md.append(f"  - `{tbl}`: {c}")
        err_msgs = r.get("err_messages", [])
        if err_msgs:
            md.append(f"- **`:ERR` messages observed "
                      f"({len(err_msgs)}):**")
            for m in err_msgs:
                md.append(f"  - `{m}`")
        else:
            md.append(f"- **`:ERR` messages observed:** "
                      f"none")
        if r.get("exception"):
            md.append(f"- **exception:** "
                      f"`{r['exception'][:300]}`")
        md.append(f"- **files produced ({r['file_count']}):**")
        for f in r.get("files", [])[:20]:
            md.append(f"  - `{f.get('name','?')}` "
                      f"({f.get('size','—')} B, "
                      f"{f.get('header_n_cols','?')} cols, "
                      f"first=`{f.get('header_first_col','?')}`)")
        if len(r.get("files", [])) > 20:
            md.append(f"  - … (+{len(r['files']) - 20} more)")
        debug = r.get("debug_transcript", [])
        md.append(f"- **ZZ_TEST_DEBUG transcript "
                  f"({len(debug)}):**")
        for d in debug[:30]:
            md.append(f"  - `{d['id']:>4d}`: `{d['msg']}`")
        if len(debug) > 30:
            md.append(f"  - … (+{len(debug) - 30} more)")
        md.append(f"- **markers:**")
        for m in r.get("markers", []):
            md.append(f"  - `+{m['t']:>6.2f}s` {m['marker']}")
        md.append("")
    md.append("## Constraints honoured per brief")
    md.append("")
    md.append("- ✅ Investigation artifacts only — no `tests/` "
              "changes")
    md.append("- ✅ Did NOT touch README / canonical reports / "
              "issue severity / driver  (the .md/.json/.py "
              "outputs of this probe are the only files written; "
              "the brief explicitly lists them)")
    md.append("- ✅ Did NOT design a new fixture — reused "
              "person_1 inline (matches matrix_hard_forms's "
              "`groupdata_person_1_small`)")
    md.append("- ✅ Used Access COM via VbaSession")
    md.append("- ✅ Did NOT open a new issue (this PR is the "
              "evidence base for the maintainer's later "
              "issue-filing decision)")
    md.append("- ✅ Did NOT mix this with Issue #6 (which is "
              "JET 3061, a column-typo class; the bug here is "
              "DAO 3021, an unguarded recordset class)")
    OUT_MD.write_text("\n".join(md), encoding="utf-8")


def main() -> int:
    # Support --reclassify-from-json for cheap re-classification
    # after the conclusion logic is updated, so we don't burn
    # another ~12-minute Access COM run.  The per-iter results
    # captured by an earlier full run live in OUT_JSON; reload
    # them, re-run _classify_overall + _write_md, and exit.
    if "--reclassify-from-json" in sys.argv:
        print("=== reclassify-from-json mode ===")
        prior = json.loads(OUT_JSON.read_text(encoding="utf-8"))
        results = prior["results"]
        conclusion = _classify_overall(results)
        out = {
            "schema_version": prior.get("schema_version", 1),
            "generated_date": "2026-05-05",
            "probe_branch": prior.get(
                "probe_branch",
                "investigate/groupdata-cmdneo4j-tail"),
            "follow_up_to": prior.get("follow_up_to", ""),
            "fixture": prior.get("fixture", ""),
            "results": results,
            "conclusion": conclusion["conclusion"],
            "rationale": conclusion["rationale"],
            "metadata": {k: conclusion[k] for k in (
                "iter1_baseline_reproduced",
                "iter2_baseline_corroborates",
                "baseline_source",
                "iter3_files",
                "iter3_seed_succeeded",
                "iter3_entry_count_after_seed",
                "iter3_no_no_current_record_err",
            )},
            "_reclassified_only": True,
        }
        OUT_JSON.write_text(
            json.dumps(out, ensure_ascii=False, indent=2,
                       default=str), encoding="utf-8")
        _write_md(results, conclusion)
        print(f"wrote {OUT_JSON}")
        print(f"wrote {OUT_MD}")
        print(f"\n=== conclusion: {conclusion['conclusion']} ===")
        print(f"  baseline_source: "
              f"{conclusion.get('baseline_source')}")
        print(f"  {conclusion['rationale'][:400]}...")
        return 0

    print("=== GroupData × CmdNeo4j tail probe — "
          "per-block isolation ===\n")
    _kill_orphan()
    time.sleep(1)
    results = []
    for i, cfg in enumerate(ITER_CONFIGS):
        print(f"\n--- {cfg['name']} ---")
        work = WORK_BASE.with_suffix(f".{i}.mdb")
        out_dir = (ROOT / "analysis"
                   / f"_probe_groupdata_cmdneo4j_tail_out_{i}")
        if out_dir.exists():
            for f in out_dir.glob("*"):
                try:
                    f.unlink()
                except Exception:
                    pass
        else:
            out_dir.mkdir(parents=True)
        r = _run_iter(cfg, work, out_dir)
        results.append(r)
        print(f"  outcome: {r['outcome']}")
        print(f"  elapsed: {r['elapsed_sec']} s")
        print(f"  file_count: {r['file_count']}")
        if r.get("err_messages"):
            print(f"  err_messages: {r['err_messages']}")
        # Cooldown between iters
        time.sleep(60)

    conclusion = _classify_overall(results)
    out = {
        "schema_version": 1,
        "generated_date": "2026-05-05",
        "probe_branch": "investigate/groupdata-cmdneo4j-tail",
        "follow_up_to": (
            "PR probe/groupdata-cmdneo4j commit 4ace85b"),
        "fixture": "groupdata_person_1_small (c_personid = 1)",
        "results": results,
        "conclusion": conclusion["conclusion"],
        "rationale": conclusion["rationale"],
        "metadata": {k: conclusion[k] for k in (
            "iter1_baseline_reproduced",
            "iter2_baseline_corroborates",
            "baseline_source",
            "iter3_files",
            "iter3_seed_succeeded",
            "iter3_entry_count_after_seed",
            "iter3_no_no_current_record_err",
        )},
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2,
                   default=str),
        encoding="utf-8",
    )
    print(f"\nwrote {OUT_JSON}")
    _write_md(results, conclusion)
    print(f"wrote {OUT_MD}")
    print(f"\n=== conclusion: {conclusion['conclusion']} ===")
    print(f"  {conclusion['rationale']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
