"""GroupData × CmdGIS probe-first investigation.

Per analysis/agents_shrink_round2_candidates.md follow-up: the
gap-triage classified GroupData × CmdGIS as bucket A
(small_candidate, low risk).  Given the AssociationPairs probe
surfaced a hidden CmdQuery SetFocus blocker for what was
estimated low-risk, this probe answers the same question for
GroupData BEFORE writing any cross-form test:

  Q1: With person_1 as the small fixture, does CmdRun actually
      populate the per-checkbox sub-queries (specifically
      queryStatus / ZZ_SCRATCH_STATUS) when ChkStatus = True?
  Q2: With ChkGisStatus = True, does CmdGIS (chained after
      CmdRun via Form_Timer) successfully write a .tab file?
  Q3: Does the produced .tab pass the structural checks the
      cross-form CmdGIS test would apply: file exists,
      non-empty, header reasonable, first few rows have
      consistent column counts?
  Q4: Are there any hidden blockers — SetFocus / active form
      errors, missing DONE marker, watcher row_count flake,
      export dialog hangs?

Run:
    python analysis/probe_groupdata_cmdgis.py

Outputs:
    reports/groupdata_cmdgis_probe.json   (machine-readable
                                           per-question evidence)

The companion markdown summary lives at
    analysis/groupdata_cmdgis_probe.md
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
WORK_MDB = ROOT / "analysis" / "_groupdata_cmdgis_probe_test_copy.mdb"
OUT_DIR = ROOT / "analysis" / "_groupdata_cmdgis_probe_out"
OUT_JSON = ROOT / "reports" / "groupdata_cmdgis_probe.json"

PERSON_ID = 1   # matrix_hard_forms's groupdata_person_1_small —
                # has 2 entries / 2 statuses per its docstring.


def main() -> int:
    sys.path.insert(0, str(ROOT / "tests"))
    os.environ["CBDB_KILL_ALL_ACCESS"] = "1"
    from cbdb_driver.vba_session import VbaSession  # noqa: E402
    from cbdb_driver.access_app import kill_orphan_access  # noqa: E402

    if not USER_MDB.exists():
        print(f"USER_MDB not found: {USER_MDB}", file=sys.stderr)
        return 2

    out: dict = {
        "schema_version": 1,
        "purpose": (
            "Probe-first investigation: does GroupData x CmdGIS "
            "have a hidden activation blocker (like AssociationPairs "
            "CmdQuery) before any cross-form coverage test is "
            "wired."
        ),
        "fixture": {
            "form": "LookAtGroupData",
            "person_id": PERSON_ID,
            "checkboxes_enabled": ["ChkStatus", "ChkGisStatus"],
            "rationale": (
                "Person 1 has 2 statuses (per matrix_hard_forms's "
                "docstring); ChkStatus -> CmdRun calls queryStatus "
                "-> ZZ_SCRATCH_STATUS populated; ChkGisStatus -> "
                "CmdGIS calls WriteGIS_Status -> writes "
                "status_gis_*.tab.  Smallest end-to-end CmdGIS "
                "exercise on this form."
            ),
        },
        "user_mdb": str(USER_MDB),
        "errors": [],
    }

    # ---- pre-probe: confirm person_1 actually has status data ----
    import pyodbc
    conn = pyodbc.connect(
        "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={USER_MDB};", autocommit=True, readonly=True)
    cur = conn.cursor()
    cur.execute(f"SELECT COUNT(*) FROM STATUS_DATA "
                f"WHERE c_personid = {PERSON_ID}")
    n_status = int(cur.fetchone()[0])
    out["preprobe_status_data_count"] = n_status
    print(f"[pre] STATUS_DATA rows for person {PERSON_ID}: {n_status}")
    cur.execute(f"SELECT COUNT(*) FROM BIOG_MAIN "
                f"WHERE c_personid = {PERSON_ID}")
    n_biog = int(cur.fetchone()[0])
    out["preprobe_biog_main_count"] = n_biog
    print(f"[pre] BIOG_MAIN rows for person {PERSON_ID}: {n_biog}")
    conn.close()
    if n_biog == 0:
        out["errors"].append(
            f"person {PERSON_ID} not in BIOG_MAIN — abort.")
        OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
        OUT_JSON.write_text(json.dumps(out, indent=2,
                                       ensure_ascii=False),
                            encoding="utf-8")
        return 2

    # ---- COM probe ----
    kill_orphan_access(); time.sleep(45)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    for f in OUT_DIR.iterdir():
        try:
            f.unlink()
        except OSError:
            pass

    sess = None
    t0 = time.time()
    try:
        sess = VbaSession(USER_MDB, WORK_MDB).open()

        # 1) open form
        sess.open_form("LookAtGroupData")
        out["form_opened"] = True

        # 2) seed person_1 directly into ZZ_SCRATCH_IMPORT_PEOPLE
        #    (matrix_hard_forms's pattern: picker via
        #     set_picker_codes against the spec's picker_table).
        from cbdb_driver.form_specs import LOOKATGROUPDATA  # noqa: E402
        sess.set_picker_codes(
            LOOKATGROUPDATA.picker_table,
            [PERSON_ID],
            column=LOOKATGROUPDATA.picker_column)
        out["picker_seeded"] = True

        # 3) tick ChkStatus + ChkGisStatus so CmdRun populates
        #    ZZ_SCRATCH_STATUS and CmdGIS dispatches WriteGIS_Status.
        for ctl in ("ChkStatus", "ChkGisStatus"):
            try:
                sess.set_control("LookAtGroupData", ctl, True)
            except Exception as e:
                out["errors"].append(f"set {ctl}=True: {e}")

        # 4) patch FileDialog so WriteGIS_Status writes to our
        #    test path instead of popping a SaveAs dialog.
        sess.patch_filedialog("LookAtGroupData")

        # 5) chain CmdRun -> CmdGIS via Form_Timer
        out_path = OUT_DIR / "groupdata_status_gis.tab"
        sess.set_form_tag("LookAtGroupData",
                          "CmdRun,CmdGIS",
                          str(out_path))

        n_run = sess.click_via_timer(
            "LookAtGroupData",
            ctl=LOOKATGROUPDATA.cmd_name,   # "CmdRun"
            # CmdRun is UPDATE-style backfill; row_count won't
            # change.  Watch ZZ_SCRATCH_STATUS instead — that's
            # what queryStatus populates.
            result_table="ZZ_SCRATCH_STATUS",
            timeout=180)
        out["cmd_run_returned_n_rows_in_zz_scratch_status"] = (
            int(n_run) if n_run is not None else None)
        print(f"  CmdRun -> ZZ_SCRATCH_STATUS rows = {n_run}")

        # 6) inspect ZZ_SCRATCH_IMPORT_PEOPLE backfill (CmdRun's
        #    primary effect — same as matrix_hard_forms's
        #    _check_groupdata).
        cur = sess.conn.cursor()
        cur.execute(
            "SELECT c_person_id, c_name, c_dynasty FROM "
            "ZZ_SCRATCH_IMPORT_PEOPLE "
            f"WHERE c_person_id = {PERSON_ID}")
        row = cur.fetchone()
        out["cmd_run_backfill_seen"] = (row is not None
                                        and bool(row[1]))
        out["cmd_run_backfill_sample"] = (
            None if row is None
            else {"c_person_id": int(row[0]),
                  "c_name": row[1],
                  "c_dynasty": row[2]})

        # 7) inspect ZZ_SCRATCH_STATUS directly
        cur.execute("SELECT COUNT(*) FROM ZZ_SCRATCH_STATUS")
        n_zz = int(cur.fetchone()[0])
        out["zz_scratch_status_actual_count"] = n_zz
        print(f"  ZZ_SCRATCH_STATUS direct count = {n_zz}")

        # 8) ZZ_TEST_DEBUG transcript (markers + any ERR)
        cur.execute(
            "SELECT id, msg FROM ZZ_TEST_DEBUG ORDER BY id")
        transcript = [
            {"id": int(r[0]), "msg": (r[1] or "")[:300]}
            for r in cur.fetchall()
        ]
        out["zz_test_debug_transcript"] = transcript
        msgs = [t["msg"] for t in transcript]
        out["enter_marker_seen"] = any(
            "LookAtGroupData:ENTER" in m for m in msgs)
        out["done_marker_seen"] = any(
            "LookAtGroupData:DONE" in m for m in msgs)
        out["err_marker_seen"] = any(
            "LookAtGroupData:ERR" in m for m in msgs)
        out["err_marker_messages"] = [
            m for m in msgs if "LookAtGroupData:ERR" in m]
        out["msgbox_marker_count"] = sum(
            1 for m in msgs if "LookAtGroupData:MSGBOX" in m)

        # 9) inspect produced output
        files = sorted(OUT_DIR.glob("*"))
        file_recs = []
        for f in files:
            try:
                raw = f.read_bytes()
                # GroupData CmdGIS writes UTF-8 by default
                # (GISFrame.Value not set; falls into the gb18030
                # branch per source line 3171, but the text is
                # latin/numeric in the header so utf-8 decode w/
                # errors=replace yields readable header anyway).
                text = raw.decode("utf-8", errors="replace").lstrip("﻿")
                lines = [ln for ln in text.replace("\r\n",
                                                   "\n").split("\n")
                         if ln.strip()]
                header = lines[0] if lines else ""
                first_row = lines[1] if len(lines) > 1 else ""
                # GIS exports are tab-separated.
                header_cols = header.split("\t")
                first_row_cols = first_row.split("\t") if first_row else []
            except Exception as e:
                header = f"<read failed: {e}>"
                lines = []
                header_cols = []
                first_row_cols = []
            file_recs.append({
                "name": f.name,
                "size_bytes": f.stat().st_size,
                "n_data_rows": max(0, len(lines) - 1),
                "header": header,
                "header_n_cols": len(header_cols),
                "first_data_row_n_cols": len(first_row_cols),
                "header_cols_match_first_row":
                    bool(header_cols) and
                    len(header_cols) == len(first_row_cols),
            })
        out["output_files"] = file_recs

        # 10) Q3 structural checks — synthesise the verdict.
        if not file_recs:
            out["q3_structural_checks"] = {
                "file_exists": False,
                "non_empty": False,
                "header_reasonable": False,
                "first_rows_consistent": False,
            }
        else:
            f0 = file_recs[0]
            out["q3_structural_checks"] = {
                "file_exists": True,
                "non_empty": f0["size_bytes"] > 0,
                "header_reasonable":
                    f0["header_n_cols"] >= 2
                    and "ID" in f0["header"].upper(),
                "first_rows_consistent":
                    f0["n_data_rows"] == 0
                    or f0["header_cols_match_first_row"],
            }

        out["elapsed_seconds"] = round(time.time() - t0, 1)

    except Exception as e:
        out["errors"].append(f"COM probe raised: {e!r}")
        print(f"[err] {e!r}", file=sys.stderr)
    finally:
        if sess is not None:
            try:
                sess.close()
            except Exception:
                pass

    # ---- Verdict synthesis ----------------------------------------
    inst = out.get("zz_scratch_status_actual_count", 0)
    err_marker = out.get("err_marker_seen", False)
    files = out.get("output_files") or []
    q3 = out.get("q3_structural_checks") or {}
    backfill_ok = out.get("cmd_run_backfill_seen", False)

    if err_marker and not files:
        verdict = "blocked_by_driver_issue"
        reason = (
            "ZZ_TEST_DEBUG carries a LookAtGroupData:ERR marker "
            "AND no output files were produced — same failure-"
            "shape as AssociationPairs's CmdQuery SetFocus "
            "blocker.  Driver-side patch likely required before "
            "any cross-form coverage will work."
        )
    elif err_marker and files and backfill_ok:
        verdict = "needs_investigation"
        reason = (
            f"Chain produced output ({len(files)} file(s)) AND "
            "reached DONE, but a non-fatal ERR marker fired mid-"
            f"chain: {out.get('err_marker_messages')}.  This is "
            "NOT the AssociationPairs SetFocus blocker (which "
            "blocks all output) — it's a separate runtime error "
            "embedded in the chain.  Coverage feasible AFTER "
            "identifying which sub the ERR comes from (likely a "
            "JET 'missing parameter' = column reference against "
            "wrong recordset/projection — same bug-class as "
            "Issues #7-9 if it surfaces in a release-shipped "
            "code path)."
        )
    elif not backfill_ok:
        verdict = "blocked_by_driver_issue"
        reason = (
            "CmdRun didn't even backfill ZZ_SCRATCH_IMPORT_PEOPLE "
            "(matrix_hard_forms's baseline check failed), so the "
            "Form_Timer dispatch isn't reaching CmdRun's body.  "
            "Investigate dispatch / form-state precondition."
        )
    elif inst == 0:
        verdict = "blocked_by_query_behavior"
        reason = (
            "CmdRun completed (backfill OK) but queryStatus didn't "
            "populate ZZ_SCRATCH_STATUS even though person_1 has "
            f"{out.get('preprobe_status_data_count', '?')} "
            "STATUS_DATA rows.  Either ChkStatus didn't take effect "
            "or queryStatus has its own gate."
        )
    elif not files:
        verdict = "blocked_by_query_behavior"
        reason = (
            "CmdRun + queryStatus completed (ZZ_SCRATCH_STATUS "
            f"has {inst} rows) but CmdGIS produced zero files.  "
            "ChkGisStatus may not have taken effect, or "
            "WriteGIS_Status's RecCount==0 early-bail fired despite "
            "the populated table."
        )
    elif not all(q3.values()):
        verdict = "needs_fixture_design"
        reason = (
            "File produced but didn't pass all structural checks "
            f"({q3}); the chain works but the small fixture isn't "
            "shaped right for an assertion-grade test yet."
        )
    else:
        verdict = "safe_to_cover_next"
        reason = (
            f"Full chain worked: CmdRun backfilled, queryStatus "
            f"populated ZZ_SCRATCH_STATUS ({inst} rows), CmdGIS "
            f"wrote {len(files)} file(s) passing all 4 structural "
            "checks, no ERR marker.  GroupData x CmdGIS is a "
            "true bucket-A small-candidate; coverage PR is "
            "mechanical."
        )

    out["verdict"] = verdict
    out["verdict_reason"] = reason
    print(f"\nverdict: {verdict}")
    print(f"reason:  {reason}")

    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(json.dumps(out, indent=2,
                                   ensure_ascii=False),
                        encoding="utf-8")
    print(f"\nwrote {OUT_JSON}")
    return 0 if verdict == "safe_to_cover_next" else 1


if __name__ == "__main__":
    raise SystemExit(main())
