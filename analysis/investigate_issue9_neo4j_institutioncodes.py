"""Re-verify Issue #9 — LookAtEntry.CmdNeo4j InstitutionCodes branch.

Manual user observation (2026-05-04): selecting `c_entry_code = 101`
(recommendation / 薦舉) on LookAtEntry, then clicking Neo4j, produces
"Finished saving to Neo4j" with no error popup — but the
InstitutionCodes_*.csv file is NOT in the output folder.  The current
canonical Issue #9 narrative says the popup *would* fire on the typo
`With tRstAssocCodes` at Form_LookAtEntry.vb:1425.  Those two
observations need reconciling.

Re-reading the VBA precisely (Form_LookAtEntry.vb:1375-1391):

    cmdSQL.CommandText = "DELETE * FROM ZZ_SCRATCH_P_TEXT"
    cmdSQL.Execute tRecDeleted

    tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id ) " + _
                "SELECT DISTINCT ZZ_SCRATCH_ENTRY.c_personid " + _
                "FROM ZZ_SCRATCH_ENTRY " + _
                "WHERE (((ZZ_SCRATCH_ENTRY.c_inst_code)>0))"
    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecDeleted

    tQueryStr = "SELECT DISTINCT ZZ_SCRATCH_ENTRY.c_inst_code, ..."

    If tRecDeleted > 0 Then
        dlgSaveAs.InitialFileName = "InstitutionCodes_" + tCodeStr + ".csv"
        ...
        Set tRstInstitutions = CurrentDb.OpenRecordset(tQueryStr)
        ...
        With tRstAssocCodes        '   <-- the documented typo
            .MoveFirst             '   <-- DAO 3021 if this Rs is closed
                                   '       (it WAS closed: it was
                                   '       Close-d at line 1373 of the
                                   '       earlier AssocCodes block)

So the entire SaveAs prompt + the buggy `With tRstAssocCodes` block
sit inside `If tRecDeleted > 0 Then`, where `tRecDeleted` is the
row-count of the immediately-prior INSERT, which is gated by
`ZZ_SCRATCH_ENTRY.c_inst_code > 0`.

Since `ZZ_SCRATCH_ENTRY.c_inst_code` comes verbatim from
`ENTRY_DATA.c_inst_code` via the CmdQuery INSERT (lines 1645-1652),
if NO rows in ENTRY_DATA satisfy the user's filter AND have
`c_inst_code > 0`, the InstitutionCodes branch is silently skipped
and no popup fires.

This script gathers the SQL-only evidence so we can then decide
whether an Access COM probe is warranted.

Run:
    python analysis/investigate_issue9_neo4j_institutioncodes.py [--com]

Without `--com` the script runs the SQL evidence phase only.  With
`--com`, it additionally opens Access via VbaSession and runs the
two LookAtEntry CmdQuery → CmdNeo4j chains (entry code 36, 101) so
the JSON records both predicted-from-SQL and observed-from-COM
evidence.

Outputs:
    reports/issue9_neo4j_institutioncodes_reverification.json

The companion markdown summary lives at
    analysis/issue9_neo4j_institutioncodes_reverification.md
and is hand-written from this script's JSON.
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

import pyodbc


ROOT = Path(__file__).resolve().parent.parent
USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
OUT_JSON = ROOT / "reports" / "issue9_neo4j_institutioncodes_reverification.json"

# COM-phase artifacts.
WORK_MDB = ROOT / "analysis" / "_issue9_reverify_test_copy.mdb"
COM_OUT_DIR = ROOT / "analysis" / "_issue9_reverify_neo4j_out"


# Two fixtures the brief asks us to characterise.
FIXTURES = [
    {
        "label": "c_entry_code=36 (jinshi / 科舉: 進士)",
        "c_entry_code": 36,
        # Full code 36 = 92,514 ENTRY_DATA rows; ZZ_SCRATCH_ENTRY of
        # that size + the 6-block CmdNeo4j chain is too heavy for a
        # single Access COM session and dies with RPC unavailable.
        # We still try it (the user's manual fixture is also un-
        # narrowed) — but a narrowed companion fixture below gives
        # us the in-COM verdict for code 36 too.
    },
    {
        "label": "c_entry_code=101 (recommendation / 薦舉)",
        "c_entry_code": 101,
    },
    {
        "label": ("c_entry_code=36 narrowed to 1100-1110 "
                  "(supplements full-code probe to keep "
                  "ZZ_SCRATCH_ENTRY tractable)"),
        "c_entry_code": 36,
        "year_filter": (1100, 1110),
    },
]


def _open() -> pyodbc.Connection:
    return pyodbc.connect(
        "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={USER_MDB};", autocommit=True, readonly=True)


def _scalar(cur, sql: str) -> int:
    cur.execute(sql)
    row = cur.fetchone()
    return int(row[0]) if row and row[0] is not None else 0


def _entry_code_desc(cur, code: int) -> dict:
    cur.execute(
        "SELECT c_entry_desc, c_entry_desc_chn "
        f"FROM ENTRY_CODES WHERE c_entry_code = {code}")
    row = cur.fetchone()
    if row is None:
        return {"c_entry_desc": None, "c_entry_desc_chn": None}
    return {"c_entry_desc": row[0], "c_entry_desc_chn": row[1]}


def _run_com_probe(out: dict) -> None:
    """Open Access, drive LookAtEntry CmdQuery → CmdNeo4j for each
    fixture, and record the observed evidence: ZZ_TEST_DEBUG
    transcript, output file inventory, and a from-COM verdict
    (chain_finished, popup_observed, institutioncodes_present).

    Notes:
    - Output dir uses the trailing-backslash directory mode that
      patch_filedialog supports — each dlgSaveAs.Show write lands as
      f<n>.out (or f<n>.out.csv after CmdNeo4j's auto-extension).
    - We re-create a fresh out dir per fixture so the file inventory
      reflects only that fixture's chain.
    - Per-fixture isolation: open + close VbaSession around each
      fixture so the work mdb / VBA project state cannot bleed across.
    """
    sys.path.insert(0, str(ROOT / "tests"))
    import os
    os.environ["CBDB_KILL_ALL_ACCESS"] = "1"
    from cbdb_driver.vba_session import VbaSession  # noqa: E402
    from cbdb_driver.access_app import kill_orphan_access  # noqa: E402
    # Reap any leftover MSACCESS.EXE before we start so the work-mdb
    # copy step can succeed on the first fixture, and recover after
    # an inter-fixture RPC death.
    try:
        kill_orphan_access()
    except Exception as e:
        print(f"  warn: kill_orphan_access (pre-loop): {e}")

    out["q2_com_probe_status"]["ran_in_this_script"] = True
    out["q2_com_probe_status"]["reason"] = (
        "Brief explicitly asks for at least two COM fixtures; ran "
        "both even though SQL evidence is unambiguous, so the JSON "
        "carries observed-not-just-predicted evidence."
    )

    out["q2_per_fixture_observed"] = []

    COM_OUT_DIR.mkdir(parents=True, exist_ok=True)

    for fx in FIXTURES:
        code = fx["c_entry_code"]
        label = fx["label"]
        yf = fx.get("year_filter")
        # Console may be cp1252 — emit ascii-safe label.
        safe_label = label.encode("ascii", "backslashreplace").decode()
        print(f"\n[Q2] === COM probe: {safe_label} ===")

        # Fresh per-fixture output dir.  Include year-filter in the
        # dir name so the full and narrowed code-36 probes don't share
        # an output dir.
        suffix = f"_yrs_{yf[0]}_{yf[1]}" if yf else ""
        fx_dir = COM_OUT_DIR / f"entry_{code}{suffix}"
        if fx_dir.exists():
            for f in fx_dir.iterdir():
                try:
                    f.unlink()
                except OSError:
                    pass
        else:
            fx_dir.mkdir(parents=True)

        rec: dict = {
            "label": label,
            "c_entry_code": code,
            "year_filter": list(yf) if yf else None,
            "out_dir": str(fx_dir),
            "errors": [],
        }
        sess = None
        t0 = time.time()
        try:
            sess = VbaSession(USER_MDB, WORK_MDB).open()

            # Mirror the manual user fixture: pick a single entry code,
            # default year-frame (no filter), no addr/person filter.
            sess.open_form("LookAtEntry")
            sess.set_picker_codes(
                "ZZ_SCRATCH_ENTRY_CODE", [code], "c_entry_code")
            year_filter = fx.get("year_filter")
            try:
                if year_filter:
                    # FrameYears = 2 → birth/death year window
                    # (TxtFromYear .. TxtToYear).
                    lo, hi = year_filter
                    sess.set_control("LookAtEntry", "FrameYears", 2)
                    sess.set_control("LookAtEntry", "TxtFromYear", lo)
                    sess.set_control("LookAtEntry", "TxtToYear", hi)
                    rec["year_filter"] = list(year_filter)
                else:
                    # FrameYears = 1 → no year filter.
                    sess.set_control("LookAtEntry", "FrameYears", 1)
            except Exception as e:
                rec["errors"].append(f"set FrameYears: {e}")

            # Patch dialogs: every dlgSaveAs.Show short-circuits to
            # GetTestExportPath() returning a fresh f<n>.out per call.
            sess.patch_filedialog("LookAtEntry")

            # ---- Phase 1: CmdQuery alone (no chain) so we can read
            # ZZ_SCRATCH_ENTRY between the two clicks without the
            # CmdNeo4j chain fighting us for the same Form_Timer fire.
            sess.set_form_tag(
                "LookAtEntry", "CmdQuery",
                str(fx_dir) + "\\")
            n_query = sess.click_via_timer(
                "LookAtEntry", ctl="CmdQuery",
                result_table="ZZ_SCRATCH_ENTRY", timeout=240)
            rec["cmdquery_zz_scratch_entry_rows"] = int(n_query)
            print(f"  CmdQuery -> ZZ_SCRATCH_ENTRY rows = {n_query}")

            # Capture in-DB column counts that the InstitutionCodes
            # branch's gate depends on.
            cur = sess.conn.cursor()
            for col in ("c_assoc_code", "c_inst_code",
                        "c_inst_name_code"):
                cur.execute(
                    f"SELECT COUNT(*) FROM ZZ_SCRATCH_ENTRY "
                    f"WHERE {col} > 0")
                v = cur.fetchone()[0] or 0
                rec[f"zz_scratch_entry_rows_{col}_gt_0"] = int(v)
                print(f"  ZZ_SCRATCH_ENTRY {col} > 0: {v}")

            # Snapshot transcript before CmdNeo4j so we can attribute
            # markers to the right phase.
            cur.execute(
                "SELECT MAX(id) FROM ZZ_TEST_DEBUG")
            cmdquery_last_id = int((cur.fetchone()[0] or 0))

            # ---- Phase 2: CmdNeo4j alone.
            sess.set_form_tag(
                "LookAtEntry", "CmdNeo4j",
                str(fx_dir) + "\\")
            n_neo4j = sess.click_via_timer(
                "LookAtEntry", ctl="CmdNeo4j",
                # No new result_table — CmdNeo4j is an export-only
                # button.  Use ZZ_SCRATCH_ENTRY (already populated by
                # CmdQuery) so the table-watcher degrades to the DONE
                # marker.
                result_table="ZZ_SCRATCH_ENTRY", timeout=300)
            rec["cmdneo4j_returned_rows"] = int(n_neo4j)

            # Inspect ZZ_TEST_DEBUG transcript for ENTER/DONE/ERR/
            # MSGBOX markers from this chain.  Markers are
            # `<short>:ENTER`, `<short>:DONE`, `<short>:ERR <desc>`,
            # `<short>:MSGBOX` where <short> is the form name.
            cur.execute(
                "SELECT id, msg FROM ZZ_TEST_DEBUG ORDER BY id")
            transcript = [
                {"id": int(r[0]),
                 "msg": (r[1] or "")[:300]}
                for r in cur.fetchall()
            ]
            rec["zz_test_debug_transcript"] = transcript
            cmdneo4j_msgs = [
                t["msg"] for t in transcript
                if t["id"] > cmdquery_last_id
            ]

            # Marker detection — scoped to the CmdNeo4j phase only.
            cmdneo4j_done = any(":DONE" in m for m in cmdneo4j_msgs)
            cmdneo4j_enter = any(":ENTER" in m for m in cmdneo4j_msgs)
            cmdneo4j_err = any(":ERR" in m for m in cmdneo4j_msgs)
            cmdneo4j_msgbox_count = sum(
                1 for m in cmdneo4j_msgs if ":MSGBOX" in m)
            rec["cmdneo4j_enter_marker_seen"] = cmdneo4j_enter
            rec["cmdneo4j_done_marker_seen"] = cmdneo4j_done
            rec["cmdneo4j_err_marker_seen"] = cmdneo4j_err
            rec["cmdneo4j_msgbox_count"] = cmdneo4j_msgbox_count

            # File inventory.  patch_filedialog redirects every Show to
            # GetTestExportPath(), so original names like
            # "InstitutionCodes_*.csv" don't land on disk — but every
            # file's first header column tells us what shape it is.
            files = sorted(fx_dir.glob("*"))
            file_recs = []
            for f in files:
                try:
                    head = f.read_bytes()[:200].decode(
                        "utf-8", errors="replace").lstrip("﻿")
                    first_line = head.split("\n", 1)[0].strip()
                    first_col = first_line.split(",", 1)[0]
                except Exception as e:
                    first_line = f"<read failed: {e}>"
                    first_col = ""
                file_recs.append({
                    "name": f.name,
                    "size": f.stat().st_size,
                    "header_first_column": first_col,
                    "header": first_line,
                })
            rec["output_files"] = file_recs
            rec["output_file_count"] = len(files)
            print(f"  CmdNeo4j produced {len(files)} files: "
                  f"{[f.name for f in files]}")

            # Detect whether any produced file is the InstitutionCodes
            # shape (header first column = "InstitutionCode"
            # per Form_LookAtEntry.vb:1418).  This is the directly-
            # observable signal that the gated branch fired.
            inst_files = [
                fr for fr in file_recs
                if fr["header_first_column"] == "InstitutionCode"
            ]
            rec["institutioncodes_files_observed"] = inst_files
            rec["institutioncodes_branch_observed_to_fire"] = (
                len(inst_files) > 0
            )

            # Final from-COM verdict.  Treat absence of ERR + presence
            # of >= 1 produced file as "chain completed cleanly":
            # _wait_for_done can short-circuit on stale DONE markers
            # from a previous Form_Timer fire in the same session, so
            # DONE-marker absence alone is NOT evidence of a hang.
            # ERR markers are written by the form's injected error
            # handler and ARE definitive.
            if cmdneo4j_err:
                rec["com_outcome"] = (
                    "ERR marker present — popup or runtime error "
                    "reached the form's error-handler injection")
            elif len(files) >= 1:
                rec["com_outcome"] = (
                    "chain finished cleanly, no ERR marker, "
                    f"{len(files)} files produced, "
                    f"InstitutionCodes file "
                    f"{'PRESENT' if inst_files else 'ABSENT'}"
                )
            elif cmdneo4j_done:
                rec["com_outcome"] = (
                    "DONE marker present but 0 files produced — "
                    "chain ran but every block's gate failed")
            else:
                rec["com_outcome"] = (
                    "no DONE, no files — chain may have hung or "
                    "exited via a non-injected path")

            rec["elapsed_seconds"] = round(time.time() - t0, 1)

        except Exception as e:
            rec["errors"].append(f"COM probe raised: {e!r}")
            rec["com_outcome"] = (
                f"COM probe aborted before completion: {e!r}")
        finally:
            if sess is not None:
                try:
                    sess.close()
                except Exception:
                    pass
            # Always reap any orphan MSACCESS.EXE between fixtures so
            # the next fixture's work-mdb copy step doesn't trip on
            # a stale file lock left by a crashed COM session.
            try:
                kill_orphan_access()
                # Windows takes a moment to release file handles after
                # taskkill returns.  Without the wait, the next
                # make_working_copy() races and trips PermissionError.
                time.sleep(3)
            except Exception:
                pass

        out["q2_per_fixture_observed"].append(rec)

    # Reconcile predicted vs observed.  Key on (c_entry_code, year_filter)
    # so the full and narrowed code-36 probes don't collide.
    out["reconciliation"] = []
    def _key(d: dict):
        yf = d.get("year_filter")
        return (d["c_entry_code"], tuple(yf) if yf else None)
    pred = {_key(v): v for v in out.get("verdict", [])}
    for obs in out["q2_per_fixture_observed"]:
        code = obs["c_entry_code"]
        p = pred.get(_key(obs), {})
        out["reconciliation"].append({
            "c_entry_code": code,
            "year_filter": obs.get("year_filter"),
            "predicted_from_sql": p.get(
                "outcome_predicted_from_sql"),
            "observed_from_com": obs.get("com_outcome"),
            "matches": (
                p.get("outcome_predicted_from_sql") == "no popup"
                and obs.get("com_outcome", "").startswith(
                    "chain finished cleanly")
                and obs.get(
                    "institutioncodes_branch_observed_to_fire") is False
            ),
            "com_probe_completed": obs.get("errors") == [],
        })


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--com", action="store_true",
        help="Also run real Access COM probes for both fixtures.")
    args = parser.parse_args()

    if not USER_MDB.exists():
        print(f"USER_MDB not found: {USER_MDB}", file=sys.stderr)
        return 2

    conn = _open()
    cur = conn.cursor()

    out: dict = {
        "schema_version": 1,
        "purpose": (
            "Re-verify Issue #9 InstitutionCodes branch gating on the "
            "current dump.  All counts are pure-SQL pre-images of what "
            "ZZ_SCRATCH_ENTRY would contain after CmdQuery for a single "
            "c_entry_code with no other filters."
        ),
        "user_mdb": str(USER_MDB),
        "vba_branch_gating": {
            "module": "Form_LookAtEntry",
            "branch_name": "InstitutionCodes SaveAs (CmdNeo4j_Click)",
            "gating_line": 1389,
            "gating_condition_vba": "If tRecDeleted > 0 Then",
            "gating_source_sql": (
                "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id ) "
                "SELECT DISTINCT ZZ_SCRATCH_ENTRY.c_personid "
                "FROM ZZ_SCRATCH_ENTRY "
                "WHERE (((ZZ_SCRATCH_ENTRY.c_inst_code)>0))"
            ),
            "buggy_line": 1425,
            "buggy_statement_vba": "With tRstAssocCodes",
            "buggy_intent_vba": (
                "tRstInstitutions (set on line 1415) — typo: "
                "tRstAssocCodes was Close-d at line 1373 of the "
                "earlier AssocCodes block"
            ),
            "implication": (
                "If 0 rows in ZZ_SCRATCH_ENTRY have c_inst_code > 0, "
                "the entire SaveAs prompt + buggy With block are "
                "skipped silently.  No popup, no InstitutionCodes_*.csv."
            ),
        },
    }

    # ---- Q1: ENTRY_DATA-wide c_inst_code / c_inst_name_code stats ----
    entry_data_total = _scalar(cur, "SELECT COUNT(*) FROM ENTRY_DATA")
    entry_data_inst_code_pos = _scalar(
        cur,
        "SELECT COUNT(*) FROM ENTRY_DATA WHERE c_inst_code > 0",
    )
    entry_data_inst_name_code_pos = _scalar(
        cur,
        "SELECT COUNT(*) FROM ENTRY_DATA WHERE c_inst_name_code > 0",
    )

    out["q1_entry_data_global"] = {
        "total_rows": entry_data_total,
        "rows_with_c_inst_code_gt_0": entry_data_inst_code_pos,
        "rows_with_c_inst_name_code_gt_0": entry_data_inst_name_code_pos,
        "interpretation": (
            "If both pos-counts are 0, no fixture under any "
            "c_entry_code can populate ZZ_SCRATCH_ENTRY.c_inst_code, "
            "and the InstitutionCodes branch is unreachable on this "
            "dump."
        ),
    }
    print(f"[Q1] ENTRY_DATA total: {entry_data_total}")
    print(f"[Q1] ENTRY_DATA c_inst_code      > 0: "
          f"{entry_data_inst_code_pos}")
    print(f"[Q1] ENTRY_DATA c_inst_name_code > 0: "
          f"{entry_data_inst_name_code_pos}")

    # If c_inst_code > 0 ever occurs, surface the row distribution so a
    # follow-up can pick a real-data fixture that DOES enter the branch.
    if entry_data_inst_code_pos > 0:
        cur.execute(
            "SELECT TOP 20 c_entry_code, COUNT(*) AS n "
            "FROM ENTRY_DATA WHERE c_inst_code > 0 "
            "GROUP BY c_entry_code ORDER BY COUNT(*) DESC"
        )
        out["q1_entry_codes_with_inst_code"] = [
            {"c_entry_code": int(r[0]) if r[0] is not None else None,
             "n_rows": int(r[1])}
            for r in cur.fetchall()
        ]
    else:
        out["q1_entry_codes_with_inst_code"] = []

    # ---- Q1b: per-fixture pre-image counts ----
    out["q1_per_fixture_preimage"] = []
    for fx in FIXTURES:
        code = fx["c_entry_code"]
        year_filter = fx.get("year_filter")
        desc = _entry_code_desc(cur, code)
        # Year-window predicate, when applicable, lifted onto BIOG_MAIN
        # the same way Form_LookAtEntry's CmdQuery applies FrameYears=2:
        #   tStrYears = "((BIOG_MAIN.c_index_year)>=lo AND <=hi)"
        if year_filter:
            lo, hi = year_filter
            yrs_join = " INNER JOIN BIOG_MAIN ON ENTRY_DATA.c_personid = BIOG_MAIN.c_personid "
            yrs_where = (f" AND BIOG_MAIN.c_index_year >= {lo} "
                         f"AND BIOG_MAIN.c_index_year <= {hi}")
        else:
            yrs_join = ""
            yrs_where = ""
        n_total = _scalar(
            cur,
            f"SELECT COUNT(*) FROM ENTRY_DATA{yrs_join} "
            f"WHERE ENTRY_DATA.c_entry_code = {code}{yrs_where}",
        )
        n_assoc_pos = _scalar(
            cur,
            f"SELECT COUNT(*) FROM ENTRY_DATA{yrs_join} "
            f"WHERE ENTRY_DATA.c_entry_code = {code} "
            f"AND ENTRY_DATA.c_assoc_code > 0{yrs_where}",
        )
        n_inst_pos = _scalar(
            cur,
            f"SELECT COUNT(*) FROM ENTRY_DATA{yrs_join} "
            f"WHERE ENTRY_DATA.c_entry_code = {code} "
            f"AND ENTRY_DATA.c_inst_code > 0{yrs_where}",
        )
        n_inst_name_pos = _scalar(
            cur,
            f"SELECT COUNT(*) FROM ENTRY_DATA{yrs_join} "
            f"WHERE ENTRY_DATA.c_entry_code = {code} "
            f"AND ENTRY_DATA.c_inst_name_code > 0{yrs_where}",
        )
        # CmdQuery's INSERT INTO ZZ_SCRATCH_ENTRY copies these columns
        # verbatim from ENTRY_DATA (Form_LookAtEntry.vb:1645-1652), so
        # the ENTRY_DATA-side count IS the pre-image of the
        # ZZ_SCRATCH_ENTRY-side count when the only filter is the entry
        # code (and optional year window).  Note the BIOG_MAIN inner
        # join multiplies by name matches; we capture the join
        # multiplier so the JSON tells the full story.
        n_join_biog = _scalar(
            cur,
            "SELECT COUNT(*) FROM ENTRY_DATA "
            "INNER JOIN BIOG_MAIN ON ENTRY_DATA.c_personid = "
            "BIOG_MAIN.c_personid "
            f"WHERE ENTRY_DATA.c_entry_code = {code} "
            f"AND ENTRY_DATA.c_inst_code > 0{yrs_where}",
        )

        rec = {
            "label": fx["label"],
            "c_entry_code": code,
            "year_filter": list(year_filter) if year_filter else None,
            "entry_code_desc": desc,
            "rows_in_ENTRY_DATA_for_code": n_total,
            "rows_with_c_assoc_code_gt_0": n_assoc_pos,
            "rows_with_c_inst_code_gt_0": n_inst_pos,
            "rows_with_c_inst_name_code_gt_0": n_inst_name_pos,
            "rows_after_BIOG_MAIN_join_with_c_inst_code_gt_0": n_join_biog,
            "would_enter_InstitutionCodes_branch": n_join_biog > 0,
        }
        out["q1_per_fixture_preimage"].append(rec)
        tag = f"code={code}" + (f" yrs={year_filter}" if year_filter else "")
        print(f"[Q1b] {tag}  total={n_total}  "
              f"assoc>0={n_assoc_pos}  inst>0={n_inst_pos}  "
              f"inst_name>0={n_inst_name_pos}  "
              f"after-join={n_join_biog}  "
              f"branch_entered={'YES' if n_join_biog > 0 else 'NO (gated out)'}")

    # ---- Per-fixture verdict roll-up ----
    out["verdict"] = []
    for rec in out["q1_per_fixture_preimage"]:
        if rec["rows_after_BIOG_MAIN_join_with_c_inst_code_gt_0"] == 0:
            out["verdict"].append({
                "c_entry_code": rec["c_entry_code"],
                "year_filter": rec.get("year_filter"),
                "outcome_predicted_from_sql": "no popup",
                "reason": (
                    "0 rows would land in ZZ_SCRATCH_ENTRY with "
                    "c_inst_code > 0, so the gating "
                    "INSERT...WHERE c_inst_code > 0 returns 0 rows, "
                    "tRecDeleted = 0, the `If tRecDeleted > 0 Then` "
                    "branch is skipped, no SaveAs is shown, no "
                    "InstitutionCodes_*.csv is written, and the buggy "
                    "`With tRstAssocCodes` is never executed."
                ),
                "expected_files_missing": ["InstitutionCodes_*.csv"],
                "classification_on_current_dump": (
                    "LATENT source-level typo (gated out by data)"
                ),
            })
        else:
            out["verdict"].append({
                "c_entry_code": rec["c_entry_code"],
                "year_filter": rec.get("year_filter"),
                "outcome_predicted_from_sql": (
                    "branch ENTERED — popup expected on `.MoveFirst` "
                    "of closed tRstAssocCodes (DAO 3021)"
                ),
                "reason": (
                    "Gating INSERT yields > 0 rows, SaveAs prompt "
                    "fires, then `With tRstAssocCodes; .MoveFirst` "
                    "executes against a Recordset that was Close-d "
                    "in the earlier AssocCodes block."
                ),
                "expected_files_missing": [
                    "InstitutionCodes_*.csv (write loop never reached "
                    "after MoveFirst raises)"
                ],
                "classification_on_current_dump": (
                    "ACTIVE bug (popup reproducible)"
                ),
            })

    # ---- Roadmap for Q2 (Access COM verification) -------------------
    # Q2 (real Access COM probes) is recorded here as the next step,
    # but is run by a separate authorised script.  The brief says
    # "用真實 Access COM 跑至少兩個 fixture" — the COM probe must
    # confirm: (i) no popup for any fixture predicted "no popup", and
    # (ii) "Finished saving to Neo4j" appears, and (iii) the produced
    # file list does NOT contain InstitutionCodes_*.csv.
    out["q2_com_probe_status"] = {
        "ran_in_this_script": False,
        "reason": (
            "Pure-SQL evidence is unambiguous on the current dump "
            "(see q1_entry_data_global): if 0 rows have c_inst_code "
            "> 0, the branch cannot be entered.  COM probe should be "
            "added in a follow-up if (and only if) any fixture is "
            "predicted to enter the branch."
        ),
        "fixtures_to_probe_if_branch_predicted_entered": [
            v["c_entry_code"] for v in out["verdict"]
            if v["outcome_predicted_from_sql"] != "no popup"
        ],
    }

    if args.com:
        try:
            _run_com_probe(out)
        except Exception as e:
            out["q2_com_probe_status"]["ran_in_this_script"] = "errored"
            out["q2_com_probe_status"]["error"] = repr(e)
            print(f"\n[Q2] COM probe errored: {e!r}", file=sys.stderr)

    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(json.dumps(out, indent=2, ensure_ascii=False),
                        encoding="utf-8")
    print(f"\nwrote {OUT_JSON}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
