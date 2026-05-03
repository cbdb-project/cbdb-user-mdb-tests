"""Deep-dive into PR L's `same_candidates_diff_winner` × 10 rows.

PR L found 10 personids whose BIOG_ADDR_DATA rows are IDENTICAL
between User MDB and SQLite, but the two pipelines store different
`c_index_addr_id` values.  All 10 winners are addr_type=1.

This script confirms what's actually going on by walking each
person's candidate set on both sides:

  - For each personid, list every BIOG_ADDR_DATA row of the
    winning addr_type, with its c_addr_id and c_sequence.
  - Show what each side stored, and what our PR L recompute
    produced.
  - Detect when multiple rows TIE on MAX(c_sequence) (or the
    sequence is 0 / NULL across the board).

Headline (writes to reports/index_addr_same_candidates_deep_dive.
json):

  Most/all 10 rows turn out to have multiple BIOG_ADDR_DATA rows
  of (person, addr_type=1) with the same c_sequence.  PHP, the
  Access maintenance UPDATE, and our PR L recompute then each
  pick a different row depending on the underlying engine's
  storage order.  That's a candidate non-determinism in the
  algorithm — neither side is "wrong", but the result is
  arbitrary.  Mitigation candidates: add an explicit secondary
  tie-break (e.g. MIN(c_addr_id)) to both implementations.

Conservative: nothing labelled as a confirmed bug.
"""
from __future__ import annotations

import json
import sqlite3
from collections import Counter
from pathlib import Path

import pyodbc

ROOT = Path(__file__).resolve().parent.parent
USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
SQLITE_DIR = ROOT / "data" / "cbdb_online_sqlite"
ADDR_CLASSIFICATION_JSON = (
    ROOT / "reports" / "index_addr_drift_classification.json")
OUT = (ROOT / "reports"
       / "index_addr_same_candidates_deep_dive.json")


def _open_user() -> pyodbc.Connection:
    return pyodbc.connect(
        "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={USER_MDB};", autocommit=True, readonly=True)


def _open_sqlite() -> sqlite3.Connection:
    paths = sorted(SQLITE_DIR.glob("*.sqlite3"))
    if not paths:
        raise SystemExit(f"no sqlite snapshot in {SQLITE_DIR}")
    return sqlite3.connect(str(paths[-1]))


def main() -> int:
    if not ADDR_CLASSIFICATION_JSON.exists():
        raise SystemExit(
            f"missing {ADDR_CLASSIFICATION_JSON}; run "
            f"`python analysis/classify_index_addr_drift.py` first")
    cls = json.loads(
        ADDR_CLASSIFICATION_JSON.read_text(encoding="utf-8"))
    rows = cls["buckets"]["same_candidates_diff_winner"]

    user = _open_user(); sql = _open_sqlite()
    cur_u = user.cursor(); cur_s = sql.cursor()

    findings = []
    tie_count = 0
    for r in rows:
        pid = r["personid"]
        winning_type = r["user"]["winner_addr_type"]
        cur_u.execute(
            "SELECT c_addr_type, c_addr_id, c_sequence "
            "FROM BIOG_ADDR_DATA "
            "WHERE c_personid=? AND c_addr_type=?",
            (pid, winning_type),
        )
        u_rows = [{"addr_type": int(r2[0]),
                    "addr_id": int(r2[1]) if r2[1] is not None else 0,
                    "sequence": int(r2[2]) if r2[2] is not None else 0}
                  for r2 in cur_u.fetchall()]
        cur_s.execute(
            "SELECT c_addr_type, c_addr_id, c_sequence "
            "FROM BIOG_ADDR_DATA "
            "WHERE c_personid=? AND c_addr_type=?",
            (pid, winning_type),
        )
        s_rows = [{"addr_type": int(r2[0]),
                    "addr_id": int(r2[1]) if r2[1] is not None else 0,
                    "sequence": int(r2[2]) if r2[2] is not None else 0}
                  for r2 in cur_s.fetchall()]

        # Detect MAX(c_sequence) tie.
        if u_rows:
            max_seq = max(rr["sequence"] for rr in u_rows)
            tied_at_max = [rr for rr in u_rows
                            if rr["sequence"] == max_seq]
        else:
            max_seq = None; tied_at_max = []
        is_tie = len(tied_at_max) >= 2
        if is_tie:
            tie_count += 1

        findings.append({
            "personid": pid,
            "name_chn": r["name_chn"],
            "name_py": r["name_py"],
            "winning_addr_type": winning_type,
            "user_stored_addr_id": r["user"]["index_addr_id"],
            "sqlite_stored_addr_id": r["sqlite"]["index_addr_id"],
            "user_recompute_addr_id": r["user"]["winner_recomputed"],
            "sqlite_recompute_addr_id":
                r["sqlite"]["winner_recomputed"],
            "user_addr_data_for_winning_type": u_rows,
            "sqlite_addr_data_for_winning_type": s_rows,
            "max_c_sequence": max_seq,
            "candidates_tied_at_max_sequence":
                len(tied_at_max),
            "is_max_sequence_tie": is_tie,
            "candidates_match_between_sides": (u_rows == s_rows),
        })

    # Detect: do the two sides agree on what to pick when there's
    # a tie, or do they pick different rows?  If user-stored,
    # sqlite-stored, and recompute are 3 distinct values within
    # the tied set, that's the strongest signal of pure storage-
    # order non-determinism.
    distinct_picks = []
    for f in findings:
        picks = {f["user_stored_addr_id"],
                 f["sqlite_stored_addr_id"],
                 f["user_recompute_addr_id"]}
        if len(picks) >= 2 and f["is_max_sequence_tie"]:
            distinct_picks.append(f["personid"])

    summary = {
        "total_examined": len(rows),
        "rows_with_max_sequence_tie": tie_count,
        "rows_with_distinct_picks_among_tied":
            len(distinct_picks),
        "interpretation": (
            "When BIOG_ADDR_DATA has multiple rows of "
            "(person, addr_type) tied on c_sequence (typically "
            "all c_sequence=1), MAX(c_sequence) returns 'one of "
            "them' and the storage engine's row order picks "
            "which one.  PHP IndexAddressRebuildService, the "
            "Access frmBaseMaintenance.CmdIndexAddress UPDATE, "
            "and our PR L recompute each see a different "
            "underlying engine (MariaDB / Microsoft JET / SQLite "
            "via pyodbc-sqlite3 ordering), so each picks a "
            "different addr_id.  Neither side is wrong — both "
            "implement the same documented rule (rank-priority + "
            "MAX c_sequence).  The shared cause is **no explicit "
            "secondary tie-break** when multiple rows share the "
            "same max c_sequence."
        ),
        "candidate_mitigation": (
            "Add an explicit secondary tie-break to BOTH "
            "implementations (e.g. MIN(c_addr_id) when multiple "
            "rows share max c_sequence).  Either side could "
            "implement it independently to reduce divergence.  "
            "Treat as candidate_release_process_or_algorithm_"
            "improvement rather than a confirmed bug."
        ),
        "is_confirmed_bug": False,
    }
    out = {"summary": summary, "rows": findings}
    OUT.write_text(
        json.dumps(out, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"wrote {OUT}")
    print(f"  total examined:                       "
          f"{summary['total_examined']}")
    print(f"  rows with MAX(c_sequence) tie:        "
          f"{summary['rows_with_max_sequence_tie']}")
    print(f"  rows with distinct picks among tied:  "
          f"{summary['rows_with_distinct_picks_among_tied']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
