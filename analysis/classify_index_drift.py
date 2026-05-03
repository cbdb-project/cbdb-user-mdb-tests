"""Classify per-personid drift between the User MDB's BIOG_MAIN and
the cbdb-online-main-server SQLite snapshot.

Goal: replace the previous "13 illustrative samples" framing of the
index-year / index-address cross-check with an actual taxonomy of
where the differences come from, so a maintainer can decide which
buckets warrant follow-up.

The algorithm sources we are comparing:

  c_index_addr_id:
    User MDB → analysis/dump/vba/Form_frmIndexAddr.vb
    SQLite   → IndexAddressRebuildService.php
                in https://github.com/cbdb-project/cbdb-online-main-server

  c_index_year:
    User MDB → not in the shipped User MDB (see notes); the column
                is read but not rewritten by user-facing forms.
                Likely produced by an Admin MDB we don't have.
    SQLite   → IndexYearRebuildService.php
                in the same repo.

For each common personid we read four fields from each side
(c_index_year, c_index_addr_id, c_birthyear, c_deathyear) and
classify:

  exact_match
  source_drift_index_diffs_too   - birthyear/deathyear differ AND
                                   at least one index differs
  source_drift_index_agrees      - source differs, both indices agree
  index_year_only_diff           - birthyear/deathyear identical;
                                   only c_index_year differs
  index_addr_only_diff           - birthyear/deathyear identical;
                                   only c_index_addr_id differs
  index_both_diff                - birthyear/deathyear identical;
                                   both indices differ

Buckets 4-6 are NOT automatically bugs — they are the unclassified
remainder that needs per-row investigation.  See
`analysis/index_drift_algorithm_notes.md` for the rationale.

Output: `reports/index_drift_classification.json` with per-bucket
counts plus up to 8 sampled examples per non-empty bucket.
"""
from __future__ import annotations

import json
import os
import sqlite3
import sys
from pathlib import Path

import pyodbc

ROOT = Path(__file__).resolve().parent.parent
USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
SQLITE_DIR = ROOT / "data" / "cbdb_online_sqlite"
OUT = ROOT / "reports" / "index_drift_classification.json"

SAMPLE_LIMIT = 8  # examples per bucket


def _open_user_mdb() -> pyodbc.Connection:
    cs = ("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
          f"DBQ={USER_MDB};")
    return pyodbc.connect(cs, autocommit=True, readonly=True)


def _open_sqlite() -> sqlite3.Connection:
    paths = sorted(SQLITE_DIR.glob("*.sqlite3"))
    if not paths:
        raise SystemExit(
            f"no sqlite snapshot in {SQLITE_DIR} — "
            f"run `python analysis/download_hf_sqlite.py` first"
        )
    print(f"using sqlite snapshot: {paths[-1].name}")
    return sqlite3.connect(str(paths[-1]))


def _normalize(v):
    """Treat 0 / None / '' as the same 'no value' for comparison
    purposes.  Both pipelines use 0 as the sentinel for missing
    year, but a snapshot bug or schema migration could leave NULLs
    in either place; collapse them."""
    if v is None:
        return 0
    if isinstance(v, str) and not v.strip():
        return 0
    return v


def _row(record: dict) -> dict:
    """Snapshot just the four fields we classify on, plus
    diagnostic context useful in the per-bucket samples."""
    return {
        "index_year": _normalize(record.get("index_year")),
        "index_addr_id": _normalize(record.get("index_addr_id")),
        "birthyear": _normalize(record.get("birthyear")),
        "deathyear": _normalize(record.get("deathyear")),
        # Diagnostic context (not used in classification):
        "index_year_source_id":
            record.get("index_year_source_id"),
        "index_year_type_code":
            record.get("index_year_type_code"),
    }


def _classify(u: dict, s: dict) -> str:
    src_drift = (u["birthyear"] != s["birthyear"]
                 or u["deathyear"] != s["deathyear"])
    year_diff = u["index_year"] != s["index_year"]
    addr_diff = u["index_addr_id"] != s["index_addr_id"]
    if not src_drift and not year_diff and not addr_diff:
        return "exact_match"
    if src_drift:
        return ("source_drift_index_diffs_too"
                if (year_diff or addr_diff)
                else "source_drift_index_agrees")
    # Source matched on the four fields we compare; index outputs differ.
    if year_diff and addr_diff:
        return "index_both_diff"
    if year_diff:
        return "index_year_only_diff"
    return "index_addr_only_diff"


def main() -> int:
    user = _open_user_mdb()
    sql = _open_sqlite()
    cur_u = user.cursor()
    cur_s = sql.cursor()

    print("loading User MDB BIOG_MAIN ...")
    cur_u.execute(
        "SELECT c_personid, c_name, c_name_chn, "
        "c_index_year, c_index_addr_id, c_birthyear, c_deathyear, "
        "c_index_year_source_id, c_index_year_type_code "
        "FROM BIOG_MAIN WHERE c_personid > 0"
    )
    user_rows: dict[int, dict] = {}
    for r in cur_u.fetchall():
        user_rows[int(r[0])] = {
            "name_py": r[1] or "",
            "name_chn": r[2] or "",
            "index_year": r[3],
            "index_addr_id": r[4],
            "birthyear": r[5],
            "deathyear": r[6],
            "index_year_source_id": r[7],
            "index_year_type_code": r[8],
        }
    print(f"  {len(user_rows):,} User MDB persons")

    print("loading SQLite snapshot BIOG_MAIN ...")
    cur_s.execute(
        "SELECT c_personid, c_name, c_name_chn, "
        "c_index_year, c_index_addr_id, c_birthyear, c_deathyear, "
        "c_index_year_source_id, c_index_year_type_code "
        "FROM BIOG_MAIN WHERE c_personid > 0"
    )
    sqlite_rows: dict[int, dict] = {}
    for r in cur_s.fetchall():
        sqlite_rows[int(r[0])] = {
            "name_py": r[1] or "",
            "name_chn": r[2] or "",
            "index_year": r[3],
            "index_addr_id": r[4],
            "birthyear": r[5],
            "deathyear": r[6],
            "index_year_source_id": r[7],
            "index_year_type_code": r[8],
        }
    print(f"  {len(sqlite_rows):,} SQLite persons")

    common = sorted(set(user_rows) & set(sqlite_rows))
    only_user = sorted(set(user_rows) - set(sqlite_rows))
    only_sqlite = sorted(set(sqlite_rows) - set(user_rows))
    print(f"\n  common personids:        {len(common):,}")
    print(f"  in User MDB only:        {len(only_user):,}")
    print(f"  in SQLite snapshot only: {len(only_sqlite):,}")

    counts: dict[str, int] = {}
    samples: dict[str, list] = {}
    bucket_order = [
        "exact_match",
        "source_drift_index_agrees",
        "source_drift_index_diffs_too",
        "index_year_only_diff",
        "index_addr_only_diff",
        "index_both_diff",
    ]
    for b in bucket_order:
        counts[b] = 0
        samples[b] = []

    for pid in common:
        u_full = user_rows[pid]
        s_full = sqlite_rows[pid]
        u = _row(u_full); s = _row(s_full)
        bucket = _classify(u, s)
        counts[bucket] += 1
        if (bucket != "exact_match"
                and len(samples[bucket]) < SAMPLE_LIMIT):
            samples[bucket].append({
                "personid": pid,
                "name_py": u_full["name_py"],
                "name_chn": u_full["name_chn"],
                "user": u,
                "sqlite": s,
            })

    # Print summary
    total = sum(counts.values())
    print(f"\n=== classification summary ({total:,} compared) ===")
    for b in bucket_order:
        pct = 100.0 * counts[b] / max(total, 1)
        print(f"  {b:35s} {counts[b]:>10,d}  ({pct:6.3f}%)")

    out_doc = {
        "summary": {
            "user_mdb_total": len(user_rows),
            "sqlite_total": len(sqlite_rows),
            "common": len(common),
            "in_user_only": len(only_user),
            "in_sqlite_only": len(only_sqlite),
            "buckets": {b: counts[b] for b in bucket_order},
        },
        "bucket_meaning": {
            "exact_match":
                "All four compared fields (c_index_year, "
                "c_index_addr_id, c_birthyear, c_deathyear) agree.",
            "source_drift_index_agrees":
                "c_birthyear or c_deathyear differs but BOTH index "
                "fields agree.  Both implementations tolerated the "
                "underlying source drift and produced compatible "
                "outputs.",
            "source_drift_index_diffs_too":
                "c_birthyear or c_deathyear differs AND at least one "
                "index field differs.  Consistent with the simple "
                "data-drift hypothesis, but doesn't prove it.",
            "index_year_only_diff":
                "c_birthyear and c_deathyear identical; only "
                "c_index_year differs.  Could be PHP IndexYearRebuild "
                "vs (Admin) VBA divergence on year picking; could "
                "also be drift in evidence tables we don't compare "
                "(ENTRY_DATA exam years, NIAN_HAO mappings, "
                "fl_earliest_year / fl_latest_year, etc.).",
            "index_addr_only_diff":
                "c_birthyear and c_deathyear identical; only "
                "c_index_addr_id differs.  Could be PHP "
                "IndexAddressRebuild vs VBA Form_frmIndexAddr "
                "divergence on rank order / tie-break; could also "
                "be drift in BIOG_ADDR_DATA (which we don't compare).",
            "index_both_diff":
                "c_birthyear and c_deathyear identical; BOTH index "
                "fields differ.  Strongest single-row signal of "
                "compound divergence; warrants per-row "
                "investigation.",
        },
        "limitations": [
            "Source-data comparison is restricted to the four "
            "fields available on both BIOG_MAIN tables (c_index_year, "
            "c_index_addr_id, c_birthyear, c_deathyear).  Drift in "
            "evidence tables (BIOG_ADDR_DATA, ENTRY_DATA, NIAN_HAO, "
            "fl_*_year, etc.) is invisible to this classifier and "
            "can show up in any of the index_*_diff buckets.",
            "We have not line-by-line audited PHP "
            "IndexYearRebuildService.php / IndexAddressRebuildService"
            ".php.  See `analysis/index_drift_algorithm_notes.md` "
            "for what we do know about each side.",
            "Buckets 4-6 are NOT automatically bugs.  They are the "
            "unclassified remainder that needs per-row "
            "investigation.",
        ],
        "samples": samples,
    }
    OUT.write_text(
        json.dumps(out_doc, ensure_ascii=False, indent=2,
                    default=str),
        encoding="utf-8",
    )
    print(f"\nwrote {OUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
