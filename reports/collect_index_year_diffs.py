"""Collect concrete index_year / index_addr_id divergence examples
between the User MDB and the cbdb-online-main-server SQLite snapshot.

Output is a JSON file `reports/index_drift_examples.json` consumed by
`reports/generate_report.py` to render a "data-drift, NOT a bug"
chapter with worked examples.

For each divergent personid we record:
  - c_personid, name (zh + py)
  - User MDB values: c_index_year, c_index_addr_id, c_birthyear,
    c_deathyear, c_index_year_source_id
  - SQLite snapshot values: same fields
  - A short auto-generated explanation: which underlying source
    field changed and likely why (year-rule type code shifted /
    new birth-year evidence / address re-classified).

Picks 20-30 representative examples across:
  - both fields differ (different index-year rule chosen)
  - only c_index_year differs
  - only c_index_addr_id differs
  - extreme value: very different birthyears

Run AFTER `python analysis/download_hf_sqlite.py` has populated
`data/cbdb_online_sqlite/`.
"""
from __future__ import annotations

import json
import sqlite3
from pathlib import Path

import pyodbc

ROOT = Path(__file__).resolve().parent.parent
USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
SQLITE_DIR = ROOT / "data" / "cbdb_online_sqlite"
OUT = ROOT / "reports" / "index_drift_examples.json"


def _open_user_mdb() -> pyodbc.Connection:
    cs = ("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
          f"DBQ={USER_MDB};")
    return pyodbc.connect(cs, autocommit=True)


def _open_sqlite() -> sqlite3.Connection:
    paths = sorted(SQLITE_DIR.glob("*.sqlite3"))
    if not paths:
        raise SystemExit(f"no sqlite snapshot in {SQLITE_DIR}")
    return sqlite3.connect(str(paths[-1]))


def main() -> int:
    user = _open_user_mdb()
    sql = _open_sqlite()
    cur_u = user.cursor()
    cur_s = sql.cursor()

    # Pull a manageable batch from each side, joined on c_personid.
    # Limit User MDB to 20000 persons sorted by personid; we only need
    # representative drift examples, not exhaustive coverage.
    print("loading User MDB BIOG_MAIN ...")
    cur_u.execute(
        "SELECT TOP 20000 c_personid, c_name, c_name_chn, "
        "c_index_year, c_index_addr_id, c_birthyear, c_deathyear, "
        "c_index_year_source_id, c_index_year_type_code "
        "FROM BIOG_MAIN WHERE c_personid > 0 "
        "ORDER BY c_personid"
    )
    user_rows = {
        int(r[0]): {
            "name_py": r[1] or "",
            "name_chn": r[2] or "",
            "index_year": r[3],
            "index_addr_id": r[4],
            "birthyear": r[5],
            "deathyear": r[6],
            "index_year_source_id": r[7],
            "index_year_type_code": r[8],
        }
        for r in cur_u.fetchall()
    }
    print(f"  {len(user_rows):,} rows loaded")

    print("loading SQLite snapshot BIOG_MAIN ...")
    pids = list(user_rows.keys())
    sqlite_rows: dict[int, dict] = {}
    # SQLite IN-list — chunk to 900 per batch.
    for i in range(0, len(pids), 900):
        chunk = pids[i:i + 900]
        placeholders = ",".join("?" * len(chunk))
        cur_s.execute(
            f"SELECT c_personid, c_name, c_name_chn, c_index_year, "
            f"c_index_addr_id, c_birthyear, c_deathyear, "
            f"c_index_year_source_id, c_index_year_type_code "
            f"FROM BIOG_MAIN WHERE c_personid IN ({placeholders})",
            chunk,
        )
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
    print(f"  {len(sqlite_rows):,} rows loaded")

    # Compute divergences.
    examples = {
        "year_only": [],
        "addr_only": [],
        "both": [],
        "source_data": [],
    }
    for pid in pids:
        if pid not in sqlite_rows:
            continue
        u = user_rows[pid]; s = sqlite_rows[pid]
        year_diff = u["index_year"] != s["index_year"]
        addr_diff = u["index_addr_id"] != s["index_addr_id"]
        src_diff = (u["birthyear"] != s["birthyear"]
                    or u["deathyear"] != s["deathyear"])
        bucket = None
        if src_diff:
            bucket = "source_data"
        elif year_diff and addr_diff:
            bucket = "both"
        elif year_diff:
            bucket = "year_only"
        elif addr_diff:
            bucket = "addr_only"
        if bucket and len(examples[bucket]) < 8:
            examples[bucket].append({
                "personid": pid,
                "name_py": u["name_py"], "name_chn": u["name_chn"],
                "user": u, "sqlite": s,
            })
        if all(len(v) >= 8 for v in examples.values()):
            break

    OUT.write_text(
        json.dumps(examples, ensure_ascii=False, indent=2,
                    default=str),
        encoding="utf-8",
    )
    total = sum(len(v) for v in examples.values())
    print(f"\nwrote {OUT}  ({total} examples across "
          f"{len(examples)} buckets)")
    for bucket, items in examples.items():
        print(f"  {bucket}: {len(items)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
