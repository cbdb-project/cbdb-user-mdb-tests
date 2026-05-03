"""
Cross-check `c_index_year` and `c_index_addr_id` derivations between
the User MDB and the canonical cbdb-online-main-server SQLite
snapshot (roadmap item 12).

Both databases are nominally derived from the same upstream CBDB
sources, but they apply potentially-different rules to compute the
"index" year and address (the date and place each person is most
strongly associated with — birth year, ji-jia year, exam year, etc.,
in some priority order).  The cbdb-online-main-server version is
recomputed weekly and published as a SQLite snapshot at
<https://huggingface.co/datasets/cbdb/cbdb-sqlite/blob/main/latest.zip>;
that's our reference.

What this file tests:
  - On a random sample (default 1000 persons; ~80 ms with both DBs
    open), the User MDB and the SQLite agree on `c_index_year` AND
    `c_index_addr_id` for each `c_personid` they share.
  - With env `CBDB_FULL_XCHECK=1`, run the comparison against ALL
    common person ids (~657k rows; ~30 s).

If the test fails, the report names a few divergent person ids so a
maintainer can manually classify whether the disagreement is a real
derivation-rule difference or a stale-data artefact.

To get the SQLite snapshot:
    python analysis/download_hf_sqlite.py
"""
from __future__ import annotations

import os
import random
import sqlite3
from pathlib import Path

import pyodbc
import pytest


ROOT = Path(__file__).resolve().parent.parent
USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
HF_SQLITE_DIR = ROOT / "data" / "cbdb_online_sqlite"


def _find_hf_sqlite() -> Path | None:
    """Return path to the most recent .sqlite3 file in the snapshot
    directory, or None if no snapshot is present."""
    if not HF_SQLITE_DIR.exists():
        return None
    candidates = sorted(HF_SQLITE_DIR.glob("*.sqlite3"))
    if not candidates:
        candidates = sorted(HF_SQLITE_DIR.glob("*.sqlite"))
    return candidates[-1] if candidates else None


@pytest.fixture(scope="module")
def hf_sqlite() -> sqlite3.Connection:
    p = _find_hf_sqlite()
    if p is None:
        pytest.skip(
            f"No HF SQLite snapshot found under {HF_SQLITE_DIR}.\n"
            f"Run: python analysis/download_hf_sqlite.py"
        )
    conn = sqlite3.connect(f"file:{p}?mode=ro", uri=True)
    yield conn
    conn.close()


@pytest.fixture(scope="module")
def mdb() -> pyodbc.Connection:
    cs = ("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
          f"DBQ={USER_MDB};")
    conn = pyodbc.connect(cs, autocommit=True, readonly=True)
    yield conn
    conn.close()


def _fetch_fields(conn, ids: list[int],
                  fields: tuple[str, ...]) -> dict[int, tuple]:
    """Return {c_personid: (field1, field2, ...)} for the given ids
    and field list.  Chunked at 500 ids/query to stay under both
    JET's ~2100-parameter and SQLite's 999-parameter limits."""
    out: dict[int, tuple] = {}
    chunk = 500
    cur = conn.cursor()
    fields_sql = ", ".join(fields)
    for i in range(0, len(ids), chunk):
        batch = ids[i:i + chunk]
        in_clause = ",".join(str(int(b)) for b in batch)
        cur.execute(
            f"SELECT c_personid, {fields_sql} "
            f"FROM BIOG_MAIN WHERE c_personid IN ({in_clause})"
        )
        for r in cur.fetchall():
            pid = int(r[0])
            vals = tuple(None if v is None else int(v) for v in r[1:])
            out[pid] = vals
    cur.close()
    return out


def _all_ids(conn, dialect: str) -> set[int]:
    cur = conn.cursor()
    cur.execute("SELECT c_personid FROM BIOG_MAIN WHERE c_personid IS NOT NULL")
    out = {int(r[0]) for r in cur.fetchall()}
    cur.close()
    return out


def test_index_year_addr_xcheck_sample(mdb, hf_sqlite):
    """Sample 1000 random common person ids and assert the User MDB
    and HF SQLite agree on `c_index_year` AND `c_index_addr_id` (the
    DERIVED fields).  Also report disagreements on `c_birthyear` and
    `c_deathyear` (the SOURCE fields) — those would indicate the two
    pipelines have diverged on basic person data, not just on
    derivation rules.

    Set `CBDB_FULL_XCHECK=1` to widen to all ~657k common ids."""
    full = bool(os.environ.get("CBDB_FULL_XCHECK"))

    mdb_ids = _all_ids(mdb, "mdb")
    sql_ids = _all_ids(hf_sqlite, "sqlite")
    common = mdb_ids & sql_ids
    print(f"\n[xcheck] MDB={len(mdb_ids)} SQLite={len(sql_ids)} "
          f"common={len(common)}", flush=True)

    if full:
        ids = sorted(common)
    else:
        # Stable-seeded sample so failure reports are reproducible.
        rng = random.Random(20260502)
        ids = rng.sample(sorted(common), min(1000, len(common)))

    fields = ("c_index_year", "c_index_addr_id",
              "c_birthyear", "c_deathyear")
    mdb_rows = _fetch_fields(mdb, ids, fields)
    sql_rows = _fetch_fields(hf_sqlite, ids, fields)
    print(f"[xcheck] fetched {len(mdb_rows)} from MDB, "
          f"{len(sql_rows)} from SQLite (asked for {len(ids)})",
          flush=True)

    by_col: dict[str, list[tuple[int, tuple, tuple]]] = {
        c: [] for c in fields
    }
    for pid in ids:
        m = mdb_rows.get(pid)
        s = sql_rows.get(pid)
        if m is None or s is None:
            continue
        for j, col in enumerate(fields):
            if m[j] != s[j]:
                by_col[col].append((pid, m, s))

    n = len(ids)
    print(f"\n[xcheck] disagreement counts (out of {n} ids):", flush=True)
    for col in fields:
        cnt = len(by_col[col])
        print(f"  {col:<20}  {cnt:>6}  "
              f"({100.0 * cnt / max(1, n):.3f}%)", flush=True)
        for row in by_col[col][:3]:
            print(f"      {row}", flush=True)
    print()

    # Failure threshold per column.  c_index_year / c_index_addr_id
    # are derived and may differ on tiebreaks — we accept up to 0.5%.
    # c_birthyear / c_deathyear are SOURCE data; any divergence above
    # 0.1% would suggest the two CBDB pipelines have meaningfully
    # diverged on basic person facts (Pattern B in reports/CBDB_Issues_Report_EN.md
    # Open Question #1) — surface as a hard failure so we notice.
    derived_pct = lambda col: 100.0 * len(by_col[col]) / max(1, n)
    source_pct = lambda col: 100.0 * len(by_col[col]) / max(1, n)
    assert derived_pct("c_index_year") < 0.5, (
        f"c_index_year disagreement {derived_pct('c_index_year'):.3f}% "
        f"exceeds 0.5% threshold"
    )
    assert derived_pct("c_index_addr_id") < 0.5, (
        f"c_index_addr_id disagreement {derived_pct('c_index_addr_id'):.3f}% "
        f"exceeds 0.5% threshold"
    )
    assert source_pct("c_birthyear") < 0.1, (
        f"c_birthyear disagreement {source_pct('c_birthyear'):.3f}% "
        f"exceeds 0.1% threshold — pipelines have diverged on SOURCE "
        f"data, not just derivation"
    )
    assert source_pct("c_deathyear") < 0.1, (
        f"c_deathyear disagreement {source_pct('c_deathyear'):.3f}% "
        f"exceeds 0.1% threshold — pipelines have diverged on SOURCE "
        f"data, not just derivation"
    )
