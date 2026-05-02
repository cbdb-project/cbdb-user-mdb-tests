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


def _fetch_index(conn, ids: list[int], dialect: str) -> dict[int, tuple]:
    """Return {c_personid: (c_index_year, c_index_addr_id)} for the
    given ids.  `dialect` is "mdb" or "sqlite" — used to format the
    IN clause within the engines' max-parameter limits.

    JET tops out at ~2100 parameters per query; SQLite at 999 by
    default.  Chunk to be safe under both."""
    out: dict[int, tuple] = {}
    chunk = 500
    cur = conn.cursor()
    for i in range(0, len(ids), chunk):
        batch = ids[i:i + chunk]
        in_clause = ",".join(str(int(b)) for b in batch)
        cur.execute(
            f"SELECT c_personid, c_index_year, c_index_addr_id "
            f"FROM BIOG_MAIN WHERE c_personid IN ({in_clause})"
        )
        for r in cur.fetchall():
            pid = int(r[0])
            iy = None if r[1] is None else int(r[1])
            ia = None if r[2] is None else int(r[2])
            out[pid] = (iy, ia)
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
    and HF SQLite agree on `c_index_year` AND `c_index_addr_id`.

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

    mdb_idx = _fetch_index(mdb, ids, "mdb")
    sql_idx = _fetch_index(hf_sqlite, ids, "sqlite")
    print(f"[xcheck] fetched {len(mdb_idx)} from MDB, "
          f"{len(sql_idx)} from SQLite (asked for {len(ids)})", flush=True)

    iy_mismatch: list[tuple[int, tuple, tuple]] = []
    ia_mismatch: list[tuple[int, tuple, tuple]] = []
    for pid in ids:
        m = mdb_idx.get(pid)
        s = sql_idx.get(pid)
        if m is None or s is None:
            continue
        if m[0] != s[0]:
            iy_mismatch.append((pid, m, s))
        if m[1] != s[1]:
            ia_mismatch.append((pid, m, s))

    print(f"[xcheck] c_index_year mismatches: {len(iy_mismatch)} "
          f"/ {len(ids)}", flush=True)
    print(f"[xcheck] c_index_addr_id mismatches: {len(ia_mismatch)} "
          f"/ {len(ids)}", flush=True)

    if iy_mismatch:
        print(f"[xcheck] first {min(10, len(iy_mismatch))} "
              f"c_index_year diffs (pid, MDB, SQLite):", flush=True)
        for row in iy_mismatch[:10]:
            print(f"    {row}", flush=True)
    if ia_mismatch:
        print(f"[xcheck] first {min(10, len(ia_mismatch))} "
              f"c_index_addr_id diffs (pid, MDB, SQLite):", flush=True)
        for row in ia_mismatch[:10]:
            print(f"    {row}", flush=True)

    # The two databases COULD legitimately disagree on a few persons
    # (different snapshot dates, derivation rule changes mid-week).
    # Treat substantial disagreement as a finding to flag.  Threshold
    # tuned conservatively: any sample-level disagreement above 0.5%
    # (so 5+ on a 1000-sample) is worth investigating.
    iy_pct = 100.0 * len(iy_mismatch) / max(1, len(ids))
    ia_pct = 100.0 * len(ia_mismatch) / max(1, len(ids))
    assert iy_pct < 0.5, (
        f"c_index_year disagreement {iy_pct:.2f}% "
        f"({len(iy_mismatch)} / {len(ids)}) exceeds 0.5% threshold; "
        f"investigate (sample examples printed above)"
    )
    assert ia_pct < 0.5, (
        f"c_index_addr_id disagreement {ia_pct:.2f}% "
        f"({len(ia_mismatch)} / {len(ids)}) exceeds 0.5% threshold; "
        f"investigate (sample examples printed above)"
    )
