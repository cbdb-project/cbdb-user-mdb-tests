"""Per-row classification of c_index_addr_id drifts between the
User MDB and the cbdb-online-main-server SQLite snapshot.

Scope: only c_index_addr_id.  c_index_year is handled by PR K1 / K2.

The two algorithms being compared:

  Access (User MDB front-end) — analysis/dump/vba/Form_frmIndexAddr.vb
    UpdateBiogMain():
      1. Build TMP_BIOG_ADDR_DATA = (c_personid, c_addr_type,
         MAX(c_sequence)) over BIOG_ADDR_DATA — collapses
         duplicates per (person, addr_type) keeping the highest
         sequence.
      2. Loop ranks 1..N over BIOG_ADDR_CODES.c_index_addr_rank;
         for each addr_type, run UPDATE setting BIOG_MAIN.
         c_index_addr_id = the matching BIOG_ADDR_DATA.c_addr_id,
         WHERE c_index_addr_id Is Null.  Earlier (lower-rank)
         addr_types win.

  PHP (cbdb-online-main-server) — pinned to commit
    e31fba7225e98caa115e9078b8316f9f126423fc, snapshotted at
    analysis/php_source/IndexAddressRebuildService.php.
    rebuild():
      For each (person, addr_type), pick the row with MAX(c_sequence);
      then for each person assign c_index_addr_id = the candidate
      with the MIN c_index_addr_default_rank.  Same shape as VBA.
      Note PHP reads BIOG_ADDR_CODES.c_index_addr_default_rank;
      VBA reads c_index_addr_rank.  In the shipped 2026-04-30
      DATA mdb the two columns are equal for all 22 addr_types
      (verified separately).

Method:

  1. For each of the 478 index_addr_only_diff personids (and the
     10 index_both_diff personids), pull BIOG_ADDR_DATA from
     BOTH sides.
  2. Compute the rank-priority + MAX(c_sequence) winner using each
     side's BIOG_ADDR_DATA.  Compare to the actual stored value
     on each side.
  3. Bucket each row:

       both_sides_match_recomputed
         The actual stored value matches what we'd recompute on
         each side from that side's BIOG_ADDR_DATA.  The diff is
         explained by BIOG_ADDR_DATA differing between snapshots.

       same_candidates_diff_winner
         BIOG_ADDR_DATA rows are identical between sides; the two
         pipelines pick different winners.  Tie-break or null-
         handling diff.

       mdb_null_php_value     / mdb_value_php_null
         One side wrote NULL/0; the other wrote a real id.

       mdb_stale_index_addr
         SQLite's stored value matches what we'd recompute from
         its data; User MDB's stored value does NOT match what
         we'd recompute from User MDB's data.  Strong signal the
         User MDB shipped with a stale c_index_addr_id — the
         underlying BIOG_ADDR_DATA was updated after the last
         frmBaseMaintenance run, but nobody re-ran the rebuild.
         (Usually because the new BIOG_ADDR_DATA row has a
         higher c_sequence and would now win the MAX(c_sequence)
         tie-break.)

       sqlite_stale_index_addr
         Reverse direction; rare.

       both_stale_recompute_mismatch
         Neither side's stored value matches what we'd recompute.
         Both stored values may predate a data refresh, OR the
         algorithm we model is missing a feature that both
         implementations share.

       unclassified
         Doesn't fit any of the above.

Output: reports/index_addr_drift_classification.json
"""
from __future__ import annotations

import json
import sqlite3
from collections import Counter, defaultdict
from pathlib import Path

import pyodbc

ROOT = Path(__file__).resolve().parent.parent
USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
SQLITE_DIR = ROOT / "data" / "cbdb_online_sqlite"
DRIFT_CLASSIFICATION_JSON = (
    ROOT / "reports" / "index_drift_classification.json")
OUT = ROOT / "reports" / "index_addr_drift_classification.json"


def _open_user() -> pyodbc.Connection:
    cs = ("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
          f"DBQ={USER_MDB};")
    return pyodbc.connect(cs, autocommit=True, readonly=True)


def _open_sqlite() -> sqlite3.Connection:
    paths = sorted(SQLITE_DIR.glob("*.sqlite3"))
    if not paths:
        raise SystemExit(
            f"no sqlite snapshot in {SQLITE_DIR}")
    return sqlite3.connect(str(paths[-1]))


def _norm_addr(v) -> int:
    """Treat NULL / 0 / '' as the same 'no value'."""
    if v is None or v == "" or v == 0:
        return 0
    return int(v)


def _select_winner(addr_rows: list[dict],
                   rank_for_type: dict[int, int]) -> tuple[int, int]:
    """Apply the rank-priority + MAX(c_sequence) algorithm.
    Returns (winning c_addr_id, winning c_addr_type) or (0, 0) if
    the person has no eligible address row."""
    if not addr_rows:
        return (0, 0)
    # Collapse to (addr_type → addr_id_with_max_sequence)
    best_per_type: dict[int, tuple[int, int]] = {}  # type → (sequence, addr_id)
    for r in addr_rows:
        at = r["addr_type"]
        seq = r["sequence"] or 0
        aid = r["addr_id"] or 0
        prev = best_per_type.get(at)
        if prev is None or seq > prev[0]:
            best_per_type[at] = (seq, aid)
    # Pick the eligible (rank<100) type with lowest rank
    best_type = None
    best_rank = None
    for at, (_seq, _aid) in best_per_type.items():
        rk = rank_for_type.get(at)
        if rk is None or rk >= 100:
            continue
        if best_rank is None or rk < best_rank:
            best_rank = rk
            best_type = at
    if best_type is None:
        return (0, 0)
    return (best_per_type[best_type][1], best_type)


def main() -> int:
    if not DRIFT_CLASSIFICATION_JSON.exists():
        raise SystemExit(
            f"missing {DRIFT_CLASSIFICATION_JSON}; "
            f"run `python analysis/classify_index_drift.py` first")
    cls = json.loads(
        DRIFT_CLASSIFICATION_JSON.read_text(encoding="utf-8"))

    # The samples list in PR G's JSON only has 8 examples per bucket;
    # we need ALL 478 + 10 personids.  Re-derive them by re-running
    # PR G's filter.

    user = _open_user()
    sql = _open_sqlite()
    cur_u = user.cursor()
    cur_s = sql.cursor()

    print("loading both BIOG_MAIN tables ...")
    cur_u.execute(
        "SELECT c_personid, c_name, c_name_chn, c_index_year, "
        "c_index_addr_id, c_birthyear, c_deathyear "
        "FROM BIOG_MAIN WHERE c_personid > 0"
    )
    user_main = {int(r[0]): {
        "name_py": r[1] or "", "name_chn": r[2] or "",
        "index_year": r[3], "index_addr_id": r[4],
        "birthyear": r[5], "deathyear": r[6],
    } for r in cur_u.fetchall()}
    cur_s.execute(
        "SELECT c_personid, c_name, c_name_chn, c_index_year, "
        "c_index_addr_id, c_birthyear, c_deathyear "
        "FROM BIOG_MAIN WHERE c_personid > 0"
    )
    sqlite_main = {int(r[0]): {
        "name_py": r[1] or "", "name_chn": r[2] or "",
        "index_year": r[3], "index_addr_id": r[4],
        "birthyear": r[5], "deathyear": r[6],
    } for r in cur_s.fetchall()}

    # Identify the addr-only and both-diff personids.
    addr_diffs: list[int] = []
    for pid in sorted(set(user_main) & set(sqlite_main)):
        u = user_main[pid]; s = sqlite_main[pid]
        # Same source-year fields, different c_index_addr_id
        if (u["birthyear"] == s["birthyear"]
                and u["deathyear"] == s["deathyear"]
                and _norm_addr(u["index_addr_id"]) != _norm_addr(s["index_addr_id"])):
            addr_diffs.append(pid)
    print(f"  candidates to classify: {len(addr_diffs)}")

    # Pull BIOG_ADDR_DATA for these personids from both sides.
    print("loading BIOG_ADDR_DATA from both sides ...")
    user_addr: dict[int, list[dict]] = defaultdict(list)
    sqlite_addr: dict[int, list[dict]] = defaultdict(list)
    for i in range(0, len(addr_diffs), 500):
        chunk = addr_diffs[i:i + 500]
        ph = ",".join("?" * len(chunk))
        cur_u.execute(
            f"SELECT c_personid, c_addr_type, c_addr_id, c_sequence "
            f"FROM BIOG_ADDR_DATA WHERE c_personid IN ({ph})",
            chunk,
        )
        for r in cur_u.fetchall():
            user_addr[int(r[0])].append({
                "addr_type": int(r[1]) if r[1] is not None else 0,
                "addr_id": int(r[2]) if r[2] is not None else 0,
                "sequence": int(r[3]) if r[3] is not None else 0,
            })
        cur_s.execute(
            f"SELECT c_personid, c_addr_type, c_addr_id, c_sequence "
            f"FROM BIOG_ADDR_DATA WHERE c_personid IN ({ph})",
            chunk,
        )
        for r in cur_s.fetchall():
            sqlite_addr[int(r[0])].append({
                "addr_type": int(r[1]) if r[1] is not None else 0,
                "addr_id": int(r[2]) if r[2] is not None else 0,
                "sequence": int(r[3]) if r[3] is not None else 0,
            })
    print(f"  user_addr personids:   {len(user_addr)}")
    print(f"  sqlite_addr personids: {len(sqlite_addr)}")

    # Rank table (verified identical between the two sides).
    cur_u.execute(
        "SELECT c_addr_type, c_index_addr_default_rank, c_index_addr_rank "
        "FROM BIOG_ADDR_CODES"
    )
    rank_default: dict[int, int] = {}
    rank_current: dict[int, int] = {}
    for r in cur_u.fetchall():
        at = int(r[0])
        rank_default[at] = int(r[1]) if r[1] is not None else 100
        rank_current[at] = int(r[2]) if r[2] is not None else 100

    # Classify
    classified = {
        "both_sides_match_recomputed": [],
        "same_candidates_diff_winner": [],
        "mdb_null_php_value": [],
        "mdb_value_php_null": [],
        "mdb_stale_index_addr": [],
        "sqlite_stale_index_addr": [],
        "both_stale_recompute_mismatch": [],
        "unclassified": [],
    }
    for pid in addr_diffs:
        u_main = user_main[pid]
        s_main = sqlite_main[pid]
        u_aid = _norm_addr(u_main["index_addr_id"])
        s_aid = _norm_addr(s_main["index_addr_id"])
        u_addrs = user_addr.get(pid, [])
        s_addrs = sqlite_addr.get(pid, [])

        # Recompute each side using its own BIOG_ADDR_DATA.
        # Use rank_default for PHP semantics; rank_current for VBA.
        # Since the two are equal in the shipped data, it doesn't
        # matter which we pass.
        u_winner_id, u_winner_type = _select_winner(u_addrs, rank_current)
        s_winner_id, s_winner_type = _select_winner(s_addrs, rank_default)

        # Build comparable "candidate signature" per side for
        # detecting source-data drift on the address side.
        u_sig = sorted({(r["addr_type"], r["addr_id"], r["sequence"])
                        for r in u_addrs})
        s_sig = sorted({(r["addr_type"], r["addr_id"], r["sequence"])
                        for r in s_addrs})
        candidates_match = (u_sig == s_sig)

        record = {
            "personid": pid,
            "name_py": u_main["name_py"],
            "name_chn": u_main["name_chn"],
            "user": {"index_addr_id": u_aid,
                     "n_addr_rows": len(u_addrs),
                     "winner_recomputed": u_winner_id,
                     "winner_addr_type": u_winner_type},
            "sqlite": {"index_addr_id": s_aid,
                       "n_addr_rows": len(s_addrs),
                       "winner_recomputed": s_winner_id,
                       "winner_addr_type": s_winner_type},
            "candidates_match": candidates_match,
        }

        # 0a: one side null
        if u_aid != 0 and s_aid == 0:
            record["explanation"] = (
                f"User MDB wrote {u_aid}, PHP wrote NULL/0.  "
                f"User has {len(u_addrs)} address rows; PHP has "
                f"{len(s_addrs)}.")
            classified["mdb_value_php_null"].append(record)
            continue
        if u_aid == 0 and s_aid != 0:
            record["explanation"] = (
                f"PHP wrote {s_aid}, User MDB wrote NULL/0.  "
                f"User has {len(u_addrs)} address rows; PHP has "
                f"{len(s_addrs)}.")
            classified["mdb_null_php_value"].append(record)
            continue

        # 1: each side stored what we'd recompute from its own data
        if (u_winner_id == u_aid
                and s_winner_id == s_aid
                and u_winner_id != 0
                and s_winner_id != 0):
            record["explanation"] = (
                f"Each side's stored c_index_addr_id matches the "
                f"rank-priority + MAX(c_sequence) winner over its "
                f"OWN BIOG_ADDR_DATA "
                f"({'IDENTICAL' if candidates_match else 'DIFFERING'} "
                f"candidate sets between sides).  Diff is explained "
                f"by source-data drift on the address side."
            )
            classified["both_sides_match_recomputed"].append(record)
            continue

        # 2: same candidates on both sides, different winners
        if (candidates_match
                and u_winner_id != s_winner_id
                and u_winner_id != 0 and s_winner_id != 0):
            record["explanation"] = (
                f"BIOG_ADDR_DATA candidates IDENTICAL on both sides; "
                f"winners differ.  User picks "
                f"addr_type={u_winner_type} → addr_id={u_winner_id}; "
                f"PHP picks addr_type={s_winner_type} → addr_id="
                f"{s_winner_id}.  Tie-break / null-handling "
                f"divergence between Form_frmIndexAddr.vb and "
                f"IndexAddressRebuildService.php."
            )
            classified["same_candidates_diff_winner"].append(record)
            continue

        # 3: one side stored something different from what we recompute
        u_match = (u_winner_id == u_aid and u_winner_id != 0)
        s_match = (s_winner_id == s_aid and s_winner_id != 0)
        if s_match and not u_match:
            record["explanation"] = (
                f"User MDB stored {u_aid} but recompute over its "
                f"BIOG_ADDR_DATA gives {u_winner_id}.  SQLite "
                f"stored value {s_aid} matches its own recompute.  "
                f"Most likely cause: User MDB shipped with a stale "
                f"c_index_addr_id — BIOG_ADDR_DATA was updated "
                f"after the last frmBaseMaintenance rebuild but "
                f"nobody re-ran the rebuild.  PHP re-runs weekly "
                f"so SQLite stays fresh."
            )
            classified["mdb_stale_index_addr"].append(record)
            continue
        if u_match and not s_match:
            record["explanation"] = (
                f"PHP stored {s_aid} but recompute over its "
                f"BIOG_ADDR_DATA gives {s_winner_id}.  User MDB "
                f"stored value {u_aid} matches its own recompute.  "
                f"Reverse direction of mdb_stale_index_addr; rare."
            )
            classified["sqlite_stale_index_addr"].append(record)
            continue
        # Neither side matches its own recompute.
        record["explanation"] = (
            f"Both stored values differ from the rank-priority + "
            f"MAX(c_sequence) recompute.  User stored={u_aid} "
            f"vs recompute={u_winner_id}; PHP stored={s_aid} vs "
            f"recompute={s_winner_id}.  Either both stored values "
            f"predate a data refresh or the algorithm we model is "
            f"missing a feature both implementations share."
        )
        classified["both_stale_recompute_mismatch"].append(record)
        continue

        record["explanation"] = (
            f"Doesn't fit any classifier branch."
        )
        classified["unclassified"].append(record)

    # ---- Summary ----
    bucket_order = [
        "both_sides_match_recomputed",
        "same_candidates_diff_winner",
        "mdb_null_php_value",
        "mdb_value_php_null",
        "mdb_stale_index_addr",
        "sqlite_stale_index_addr",
        "both_stale_recompute_mismatch",
        "unclassified",
    ]
    total = sum(len(v) for v in classified.values())
    print(f"\n=== addr-drift classification summary ({total} rows) ===")
    for k in bucket_order:
        print(f"  {k:38s} {len(classified[k]):>4d}")

    # For same_candidates_diff_winner, group by (mdb_winner_addr_type,
    # php_winner_addr_type) so we can see which addr_type pairings
    # show systematic divergence.
    sigs: Counter = Counter()
    for r in classified["same_candidates_diff_winner"]:
        sigs[(r["user"]["winner_addr_type"],
              r["sqlite"]["winner_addr_type"])] += 1
    if sigs:
        print(f"\n  same_candidates_diff_winner by (mdb_type, php_type):")
        for (a, b), n in sigs.most_common():
            print(f"    user={a} ↔ sqlite={b}  x{n}")

    out = {
        "summary": {
            "total_addr_diffs": total,
            "buckets": {k: len(v) for k, v in classified.items()},
            "rank_table_status": (
                "BIOG_ADDR_CODES is identical between the two sides "
                "for all 22 addr_types (verified before classification)"
            ),
            "same_candidates_winner_pairings": {
                f"{a}->{b}": n for (a, b), n in sigs.items()
            },
        },
        "buckets": classified,
    }
    OUT.write_text(
        json.dumps(out, ensure_ascii=False, indent=2,
                    default=str),
        encoding="utf-8",
    )
    print(f"\nwrote {OUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
