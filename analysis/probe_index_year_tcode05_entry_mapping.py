"""Probe Y's "cleanest next investigation" — does the
`access_tcode='05'` × 7 cluster in `php_did_not_compute` really
arise from a missing entry-code → '040101' mapping in PHP's
`ENTRY_CODE_TYPE_REL`?

Background
----------
PR Y's cause analysis flagged this group as
**candidate_php_entry_code_mapping_gap** (medium-high
confidence): runtime Access fires Rule 05 (`c_year - 30` for
`c_entry_type='040101'` per PR N) and writes a real
c_index_year for these 7 personids; PHP's
`sqlEntryRule('040101', 30, '05')` returns 0 / NULL.  The
working hypothesis is that PHP's `ENTRY_CODE_TYPE_REL` doesn't
classify these persons' `c_entry_code` into `'040101'` on the
SQLite snapshot side.

This probe pulls each of the 7 personids and answers:

  1. What `c_entry_code` rows does ENTRY_DATA carry on the User
     MDB side?
  2. Are those `c_entry_code` values mapped to `c_entry_type =
     '040101'` on the User MDB side?
  3. Same questions on the SQLite snapshot side.
  4. Does PHP's rule shape (`c_year - 30`) reconstruct the
     Access-stored `c_index_year` from the joined ENTRY_DATA
     row?

Outcome decisions
-----------------
- If User MDB DOES have an `ENTRY_CODE_TYPE_REL` row mapping
  the person's `c_entry_code` to `'040101'` AND SQLite does
  NOT — that's the **mapping gap confirmed** path.  PR Y's
  confidence on this group can be promoted from medium-high
  → **supported by focused probe**.
- If both sides have the mapping and PHP still wrote NULL —
  the cause is elsewhere (e.g. PHP rule predicate beyond the
  type-code join, or evidence-row pick).  Hypothesis is **not
  supported**; PR Y's bucket label needs revision.
- If neither side has the mapping — Access shouldn't have
  fired Rule 05 either; the cause is upstream and the bucket
  label needs revision.

This probe is pure pyodbc + sqlite3.  No Access COM.

Outputs:
  - reports/index_year_tcode05_entry_mapping_probe.json
  - analysis/index_year_tcode05_entry_mapping_probe.md (brief)
"""
from __future__ import annotations

import json
import sqlite3
import sys
from pathlib import Path

import pyodbc

ROOT = Path(__file__).resolve().parent.parent
USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
SQLITE_DIR = ROOT / "data" / "cbdb_online_sqlite"
K1_JSON = ROOT / "reports" / "index_year_drift_rule_classification.json"
OUT_JSON = ROOT / "reports" / "index_year_tcode05_entry_mapping_probe.json"
OUT_MD = ROOT / "analysis" / "index_year_tcode05_entry_mapping_probe.md"

PHP_TARGET_TYPE = "040101"  # the type sqlEntryRule('040101', 30, '05') keys on
PHP_OFFSET = 30             # the -N offset


def _open_user() -> pyodbc.Connection:
    return pyodbc.connect(
        "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={USER_MDB};", autocommit=True, readonly=True)


def _open_sqlite() -> sqlite3.Connection:
    paths = sorted(SQLITE_DIR.glob("*.sqlite3"))
    if not paths:
        raise SystemExit(f"no sqlite snapshot in {SQLITE_DIR}")
    conn = sqlite3.connect(str(paths[-1]))
    conn.row_factory = sqlite3.Row
    return conn


def _user_entry_data(cur, pid: int) -> list[dict]:
    cur.execute(
        "SELECT c_entry_code, c_year, c_sequence "
        "FROM ENTRY_DATA WHERE c_personid = ?", (pid,))
    return [{"c_entry_code": int(r[0]) if r[0] is not None else None,
             "c_year": int(r[1]) if r[1] is not None else None,
             "c_sequence": int(r[2]) if r[2] is not None else None}
            for r in cur.fetchall()]


def _user_entry_type_map(cur, codes: list[int]) -> dict[int, list[str]]:
    if not codes:
        return {}
    in_clause = ",".join(str(c) for c in codes if c is not None)
    if not in_clause:
        return {}
    cur.execute(
        f"SELECT c_entry_code, c_entry_type FROM ENTRY_CODE_TYPE_REL "
        f"WHERE c_entry_code IN ({in_clause})")
    out: dict[int, list[str]] = {}
    for r in cur.fetchall():
        out.setdefault(int(r[0]), []).append(r[1])
    return out


def _sqlite_entry_data(conn: sqlite3.Connection, pid: int) -> list[dict]:
    rows = conn.execute(
        "SELECT c_entry_code, c_year, c_sequence "
        "FROM ENTRY_DATA WHERE c_personid = ?", (pid,)).fetchall()
    return [dict(r) for r in rows]


def _sqlite_entry_type_map(conn: sqlite3.Connection,
                            codes: list[int]) -> dict[int, list[str]]:
    if not codes:
        return {}
    in_clause = ",".join(str(c) for c in codes if c is not None)
    if not in_clause:
        return {}
    rows = conn.execute(
        f"SELECT c_entry_code, c_entry_type FROM ENTRY_CODE_TYPE_REL "
        f"WHERE c_entry_code IN ({in_clause})").fetchall()
    out: dict[int, list[str]] = {}
    for r in rows:
        out.setdefault(int(r["c_entry_code"]), []).append(r["c_entry_type"])
    return out


def _classify_row(user_entries: list[dict],
                   user_map: dict[int, list[str]],
                   sqlite_entries: list[dict],
                   sqlite_map: dict[int, list[str]],
                   access_year: int | None) -> dict:
    """For one personid, decide which outcome the row supports."""
    user_has_target = any(
        PHP_TARGET_TYPE in user_map.get(e["c_entry_code"], [])
        for e in user_entries
    )
    sqlite_has_target = any(
        PHP_TARGET_TYPE in sqlite_map.get(e["c_entry_code"], [])
        for e in sqlite_entries
    )
    # Reconstruct Access via PHP-shape: any user-side entry whose
    # c_entry_code maps to '040101' AND c_year - 30 == access_year?
    reconstructs_access_via_php_shape = False
    reconstructing_entry = None
    for e in user_entries:
        if PHP_TARGET_TYPE in user_map.get(e["c_entry_code"], []) \
                and e["c_year"] is not None and access_year is not None \
                and e["c_year"] - PHP_OFFSET == access_year:
            reconstructs_access_via_php_shape = True
            reconstructing_entry = e
            break

    # Decision tree.
    if user_has_target and not sqlite_has_target:
        outcome = "mapping_gap_confirmed_user_has_target_sqlite_missing"
    elif user_has_target and sqlite_has_target:
        outcome = "both_sides_have_target_cause_elsewhere"
    elif not user_has_target and not sqlite_has_target:
        outcome = "neither_side_has_target_access_path_unexplained"
    else:
        outcome = "sqlite_has_target_user_does_not_unexpected"

    return {
        "user_has_target_040101_mapping": user_has_target,
        "sqlite_has_target_040101_mapping": sqlite_has_target,
        "user_reconstructs_access_via_php_shape":
            reconstructs_access_via_php_shape,
        "user_reconstructing_entry": reconstructing_entry,
        "outcome": outcome,
    }


def main() -> int:
    if not K1_JSON.exists():
        raise SystemExit(
            f"missing {K1_JSON}; run "
            f"`python analysis/classify_index_year_drift_by_rule.py` first")
    k1 = json.loads(K1_JSON.read_text(encoding="utf-8"))
    rows = [r for r in k1["buckets"]["php_did_not_compute"]
            if r["user"]["index_year_type_code"] == "05"]
    print(f"loaded {len(rows)} tcode='05' rows from "
          f"php_did_not_compute")

    user = _open_user(); sql = _open_sqlite()
    cur_u = user.cursor()

    findings = []
    for r in rows:
        pid = r["personid"]
        access_year = r["user"]["index_year"]
        u_entries = _user_entry_data(cur_u, pid)
        u_map = _user_entry_type_map(
            cur_u, [e["c_entry_code"] for e in u_entries])
        s_entries = _sqlite_entry_data(sql, pid)
        s_map = _sqlite_entry_type_map(
            sql, [e["c_entry_code"] for e in s_entries])
        cls = _classify_row(u_entries, u_map, s_entries, s_map,
                             access_year)
        findings.append({
            "personid": pid,
            "name_chn": r["name_chn"],
            "name_py": r["name_py"],
            "access_index_year": access_year,
            "php_index_year": r["sqlite"]["index_year"],
            "user_entry_data": u_entries,
            "user_entry_type_map": {
                str(k): v for k, v in u_map.items()},
            "sqlite_entry_data": s_entries,
            "sqlite_entry_type_map": {
                str(k): v for k, v in s_map.items()},
            **cls,
        })

    # Tally outcomes.
    outcome_counts: dict[str, int] = {}
    for f in findings:
        outcome_counts[f["outcome"]] = outcome_counts.get(
            f["outcome"], 0) + 1

    n_mapping_gap = outcome_counts.get(
        "mapping_gap_confirmed_user_has_target_sqlite_missing", 0)
    n_both = outcome_counts.get(
        "both_sides_have_target_cause_elsewhere", 0)
    n_neither = outcome_counts.get(
        "neither_side_has_target_access_path_unexplained", 0)
    n_unexpected = outcome_counts.get(
        "sqlite_has_target_user_does_not_unexpected", 0)

    if n_mapping_gap == len(findings):
        verdict = "hypothesis_fully_supported"
        verdict_note = (
            f"All {len(findings)} rows have the '040101' mapping on "
            f"the User MDB side and lack it on the SQLite side.  "
            f"PR Y's `candidate_php_entry_code_mapping_gap` label is "
            f"directly evidenced; confidence promotable to "
            f"`supported_by_focused_probe`."
        )
    elif n_mapping_gap >= len(findings) * 0.6:
        verdict = "hypothesis_mostly_supported"
        verdict_note = (
            f"{n_mapping_gap}/{len(findings)} rows fit the mapping-gap "
            f"hypothesis cleanly.  Remaining "
            f"{len(findings) - n_mapping_gap} need separate triage.  "
            f"PR Y's confidence stays at medium-high; the bucket may "
            f"split."
        )
    elif n_neither >= 1:
        verdict = "hypothesis_partially_undermined"
        verdict_note = (
            f"{n_neither}/{len(findings)} rows have NO '040101' "
            f"mapping on either side, which means Access's Rule 05 "
            f"shouldn't have fired either.  Cause is upstream of the "
            f"PHP join.  PR Y's bucket label needs revision for "
            f"those rows."
        )
    else:
        verdict = "hypothesis_not_supported"
        verdict_note = (
            f"{n_both}/{len(findings)} rows have '040101' mapping on "
            f"both sides, so PHP should fire Rule 05.  The reason "
            f"PHP wrote NULL is elsewhere (e.g. PHP rule predicate, "
            f"evidence-row selection).  PR Y's bucket label needs "
            f"revision."
        )

    out = {
        "summary": {
            "n_rows": len(findings),
            "outcome_counts": outcome_counts,
            "verdict": verdict,
            "verdict_note": verdict_note,
            "php_target_type": PHP_TARGET_TYPE,
            "php_offset": PHP_OFFSET,
        },
        "findings": findings,
        "is_confirmed_bug": False,
    }
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"wrote {OUT_JSON}")
    print(f"  outcome counts: {outcome_counts}")
    print(f"  verdict: {verdict}")
    print(f"  note: {verdict_note}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
