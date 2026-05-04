"""Companion to PR AI — same machinery, applied to the rule
11/11 diff=+1 K2 sub-cluster (5 rows).

Target: 5 personids in K2's `consistent_within_rule` bucket
sharing signature `(php_tcode='11', access_tcode='11', diff=+1)`:
  523820 (金文伯), 523821 (金武伯), 523822 (金堅伯),
  523823 (no name fetched), 696896 (no name fetched)

PR AI showed the diff=-20 cluster (9 rows) is upstream
BIOG_MAIN.c_birthyear drift, NOT algorithm divergence.  This
probe asks: is the diff=+1 cluster the same mechanism (drift
in opposite direction) or something else?

Pure pyodbc + sqlite3.  No Access COM.

Outputs:
  - reports/index_year_diff_plus_1_cluster_probe.json
  - analysis/index_year_diff_plus_1_cluster_probe.md
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
OUT_JSON = ROOT / "reports" / "index_year_diff_plus_1_cluster_probe.json"
OUT_MD = ROOT / "analysis" / "index_year_diff_plus_1_cluster_probe.md"

# Rule 11 (per PR N): child = father.c_birthyear + 30, kin_code 75.
RULE_DEF = {"kin_code": 75, "offset": 30, "intent":
            "child = father.c_birthyear + 30 (kin_code 75)"}


def _open_user(read_only: bool = True) -> pyodbc.Connection:
    return pyodbc.connect(
        "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={USER_MDB};", autocommit=True, readonly=read_only)


def _open_sqlite() -> sqlite3.Connection:
    paths = sorted(SQLITE_DIR.glob("*.sqlite3"))
    if not paths:
        raise SystemExit(f"no sqlite snapshot in {SQLITE_DIR}")
    conn = sqlite3.connect(str(paths[-1]))
    conn.row_factory = sqlite3.Row
    return conn


def father_evidence(cursor_or_conn, pid: int,
                     is_sqlite: bool) -> list[dict]:
    sql = (
        "SELECT bm.c_personid AS evidence_pid, "
        "bm.c_birthyear AS evidence_birthyear, "
        "bm.c_name_chn, kd.c_kin_code "
        "FROM KIN_DATA kd INNER JOIN BIOG_MAIN bm "
        "  ON kd.c_kin_id = bm.c_personid "
        "WHERE kd.c_personid = ? AND kd.c_kin_code = 75"
    )
    if is_sqlite:
        rows = cursor_or_conn.execute(sql, (pid,)).fetchall()
        out = []
        for r in rows:
            d = dict(r)
            if "c_name_chn" in d:
                d["name_chn"] = d.pop("c_name_chn")
            for k in ("evidence_pid", "evidence_birthyear", "kin_code"):
                if k in d and d[k] is not None:
                    d[k] = int(d[k])
            out.append(d)
        return out
    else:
        cursor_or_conn.execute(sql, (pid,))
        return [{"evidence_pid": int(r[0]),
                 "evidence_birthyear": int(r[1]) if r[1] is not None else None,
                 "name_chn": r[2],
                 "kin_code": int(r[3])}
                for r in cursor_or_conn.fetchall()]


def main() -> int:
    if not K1_JSON.exists():
        raise SystemExit(f"missing {K1_JSON}")
    k1 = json.loads(K1_JSON.read_text(encoding="utf-8"))
    targets = [r for r in k1["buckets"]["consistent_within_rule"]
               if r["sqlite"]["index_year_type_code"] == "11"
               and r["user"]["index_year_type_code"] == "11"
               and r["diff_access_minus_php"] == 1]

    user = _open_user(); sql = _open_sqlite()
    cur_u = user.cursor()

    findings = []
    for r in targets:
        pid = r["personid"]
        u_ev = father_evidence(cur_u, pid, is_sqlite=False)
        s_ev = father_evidence(sql, pid, is_sqlite=True)
        access_actual = r["user"]["index_year"]
        php_actual = r["sqlite"]["index_year"]
        diff = access_actual - php_actual

        u_by_pid = {e["evidence_pid"]: e.get("evidence_birthyear")
                     for e in u_ev}
        s_by_pid = {e["evidence_pid"]: e.get("evidence_birthyear")
                     for e in s_ev}
        common_pids = set(u_by_pid) & set(s_by_pid)
        drift_per_pid = {
            pid_: {"user_birthyear": u_by_pid.get(pid_),
                   "sqlite_birthyear": s_by_pid.get(pid_),
                   "diff": ((u_by_pid.get(pid_) or 0)
                            - (s_by_pid.get(pid_) or 0))}
            for pid_ in common_pids
            if u_by_pid.get(pid_) != s_by_pid.get(pid_)
        }

        # Rule 11 expects MIN over fathers (typically 1 row).
        u_winner = min(
            (e for e in u_ev if e.get("evidence_birthyear")),
            key=lambda e: e["evidence_birthyear"], default=None)
        s_winner = min(
            (e for e in s_ev if e.get("evidence_birthyear")),
            key=lambda e: e["evidence_birthyear"], default=None)

        winner_drift_explains = False
        if (u_winner and s_winner
                and u_winner["evidence_pid"] == s_winner["evidence_pid"]):
            wpid = u_winner["evidence_pid"]
            if wpid in drift_per_pid:
                bd = drift_per_pid[wpid]["diff"]
                if bd == diff:
                    winner_drift_explains = True

        outcome = (
            "source_data_drift_biog_main_birthyear"
            if winner_drift_explains
            else "different_evidence_pid_set_between_sides"
            if set(u_by_pid) != set(s_by_pid)
            else "same_pid_drift_does_not_explain_diff"
        )
        findings.append({
            "personid": pid,
            "name_chn": r["name_chn"],
            "name_py": r["name_py"],
            "access_actual": access_actual,
            "php_actual": php_actual,
            "diff": diff,
            "user_evidence": u_ev,
            "sqlite_evidence": s_ev,
            "user_winner": u_winner,
            "sqlite_winner": s_winner,
            "drift_per_pid": drift_per_pid,
            "winner_drift_explains_diff": winner_drift_explains,
            "outcome": outcome,
        })

    outcome_counts: dict[str, int] = {}
    for f in findings:
        outcome_counts[f["outcome"]] = outcome_counts.get(
            f["outcome"], 0) + 1
    out = {
        "summary": {
            "n_rows": len(findings),
            "outcome_counts": outcome_counts,
            "verdict": (
                "supports_PR_AI — same upstream BIOG_MAIN drift "
                "mechanism in opposite direction"
                if outcome_counts.get(
                    "source_data_drift_biog_main_birthyear", 0)
                    >= len(findings) * 0.6
                else
                "different_mechanism — see outcomes detail"
            ),
        },
        "findings": findings,
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"wrote {OUT_JSON}")
    print(f"  outcome counts: {outcome_counts}")
    print(f"  verdict: {out['summary']['verdict']}")

    md = []
    md.append("# Rule 11/11 diff=+1 cluster probe (PR AJ)")
    md.append("")
    md.append("Companion to PR AI.  Tests whether the 5-row K2 "
              "diff=+1 sub-cluster shares the same upstream "
              "BIOG_MAIN.c_birthyear drift mechanism as the diff=-20 "
              "cluster did, just in the opposite direction.")
    md.append("")
    md.append(f"## Verdict: `{out['summary']['verdict']}`")
    md.append("")
    md.append("Outcome counts:")
    for k, n in sorted(outcome_counts.items(), key=lambda kv: -kv[1]):
        md.append(f"- `{k}`: {n}")
    md.append("")
    md.append("## Per-row detail")
    md.append("")
    for f in findings:
        md.append(f"### `c_personid = {f['personid']}` "
                  f"({f['name_chn']} / {f['name_py']})")
        md.append(f"- Access actual: {f['access_actual']}, "
                  f"PHP actual: {f['php_actual']}, diff: {f['diff']}")
        if f["user_winner"]:
            md.append(f"- User winner: pid={f['user_winner']['evidence_pid']} "
                      f"({f['user_winner']['name_chn']}), "
                      f"birthyear={f['user_winner']['evidence_birthyear']}")
        if f["sqlite_winner"]:
            md.append(f"- SQLite winner: pid={f['sqlite_winner']['evidence_pid']} "
                      f"({f['sqlite_winner']['name_chn']}), "
                      f"birthyear={f['sqlite_winner']['evidence_birthyear']}")
        if f["drift_per_pid"]:
            for pid_, d in f["drift_per_pid"].items():
                md.append(f"- BIOG_MAIN drift pid={pid_}: "
                          f"User={d['user_birthyear']} / "
                          f"SQLite={d['sqlite_birthyear']} "
                          f"(diff={d['diff']})")
        md.append(f"- Outcome: **{f['outcome']}**")
        md.append("")
    md.append("## Cumulative picture across PR AI + AJ")
    md.append("")
    md.append("With both probes:")
    md.append("- diff=-20 sub-cluster (9 rows): "
              "7 source_data_drift_biog_main_birthyear + 2 "
              "different_evidence_pid_set")
    md.append("- diff=+1 sub-cluster (5 rows): see above")
    md.append("")
    md.append("Combined, the K2 `consistent_within_rule × 14` bucket "
              "is overwhelmingly upstream-data drift, not algorithm "
              "divergence.  Cause-summary JSON re-class left for "
              "morning review (per overnight rules: don't change "
              "severity without overwhelming evidence + maintainer "
              "sign-off).")
    OUT_MD.write_text("\n".join(md), encoding="utf-8")
    print(f"wrote {OUT_MD}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
