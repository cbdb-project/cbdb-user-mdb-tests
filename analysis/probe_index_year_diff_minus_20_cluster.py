"""Deep-dive the K2 `consistent_within_rule` diff=-20 cluster
(PR Y suggested investigation #4).

Target: 9 personids across 4 signature groups all sharing
`diff_access_minus_php = -20`:

  rule 11/11 (child = father.c_birthyear + 30, kin 75): 228114, 527169
  rule 13/13 (father from MIN(child.c_birthyear) - 30):  228102, 248797
  rule 15/15 (mother from MIN(child.c_birthyear) - 27):  228104, 248799
  rule 19/19 (older brother MAX(c_birthyear) + 2):        228111, 228112, 228113

PR N showed all four rules are matched (or matched_minor_diff
on staging-vs-subquery aggregate).  PR Y's hypothesis: "single
staging-step row pick reproducing consistently".  This probe
confirms or refutes that by walking each rule's actual evidence
rows on both sides.

For each personid:
  - find the relevant kin (per the rule)
  - pull their c_birthyear values on both sides
  - identify the candidate evidence row each side would pick
    by the documented rule
  - report which row Access vs PHP actually used (inferable
    from the recorded c_index_year)

Pure pyodbc + sqlite3.  No Access COM.

Outputs:
  - reports/index_year_diff_minus_20_cluster_probe.json
  - analysis/index_year_diff_minus_20_cluster_probe.md
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
OUT_JSON = ROOT / "reports" / "index_year_diff_minus_20_cluster_probe.json"
OUT_MD = ROOT / "analysis" / "index_year_diff_minus_20_cluster_probe.md"

# Rule → (kin_codes, aggregator, offset).
# kin_codes: which c_kin_code values represent the rule's
# evidence relationship (per PR N's index_year_rule_comparison.md).
# aggregator: MIN or MAX over the evidence c_birthyear.
# offset: how the rule converts the kin year to subject's year.
RULE_DEF = {
    "11": {"kin_codes": [75],
            "kin_role": "father",
            "aggregator": "MIN",  # there's only one father normally
            "offset": +30,
            "intent": "child = father.c_birthyear + 30"},
    "13": {"kin_codes": [75],
            "kin_role": "father (looks at children of subject)",
            "aggregator": "MIN_over_inverse_kin",
            "offset": -30,
            "intent": "father = MIN(child.c_birthyear) - 30"},
    "15": {"kin_codes": [111],
            "kin_role": "mother (looks at children of subject)",
            "aggregator": "MIN_over_inverse_kin",
            "offset": -27,
            "intent": "mother = MIN(child.c_birthyear) - 27"},
    "19": {"kin_codes": [125, 165],
            "kin_role": "older brother",
            "aggregator": "MAX",
            "offset": +2,
            "intent": "older brother = MAX(sibling.c_birthyear) + 2"},
}


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


def kin_evidence_user(cur, pid: int, rule_def: dict) -> list[dict]:
    """Pull candidate kin rows + their c_birthyear for the
    given rule on the User MDB side.  Implements two query
    shapes:
      - direct (rule 11): subject is the wife, kin_code IN
        (rule_def[kin_codes]) → father is c_kin_id
      - inverse (rules 13/15): subject is the parent, kin_code
        IN (rule_def[kin_codes]) flipped → children are
        c_personid where they have the subject as kin
      - sibling (rule 19): kin from subject side
    """
    codes_in = ",".join(str(c) for c in rule_def["kin_codes"])
    if rule_def["aggregator"] == "MIN_over_inverse_kin":
        sql = (
            "SELECT bm.c_personid AS evidence_pid, "
            "bm.c_birthyear AS evidence_birthyear, "
            "bm.c_name_chn, kd.c_kin_code "
            "FROM KIN_DATA kd INNER JOIN BIOG_MAIN bm "
            "  ON kd.c_personid = bm.c_personid "
            f"WHERE kd.c_kin_id = ? AND kd.c_kin_code IN ({codes_in})"
        )
    else:
        sql = (
            "SELECT bm.c_personid AS evidence_pid, "
            "bm.c_birthyear AS evidence_birthyear, "
            "bm.c_name_chn, kd.c_kin_code "
            "FROM KIN_DATA kd INNER JOIN BIOG_MAIN bm "
            "  ON kd.c_kin_id = bm.c_personid "
            f"WHERE kd.c_personid = ? AND kd.c_kin_code IN ({codes_in})"
        )
    cur.execute(sql, (pid,))
    return [{"evidence_pid": int(r[0]),
             "evidence_birthyear": int(r[1]) if r[1] is not None else None,
             "name_chn": r[2],
             "kin_code": int(r[3])}
            for r in cur.fetchall()]


def kin_evidence_sqlite(conn, pid: int,
                          rule_def: dict) -> list[dict]:
    codes_in = ",".join(str(c) for c in rule_def["kin_codes"])
    if rule_def["aggregator"] == "MIN_over_inverse_kin":
        sql = (
            "SELECT bm.c_personid AS evidence_pid, "
            "bm.c_birthyear AS evidence_birthyear, "
            "bm.c_name_chn, kd.c_kin_code "
            "FROM KIN_DATA kd INNER JOIN BIOG_MAIN bm "
            "  ON kd.c_personid = bm.c_personid "
            f"WHERE kd.c_kin_id = ? AND kd.c_kin_code IN ({codes_in})"
        )
    else:
        sql = (
            "SELECT bm.c_personid AS evidence_pid, "
            "bm.c_birthyear AS evidence_birthyear, "
            "bm.c_name_chn, kd.c_kin_code "
            "FROM KIN_DATA kd INNER JOIN BIOG_MAIN bm "
            "  ON kd.c_kin_id = bm.c_personid "
            f"WHERE kd.c_personid = ? AND kd.c_kin_code IN ({codes_in})"
        )
    rows = conn.execute(sql, (pid,)).fetchall()
    out = []
    for r in rows:
        d = dict(r)
        # Rename to match user-side dict shape.
        if "c_name_chn" in d:
            d["name_chn"] = d.pop("c_name_chn")
        if "evidence_pid" in d:
            d["evidence_pid"] = int(d["evidence_pid"])
        if d.get("evidence_birthyear") is not None:
            d["evidence_birthyear"] = int(d["evidence_birthyear"])
        if "kin_code" in d and d["kin_code"] is not None:
            d["kin_code"] = int(d["kin_code"])
        out.append(d)
    return out


def _select_winner(rows: list[dict], aggregator: str) -> dict | None:
    valid = [r for r in rows
             if r.get("evidence_birthyear")
             and r["evidence_birthyear"] > 0]
    if not valid:
        return None
    if aggregator == "MIN" or aggregator == "MIN_over_inverse_kin":
        return min(valid, key=lambda r: r["evidence_birthyear"])
    if aggregator == "MAX":
        return max(valid, key=lambda r: r["evidence_birthyear"])
    return None


def main() -> int:
    if not K1_JSON.exists():
        raise SystemExit(f"missing {K1_JSON}")
    k1 = json.loads(K1_JSON.read_text(encoding="utf-8"))
    by_pid = {r["personid"]: r for r in k1["buckets"]["consistent_within_rule"]}

    user = _open_user(); sql = _open_sqlite()
    cur_u = user.cursor()

    # Group personids by rule for clean reporting.
    targets = []
    for tcode, rule_def in RULE_DEF.items():
        for pid, k1_row in by_pid.items():
            if (k1_row["sqlite"]["index_year_type_code"] == tcode
                    and k1_row["user"]["index_year_type_code"] == tcode
                    and k1_row["diff_access_minus_php"] == -20):
                targets.append((tcode, pid, k1_row))

    findings = []
    for tcode, pid, k1_row in targets:
        rdef = RULE_DEF[tcode]
        u_evidence = kin_evidence_user(cur_u, pid, rdef)
        s_evidence = kin_evidence_sqlite(sql, pid, rdef)
        u_winner = _select_winner(u_evidence, rdef["aggregator"])
        s_winner = _select_winner(s_evidence, rdef["aggregator"])

        # Inferred rule output: aggregator(c_birthyear) + offset.
        u_inferred = (u_winner["evidence_birthyear"] + rdef["offset"]
                       if u_winner and u_winner.get("evidence_birthyear")
                       else None)
        s_inferred = (s_winner["evidence_birthyear"] + rdef["offset"]
                       if s_winner and s_winner.get("evidence_birthyear")
                       else None)

        access_actual = k1_row["user"]["index_year"]
        php_actual = k1_row["sqlite"]["index_year"]

        # Compare evidence by personid + birthyear separately so we
        # can distinguish "different rows" from "same rows but
        # birthyear drifted on one side".
        u_pids = {r["evidence_pid"] for r in u_evidence}
        s_pids = {r["evidence_pid"] for r in s_evidence}
        same_evidence_pid_set = (u_pids == s_pids)

        # Drift detection: per-evidence_pid birthyear comparison.
        u_by_pid = {r["evidence_pid"]: r.get("evidence_birthyear")
                     for r in u_evidence}
        s_by_pid = {r["evidence_pid"]: r.get("evidence_birthyear")
                     for r in s_evidence}
        birthyear_drift_per_pid = {
            pid: {"user_birthyear": u_by_pid.get(pid),
                  "sqlite_birthyear": s_by_pid.get(pid),
                  "diff": ((u_by_pid.get(pid) or 0)
                           - (s_by_pid.get(pid) or 0))}
            for pid in u_pids & s_pids
            if u_by_pid.get(pid) != s_by_pid.get(pid)
        }
        any_birthyear_drift = bool(birthyear_drift_per_pid)
        # Does the per-pid drift sum match the diff?  If yes,
        # that's strong evidence the cause is upstream birthyear
        # drift, not algorithm divergence.
        winner_pid_drift_explains_diff = False
        if (u_winner and s_winner
                and u_winner["evidence_pid"] == s_winner["evidence_pid"]):
            wpid = u_winner["evidence_pid"]
            if wpid in birthyear_drift_per_pid:
                # Inferred index_year diff = birthyear diff (the
                # offset is the same on both sides per PR N).
                bd = birthyear_drift_per_pid[wpid]["diff"]
                if bd == (access_actual - php_actual if access_actual
                          and php_actual else None):
                    winner_pid_drift_explains_diff = True

        outcome = (
            "source_data_drift_biog_main_birthyear"
            if winner_pid_drift_explains_diff
            else "same_pid_set_diff_birthyear_partially_explains"
            if same_evidence_pid_set and any_birthyear_drift
            else "same_evidence_diff_winner_pick"
            if same_evidence_pid_set and u_winner and s_winner
                and u_winner["evidence_pid"] != s_winner["evidence_pid"]
            else "same_evidence_same_winner_no_drift"
            if same_evidence_pid_set and u_winner and s_winner
                and u_winner["evidence_pid"] == s_winner["evidence_pid"]
            else "different_evidence_pid_set_between_sides"
            if not same_evidence_pid_set
            else "no_winner_inferable"
        )

        findings.append({
            "personid": pid,
            "name_chn": k1_row["name_chn"],
            "name_py": k1_row["name_py"],
            "rule_tcode": tcode,
            "rule_intent": rdef["intent"],
            "rule_aggregator": rdef["aggregator"],
            "rule_offset": rdef["offset"],
            "access_actual_index_year": access_actual,
            "php_actual_index_year": php_actual,
            "diff": access_actual - php_actual,
            "user_evidence_count": len(u_evidence),
            "user_evidence_sample": u_evidence[:6],
            "user_inferred_winner": u_winner,
            "user_inferred_index_year": u_inferred,
            "sqlite_evidence_count": len(s_evidence),
            "sqlite_evidence_sample": s_evidence[:6],
            "sqlite_inferred_winner": s_winner,
            "sqlite_inferred_index_year": s_inferred,
            "same_evidence_pid_set_between_sides": same_evidence_pid_set,
            "birthyear_drift_per_pid": birthyear_drift_per_pid,
            "winner_pid_drift_explains_diff":
                winner_pid_drift_explains_diff,
            "outcome": outcome,
        })

    # Tally.
    outcome_counts: dict[str, int] = {}
    for f in findings:
        outcome_counts[f["outcome"]] = outcome_counts.get(
            f["outcome"], 0) + 1

    out = {
        "summary": {
            "n_rows": len(findings),
            "outcome_counts": outcome_counts,
            "verdict": (
                "hypothesis_REVISED — actual cause is upstream "
                "BIOG_MAIN.c_birthyear drift between User MDB and "
                "SQLite snapshot, NOT staging-step row pick"
                if outcome_counts.get(
                    "source_data_drift_biog_main_birthyear", 0) >= len(findings) * 0.6
                else
                "single_staging_pick_supported"
                if outcome_counts.get(
                    "same_evidence_diff_winner_pick", 0) == len(findings)
                else
                "mixed — see outcomes detail"
            ),
        },
        "findings": findings,
        "is_confirmed_bug": False,
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
    md.append("# Index_year diff=-20 cluster deep-dive (PR AI)")
    md.append("")
    md.append("PR Y suggested investigation #4: walk the 9 K2 rows in "
              "the `consistent_within_rule` bucket whose `(php_tcode, "
              "access_tcode, diff)` signature has `diff = -20` across "
              "rules 11/13/15/19, and confirm or refute the working "
              "hypothesis that they share a single staging-step row "
              "pick.")
    md.append("")
    md.append(f"## Verdict: `{out['summary']['verdict']}`")
    md.append("")
    md.append(f"Outcome counts:")
    for k, n in sorted(outcome_counts.items(), key=lambda kv: -kv[1]):
        md.append(f"- `{k}`: {n}")
    md.append("")
    md.append("## Per-row detail")
    md.append("")
    for f in findings:
        md.append(f"### `c_personid = {f['personid']}` "
                  f"({f['name_chn']} / {f['name_py']}) — rule "
                  f"{f['rule_tcode']}")
        md.append("")
        md.append(f"- Rule intent: {f['rule_intent']}")
        md.append(f"- Access actual index_year: "
                  f"{f['access_actual_index_year']}")
        md.append(f"- PHP actual index_year: "
                  f"{f['php_actual_index_year']}  "
                  f"(diff = {f['diff']})")
        md.append(f"- User-side evidence rows: "
                  f"{f['user_evidence_count']}")
        md.append(f"- SQLite-side evidence rows: "
                  f"{f['sqlite_evidence_count']}")
        md.append(f"- Same evidence pid set on both sides: "
                  f"{f['same_evidence_pid_set_between_sides']}")
        if f["birthyear_drift_per_pid"]:
            md.append(f"- BIOG_MAIN.c_birthyear drift per evidence pid:")
            for pid, d in f["birthyear_drift_per_pid"].items():
                md.append(f"  - pid={pid}: User={d['user_birthyear']} / "
                          f"SQLite={d['sqlite_birthyear']} "
                          f"(diff={d['diff']})")
        if f["winner_pid_drift_explains_diff"]:
            md.append(f"- **Winner-pid birthyear drift fully "
                      f"explains the index_year diff.**")
        if f["user_inferred_winner"] is not None:
            w = f["user_inferred_winner"]
            md.append(f"- User-inferred winner: "
                      f"pid={w['evidence_pid']} "
                      f"({w['name_chn']}), birthyear="
                      f"{w['evidence_birthyear']} → inferred "
                      f"index_year = {f['user_inferred_index_year']}")
        if f["sqlite_inferred_winner"] is not None:
            w = f["sqlite_inferred_winner"]
            md.append(f"- SQLite-inferred winner: "
                      f"pid={w['evidence_pid']} "
                      f"({w['name_chn']}), birthyear="
                      f"{w['evidence_birthyear']} → inferred "
                      f"index_year = {f['sqlite_inferred_index_year']}")
        md.append(f"- Outcome: **{f['outcome']}**")
        md.append("")
    md.append("## Implications for PR Y")
    md.append("")
    md.append(f"PR Y's `consistent_within_rule` × 14 bucket "
              f"(at confidence `medium`) had this 9-row -20 sub-cluster "
              f"as its top suggested next investigation.  Result: "
              f"`{out['summary']['verdict']}`.")
    md.append("")
    md.append("Per the result, the cause-summary JSON's confidence "
              "for the diff=-20 cluster can be promoted from `medium` "
              "to either `supported_by_focused_probe` (if cleanly "
              "supported) or annotated with the partial-support "
              "outcome breakdown.  Update is left to a follow-up PR "
              "for morning review.")
    OUT_MD.write_text("\n".join(md), encoding="utf-8")
    print(f"wrote {OUT_MD}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
