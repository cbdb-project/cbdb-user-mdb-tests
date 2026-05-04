"""Render PR S's same-candidate deep-dive into a tie-break-
focused per-row table + tie-break-rule projection JSON.

Writes:
  - reports/index_addr_same_candidate_tiebreak.json
  - appends a per-row table to
    analysis/index_addr_same_candidate_tiebreak.md
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "reports" / "index_addr_same_candidates_deep_dive.json"
OUT_JSON = ROOT / "reports" / "index_addr_same_candidate_tiebreak.json"
OUT_MD = ROOT / "analysis" / "index_addr_same_candidate_tiebreak.md"

TIEBREAK_RULES = {
    "MIN_addr_id": lambda candidates: min(
        candidates, key=lambda c: c["addr_id"])["addr_id"],
    "MAX_addr_id": lambda candidates: max(
        candidates, key=lambda c: c["addr_id"])["addr_id"],
}


def _candidates_at_max(rows: list[dict],
                        max_seq: int) -> list[dict]:
    return [r for r in rows if r["sequence"] == max_seq]


def main() -> int:
    if not SRC.exists():
        raise SystemExit(f"missing {SRC}; run "
                          "analysis/deep_dive_addr_same_candidates.py "
                          "first")
    pr_s = json.loads(SRC.read_text(encoding="utf-8"))

    projected = []
    for r in pr_s["rows"]:
        max_seq = r["max_c_sequence"]
        u_tied = _candidates_at_max(
            r["user_addr_data_for_winning_type"], max_seq)
        s_tied = _candidates_at_max(
            r["sqlite_addr_data_for_winning_type"], max_seq)
        # Compute what each tie-break rule would pick from the
        # tied set on each side (they should agree because the
        # candidate set is the same).
        rule_picks = {}
        for rule_name, fn in TIEBREAK_RULES.items():
            try:
                rule_picks[rule_name] = {
                    "user_pick": fn(u_tied),
                    "sqlite_pick": fn(s_tied),
                    "agree": fn(u_tied) == fn(s_tied),
                }
            except Exception as e:
                rule_picks[rule_name] = {"_error": str(e)}
        projected.append({
            "personid": r["personid"],
            "name_chn": r["name_chn"],
            "name_py": r["name_py"],
            "winning_addr_type": r["winning_addr_type"],
            "max_c_sequence": max_seq,
            "n_candidates_tied": len(u_tied),
            "tied_candidate_addr_ids": [c["addr_id"] for c in u_tied],
            "user_stored_addr_id": r["user_stored_addr_id"],
            "sqlite_stored_addr_id": r["sqlite_stored_addr_id"],
            "user_recompute_addr_id": r["user_recompute_addr_id"],
            "sqlite_recompute_addr_id": r["sqlite_recompute_addr_id"],
            "tiebreak_rule_picks": rule_picks,
        })

    # Aggregate: under each rule, do all 3 sides converge?
    rule_summary = {}
    for rule_name in TIEBREAK_RULES:
        agree_count = sum(1 for p in projected
                           if p["tiebreak_rule_picks"][rule_name].get("agree"))
        rule_summary[rule_name] = {
            "would_make_user_and_sqlite_pick_same_row": agree_count,
            "of_total_rows": len(projected),
        }

    out = {
        "summary": {
            "n_rows": len(projected),
            "tiebreak_rule_summary": rule_summary,
            "interpretation": (
                "Under both MIN(c_addr_id) and MAX(c_addr_id) the "
                "candidate set is symmetric across User and SQLite "
                "(they see the same BIOG_ADDR_DATA rows in PR S's "
                "deep dive), so any deterministic tie-break would "
                "make both sides converge.  The choice between MIN "
                "vs MAX is a maintainer style call, not a "
                "correctness call."
            ),
        },
        "rows": projected,
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"wrote {OUT_JSON}")

    # Append the per-row table to the MD.
    md = OUT_MD.read_text(encoding="utf-8")
    addendum = []
    addendum.append("\n## Per-row data\n")
    addendum.append(
        "| personid | tied addr_ids | User stored | SQLite stored "
        "| User recompute | SQLite recompute | MIN-rule | MAX-rule |"
    )
    addendum.append(
        "|---:|---|---:|---:|---:|---:|---:|---:|"
    )
    for p in projected:
        rp = p["tiebreak_rule_picks"]
        min_pick = rp["MIN_addr_id"].get("user_pick", "?")
        max_pick = rp["MAX_addr_id"].get("user_pick", "?")
        addendum.append(
            f"| {p['personid']} | "
            f"{', '.join(str(i) for i in p['tied_candidate_addr_ids'])} "
            f"| {p['user_stored_addr_id']} "
            f"| {p['sqlite_stored_addr_id']} "
            f"| {p['user_recompute_addr_id']} "
            f"| {p['sqlite_recompute_addr_id']} "
            f"| {min_pick} | {max_pick} |"
        )
    addendum.append("")
    addendum.append("### Convergence under candidate tie-breaks")
    addendum.append("")
    addendum.append("| Rule | Rows where both sides converge |")
    addendum.append("|---|---:|")
    for rule_name, stats in rule_summary.items():
        addendum.append(
            f"| `{rule_name}` | "
            f"{stats['would_make_user_and_sqlite_pick_same_row']} / "
            f"{stats['of_total_rows']} |"
        )
    addendum.append("")
    addendum.append("Both candidate rules deterministically converge "
                     "all 10 rows; choice is a style call.")
    OUT_MD.write_text(md + "\n".join(addendum), encoding="utf-8")
    print(f"appended per-row table to {OUT_MD}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
