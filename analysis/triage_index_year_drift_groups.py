"""Deep triage of the c_index_year drift buckets that PR K1 left
under-named.

Updated 2026-05-03 (PR X) — `PHP_RULE_BY_TCODE` below was
re-derived from PR N's `analysis/index_year_rule_comparison.md`
(runtime `GetBirthIndexYearSQL` paired against PHP
`IndexYearRebuildService.php` by emitted type_code).  Per PR N
runtime Access agrees with PHP on most type_codes (22 matched
/ 8 matched_minor_diff / 0 logic_diff / 3 access_only).  The
buckets this script outputs (consistent_within_rule,
iteration_order_diff, php_did_not_compute, etc.) are unchanged
— they're indexed by type_codes only — but each bucket's
rationale text now references PR N's runtime comparison
rather than PR I's vestigial-QueryDef comparison.

The blocker label `blocked_by_missing_frmBaseMaintenance_vba`
that appeared in earlier output is also stale: PR M dumped
`frmBaseMaintenance` via `SaveAsText` in 2026-05-03.  This
script now emits `blocked_by_runtime_priority_triage_pending`
for the same rows — the blocker is no longer "we don't have
the source", it's "we have the source but need to walk the
priority/iteration order to decide which side's choice was
intentional".

Reads:
  - reports/index_year_drift_rule_classification.json (PR K1 output)

Three things this script does:

  1. `consistent_within_rule` — group by (php_tcode,
     access_tcode, diff) signature, then label each group
     using PR N's rule mapping.  diff = -20 across rules 11 /
     13 / 15 / 19 is the standout pattern; per PR N rules 13
     / 15 / 19 are `matched_minor_diff` (staging vs subquery
     aggregate equivalent on the happy path), so the -20
     cluster needs a runtime-rule / phase-order triage to
     decide whether the staging step picks a consistently
     wrong row.

  2. `unclassified` — for each row, emit either a candidate
     cause or an explicit blocker label.  Blocker now means
     "needs runtime priority/iteration-order triage against
     `GetBirthIndexYearSQL`".

  3. `php_did_not_compute` — group by Access tcode and label
     each group with the PHP-side coverage gap suggested by
     the runtime rule.  Biggest group (`access_tcode='05'`)
     still reads as `candidate_php_entry_code_mapping_gap` —
     the row evidence supports PHP missing the entry-code →
     '040101' map for those persons.

Conservative: nothing here is labelled as a confirmed bug.
The output uses `candidate_*` / `blocked_by_*` labels.

Writes `reports/index_year_drift_rule_groups.json`.
"""
from __future__ import annotations

import json
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
K1_JSON = ROOT / "reports" / "index_year_drift_rule_classification.json"
OUT = ROOT / "reports" / "index_year_drift_rule_groups.json"


# Rule-name lookup, derived from PR N's
# `analysis/index_year_rule_comparison.md` (runtime
# GetBirthIndexYearSQL ↔ PHP IndexYearRebuildService pairing).
# Keyed by PHP type_code; value = (rule_intent, php_method).
PHP_RULE_BY_TCODE: dict[str, tuple[str, str]] = {
    "01": ("c_index_year = c_birthyear (raw birthyear)", "sqlRule01"),
    "02": ("c_index_year = c_deathyear - c_death_age + 1",
           "sqlRule02"),
    "03": ("wife from husband.c_index_year + 3 (kin 134)",
           "sqlRule03"),
    "11": ("child = father.c_birthyear + 30 (kin 75)", "sqlRule11"),
    "13": ("father from MIN(child.c_birthyear) - 30",
           "sqlAggregateRule13"),
    "15": ("mother from MIN(child.c_birthyear) - 27",
           "sqlAggregateRule15"),
    "17": ("wife from husband (concubine variant)", "sqlRule03"),
    "19": ("older brother MAX(c_birthyear) + 2 (kin 125,165)",
           "sqlAggregateSiblingRule"),
    "21": ("younger brother MIN(c_birthyear) - 2 (kin 126,166)",
           "sqlAggregateSiblingRule"),
    "23": ("son-in-law male target -27 (kin 181,201,224,332)",
           "sqlAggregateSonInLawRule"),
    "25": ("son-in-law female target -24",
           "sqlAggregateSonInLawRule"),
    "27": ("descendant from grandfather.c_birthyear + 60 (kin 62)",
           "sqlRule27"),
    "29": ("c_index_year = c_deathyear - 63 (male, no death_age)",
           "sqlRule29Or30"),
    "30": ("c_index_year = c_deathyear - 56 (female, no death_age)",
           "sqlRule29Or30"),
    "05": ("entry: jinshi - 30 (entry_type 040101)",
           "sqlEntryRule"),
    "06": ("wife from jinshi husband - 27", "sqlWifeFromEntryRule"),
    "07": ("entry: juren - 27 (entry_type 040102)",
           "sqlEntryRule"),
    "08": ("wife from juren husband - 24",
           "sqlWifeFromEntryRule"),
    "09": ("entry: 040103 - 21", "sqlEntryRule"),
    "10": ("wife from 040103 husband - 18",
           "sqlWifeFromEntryRule"),
    # Phase C loop tcodes (CONCAT'd onto parent's)
    "04": ("Phase C: husband propagation", "sqlLoopHusbandPropagationRule"),
    "12": ("Phase C: father propagation +30",
           "sqlLoopFatherPropagationRule"),
    "14": ("Phase C: father from MIN(child.c_index_year) - 30",
           "sqlLoopOldestChildIndexToFatherRule"),
    "16": ("Phase C: mother from MIN(child.c_index_year) - 27",
           "sqlLoopOldestChildIndexToMotherRule"),
    "18": ("Phase C: husband (concubine variant)",
           "sqlLoopHusbandPropagationRule"),
    "20": ("Phase C: older brother +2",
           "sqlLoopSiblingRule"),
    "22": ("Phase C: younger brother -2",
           "sqlLoopSiblingRule"),
    "24": ("Phase C: son-in-law male -27",
           "sqlLoopSonInLawRule"),
    "26": ("Phase C: son-in-law female -24",
           "sqlLoopSonInLawRule"),
    "28": ("Phase C: grandfather propagation +60",
           "sqlLoopGrandfatherPropagationRule"),
}


def _decompose_tcode(tcode: str) -> list[str]:
    """Split a CONCAT'd type_code (e.g. '1112') into its 2-char
    components ['11','12'].  Empty / odd-length strings return [tcode]
    unchanged so downstream lookups still work."""
    if not tcode or len(tcode) % 2 != 0:
        return [tcode] if tcode else []
    return [tcode[i:i+2] for i in range(0, len(tcode), 2)]


def _name_by_signature(php_t: str, access_t: str, diff: int) -> dict:
    """Best-effort name for a single (php_tcode, access_tcode, diff)
    triple.  Returns a dict with keys:
        label    - candidate_*  or  blocked_by_*  or  needs_manual_review
        rationale
    """
    php_steps = _decompose_tcode(php_t)
    access_steps = _decompose_tcode(access_t)

    # Iteration-order divergence (one is a prefix of the other)
    if (php_t and access_t and php_t != access_t
            and (access_t.startswith(php_t)
                 or php_t.startswith(access_t))):
        extra = (access_t[len(php_t):] if access_t.startswith(php_t)
                 else php_t[len(access_t):])
        return {
            "label": "candidate_phase_c_iteration_order",
            "rationale": (
                f"One tcode is a prefix of the other; the extra "
                f"suffix {extra!r} corresponds to "
                f"{PHP_RULE_BY_TCODE.get(extra, ('unknown rule', ''))[0]}.  "
                f"Whichever side ran the extra Phase-C step produced "
                f"the larger-magnitude index_year."
            ),
        }

    # Same tcode on both sides — different rule outcome from the
    # same rule.  Per PR N this rule is matched (or matched_minor_
    # diff) at the source level, so the divergence almost
    # certainly comes from row-pick (tie-break) or aggregation
    # detail rather than a logic divergence.
    if php_t == access_t and php_t:
        intent = PHP_RULE_BY_TCODE.get(php_t, ("unknown rule", ""))[0]
        return {
            "label": "candidate_same_rule_tie_break_or_aggregation_diff",
            "rationale": (
                f"Both sides claim type_code {php_t!r} ({intent}) "
                f"but produce different values (diff={diff}).  PR "
                f"N has this rule as matched (or matched_minor_"
                f"diff) at the source level, so the divergence is "
                f"most likely a tie-break / aggregation detail — "
                f"PHP and runtime Access can pick different "
                f"evidence rows when multiple candidates have the "
                f"same priority (NULL handling, MIN-vs-first-row "
                f"under different storage orders, staging-step "
                f"row pick, etc.).  diff sign tells which side "
                f"picked the higher value.  Needs runtime-rule / "
                f"phase-order triage against GetBirthIndexYearSQL "
                f"to decide whether the divergence is intentional."
            ),
        }

    # Different rules picked on each side — priority order or
    # evidence preference diverged at the rule-selection layer.
    if php_t and access_t and php_t != access_t:
        return {
            "label": "candidate_priority_order_or_rule_selection_diff",
            "rationale": (
                f"PHP ran rule {php_t!r} ({PHP_RULE_BY_TCODE.get(php_t,('?', ''))[0]}); "
                f"Access ran rule {access_t!r} ({PHP_RULE_BY_TCODE.get(access_t,('?', ''))[0]}).  "
                f"Different priority order or evidence preference.  "
                f"PR M dumped the runtime VBA "
                f"(GetBirthIndexYearSQL inside frmBaseMaintenance), "
                f"so the source is in repo, but resolving which "
                f"side's choice is intentional still needs a "
                f"per-row walk of the runtime priority/iteration "
                f"order.  Blocker: blocked_by_runtime_priority_"
                f"triage_pending."
            ),
        }

    return {
        "label": "needs_manual_review",
        "rationale": (
            f"php_tcode={php_t!r} access_tcode={access_t!r} diff={diff}; "
            f"no signature match."
        ),
    }


def main() -> int:
    if not K1_JSON.exists():
        raise SystemExit(
            f"missing {K1_JSON}; run "
            f"`python analysis/classify_index_year_drift_by_rule.py` first")
    k1 = json.loads(K1_JSON.read_text(encoding="utf-8"))

    # --- 1) consistent_within_rule grouping ---
    groups: dict[tuple[str, str, int], list[dict]] = defaultdict(list)
    for r in k1["buckets"]["consistent_within_rule"]:
        sig = (
            r["sqlite"]["index_year_type_code"],
            r["user"]["index_year_type_code"],
            r["diff_access_minus_php"],
        )
        groups[sig].append(r)
    consistent_groups = []
    for sig, rows in sorted(groups.items(), key=lambda kv: -len(kv[1])):
        php_t, access_t, diff = sig
        named = _name_by_signature(php_t, access_t, diff)
        consistent_groups.append({
            "signature": {
                "php_tcode": php_t,
                "access_tcode": access_t,
                "diff_access_minus_php": diff,
            },
            "size": len(rows),
            "named_as": named["label"],
            "rationale": named["rationale"],
            "examples": [
                {"personid": r["personid"],
                 "name_chn": r["name_chn"],
                 "name_py": r["name_py"]}
                for r in rows[:3]
            ],
            "all_personids": [r["personid"] for r in rows],
        })

    # --- 2) unclassified per-row diagnostics ---
    unclassified_diag = []
    for r in k1["buckets"]["unclassified"]:
        php_t = r["sqlite"]["index_year_type_code"]
        access_t = r["user"]["index_year_type_code"]
        diff = r["diff_access_minus_php"]
        named = _name_by_signature(php_t, access_t, diff)
        unclassified_diag.append({
            "personid": r["personid"],
            "name_chn": r["name_chn"],
            "name_py": r["name_py"],
            "php_tcode": php_t,
            "access_tcode": access_t,
            "diff_access_minus_php": diff,
            "named_as": named["label"],
            "rationale": named["rationale"],
            "blocker": (
                "blocked_by_runtime_priority_triage_pending"
                if named["label"] in (
                    "candidate_priority_order_or_rule_selection_diff",
                    "needs_manual_review",
                )
                else None
            ),
        })

    # --- 3) php_did_not_compute grouping by Access tcode ---
    by_access_tcode: dict[str, list[dict]] = defaultdict(list)
    for r in k1["buckets"]["php_did_not_compute"]:
        by_access_tcode[r["user"]["index_year_type_code"]].append(r)
    php_missing_groups = []
    for tcode, rows in sorted(by_access_tcode.items(),
                               key=lambda kv: -len(kv[1])):
        # Use PR N's intent string as the candidate "what PHP missed".
        intent_first = _decompose_tcode(tcode)[0] if tcode else ""
        intent = PHP_RULE_BY_TCODE.get(intent_first,
                                        ("unknown rule",
                                         "unknown method"))[0]
        # Per-tcode rationales aligned with PR N's runtime-vs-PHP
        # comparison.  Conservative: each is `candidate_*`; row
        # evidence cited where it specifically supports the gap.
        notes = {
            "05": (
                "Access wrote via runtime Rule 05 (entry c_year - "
                "30 with ENTRY_CODE_TYPE_REL.c_entry_type = "
                "'040101', per PR N matched against PHP "
                "sqlEntryRule).  For PHP to NOT fire, the SQLite "
                "snapshot's ENTRY_CODE_TYPE_REL must not map this "
                "person's c_entry_code to '040101'.  Row evidence "
                "in K1 supports this for the 7 affected rows.  "
                "candidate_php_entry_code_mapping_gap."
            ),
            "07": (
                "Access wrote via Rule 07 XC (writes to tmpBM_NIY "
                "with c_year+39, c_entry_code=257).  PHP has no "
                "obvious entry-code-257 path in PR N's runtime "
                "comparison.  Needs runtime-rule triage to confirm "
                "whether the Access path is intentional or a "
                "vestigial branch."
            ),
            "11": (
                "Access wrote via Rule 11 (child = father.c_birthyear "
                "+ 30).  PHP sqlRule11 has the same shape — for PHP "
                "to NOT fire, the father's c_birthyear must fail "
                "validYearExpr (c_birthyear=0 AND c_dy NOT IN special "
                "dynasty list [2,25,29,46,83]).  candidate_php_"
                "validyear_gate_excluded_father."
            ),
            "14": (
                "Access wrote via Phase C Rule 14 (father from "
                "MIN(child.c_index_year) - 30).  PHP "
                "sqlLoopOldestChildIndexToFatherRule has the same "
                "shape but only fires inside the Phase-C loop; "
                "could be an iteration-count difference."
            ),
            "20": (
                "Access wrote via Phase C Rule 20 (older brother +2).  "
                "PHP equivalent fires inside Phase-C loop; iteration "
                "count diff suspected."
            ),
            "2304": (
                "Access tcode CONCAT '23' + '04' suggests Phase A "
                "Rule 23 (son-in-law male) followed by Phase C Rule "
                "04 (husband propagation).  PHP would need both "
                "phases to reach a value; if its loop terminated "
                "earlier this is the PHP coverage gap."
            ),
        }.get(tcode, (
            f"Access wrote via type_code {tcode!r} ({intent}).  "
            f"PHP rule chain didn't reach a value — needs per-row "
            f"investigation against PHP rule for that tcode."
        ))
        php_missing_groups.append({
            "access_tcode": tcode,
            "size": len(rows),
            "candidate_named_as": (
                "candidate_php_rule_coverage_gap"
                if tcode != "07"
                else "candidate_missing_php_rule_07_xc"
            ),
            "rationale": notes,
            "examples": [
                {"personid": r["personid"],
                 "name_chn": r["name_chn"],
                 "name_py": r["name_py"]}
                for r in rows[:3]
            ],
            "all_personids": [r["personid"] for r in rows],
        })

    # ---- Aggregate "remaining truly unclassified" count ----
    remaining_unclassified = sum(
        1 for d in unclassified_diag if d["named_as"] == "needs_manual_review"
    )

    out = {
        "summary": {
            "k1_input": str(K1_JSON.relative_to(ROOT)),
            "consistent_within_rule": {
                "groups": len(consistent_groups),
                "rows": sum(g["size"] for g in consistent_groups),
            },
            "unclassified": {
                "total": len(unclassified_diag),
                "named_after_triage": sum(
                    1 for d in unclassified_diag
                    if d["named_as"] != "needs_manual_review"
                ),
                "still_unclassified": remaining_unclassified,
                "blocked_by_runtime_priority_triage_pending": sum(
                    1 for d in unclassified_diag
                    if d["blocker"]
                ),
            },
            "php_did_not_compute": {
                "groups": len(php_missing_groups),
                "rows": sum(g["size"] for g in php_missing_groups),
            },
        },
        "consistent_within_rule_groups": consistent_groups,
        "unclassified_diagnostics": unclassified_diag,
        "php_did_not_compute_groups": php_missing_groups,
    }

    OUT.write_text(
        json.dumps(out, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"wrote {OUT}")
    print()
    print(f"=== consistent_within_rule: {len(consistent_groups)} groups ===")
    for g in consistent_groups:
        print(f"  {g['signature']['php_tcode']!r:>6}/"
              f"{g['signature']['access_tcode']!r:<6} "
              f"diff={g['signature']['diff_access_minus_php']:>4}  "
              f"x{g['size']}  {g['named_as']}")
    print()
    print(f"=== unclassified: {len(unclassified_diag)} rows ===")
    print(f"  named after triage:        {len(unclassified_diag) - remaining_unclassified}")
    print(f"  still needs manual review: {remaining_unclassified}")
    print(f"  blocked_by_runtime_priority_triage_pending: {sum(1 for d in unclassified_diag if d['blocker'])}")
    print()
    print(f"=== php_did_not_compute: {len(php_missing_groups)} groups ===")
    for g in php_missing_groups:
        print(f"  access_tcode={g['access_tcode']!r:>10s}  x{g['size']}  "
              f"{g['candidate_named_as']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
