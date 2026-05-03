"""Compare Access BM IY Rule QueryDefs against
cbdb-online-main-server's PHP IndexYearRebuildService.

Reads:
  - analysis/dump_data/querydefs_index/*.sql  (extracted in PR H from
    data/CBDB_<YYYYMMDD>_DATA.mdb)
  - analysis/php_source/IndexYearRebuildService.php  (pinned to commit
    a642f7a, fetched in PR I from cbdb-project/cbdb-online-main-server)

Writes:
  - analysis/index_year_rule_comparison.json
    Machine-readable per-rule snapshot from both sides plus the
    hand-curated mapping and per-pair status verdict.

The companion human-readable narrative lives in
`analysis/index_year_rule_comparison.md` and is hand-written; this
script only produces the JSON evidence + counts that the markdown
references.

Important: the Access "Rule N" numbering and the PHP "Rule N"
numbering are NOT the same numbering scheme.  For example Access
"Rule 03 BY" is `c_birthyear + 59` (no spouse), while PHP "Rule 03"
is the wife-from-husband propagation.  The mapping below pairs them
by *semantic intent*, not by number.  Pairs flagged as
`needs_manual_review` are ones the maintainer should look at first.

This script is intentionally conservative: it does NOT auto-classify
'matched' for any non-trivial pair.  The verdict either comes from
explicit pre-computed evidence (e.g. literal SQL substring match) or
from the hand-curated map; otherwise it stays `needs_manual_review`.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ACCESS_DIR = ROOT / "analysis" / "dump_data" / "querydefs_index"
PHP_FILE = ROOT / "analysis" / "php_source" / "IndexYearRebuildService.php"
OUT_JSON = ROOT / "analysis" / "index_year_rule_comparison.json"

# Pinned PHP source provenance.  Fetched via:
#   gh api repos/cbdb-project/cbdb-online-main-server/contents/
#       app/Services/IndexYearRebuildService.php?ref=<sha>
PHP_SOURCE_REPO = "cbdb-project/cbdb-online-main-server"
PHP_SOURCE_PATH = "app/Services/IndexYearRebuildService.php"
PHP_SOURCE_COMMIT = "a642f7ab6552ac48e5e98b867b155121d5b0fe3a"
PHP_SOURCE_DATE = "2026-03-13"


# ----------------------------------------------------------------------
# Access side: load every BM IY *.sql and pull a few structured fields
# ----------------------------------------------------------------------

def load_access_rules() -> list[dict]:
    rules = []
    for f in sorted(ACCESS_DIR.glob("BM_IY_*.sql")):
        text = f.read_text(encoding="utf-8")
        # Header: "-- QueryDef name: BM IY Rule 03 BY Query"
        m = re.search(r"-- QueryDef name: (.+)", text)
        name = m.group(1).strip() if m else f.stem
        body = re.sub(r"^--.*\n", "", text, flags=re.MULTILINE).strip()
        rules.append({
            "file": str(f.relative_to(ROOT)),
            "name": name,
            "is_phase_1": "Phase 1" in name,
            "uses_tmpBM_NIY": "tmpBM_NIY" in body,
            "is_dynasty_query": "Dynasty" in name,
            # Pull the SET ... = ... expression for c_index_year if present
            "set_index_year_expr": _extract_set_expr(body, "c_index_year"),
            "where_short": _extract_where_first_line(body),
            "joins_kin_data": "KIN_DATA" in body,
            "joins_entry_data": "ENTRY_DATA" in body,
            "kin_codes_used": _extract_kin_codes(body),
            "sets_type_code": ("c_index_year_type_code" in body
                                or "c_rule" in body),
        })
    return rules


def _extract_set_expr(sql: str, target: str) -> str | None:
    # Match e.g.  BIOG_MAIN.c_index_year = [BIOG_MAIN].[c_birthyear]+59
    m = re.search(
        rf"\.{re.escape(target)}\s*=\s*([^,\n]+?)\s*,",
        sql, flags=re.IGNORECASE,
    )
    if m:
        return m.group(1).strip()
    return None


def _extract_where_first_line(sql: str) -> str | None:
    m = re.search(r"WHERE\s+(.+)", sql, flags=re.IGNORECASE | re.DOTALL)
    if not m:
        return None
    return m.group(1).strip().split("\n")[0][:160]


def _extract_kin_codes(sql: str) -> list[int]:
    # Catch  c_kin_code = 134  or  c_kin_code IN (125,165)
    out = set()
    for m in re.finditer(r"c_kin_code\s*=\s*(\d+)", sql,
                          flags=re.IGNORECASE):
        out.add(int(m.group(1)))
    for m in re.finditer(
        r"c_kin_code\s+IN\s*\(([^)]+)\)", sql, flags=re.IGNORECASE
    ):
        for tok in m.group(1).split(","):
            tok = tok.strip()
            if tok.isdigit():
                out.add(int(tok))
    return sorted(out)


# ----------------------------------------------------------------------
# PHP side: read the file, slice each method body, capture a snapshot
# ----------------------------------------------------------------------

def load_php_rules() -> dict[str, dict]:
    src = PHP_FILE.read_text(encoding="utf-8")
    # Pull each `protected function sqlXxx(...)` body.
    rules: dict[str, dict] = {}
    for m in re.finditer(
        r"protected function (sql\w+)\s*\(([^)]*)\)\s*:\s*string\s*\{(.+?)\n\s*\}\n",
        src, flags=re.DOTALL,
    ):
        name = m.group(1)
        params = m.group(2).strip()
        body = m.group(3)
        rules[name] = {
            "method": name,
            "params": params,
            # Grab the SET clauses and WHERE clauses as strings
            "set_index_year_expr": _php_first_set(body, "c_index_year"),
            "set_type_code_expr": _php_first_set(
                body, "c_index_year_type_code"),
            "set_source_id_expr": _php_first_set(
                body, "c_index_year_source_id"),
            "uses_kin_data": "KIN_DATA" in body,
            "uses_entry_data": "ENTRY_DATA" in body,
            "kin_codes_used": _php_kin_codes(body),
        }
    # Also grab the rebuild()-driver phase ordering.
    rebuild_match = re.search(
        r"public function rebuild\(\):\s*array\s*\{(.+?)\n\s*\}\n",
        src, flags=re.DOTALL,
    )
    phases = {"phase_a": [], "phase_b": [], "loop": []}
    if rebuild_match:
        body = rebuild_match.group(1)
        for phase_key, phase_marker in [
            ("phase_a", "phaseARules"),
            ("phase_b", "phaseBRules"),
            ("loop",    "loopRules"),
        ]:
            sec = re.search(
                rf"\${phase_marker}\s*=\s*\[(.+?)\];",
                body, flags=re.DOTALL,
            )
            if not sec:
                continue
            for m in re.finditer(
                r"\['(\w+)',\s*\$this->(sql\w+)\s*\(([^)]*)\)\]",
                sec.group(1),
            ):
                phases[phase_key].append({
                    "type_code": m.group(1),
                    "method": m.group(2),
                    "args": m.group(3).strip(),
                })
    return {"methods": rules, "phases": phases}


def _php_first_set(body: str, col: str) -> str | None:
    m = re.search(
        rf"\.\s*{re.escape(col)}\s*=\s*([^,\n]+?)(?:,|\n|\s+WHERE)",
        body, flags=re.IGNORECASE,
    )
    if m:
        return m.group(1).strip()
    return None


def _php_kin_codes(body: str) -> list[int]:
    out = set()
    for m in re.finditer(r"c_kin_code\s*=\s*(\d+)", body):
        out.add(int(m.group(1)))
    for m in re.finditer(r"c_kin_code\s+IN\s*\(([^)]+)\)", body):
        for tok in m.group(1).split(","):
            tok = tok.strip().rstrip("'").lstrip("'")
            if tok.isdigit():
                out.add(int(tok))
    return sorted(out)


# ----------------------------------------------------------------------
# Hand-curated semantic mapping
#
# Format: each item is one *pair* of candidate equivalents, with a
# verdict the maintainer should sanity-check.  The verdict is one of:
#   matched                 - identical formula + preconditions
#   matched_minor_diff      - same intent, minor implementation diff
#   logic_diff              - candidate algorithm divergence
#   missing_in_access       - PHP rule with no Access analogue
#   missing_in_php          - Access rule with no PHP analogue
#   needs_manual_review     - default; not yet checked
# ----------------------------------------------------------------------

PAIR_MAP: list[dict] = [
    # ---- Access "01A" (deathyear when birthyear≥deathyear-60) ----
    {
        "access_name": "BM IY Rule 01A DY Query",
        "php_method": None,
        "verdict": "missing_in_php",
        "notes": (
            "Access has TWO Rule-01 variants: 01A copies "
            "c_deathyear when c_birthyear>=c_deathyear-60, 01B uses "
            "c_birthyear+59 when c_birthyear<=c_deathyear-59.  PHP's "
            "Rule 01 is just `c_index_year = c_birthyear` whenever "
            "birthyear is valid — no death-year branch.  The Access "
            "01A path produces a different value (deathyear, not "
            "birthyear) for old people whose recorded death age is "
            "≤60 — that is potential algorithm divergence."
        ),
    },
    {
        "access_name": "BM IY Rule 01B DY Query",
        "php_method": "sqlRule01",
        "verdict": "logic_diff",
        "notes": (
            "Both fire when birthyear is the strongest evidence, but "
            "the *value* differs: Access 01B writes c_birthyear+59, "
            "PHP sqlRule01 writes c_birthyear.  PHP's Phase B Rule 03 "
            "(spouse propagation) also uses husband.c_index_year+3 "
            "rather than +62 — see Rule 04W.  The +59 / +62 offsets "
            "in Access look like a 'midpoint of life' adjustment "
            "that PHP has dropped in favour of using the raw "
            "birthyear and letting Phase C propagation do the "
            "smoothing."
        ),
    },
    # ---- Access "02 DY" (deathyear when no birthyear) ----
    {
        "access_name": "BM IY Rule 02 DY Query",
        "php_method": "sqlRule02",
        "verdict": "logic_diff",
        "notes": (
            "Access Rule 02 DY: c_index_year = c_deathyear whenever "
            "birthyear is unknown and deathyear>0.  PHP sqlRule02: "
            "c_index_year = c_deathyear - c_death_age + 1, gated on "
            "c_death_age > 0.  Different formulas AND different "
            "preconditions (PHP requires a recorded death age; "
            "Access doesn't).  PHP also has sqlRule29Or30 to cover "
            "the case where deathyear is known but death_age isn't, "
            "with c_deathyear - 56 / -63 by gender.  Together those "
            "cover similar ground but with different numerical "
            "outputs."
        ),
    },
    # ---- Access "03 BY" (birthyear+59 when no deathyear) ----
    {
        "access_name": "BM IY Rule 03 BY Query",
        "php_method": "sqlRule01",
        "verdict": "logic_diff",
        "notes": (
            "Access Rule 03 BY: c_birthyear+59, fires only when "
            "deathyear is null/zero.  PHP sqlRule01: c_birthyear "
            "(no offset), fires whenever birthyear is valid.  Same "
            "input column, different output.  Note: the Access rule "
            "*number* '03' is a different rule entirely from PHP's "
            "Rule 03 (which is wife-from-husband propagation).  The "
            "two numbering schemes do not align."
        ),
    },
    # ---- Access "04W" wife from husband ----
    {
        "access_name": "BM IY Rule 04W Query",
        "php_method": "sqlRule03",
        "verdict": "logic_diff",
        "notes": (
            "Both compute wife's index year from husband (kin_code="
            "134), excluding concubine relationships.  But Access "
            "uses husband.c_birthyear+62, PHP uses husband."
            "c_index_year+3.  Access also explicitly tests "
            "KIN_DATA_1.c_kin_code<>168 (only one concubine code); "
            "PHP excludes a wider set [168, 163, 344, 467, 585]."
        ),
    },
    # ---- Access "05 JS" jinshi entry ----
    {
        "access_name": "BM IY Rule 05 JS Query",
        "php_method": "sqlEntryRule",  # called with '040101', 30, '05'
        "verdict": "logic_diff",
        "notes": (
            "**Sign flip!**  Access Rule 05 JS: c_index_year = "
            "ENTRY_DATA.c_year + 30 (gates on c_entry_code IN (36, "
            "165) directly).  PHP sqlEntryRule('040101', 30, '05'): "
            "c_index_year = entry_year - 30, gated by ENTRY_CODE_"
            "TYPE_REL.c_entry_type='040101' (a category mapping "
            "table).  Both also use MIN(c_year) on the PHP side vs "
            "first matching row on the Access side.  This is the "
            "single biggest divergence we found and is a likely "
            "explanation for many of the index_year_only_diff "
            "observations from PR G."
        ),
    },
    # ---- Access "06 JR" juren ----
    {
        "access_name": "BM IY Rule 06 JR Query",
        "php_method": "sqlEntryRule",  # called with '040102', 27, '07'
        "verdict": "logic_diff",
        "notes": (
            "Same sign flip as Rule 05.  Access uses +27 on raw "
            "entry year, PHP uses -27.  Naming also drifts: Access "
            "calls this 'JR' (juren), PHP doesn't expose the type "
            "mnemonic directly — the entry code mapping is in "
            "ENTRY_CODE_TYPE_REL (040102)."
        ),
    },
    # ---- Access "07 XC" -> tmpBM_NIY (different storage layer) ----
    {
        "access_name": "BM IY Rule 07 XC Query",
        "php_method": None,
        "verdict": "needs_manual_review",
        "notes": (
            "Access Rule 07 XC writes to tmpBM_NIY (a temp staging "
            "table) rather than directly to BIOG_MAIN, with "
            "c_year+39 and c_entry_code=257.  PHP doesn't appear "
            "to have a direct equivalent for entry code 257; "
            "sqlEntryRule covers '040103' with -21 / -18 offsets "
            "for the wife variants, no obvious +39 path.  Need to "
            "compare against ENTRY_CODE_TYPE_REL contents to see "
            "where 257 maps."
        ),
    },
    # ---- Spouse from husband entry (W variants) ----
    {
        "access_name": "BM IY Rule 05W Husband JS Query",
        "php_method": "sqlWifeFromEntryRule",  # '040101', 27, '06'
        "verdict": "logic_diff",
        "notes": (
            "Same sign-flip pattern.  Access wife-from-husband-jinshi "
            "uses husband_entry_year + N; PHP uses entry_year - N."
        ),
    },
    {
        "access_name": "BM IY Rule 06W Husband JR Query",
        "php_method": "sqlWifeFromEntryRule",  # '040102', 24, '08'
        "verdict": "logic_diff",
        "notes": "Same sign-flip pattern as 05W."
    },
    {
        "access_name": "BM IY Rule 07W Husband XC Query",
        "php_method": None,
        "verdict": "needs_manual_review",
        "notes": (
            "Wife variant of Rule 07 XC; same tmpBM_NIY routing.  "
            "No obvious PHP equivalent."
        ),
    },
    # ---- Father from child birthyear (Rule 08 / sqlRule11 / sqlAggregateRule13) ----
    {
        "access_name": "BM IY Rule 08 Father BY Query",
        "php_method": "sqlAggregateRule13",
        "verdict": "needs_manual_review",
        "notes": (
            "Both compute father's index year from child birthyear, "
            "kin_code=75.  Access uses c_birthyear-30 in Phase 1 "
            "Doublecheck path; PHP picks MIN(child.c_birthyear)-30 "
            "with deterministic source-personid via the aggregate "
            "subquery.  Need to inspect the Access Phase 1 + "
            "Doublecheck pair to confirm the aggregation behaviour "
            "matches MIN."
        ),
    },
    # ---- The remaining BM IY rules (09 Oldest Child BY Father,
    #      10 Older Brother BY, 12 SIL BY, 13 Grandfather BY,
    #      14 Father IY, 15 Part 1 Father, 16 Older Brother IY,
    #      17 Younger Brother IY, 18 SIL IY, 19 Grandfather IY) ----
    # Each of these has a Phase 1 (and sometimes Phase 1 Doublecheck or
    # Phase 2 Null) helper plus a final UPDATE.  PHP collapses each
    # family to a single aggregate query.  We mark them all as
    # needs_manual_review for now — the next pass should walk each one.
    {
        "access_name": "BM IY Rule 09 Oldest Child BY Father (group)",
        "php_method": "sqlAggregateRule13",
        "verdict": "needs_manual_review",
        "notes": (
            "Aggregate father from oldest child birthyear.  Both "
            "sides aim at MIN(child.c_birthyear) - 30 logically, "
            "but Access splits the work across multiple Phase queries "
            "and a temp table.  Worth confirming the MIN selection "
            "agrees on tie-break and source attribution."
        ),
    },
    {
        "access_name": "BM IY Rule 10 Older Brother BY (group)",
        "php_method": "sqlAggregateSiblingRule",
        "verdict": "needs_manual_review",
        "notes": (
            "Older brother propagation.  PHP uses kin_codes "
            "[125, 165] with MAX +2; Access kin codes need "
            "verification."
        ),
    },
    {
        "access_name": "BM IY Rule 12 SIL BY (group)",
        "php_method": "sqlAggregateSonInLawRule",
        "verdict": "needs_manual_review",
        "notes": (
            "Son-in-law propagation.  PHP kin codes (181,201,224,332).  "
            "Access kin codes need verification.  Two Access queries "
            "(Phase 1 + Phase 1 Doublecheck) suggest a sanity check "
            "loop that PHP doesn't have."
        ),
    },
    {
        "access_name": "BM IY Rule 13 Grandfather BY (group)",
        "php_method": "sqlRule27",
        "verdict": "needs_manual_review",
        "notes": (
            "Grandfather propagation, kin_code=62.  PHP uses "
            "grandfather.c_birthyear + 60.  Access has THREE queries "
            "in this group (Phase 1, Phase 2, Phase 2 Null) — likely "
            "covers null handling explicitly that PHP collapses into "
            "the WHERE clause."
        ),
    },
    {
        "access_name": "BM IY Rule 14 Father IY Phase 1 Query",
        "php_method": "sqlLoopFatherPropagationRule",
        "verdict": "needs_manual_review",
        "notes": (
            "Phase C loop: child.c_index_year = father.c_index_year + "
            "30.  PHP loops up to 2 times; Access' loop driver is in "
            "frmBaseMaintenance VBA which we have NOT extracted yet."
        ),
    },
    {
        "access_name": "BM IY Rule 15 Part 1 Father Phase 1 Query",
        "php_method": "sqlLoopOldestChildIndexToFatherRule",
        "verdict": "needs_manual_review",
        "notes": (
            "Father from oldest child index_year (Phase C).  Same "
            "MIN aggregation question as Rule 13."
        ),
    },
    {
        "access_name": "BM IY Rule 16 Older Brother IY Phase 1 Query",
        "php_method": "sqlLoopSiblingRule",  # MAX with +2
        "verdict": "needs_manual_review",
        "notes": "Older brother index_year propagation.",
    },
    {
        "access_name": "BM IY Rule 17 Younger Brother IY Phase 1 Query",
        "php_method": "sqlLoopSiblingRule",  # MIN with -2
        "verdict": "needs_manual_review",
        "notes": "Younger brother index_year propagation.",
    },
    {
        "access_name": "BM IY Rule 18 SIL IY Phase 1 Query",
        "php_method": "sqlLoopSonInLawRule",
        "verdict": "needs_manual_review",
        "notes": "Son-in-law index_year propagation.",
    },
    {
        "access_name": "BM IY Rule 18 SIL IY Mother Phase 1 Query",
        "php_method": "sqlLoopOldestChildIndexToMotherRule",
        "verdict": "needs_manual_review",
        "notes": (
            "Mother from oldest child index_year, Phase C.  Naming "
            "is confusing (Access calls this '18 ... Mother' even "
            "though it's a mother-not-SIL flow)."
        ),
    },
    {
        "access_name": "BM IY Rule 19 Grandfather IY Phase 1 Query",
        "php_method": "sqlLoopGrandfatherPropagationRule",
        "verdict": "needs_manual_review",
        "notes": "Grandfather index_year propagation, Phase C.",
    },
    # ---- PHP rules with no Access analogue we identified ----
    {
        "access_name": None,
        "php_method": "sqlRule29Or30",
        "verdict": "missing_in_access",
        "notes": (
            "PHP gender-default-from-deathyear (c_deathyear - 56 for "
            "female, -63 for male, no death_age required).  Access "
            "Rule 02 DY is similar but writes c_deathyear with no "
            "offset.  These are not equivalent."
        ),
    },
    {
        "access_name": None,
        "php_method": "sqlRule11",
        "verdict": "missing_in_access",
        "notes": (
            "PHP child = father.c_birthyear + 30, kin_code=75.  "
            "Access has no obvious 'child from father birthyear' rule "
            "in the BM IY family — the analogous Rule 14 is in Phase "
            "C and uses father.c_index_year + 30.  Could be that "
            "Access intentionally relies on Phase C propagation to "
            "cover this; would benefit from confirmation."
        ),
    },
    {
        "access_name": None,
        "php_method": "sqlAggregateRule15",
        "verdict": "missing_in_access",
        "notes": (
            "PHP mother from oldest child birthyear (kin_code=111).  "
            "Access has Rule 15 Part 1 Father and Rule 18 ... Mother, "
            "but those are in Phase C using c_index_year not "
            "c_birthyear — so PHP's Phase B mother-from-child-"
            "birthyear path looks unmatched in Access."
        ),
    },
    # PHP husband propagation in Phase C (rules 04 and 18 by type code).
    # Access doesn't appear to have a Phase-C husband flow — Rule 04W
    # is in Phase A using birthyear+62, then no further husband
    # propagation seems to happen.  Worth confirming.
    {
        "access_name": None,
        "php_method": "sqlLoopHusbandPropagationRule",
        "verdict": "needs_manual_review",
        "notes": (
            "PHP Phase C re-applies the wife-from-husband rule using "
            "husband.c_index_year+3 (after husbands have themselves "
            "been resolved by other rules).  Access' Rule 04W is in "
            "Phase A (birthyear+62) and we don't see an obvious "
            "Phase C husband-propagation step.  Could mean PHP "
            "smooths through more chains than Access; could also "
            "be that Access' driver VBA (which we haven't extracted "
            "yet) re-runs the Phase A rules in a loop and effectively "
            "covers it.  Confirming requires the frmBaseMaintenance "
            "VBA source."
        ),
    },
]


def main() -> int:
    access = load_access_rules()
    php_doc = load_php_rules()
    php_methods = php_doc["methods"]
    php_phases = php_doc["phases"]

    # Build lookup tables for resolving names.
    access_by_name = {r["name"]: r for r in access}

    enriched_pairs = []
    for pair in PAIR_MAP:
        a = access_by_name.get(pair["access_name"]) if pair["access_name"] else None
        p = (php_methods.get(pair["php_method"]) if pair["php_method"]
             else None)
        enriched_pairs.append({
            **pair,
            "access": a,
            "php": p,
        })

    # Pair coverage summary
    verdict_counts: dict[str, int] = {}
    for pair in enriched_pairs:
        verdict_counts[pair["verdict"]] = (
            verdict_counts.get(pair["verdict"], 0) + 1)

    # Cross-check: any Access rule we never paired up?
    paired_access_names = {p["access_name"] for p in PAIR_MAP
                            if p["access_name"]}
    unpaired_access = [r["name"] for r in access
                       if r["name"] not in paired_access_names
                       and not r["is_dynasty_query"]]

    paired_php_methods = {p["php_method"] for p in PAIR_MAP
                          if p["php_method"]}
    unpaired_php = [m for m in php_methods if m not in paired_php_methods]

    out = {
        "php_source": {
            "repo": PHP_SOURCE_REPO,
            "path": PHP_SOURCE_PATH,
            "commit": PHP_SOURCE_COMMIT,
            "commit_date": PHP_SOURCE_DATE,
            "local_copy": str(PHP_FILE.relative_to(ROOT)),
        },
        "summary": {
            "access_rules_total": len(access),
            "php_methods_total": len(php_methods),
            "pairs_in_map": len(PAIR_MAP),
            "verdict_counts": verdict_counts,
            "unpaired_access_rules": unpaired_access,
            "unpaired_php_methods": unpaired_php,
        },
        "php_phase_order": php_phases,
        "pairs": enriched_pairs,
        "access_rules": access,
        "php_methods": php_methods,
    }
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"wrote {OUT_JSON}")
    print(f"  access rules:  {len(access)}")
    print(f"  php methods:   {len(php_methods)}")
    print(f"  pairs in map:  {len(PAIR_MAP)}")
    print("  verdict counts:")
    for v, n in sorted(verdict_counts.items()):
        print(f"    {v:30s} {n}")
    print(f"  unpaired access (excl. Dynasty): {len(unpaired_access)}")
    for n in unpaired_access:
        print(f"    - {n}")
    print(f"  unpaired php methods: {len(unpaired_php)}")
    for n in unpaired_php:
        print(f"    - {n}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
