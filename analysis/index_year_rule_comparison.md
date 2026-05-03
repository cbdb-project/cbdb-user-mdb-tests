# Access BM IY rules vs PHP IndexYearRebuildService — rule-by-rule comparison

**What this is.** A reviewable side-by-side of the two index-year
rebuild implementations.  This is the "rule-layer difference map"
the maintainer should consult before triaging the 547 unclassified
per-row diffs from PR G; the goal is to know **which rule is the
likely cause** of any given divergence before walking the row.

**What this is not.** Not a verdict on whether either side is
correct.  Verdicts of `logic_diff` and `needs_manual_review` are
*candidate* divergences — they need confirmation from the
maintainer (or per-row data evidence) before being filed as bugs
against either side.

## Sources

- **Access (User MDB DATA backend)** —
  `data/CBDB_<YYYYMMDD>_DATA.mdb`, 37 saved QueryDefs named
  `BM IY Rule …`, extracted in PR H to
  [`analysis/dump_data/querydefs_index/*.sql`](dump_data/querydefs_index/).
  The driver VBA (which orders the rules and runs them in phases)
  lives on `frmBaseMaintenance` and in 4 modules in the same .mdb
  file; their **source code has not yet been extracted** (needs
  `Access.Application.SaveAsText`).  We can see *what* each
  individual rule does; we cannot yet see the *order* the Access
  side runs them in.

- **PHP (cbdb-online-main-server)** —
  [`app/Services/IndexYearRebuildService.php`](https://github.com/cbdb-project/cbdb-online-main-server/blob/a642f7ab6552ac48e5e98b867b155121d5b0fe3a/app/Services/IndexYearRebuildService.php)
  pinned to commit `a642f7a` (2026-03-13).  Local copy at
  [`analysis/php_source/IndexYearRebuildService.php`](php_source/IndexYearRebuildService.php).
  Three phases:
  - **Phase A** (direct rules): RESET → 01 → 02 → 03 → 29 → 30 → 05
    → 06 → 07 → 08 → 09 → 10 → 11.
  - **Phase B** (aggregate rules): 13, 15, 17, 19, 21, 23, 25, 27.
  - **Phase C** (loop, max 2 iterations): 04, 12, 14, 16, 18, 20,
    22, 24, 26, 28.

The `compare_index_year_rules.py` script writes the structured
counterpart of this document to
[`analysis/index_year_rule_comparison.json`](index_year_rule_comparison.json).

## Numbering scheme — the rule numbers do NOT line up

Access "Rule N" and PHP "Rule N" are different numbering schemes.
The two were paired by *semantic intent* (what kin / source / formula
they use), not by number.  Examples:

| Number | Access "Rule N" means | PHP "Rule N" means |
|--------|-----------------------|--------------------|
| 03     | birthyear + 59 when no deathyear | wife from husband (kin 134), husband.c_index_year + 3 |
| 04     | wife from husband (Access naming: 04W), husband.c_birthyear + 62 | husband propagation in Phase C (loop), husband.c_index_year + 3 |
| 05     | jinshi entry: c_year + 30 | jinshi entry: entry_year - 30 |

**Implication.** Don't read either side's `c_index_year_type_code`
column as if it referenced the other side's rule numbering.

## Pair-by-pair table

Status legend:
- ✅ matched — identical formula AND preconditions
- 🟡 matched_minor_diff — same intent, minor implementation difference (e.g. one side groups joins differently but result equivalent)
- 🔴 logic_diff — candidate algorithm divergence (different formula, different precondition, etc.)
- ⚠ needs_manual_review — paired but not yet checked in detail
- ➖ missing_in_access / missing_in_php — no corresponding rule paired

Verdict counts (from the script's JSON output):

| Verdict | Count |
|---------|------:|
| 🔴 logic_diff           | 8 |
| ➖ missing_in_access     | 3 |
| ➖ missing_in_php        | 1 |
| ⚠ needs_manual_review  | 15 |
| ✅ matched               | 0 |
| 🟡 matched_minor_diff    | 0 |

Zero pairs land in `matched` so far — that doesn't necessarily mean
everything diverges, just that we haven't confirmed any pair as
identical yet.  Most are `needs_manual_review`.

### Phase A — direct rules

| Access rule | PHP method | Status | Why |
|-------------|------------|--------|-----|
| `BM IY Rule 01A DY Query` (c_deathyear when c_birthyear≥c_deathyear-60) | _no PHP analogue_ | ➖ missing_in_php | Access splits Rule 01 into two variants by birth-vs-death gap; PHP just uses raw birthyear. |
| `BM IY Rule 01B DY Query` (c_birthyear+59 when c_birthyear≤c_deathyear-59) | `sqlRule01` (c_birthyear, no offset, fires whenever birthyear valid) | 🔴 logic_diff | **Different output value.** Access adds +59 ("midpoint of life"); PHP writes raw birthyear. |
| `BM IY Rule 02 DY Query` (c_deathyear when no birthyear) | `sqlRule02` (c_deathyear - c_death_age + 1, requires c_death_age>0) | 🔴 logic_diff | Access: no offset, no death-age requirement.  PHP: subtracts death age, requires it.  PHP's `sqlRule29Or30` covers the no-death-age case with c_deathyear ± fixed offset by gender. |
| `BM IY Rule 03 BY Query` (c_birthyear+59 when no deathyear) | `sqlRule01` (c_birthyear, no offset) | 🔴 logic_diff | Same shape as 01B vs sqlRule01 — Access offsets by +59, PHP doesn't. |
| `BM IY Rule 04W Query` (wife from husband.c_birthyear+62, kin=134, exclude kin_code=168) | `sqlRule03` (wife from husband.c_index_year+3, kin=134, exclude concubine kin set) | 🔴 logic_diff | **Two differences**: Access uses husband's birthyear+62, PHP uses husband's index_year+3.  Access excludes only one concubine code (168); PHP excludes 5 (`168, 163, 344, 467, 585`). |
| `BM IY Rule 05 JS Query` (c_year+30 from ENTRY_DATA.c_entry_code IN (36, 165)) | `sqlEntryRule('040101', 30, '05')` (c_year-30, joins ENTRY_CODE_TYPE_REL.c_entry_type='040101', uses MIN(c_year)) | 🔴 logic_diff | **Sign flip + indirection.** Access reads c_entry_code directly, adds +30.  PHP joins through ENTRY_CODE_TYPE_REL and subtracts -30 from MIN(c_year).  Probably the single biggest single-rule divergence in the index_year_only_diff bucket. |
| `BM IY Rule 06 JR Query` | `sqlEntryRule('040102', 27, '07')` | 🔴 logic_diff | Same sign-flip pattern.  Access +27, PHP -27. |
| `BM IY Rule 07 XC Query` (writes to `tmpBM_NIY` not BIOG_MAIN, c_year+39, c_entry_code=257) | _no obvious analogue_ | ⚠ needs_manual_review | Access uses an intermediate temp table.  PHP doesn't appear to have a +39 path or an entry_code=257 mapping. |
| `BM IY Rule 05W Husband JS Query` | `sqlWifeFromEntryRule('040101', 27, '06')` | 🔴 logic_diff | Same sign-flip as 05 / 05W. |
| `BM IY Rule 06W Husband JR Query` | `sqlWifeFromEntryRule('040102', 24, '08')` | 🔴 logic_diff | Same sign-flip. |
| `BM IY Rule 07W Husband XC Query` | _no obvious analogue_ | ⚠ needs_manual_review | Same temp-table issue as 07 XC. |
| _no Access analogue_ | `sqlRule29Or30` (c_deathyear - 56/63 by gender) | ➖ missing_in_access | PHP gender-default-from-deathyear rule.  Access Rule 02 DY covers similar ground but without the offset. |
| _no Access analogue_ | `sqlRule11` (child = father.c_birthyear + 30, kin 75) | ➖ missing_in_access | No Phase-A child-from-father-birthyear rule on the Access side that we found.  Rule 14 (Phase C) uses father.c_index_year + 30, which is a different path. |

### Phase B / Phase C — aggregate + loop rules

(All paired entries below are flagged `needs_manual_review` — the
shape is similar between the two sides but per-rule inputs and
tie-break behaviour need confirmation.)

| Access rule(s) | PHP method | Why review |
|----------------|------------|-----------|
| `BM IY Rule 08 Father BY Query` (+ Phase 1 helper) | `sqlAggregateRule13` (father from MIN(child.c_birthyear) - 30, kin 75) | Confirm Access uses MIN, not first-row, and that source-personid attribution agrees |
| `BM IY Rule 09 Oldest Child BY Father …` (group of Phase 1 + Doublecheck queries) | `sqlAggregateRule13` (same as above; the Access "Oldest Child" naming hints they collapse to MIN as well) | Confirm aggregation on tie-break |
| `BM IY Rule 10 Older Brother BY …` | `sqlAggregateSiblingRule(kin=[125,165], MAX, +2)` | Confirm Access uses the same kin codes (125, 165) and MAX, +2 offset |
| `BM IY Rule 12 SIL BY …` (Phase 1 + Phase 1 Doublecheck) | `sqlAggregateSonInLawRule(kin=[181,201,224,332], MIN, -27 male / -24 female)` | Confirm Access kin code list |
| `BM IY Rule 13 Grandfather BY …` (Phase 1 + Phase 2 + Phase 2 Null) | `sqlRule27` (descendant from grandfather.c_birthyear + 60, kin 62) | Three Access queries vs one PHP method — Access likely handles null cases explicitly that PHP folds into WHERE |
| `BM IY Rule 14 Father IY Phase 1 Query` | `sqlLoopFatherPropagationRule` (Phase C: child = father.c_index_year + 30) | Loop iteration order; Access driver loop not yet extracted |
| `BM IY Rule 15 Part 1 Father Phase 1 Query` | `sqlLoopOldestChildIndexToFatherRule` (Phase C: father from MIN(child.c_index_year) - 30) | Aggregation tie-break |
| `BM IY Rule 16 Older Brother IY Phase 1 Query` | `sqlLoopSiblingRule(MAX, +2)` | Confirm kin codes |
| `BM IY Rule 17 Younger Brother IY Phase 1 Query` | `sqlLoopSiblingRule(MIN, -2)` | Confirm kin codes |
| `BM IY Rule 18 SIL IY Phase 1 Query` | `sqlLoopSonInLawRule` | Loop iteration; gender split |
| `BM IY Rule 18 SIL IY Mother Phase 1 Query` | `sqlLoopOldestChildIndexToMotherRule` | Naming is confusing on the Access side ("18 ... Mother" is the mother flow, not SIL) |
| `BM IY Rule 19 Grandfather IY Phase 1 Query` | `sqlLoopGrandfatherPropagationRule` | Loop iteration |
| _no Access analogue_ | `sqlAggregateRule15` (mother from MIN(child.c_birthyear) - 27, kin 111) | Looks like PHP has an additional Phase B mother-from-child-birthyear rule that Access either doesn't run, or covers via a different path |
| _no Access analogue_ | `sqlLoopHusbandPropagationRule` (Phase C: wife from husband.c_index_year + 3) | Access' wife rule (04W) is in Phase A using birthyear; we don't see a separate Phase C husband-propagation step.  Could be that Access' driver VBA re-runs Rule 04W in a loop with later husband index values; need to confirm by reading frmBaseMaintenance source |

## What this means for the 547 unclassified diffs from PR G

PR G's classifier surfaced 547 per-row diffs with matching
birthyear+deathyear that couldn't be attributed to source-data
drift.  Of those:

- **478** were `index_addr_only_diff` — these have **nothing to
  do with the index-year rules above**.  They depend on
  `Form_frmIndexAddr.vb` (User MDB front end) vs PHP
  `IndexAddressRebuildService.php` (not yet extracted or
  compared).  Address comparison is a separate piece of work.
- **59** were `index_year_only_diff` — these are the ones the
  rule comparison above is most relevant to.  The 8 `logic_diff`
  pairs (especially the **+30 / -30 sign flip in entry-based
  rules**) are the most likely contributors.
- **10** were `index_both_diff` — both index fields disagree.
  Could overlap with any of the above.

Concretely, suggested next-step triage order (per-row):

1. Pick a handful of `index_year_only_diff` personids whose PHP
   `c_index_year_type_code` is in `{'05', '06', '07', '08', '09',
   '10'}` (entry-based rules).  Compare the User MDB and SQLite
   c_index_year values; the difference should be ~60 years apart
   if the sign flip is the actual cause.
2. Pick `index_year_only_diff` personids whose PHP type_code is
   `'01'` (birthyear) — should differ by 59 if Rule 01B vs
   sqlRule01 is the cause.
3. For remaining unmatched cases, fall back to per-row source
   walking.

## Limitations

- **Access driver VBA still missing.**  Without
  `frmBaseMaintenance.cls` source we don't actually know in what
  order Access runs the 37 BM IY queries, whether it loops, or
  whether some queries are unused.  PR J (interactive
  `Access.Application.SaveAsText`) is the obvious fix.
- **`tmpBM_NIY` flow not analysed.**  Several Access rules write to
  this temp table; PHP has no equivalent.  We don't know what
  reads from `tmpBM_NIY` later.  Could be a staging step that gets
  copied back to BIOG_MAIN, or could be an unused leftover.
- **`c_index_year_type_code` and `c_index_year_source_id` writes
  on the Access side were not surveyed in detail.**  PHP carefully
  sets both for every rule; Access's `BM IY Rule …` queries we
  read mostly write `c_notes` instead, with no obvious type-code
  assignment.  This itself could account for some `_type_code` /
  `_source_id` divergence.
- **Pinned PHP commit may have moved.**  We pinned at `a642f7a`
  (2026-03-13).  The latest weekly snapshot the SQLite was built
  from could have a newer version of `IndexYearRebuildService.php`
  with different behaviour.  Re-pin and re-run the comparison
  whenever a fresh snapshot is downloaded.
- **No `matched` verdicts yet.**  The script and this document
  default to conservative (`needs_manual_review`); a future pass
  with the maintainer or with side-by-side test data should
  upgrade pairs to `matched` / `matched_minor_diff` where
  warranted.
