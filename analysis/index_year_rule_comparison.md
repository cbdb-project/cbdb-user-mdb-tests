# Runtime Access index-year rebuild vs PHP IndexYearRebuildService — rule-by-rule comparison

**Updated 2026-05-03 (PR N).**  This document **supersedes** the
earlier (PR I) comparison that paired PHP against the 37 saved
`BM IY Rule …` QueryDefs.  PR M's `SaveAsText` extraction of
`frmBaseMaintenance` showed those QueryDefs are **not** what
`CmdIndexYear_Click` runs at maintenance time — the actual
runtime path is the inline VBA `GetBirthIndexYearSQL`, which
turns out to track PHP much more closely than PR I's flagged
`logic_diff` count suggested.

The earlier `analysis/index_year_rule_comparison.{md,json}`
contents are kept in git history for reference, but the JSON file
in HEAD now reflects the runtime-vs-PHP pairing.

## Sources

| Side | File | Sub / function |
|------|------|---------------|
| **Access (runtime)** | [`analysis/dump_data/vba/Form_frmBaseMaintenance.vb`](dump_data/vba/Form_frmBaseMaintenance.vb) | `GetBirthIndexYearSQL` (called from `CmdIndexYear_Click`) |
| **Access (vestigial, NOT runtime)** | [`analysis/dump_data/querydefs_index/*.sql`](dump_data/querydefs_index/) | 37 saved `BM IY Rule …` QueryDefs.  Present in DATA mdb but `CmdIndexYear_Click` does not call them.  PR I compared against these by mistake. |
| **PHP** | [`analysis/php_source/IndexYearRebuildService.php`](php_source/IndexYearRebuildService.php) | pinned commit [`a642f7a`](https://github.com/cbdb-project/cbdb-online-main-server/blob/a642f7ab6552ac48e5e98b867b155121d5b0fe3a/app/Services/IndexYearRebuildService.php) (2026-03-13) |

## Pairing strategy

Both runtime Access and PHP emit a `c_index_year_type_code` per
rule (e.g. `'01'`, `'05'`, `'12'`).  We pair rules by emitted
type_code rather than by rule-number / name.  This avoids the
name-based mismatch that misled PR I (Access "Rule 03 BY" and
PHP "Rule 03" are completely different rules; both, however,
emit different type_codes, so type-code-based pairing is
unambiguous).

`compare_index_year_rules.py` now does this automatically.

## Verdict counts

(From the current run of `compare_index_year_rules.py`; re-run
after every fresh PHP-side commit pull or DATA-mdb refresh.)

| Verdict | Count |
|---------|------:|
| ✅ matched               | 22 |
| 🟡 matched_minor_diff    | 8 |
| 🔴 logic_diff            | 0 |
| ➖ access_only           | 3 |
| ➖ php_only              | 0 |
| ⚠ needs_manual_review  | 0 |

**No `logic_diff` flags.** PR I's marquee finding (Access uses
`+N`, PHP uses `-N`) was an artefact of comparing PHP against
the wrong Access source.  The runtime Access uses `-N` like PHP.

## Per-rule pairings

### Phase A: direct rules

| type_code | Access source line ish | PHP method | Verdict |
|-----------|----------------------|------------|---------|
| `01` | `c_index_year = c_birthyear` | `sqlRule01` | matched |
| `02` | `c_deathyear - c_death_age + 1` | `sqlRule02` | matched |
| `03` | `husband.c_index_year + 3`, kin 134 | `sqlRule03` (concubine=false) | matched |
| `05` | `ENTRY_DATA.c_year - 30` (entry_type 040101) | `sqlEntryRule('040101', 30, '05')` | matched |
| `06` | wife from husband entry `-27` | `sqlWifeFromEntryRule('040101', 27, '06')` | matched |
| `07` | `ENTRY_DATA.c_year - 27` (entry_type 040102) | `sqlEntryRule('040102', 27, '07')` | matched |
| `08` | wife `-24` | `sqlWifeFromEntryRule('040102', 24, '08')` | matched |
| `09` | `ENTRY_DATA.c_year - 21` (entry_type 040103) | `sqlEntryRule('040103', 21, '09')` | matched |
| `10` | wife `-18` | `sqlWifeFromEntryRule('040103', 18, '10')` | matched |
| `11` | `father.c_birthyear + 30`, kin 75 | `sqlRule11` | matched |
| `17` | wife concubine variant of 03 | `sqlRule03(concubine=true)` | matched (kin filter slightly broader on Access; low impact) |
| `29` | `c_deathyear - 64` (male) | `sqlRule29Or30(female=false)`, `-63` | **matched_minor_diff** (off-by-1) |
| `30` | `c_deathyear - 53` (female) | `sqlRule29Or30(female=true)`, `-56` | **matched_minor_diff** (off-by-3) |
| `31` | wife jinshi concubine variant | (no PHP equivalent) | access_only |
| `32` | wife juren concubine variant | (no PHP equivalent) | access_only |
| `33` | wife 040103 concubine variant | (no PHP equivalent) | access_only |

### Phase B: aggregate rules

| type_code | Access | PHP method | Verdict |
|-----------|--------|------------|---------|
| `13` | father from MIN(child.c_birthyear) - 30, kin 75 | `sqlAggregateRule13` | matched_minor_diff (Access uses staging step + final UPDATE; PHP uses single subquery aggregate; equivalent on the happy path) |
| `15` | mother from MIN(child.c_birthyear) - 27, kin 111 | `sqlAggregateRule15` | matched_minor_diff (same staging-vs-subquery distinction) |
| `19` | older brother MAX(c_birthyear) + 2 (kin 125, 165) | `sqlAggregateSiblingRule(MAX, +2)` | matched_minor_diff |
| `21` | younger brother MIN - 2 (kin 126, 166) | `sqlAggregateSiblingRule(MIN, -2)` | matched_minor_diff |
| `23` | son-in-law male target -27 | `sqlAggregateSonInLawRule(false)` | matched_minor_diff |
| `25` | son-in-law female target -24 | `sqlAggregateSonInLawRule(true)` | matched_minor_diff |
| `27` | grandfather.c_birthyear + 60, kin 62 | `sqlRule27` | matched |

### Phase C: loop rules (with CONCAT'd type_codes)

Both sides CONCAT the loop-rule type_code onto the parent's
type_code (so e.g. PHP's `'11' + '12' = '1112'` corresponds to
Access's `iif(parent='01','12',parent + '12')`).

| type_code | Access | PHP method | Verdict |
|-----------|--------|------------|---------|
| `04` | husband propagation, husband.c_index_year + 3 | `sqlLoopHusbandPropagationRule(false)` | matched |
| `12` | child = father.c_index_year + 30 | `sqlLoopFatherPropagationRule` | matched |
| `14` | father = MIN(child.c_index_year) - 30 | `sqlLoopOldestChildIndexToFatherRule` | matched |
| `16` | mother = MIN(child.c_index_year) - 27 | `sqlLoopOldestChildIndexToMotherRule` | matched |
| `18` | husband (concubine) | `sqlLoopHusbandPropagationRule(true)` | matched |
| `20` | older brother +2 | `sqlLoopSiblingRule(MAX, +2)` | matched |
| `22` | younger brother -2 | `sqlLoopSiblingRule(MIN, -2)` | matched |
| `24` | son-in-law male -27 | `sqlLoopSonInLawRule(false)` | matched |
| `26` | son-in-law female -24 | `sqlLoopSonInLawRule(true)` | matched |
| `28` | grandfather +60 | `sqlLoopGrandfatherPropagationRule` | matched |

## Implications for K1 / K2 (per-row year-drift classification)

The K1 / K2 buckets were built on the assumption that PR I's
`logic_diff` flags (especially the `+N`/`-N` sign-flip) explained
some of the year drifts.  PR N now removes that assumption: at
the rule level the runtime Access path matches PHP everywhere
*except* the off-by-1 / off-by-3 for type_codes `29` / `30`
(`matched_minor_diff`).

That doesn't invalidate K1 / K2's per-row buckets — those came
from PR G's per-personid diff list and only reference type_codes,
not the wrong rule labels.  Reading them now:

  - **`php_did_not_compute` × 19** — biggest sub-bucket from K2
    is `access_tcode='05'` × 7.  With PR N's finding that Access
    Rule 05 is `c_year - 30 with ENTRY_CODE_TYPE_REL.c_entry_type
    = '040101'` (matching PHP), the gap is more likely upstream
    in `ENTRY_CODE_TYPE_REL` membership — the entries aren't
    classified into `'040101'` on whichever side, so neither
    rule fires for those personids.  Worth a quick PHP-side
    check.
  - **`consistent_within_rule` × 14** — the diff = -20 cluster
    across PHP type_codes 11 / 13 / 15 / 19 is interesting given
    that Rules 13, 15, 19 are all `matched_minor_diff` (staging
    vs subquery aggregate).  The shared `-20` could reflect a
    consistent staging-step pick that picks the same wrong row
    each time.  Maintainer-investigatable.
  - **`iteration_order_diff` × 5** — both sides do CONCAT the
    same way, so this is most likely a Phase-C iteration-count
    difference (PHP caps at 2 loops; Access also caps at
    "tLoopCount < 3" i.e. 2 loops — both same!  Yet the data
    show divergence).  The most likely remaining cause is the
    order rules fire WITHIN each loop pass.  Worth confirming.

## Limitations

- **Type_code-based pairing presumes both sides use the same
  type_codes consistently.**  We confirmed this empirically
  (33 paired codes, 0 `php_only`, 3 Access-only concubine
  variants).  If a future cbdb-online-main-server commit renames
  a code, this comparator would silently mis-pair.  Re-run the
  comparator + spot-check after every PHP pull.
- **`needs_manual_review` is currently 0** because the
  hand-curated `VERDICT_BY_TYPE_CODE` table covers every paired
  code.  When a new code appears (Access or PHP), it'll default
  to `needs_manual_review` until added to the table.
- **PR I's `logic_diff` flags are not actively in repo any more,
  but the K1 / K2 outputs are unchanged** (their per-row buckets
  don't depend on PR I's verdict labels).  K1 / K2 docs gain a
  pointer at this updated comparator.
- **Off-by-1 / off-by-3 on Rules 29 / 30 (matched_minor_diff) is
  the closest thing to a real divergence we have at the rule
  level.**  Worth maintainer confirmation that the difference is
  intentional (Access's `-64` / `-53` may be older constants
  PHP intentionally updated).
