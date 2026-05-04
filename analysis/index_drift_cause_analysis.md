# Index drift — cause analysis (PR Y)

This document upgrades the drift narrative from **"here are the
buckets"** to **"here is what we currently believe causes each
bucket, what evidence supports it, and what the next concrete
investigation step is"**.  It synthesises the existing outputs;
no new conclusions were drawn beyond what those outputs already
support.

Inputs synthesised:

  - PR N — `analysis/index_year_rule_comparison.md` (runtime
    `GetBirthIndexYearSQL` ↔ PHP `IndexYearRebuildService.php`
    pairing by emitted `c_index_year_type_code`).
  - K1 — `reports/index_year_drift_rule_classification.json`
    (per-row buckets for the 69 year-only diffs).
  - K2 — `reports/index_year_drift_rule_groups.json` (per-bucket
    triage groups, signature naming, blocker labels).
  - PR L — `reports/index_addr_drift_classification.json`
    (per-row buckets for the 488 c_index_addr_id diffs).
  - PR S — `reports/index_addr_same_candidates_deep_dive.json`
    (root-cause walk of the 10 same-candidate-different-winner
    rows).
  - PR G/M/X — `analysis/index_drift_algorithm_notes.md`
    (algorithm narrative + maintenance-trigger-path notes).
  - Cause summary mirror (machine-readable):
    `reports/index_drift_cause_summary.json`.

## Top-level framing

After PR N's rule-level comparator, **runtime Access and PHP
agree at the rule-definition level on every paired type_code
except the off-by-1/-3 minor variants on `c_deathyear` defaults
(rules 29 / 30) and the three Access-only concubine variants
(rules 31 / 32 / 33).**  Verdict counts: 22 matched / 8
matched_minor_diff / 0 logic_diff / 3 access_only / 0 php_only.

There is **no live `+N` / `-N` sign-flip finding**.  The PR I
hypothesis that the entry rules disagreed on sign was an
artefact of comparing PHP against the wrong Access source (the
vestigial `BM IY Rule …` QueryDefs).  The runtime path
(`GetBirthIndexYearSQL`) uses `-N` like PHP.

What that leaves for the remaining drift is, by elimination, a
small number of architectural seams rather than rule-logic
divergences:

  1. **Rule coverage gaps** — one side's rule chain doesn't
     reach a value for some persons (NULL / 0 result).
  2. **Runtime priority / iteration order** — both sides have
     the rule, but pick different rules (or different rule
     iterations) for the same person.
  3. **Tie-break / aggregation detail** — both sides pick the
     same rule but apply it to different evidence rows when
     candidates tie on priority.
  4. **Sentinel / null handling** — one side writes a sentinel
     (e.g. PHP `32767`) where the other writes nothing.
  5. **Evidence / source-data drift** — input tables
     (BIOG_ADDR_DATA / ENTRY_DATA / NIAN_HAO etc.) are not
     fully compared across sides; differences there propagate
     into derived columns invisibly to a rule-by-rule check.
  6. **Stale derived state** — User MDB ships with derived
     `c_index_addr_id` that wasn't re-run after the most
     recent BIOG_ADDR_DATA refresh; PHP re-runs weekly so
     SQLite stays current.

The buckets below are organised by which seam they sit on.

---

## A. `c_index_year` cause buckets

69 year-only diffs (59 `index_year_only_diff` + 10
`index_both_diff`).  Each row counted exactly once.

### A1. `php_returned_sentinel` × 1

- **Likely cause.** PHP wrote a sentinel / overflow value
  (≥ 9999, signed-int overflow shape).  Access wrote a real
  computed year.
- **Algorithm evidence.** Per PR N runtime Access has no
  matching sentinel branch; the value is below any plausible
  computed `c_birthyear + offset` ceiling.
- **Row evidence.** K1 bucket holds 1 row; the `php_index_
  year` field carries the sentinel, the `user_index_year`
  field carries a normal year.
- **Confidence.** **High** — binary signal, doesn't depend on
  anything else.
- **Next action.** Forward the PHP value to the cbdb-online-
  main-server maintainers as a candidate sentinel-write upstream
  bug.  Out of scope for this repo.

### A2. `php_did_not_compute` × 19

K2 sub-groups by Access tcode (largest first):

| Access tcode | Count | K2 candidate label | Likely cause |
| ---: | ---: | --- | --- |
| `'05'` | 7 | `candidate_php_entry_code_mapping_gap` (6) + `candidate_php_entry_data_year_missing` (1) — **supported_by_focused_probe** (PR Z) | PHP `sqlEntryRule('040101', 30, '05')` joins `ENTRY_CODE_TYPE_REL` on `c_entry_type='040101'`; PR Z confirmed for 6/7 rows that the SQLite snapshot's `ENTRY_CODE_TYPE_REL` lacks the mapping User MDB has, and for the 1 outlier (pid 93384 Zhang Wenfu) that the SQLite `ENTRY_DATA.c_year` is 0 while User MDB's is 926.  All 7 are PHP-side upstream data gaps; Access fires Rule 05 correctly.  Detail: `analysis/index_year_tcode05_entry_mapping_probe.md`. |
| `'11'` | 5 | `candidate_php_rule_coverage_gap` | PHP `sqlRule11` requires the father to pass `validYearExpr` (`c_birthyear=0 AND c_dy NOT IN [2,25,29,46,83]`).  Likely the father was excluded by that gate on PHP's side. |
| `'14'` | 4 | `candidate_php_rule_coverage_gap` | Phase-C `sqlLoopOldestChildIndexToFatherRule` only fires inside the Phase-C loop; possibly an iteration-count difference. |
| `'20'` | 1 | `candidate_php_rule_coverage_gap` | Phase-C older-brother propagation; iteration-count diff suspected. |
| `'2304'` | 1 | `candidate_php_rule_coverage_gap` | CONCAT `'23' + '04'` — Access ran Phase A Rule 23 (son-in-law male) then Phase C Rule 04 (husband propagation); PHP's loop terminated before reaching that combination. |
| `'07'` | 1 | `candidate_missing_php_rule_07_xc` | Access wrote via Rule 07 XC (`tmpBM_NIY` with `c_year+39`, `c_entry_code=257`).  PHP has no obvious entry-code-257 path in PR N's runtime comparison; either intentional vestigial Access branch or genuine gap. |

- **Algorithm evidence.** PR N pairs each Access tcode with a
  matched PHP rule for tcodes 05 / 11 / 14 / 20.  For
  `'2304'` and `'07'` the pairing is partial — the
  CONCAT-suffix and Rule 07 XC respectively need runtime-rule
  triage to confirm the PHP-side gap.
- **Row evidence.** K1 / K2 group rows by Access tcode; per
  group the `examples` field carries 3 representative
  personids, `all_personids` the full list.
- **Confidence.** **Supported by focused probe (PR Z)** for
  `'05'` × 7 — 6 are confirmed `ENTRY_CODE_TYPE_REL` mapping
  gaps, 1 is a PHP-side `ENTRY_DATA.c_year = 0` gap; both are
  PHP-side upstream data issues, not algorithm divergence.
  **Medium** for `'11'` (PHP gate logic explicit, would need
  PHP-side pull); **medium** for the Phase-C groups (need
  iteration trace).
- **Next action.**
  - For `'05'` × 7 — **done by PR Z**.  Forward both the 6
    mapping-gap rows and the 1 c_year=0 row to the
    cbdb-online-main-server / SQLite-snapshot-build team as
    upstream data candidates; out-of-scope for this repo.
  - For `'11'` × 5 — pull each father's `c_birthyear` /
    `c_dy` from SQLite and confirm the validYear gate fires.
  - For `'14'` / `'20'` / `'2304'` — runtime-rule trace
    against `GetBirthIndexYearSQL`'s loop count.
  - For `'07'` × 1 — runtime-rule triage to decide whether
    Rule 07 XC is intentional or vestigial.

### A3. `access_did_not_compute` × 7

- **Likely cause.** Mirror of A2 from the other direction —
  PHP has a value, runtime Access wrote 0 / null.  Access's
  rule chain didn't reach a value for these persons, most
  likely because evidence available to PHP (e.g. an
  `ENTRY_DATA` row with a particular `c_entry_code` mapping)
  wasn't surfaced by the staging-step inserts in
  `GetBirthIndexYearSQL`.
- **Algorithm evidence.** PR N rule-level pairing shows PHP
  has matching rule definitions; the gap is upstream in the
  evidence-table path, not in the rule body.
- **Row evidence.** K1 bucket holds 7 rows.  No K2 sub-grouping
  yet (smaller bucket).
- **Confidence.** **Medium** — the evidence-path-mismatch
  story is consistent with PR N but not yet row-confirmed.
- **Next action.** Per-row evidence walk: pull each person's
  ENTRY_DATA / KIN_DATA on the User-MDB side and check why
  the runtime UPDATE didn't surface a row.

### A4. `iteration_order_diff` × 5

- **Likely cause.** Both sides CONCAT loop-rule type_codes the
  same way and both cap at 2 Phase-C loops; the residual
  divergence is the **order rules fire WITHIN each loop pass**
  (which rule runs first when multiple are eligible for the
  same person in the same pass).
- **Algorithm evidence.** PR N has the loop structure and cap
  matched; the within-loop rule-firing order is not
  documented either side and is the only architectural
  remaining seam for this bucket.
- **Row evidence.** 5 rows whose Access tcode is a strict
  prefix-extension of PHP's (`'1112'` vs `'11'`, etc.).
- **Confidence.** **Medium**.
- **Next action.** Runtime-rule triage of `GetBirthIndexYearSQL`'s
  Phase-C body to enumerate within-loop firing order; cross-
  reference against PHP's `IndexYearRebuildService` Phase-C
  loop body.

### A5. `consistent_within_rule` × 14 — **REVERSED by PR AI + AJ**

K2 sub-groups by `(php_tcode, access_tcode, diff)` signature,
with the per-row probe verdict from PR AI / AJ:

| Signature | Count | Probe verdict |
| --- | ---: | --- |
| `('11','11', 1)` | 5 | PR AJ — 4 KIN_DATA evidence-pid drift + 1 BIOG_MAIN birthyear drift |
| `('19','19', -20)` | 3 | PR AI — BIOG_MAIN.c_birthyear drift, winner-pid birthyear differs by exactly the index_year diff |
| `('13','13', -20)` | 2 | PR AI — same |
| `('15','15', -20)` | 2 | PR AI — same |
| `('11','11', -20)` | 2 | PR AI — same |

- **Likely cause.** **Upstream source-data drift between
  User MDB and SQLite snapshot for the EVIDENCE persons that
  the rule reads from — NOT algorithm divergence.**  PR AI's
  per-row probe of the -20 sub-cluster (9 rows) and PR AJ's
  per-row probe of the +1 sub-cluster (5 rows) confirmed all
  14 / 14 rows fit one of two upstream-data shapes:
  **BIOG_MAIN.c_birthyear differs between sides for the
  evidence person** (8 rows) or **KIN_DATA evidence-pid set
  differs between sides** (6 rows).  Concrete example:
  c_personid=228114 (王淑抃), father 王圖 (123710); User MDB
  BIOG_MAIN.c_birthyear(123710)=1557 vs SQLite 1577; Access
  fires 1557+30=1587, PHP fires 1577+30=1607, diff=−20
  explained 100% by upstream birthyear drift on the same
  evidence person.
- **Algorithm evidence.** PR N pairs each rule as matched
  (or matched_minor_diff on staging-vs-subquery aggregate);
  the rule body is identical and runs correctly on both
  sides.  PR AI / AJ probe scripts walked every row and
  inferred the rule output from the evidence on each side.
- **Row evidence.** Per-personid breakdown in
  `reports/index_year_diff_minus_20_cluster_probe.json` (PR
  AI) and `reports/index_year_diff_plus_1_cluster_probe.json`
  (PR AJ).  All 14 rows accounted for; 0 unexplained.
- **Confidence.** **Supported by focused probe (PR AI + AJ).**
  Same general class as PR Z's tcode='05' candidate_php_entry_
  code_mapping_gap.  Not a CBDB User MDB algorithm bug.
- **Next action.** Forward the affected personids + their
  evidence-row mismatches to the cbdb-online-main-server /
  SQLite-snapshot-build team as upstream data-drift
  candidates.  Out-of-scope for this repo.

(Historical note: prior to PR AI/AJ this bucket was labelled
`candidate_same_rule_tie_break_or_aggregation_diff` at
confidence "medium" with the working hypothesis that a
single staging-step row pick reproduced consistently.  Both
parts of that hypothesis are now invalidated.)

### A6. `candidate_algorithm_divergence` × 5

- **Likely cause.** Row's `(php_tcode, |diff|)` signature
  matches one of K1's dormant historical probes (e.g. the
  entry-rule 2N-diff pattern for tcodes 05–10) but the
  supporting evidence row can't be reconstructed.  Almost
  certainly an evidence-row-pick disagreement: per PR N both
  sides apply `-N`, so a 2N gap most plausibly means each side
  picked a different ENTRY_DATA row whose c_years differ by 2N.
- **Algorithm evidence.** PR N rule definition matched; the
  shape signature still fires because the diff happens to
  match `±2N` numerically.
- **Row evidence.** 5 rows; K1 records the matched signature
  but flags the evidence-row reconstruction failed.
- **Confidence.** **Low–medium** — the signature is
  suggestive but not confirmatory.
- **Next action.** Same per-row evidence walk as A5 (pull
  ENTRY_DATA candidates for each side and identify which row
  was picked).

### A7. `blocked_by_runtime_priority_triage_pending` × 17

(Subset of K1's `unclassified` × 18.)

- **Likely cause.** Different *rules* picked on each side
  (i.e. PHP's `c_index_year_type_code` ≠ Access's), which
  means the divergence is at the **rule-selection / priority**
  layer rather than at the rule-body layer.  PR N has the
  rule-body comparison; the rule-selection priority order is
  encoded in `GetBirthIndexYearSQL`'s sequence of UPDATE
  statements but has not yet been walked per-row to decide
  which side's choice is intentional.
- **Algorithm evidence.** PR M dumped `frmBaseMaintenance`
  including `GetBirthIndexYearSQL` — the source IS in repo
  (`analysis/dump_data/vba/Form_frmBaseMaintenance.vb`); the
  blocker is now the per-row priority/iteration walk, not
  missing source.  PR X renamed the blocker label from
  `blocked_by_missing_frmBaseMaintenance_vba` (factually
  stale) to `blocked_by_runtime_priority_triage_pending`.
- **Row evidence.** 17 rows; K2 lists each with both sides'
  tcodes and intent.  Examples: PHP `'15'` (mother) vs Access
  `'03'` (wife) — different relationship anchors picked.
- **Confidence.** **Low** as causes for individual rows
  (haven't walked them); **high** as a category label.
- **Next action.** This is the largest remaining unknown.
  Pick the most-recurring (PHP_tcode, Access_tcode) pairs and
  walk `GetBirthIndexYearSQL`'s priority order; document
  whether each side's choice is intentional or a coverage
  miss.

### A8. Dormant historical probes

K1 emits three sub-buckets that came back at 0 rows in this
sweep:

  - `explained_by_birthyear_offset` × 0 — PR I-era hypothesis
    (Access used `c_birthyear + 59` for tcode 01); PR N showed
    runtime Access uses raw `c_birthyear`.
  - `explained_by_entry_sign_flip` × 0 — PR I-era hypothesis
    (Access `+N` vs PHP `-N`); PR N showed both use `-N`.
  - `explained_by_husband_formula` × 0 — PR I-era hypothesis
    (Access `husband.c_birthyear + 62` vs PHP `husband.c_index_
    year + 3`); PR N showed runtime Access uses
    `husband.c_index_year + 3`.

- **Confidence.** **High** that the probes are dormant in this
  data sweep (they can be reactivated if a future commit
  reintroduces those patterns).
- **Next action.** Keep as historical probes; they cost
  nothing to run.

---

## B. `c_index_addr_id` cause buckets

488 address-only diffs (478 `index_addr_only_diff` + 10
`index_both_diff`).  PR L verified separately that
`BIOG_ADDR_CODES` (the rank table) is identical between User
MDB and SQLite for all 22 addr_types — so **none of these
buckets are explained by rank-table mismatch**.  The
divergence has to come from somewhere else.

### B1. `mdb_stale_index_addr` × 412

**The headline bucket.**

- **Likely cause.** The User MDB shipped with a
  `c_index_addr_id` that was **never re-run after BIOG_ADDR_
  DATA was updated**.  PR L's recompute (rank+MAX(c_sequence)
  walked over current BIOG_ADDR_DATA) matches what SQLite
  stores; User MDB stores something else.  PHP re-runs
  weekly via the cbdb-online-main-server cron; the User MDB
  release process apparently doesn't re-run the equivalent
  Access maintenance step before shipping.
- **Algorithm evidence.** PR L's recompute on User-MDB-side
  BIOG_ADDR_DATA produces exactly SQLite's value for all 412
  rows.  The Access maintenance code that *would* re-run this
  is `frmBaseMaintenance.CmdIndexAddress_Click` (PR M dump)
  — it exists, just isn't in the release checklist.
- **Row evidence.** 412 rows in PR L's `mdb_stale_index_addr`
  bucket.
- **Confidence.** **High** — algorithm + row evidence both
  clear; the only uncertainty is whether the missing release
  step is the *only* explanation (vs. e.g. some persons
  intentionally pinned to an older addr_id), but the breadth
  argues for the simpler story.
- **Next action.** Add a release-process step to the User MDB
  build: run `frmBaseMaintenance.CmdIndexYear` then
  `CmdIndexAddress` against the DATA mdb before shipping.
  This is `candidate_release_process_step`; not yet labelled
  a confirmed bug because the maintainer may have an
  intentional reason not to re-run on each release.

### B2. `mdb_value_php_null` × 47

- **Likely cause.** User MDB has a stored `c_index_addr_id`
  that PR L's PHP recompute returns NULL for, which means PHP
  has no eligible BIOG_ADDR_DATA row of any rank-priority
  addr_type for this person.  Mirror direction of A3 / A2:
  one side's evidence chain reaches a value, the other's
  doesn't.
- **Algorithm evidence.** PR L's recompute returns NULL for
  all 47.  Either PHP is stricter on which BIOG_ADDR_DATA rows
  count, or the User MDB stored a value before the
  BIOG_ADDR_DATA row was removed/updated.
- **Row evidence.** 47 rows in PR L's `mdb_value_php_null`
  bucket.
- **Confidence.** **Medium**.
- **Next action.** Sample 5 rows; check whether the User-MDB-
  stored addr_id still exists in the User-MDB BIOG_ADDR_DATA;
  if so, why does PR L's recompute (which mirrors PHP's
  rank-priority + MAX(c_sequence)) reject it?  Could be a
  rank exclusion not represented in BIOG_ADDR_CODES.

### B3. `same_candidates_diff_winner` × 10

**Root-caused by PR S.**

- **Likely cause.** All 10 are addr_type=1 winners.  PR S's
  deep-dive walked each row's BIOG_ADDR_DATA candidate set
  and found:
  - 10 / 10 have a **MAX(c_sequence) tie** — multiple
    BIOG_ADDR_DATA rows of the same (person, addr_type) share
    the same max c_sequence.
  - 10 / 10 have **3 distinct picks** among the tied rows
    (User-MDB-stored, SQLite-stored, and PR L recompute each
    pick a different row).
  - **Shared cause**: no explicit secondary tie-break in
    either implementation; each picks whichever row its
    storage engine surfaces first (MariaDB / Microsoft JET /
    SQLite via pyodbc all order rows differently).
- **Algorithm evidence.** PR S byte-level deep-dive
  (`reports/index_addr_same_candidates_deep_dive.json`); both
  sides implement the documented rule (rank-priority +
  MAX(c_sequence)) — neither is "wrong".
- **Row evidence.** 10 / 10 rows show the MAX-tie + distinct-
  picks pattern explicitly.
- **Confidence.** **High** — root cause identified at the
  algorithm level.
- **Next action.** **Candidate algorithm improvement**: add
  an explicit secondary tie-break (e.g. MIN(c_addr_id)) to
  both implementations.  Either side could implement
  independently to reduce divergence.  Treat as
  `candidate_release_process_or_algorithm_improvement`; not
  a confirmed bug because the documented rule is satisfied
  on both sides.

### B4. `both_stale_recompute_mismatch` × 10

- **Likely cause.** Both sides store a value that doesn't
  match the rank+MAX(c_sequence) recompute on either side —
  i.e. **both are stale**.  Whatever maintenance step
  produced these is older than the current BIOG_ADDR_DATA on
  both sides.
- **Algorithm evidence.** PR L's recompute disagrees with
  both stored values.
- **Row evidence.** 10 rows.
- **Confidence.** **Medium-high** — the "both stale" framing
  is data-evidenced; what's not certain is whether they were
  stale from the same maintenance run or independently.
- **Next action.** Re-run the maintenance step on both sides
  and re-classify; if these rows then move into one of the
  cleaner buckets, the framing is confirmed.

### B5. `both_sides_match_recomputed` × 6

- **Likely cause.** Both User MDB and SQLite store values
  that match the rank+MAX(c_sequence) recompute on both
  sides — meaning **both sides agree on what the value
  should be**, and yet PR L still flagged them as differing.
  The cause is that the *stored* values differ even though
  both recomputes agree, which most plausibly means **one
  side stores a value that's the recompute of the OTHER
  side's BIOG_ADDR_DATA** (cross-pollination from a previous
  data exchange or a shared upstream).
- **Algorithm evidence.** PR L's recompute on each side gives
  the same answer; both stored values are valid recomputes
  but of different underlying datasets.
- **Row evidence.** 6 rows.
- **Confidence.** **Low–medium** — the cross-pollination
  story is plausible but speculative; could also be a stale
  release on one side stored against the other side's
  evidence at the time.
- **Next action.** Pick 2 rows; trace the BIOG_ADDR_DATA
  row-set difference between sides for those persons; see
  whether one side's stored value matches the *other* side's
  current BIOG_ADDR_DATA recompute.

### B6. `sqlite_stale_index_addr` × 2

- **Likely cause.** Mirror of B1 from the SQLite side — User
  MDB recompute matches User-stored, SQLite stored doesn't
  match SQLite recompute.  Possibly a transient state in the
  weekly PHP rebuild (a row that was in flight when the
  snapshot was taken).
- **Algorithm evidence.** PR L's recompute on SQLite-side
  BIOG_ADDR_DATA disagrees with what SQLite stored.
- **Row evidence.** 2 rows.
- **Confidence.** **Medium** — small sample.
- **Next action.** Re-pull a fresh SQLite snapshot and see
  whether these 2 rows resolve.  If they persist, escalate
  to cbdb-online-main-server.

### B7. `mdb_null_php_value` × 1

- **Likely cause.** Single row where User MDB stored NULL
  but PHP has a value.  PR L's recompute on User-MDB-side
  BIOG_ADDR_DATA returns NULL, so the User MDB has no
  eligible row for this person; PHP must be drawing on a
  BIOG_ADDR_DATA row that exists on the SQLite side but not
  on the User MDB side.  Source-data drift.
- **Algorithm evidence.** PR L's per-side recompute
  divergence.
- **Row evidence.** 1 row.
- **Confidence.** **Medium-high**.
- **Next action.** Compare the two sides' BIOG_ADDR_DATA for
  the affected personid — confirm the missing row.

### B8. `unclassified` × 0

PR L bucketed every row.  No residual.

---

## What we're NOT claiming

- **No bucket is labelled a confirmed CBDB bug.**  Each
  bucket either points at an algorithm-improvement candidate
  (B3), a release-process candidate (B1), a coverage-gap
  candidate (A2 / A3 / A6), or an architectural seam needing
  a runtime trace (A4 / A5 / A7).
- **No claim that the totals are exhaustive of all possible
  cause classes.**  The big unknown remains source-data drift
  in BIOG_ADDR_DATA / ENTRY_DATA / KIN_DATA / NIAN_HAO that
  isn't compared head-to-head between sides.  Several of the
  "low" / "medium" confidence next-actions above would
  produce evidence narrowing this further.

## Suggested next investigations (in rough priority order)

1. **B1 release-process step** — adding the `CmdIndexYear` →
   `CmdIndexAddress` step to the User MDB build closes the
   single largest bucket (412 rows) at zero engineering cost.
   Confirm with the maintainer.
2. **B3 secondary tie-break** — small algorithm improvement
   on both sides; closes 10 rows and stops the bucket
   reappearing on every snapshot.
3. ~~**A2 `'05'` × 7 entry-code-mapping check**~~ — **done by
   PR Z**: 6 confirmed mapping gaps, 1 confirmed
   `ENTRY_DATA.c_year = 0` gap; all 7 PHP-side upstream data
   issues.  Forward to cbdb-online-main-server team.
4. **A5 `-20` cluster deep-dive** — single staging-step row
   pick may explain 9 rows across 4 signature groups.
5. **A7 priority-order triage** — largest remaining unknown
   (17 rows); will need a proper runtime trace.

## Re-running

```
python analysis/compare_index_year_rules.py
python analysis/classify_index_year_drift_by_rule.py
python analysis/triage_index_year_drift_groups.py
python analysis/classify_index_addr_drift.py
python analysis/deep_dive_addr_same_candidates.py
```

Each writes to its own JSON; this analysis is a static
synthesis that follows from those outputs.  Re-derive after
any fresh SQLite snapshot or User MDB refresh.
