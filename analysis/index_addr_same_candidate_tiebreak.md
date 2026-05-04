# Index-addr same-candidate tie-break note (PR AG)

PR S deep-dived the 10 `same_candidates_diff_winner` rows from
PR L's `c_index_addr_id` classification and found a clean root
cause: **all 10 have a `MAX(c_sequence)` tie among the
candidate BIOG_ADDR_DATA rows**, and the three pipelines
(User MDB Access, SQLite PHP, our PR L recompute) each pick a
different row from the tied set because **none of the three
implementations specifies an explicit secondary tie-break**.

This note formalises the per-row picture and explores what a
deterministic secondary tie-break would do — without
prescribing a fix.  The per-row data comes from
`reports/index_addr_same_candidates_deep_dive.json` (PR S).

## Per-row table

The script that produced this (`analysis/render_index_addr_
tiebreak_note.py`) walks each row, lists the tied candidate
addr_ids, and shows what each side picked.  The
"would-deterministic-X-pick" columns answer: if the tie-break
were "MIN(c_addr_id)" / "MAX(c_addr_id)", which row would each
pipeline pick?

(Render order = personid ascending.)

See `reports/index_addr_same_candidate_tiebreak.json` for the
machine-readable per-row + per-tie-break-rule projection.

## Candidate tie-break options

Three candidate secondary keys, ordered by simplicity:

  1. **`MIN(c_addr_id)`** — smallest addr_id wins ties.  Most
     common convention; deterministic across storage engines.
     Stable as long as `c_addr_id` is never re-numbered.
  2. **`MAX(c_addr_id)`** — largest addr_id wins.  Same
     properties as #1, opposite direction.  Slightly more
     biased toward "newest known address" if addr_ids are
     issued chronologically.
  3. **First-by-c_created_date** — earliest BIOG_ADDR_DATA
     row by `c_created_date` wins.  More semantically
     meaningful but requires every staging path to surface
     `c_created_date`, which the current `ZZ_SCRATCH_*`
     tables don't all carry.

PR S's "candidate mitigation" wording stays: this is a
**candidate algorithm improvement**, not a confirmed bug.
The rule both sides implement (rank-priority + MAX(c_sequence))
is documented and correctly implemented; the tie-break is just
an under-specification.

## What this PR does NOT do

- Doesn't propose a final fix.  Picking between the three
  candidate tie-breaks is a maintainer decision (#1 is the
  obvious default but the maintainer may have a reason to
  prefer #3).
- Doesn't promote the bucket to a confirmed CBDB bug.
- Doesn't change `reports/index_drift_cause_summary.json`
  beyond what PR S/Y already set.

## Re-running

```
python analysis/deep_dive_addr_same_candidates.py    # PR S
python analysis/render_index_addr_tiebreak_note.py    # this PR
```

Pure pyodbc.

## Per-row data

| personid | tied addr_ids | User stored | SQLite stored | User recompute | SQLite recompute | MIN-rule | MAX-rule |
|---:|---|---:|---:|---:|---:|---:|---:|
| 149797 | 400539, 14692, 15791 | 15791 | 400539 | 400539 | 14692 | 14692 | 400539 |
| 151289 | 16737, 14692 | 14692 | 16737 | 16737 | 14692 | 14692 | 16737 |
| 153923 | 15910, 14755 | 14755 | 15910 | 15910 | 14755 | 14755 | 15910 |
| 154902 | 15285, 14750, 14932 | 14932 | 15285 | 15285 | 14750 | 14750 | 15285 |
| 155767 | 400467, 15691 | 15691 | 400467 | 400467 | 15691 | 15691 | 400467 |
| 156434 | 400398, 14965 | 14965 | 400398 | 400398 | 14965 | 14965 | 400398 |
| 157177 | 15768, 400402, 14692 | 14692 | 400402 | 15768 | 14692 | 14692 | 400402 |
| 160450 | 15970, 14692 | 14692 | 15970 | 15970 | 14692 | 14692 | 15970 |
| 174879 | 14507, 14492 | 14492 | 14507 | 14507 | 14492 | 14492 | 14507 |
| 177165 | 15237, 14693, 14764, 15003 | 15003 | 15237 | 15237 | 14693 | 14693 | 15237 |

### Convergence under candidate tie-breaks

| Rule | Rows where both sides converge |
|---|---:|
| `MIN_addr_id` | 10 / 10 |
| `MAX_addr_id` | 10 / 10 |

Both candidate rules deterministically converge all 10 rows; choice is a style call.