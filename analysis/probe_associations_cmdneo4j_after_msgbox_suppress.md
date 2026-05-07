# LookAtAssociations × CmdNeo4j verification probe — post debug-MsgBox suppress patch

**Date:** 2026-05-07  ·  **Branch:** `driver/associations-cmdneo4j-debug-msgbox-suppress`

Verification probe for the narrow scoped driver patch `_suppress_associations_cmdneo4j_debug_msgbox` in `tests/cbdb_driver/vba_session.py`.  The patch comments out the 5 concat-form debug MsgBox calls inside `Form_LookAtAssociations.CmdNeo4j_Click` (lines 1130 / 1232 / 1315 / 1398 / 1481 in the current dump).  These are the watchdog-dismissed dialogs PR #116's post-c_addr_type-rewrite verification probe surfaced (outcome = patch_resolved_issue23_but_exposed_msgbox_blocker).

## Setup

- **Fixture:** `assoc_437_unfiltered` (matrix `_make_assoc_fixtures` first fixture; same as PR #112 and PR #116)
- **Chain:** `CmdQuery,CmdNeo4j` via Form.Tag, directory mode
- **Watchdog:** records and dismisses any MsgBox; after this patch, count is expected to drop from 5 (PR #116 baseline) to 0.
- **click_via_timer cap:** 180 s  ·  **outer cap:** 300 s

## Raw observed facts

- **chain_elapsed_sec:** 8.04
- **file_count:** 8
- **chain_observed_done:** True
- **click_via_timer_returned:** 11867
- **msgbox_watchdog_count:** 0  (PR #116 baseline: 5; expected after this patch: 0)
- **total_wall_elapsed_sec:** 16.86

### Scratch row counts (post-chain)

- `ZZ_SCRATCH_ASSOC`: 11867
- `ZZ_SCRATCH_P_ASSOC`: 8087
- `ZZ_SCRATCH_PEOPLE`: 8088
- `ZZ_TEST_DEBUG`: 3

### ZZ_TEST_DEBUG content

- `LookAtAssociations:ENTER`
- `LookAtAssociations:MSGBOX`
- `LookAtAssociations:DONE`

### Watchdog MsgBox observations

(none observed — suppress patch took effect)

## Q1-Q4 answers

**Q1 — watchdog dialogs to zero?**

- watchdog_count = 0 (expected: 0; PR #116 baseline: 5)
- dialogs observed: (none)

PR #116's pre-suppress probe observed 5 watchdog-dismissed dialogs (concat-form `MsgBox "<lit>" + Trim(Str(...))`).  This probe expects 0 — the 5 prefixes are now comment-prefixed by the per-form suppress.

**Q2 — :ERR markers still zero?**

- err_marker_count = 0

PR #116's post-c_addr_type-rewrite probe had 0 :ERR markers; this probe expects the same (no new runtime error class introduced by the MsgBox suppress).

**Q3 — file shape stable vs PR #116?**

- file_count = 8 (PR #116: 8)

| name | size | n_cols | first_col | data_rows |
|---|---:|---:|---|---:|
| f12.out.csv | 692366 | 13 | `Person1_ID` | 11867 |
| f15.out.csv | 134 | 4 | `AssociationCode` | 1 |
| f18.out.csv | 51 | 3 | `KinshipCode` | 1 |
| f21.out.csv | 148 | 3 | `OccasionCode` | 4 |
| f24.out.csv | 47 | 3 | `TopicCode` | 1 |
| f3.out.csv | 309113 | 6 | `nameID` | 8088 |
| f6.out.csv | 59462 | 5 | `placeID` | 1263 |
| f9.out.csv | 46975 | 4 | `nameID` | 3102 |

PR #116 saw an 8-file shape (People / Places / PeoplePlaces / PeopleAssociations / AssociationCodes / KinshipCodes / OccasionCodes / TopicCodes).  This probe expects the same family — the MsgBox suppress should not change the SaveAs blocks themselves.

**Q4 — next blocker exposed?**

- outcome_bucket = `patch_verified_chain_clean`

patch_verified_chain_clean = no next blocker on this fixture; coverage PR is feasible.  msgbox_suppress_partial_dialogs_still_observed = one or more of the 5 prefixes did not match.  next_blocker_exposed_runtime_err = a runtime :ERR appeared (regression vs PR #116).

## Verdict: `patch_verified_chain_clean`

Strict clean: chain produced 8 files in 8.04s with 0 :ERR markers AND 0 watchdog-dismissed dialogs.  This is the strongest available signal that the combined c_addr_type rewrite (PR #116) + this MsgBox suppress are sufficient driver-side workarounds for unattended LookAtAssociations × CmdNeo4j coverage on this fixture.  Per the brief, do NOT open a coverage PR off this verdict alone — the coverage PR is a separate brief.

## Markers (timeline, 14 entries)

  - `+  0.00s` constructing_session
  - `+  5.31s` session_opened_attempt_1
  - `+  5.66s` filedialog_patched
  - `+  6.38s` form_opened
  - `+  6.39s` picker_seeded_1_codes
  - `+  6.39s` fixture_seeded
  - `+  6.40s` form_tag_set_chain_CmdQuery_CmdNeo4j
  - `+  6.40s` chain_fire_t_start
  - `+  9.43s` click_via_timer_returned_11867
  - `+ 14.43s` chain_quiescent_files_8_stable_for_5s
  - `+ 14.43s` chain_elapsed_8.04s
  - `+ 14.45s` files_inventoried_8
  - `+ 14.45s` row_counts_captured
  - `+ 14.45s` zz_test_debug_captured

## Constraints honoured per brief

- ✅ Did NOT open a coverage PR
- ✅ Driver patch is narrow-scoped: only `Form_LookAtAssociations`, only `CmdNeo4j_Click`, only the 5 concat-form debug MsgBox prefixes
- ✅ Did NOT touch `Bad file Name.` (file-save errors), `Err.Description` (error trap), or `Finished saving to Neo4j` (terminal — already neutralized by generic literal rewriter)
- ✅ Did NOT introduce a generic auto-dismiss policy
- ✅ Did NOT touch other forms
- ✅ Did NOT change canonical reports / issue severity / README / triage docs
- ✅ Did NOT change tests/