# LookAtAssociations × CmdNeo4j verification probe — post target-column rewrite patch

**Date:** 2026-05-07  ·  **Branch:** `driver/associations-cmdneo4j-c-addr-type-rewrite`

Verification probe for the narrow scoped driver patch `_rewrite_associations_cmdneo4j_target_column` in `tests/cbdb_driver/vba_session.py`.  The patch rewrites the single literal anchor `c_index_addr_type_code, c_female` -> `c_addr_type, c_female` inside `Form_LookAtAssociations.CmdNeo4j_Click` to work around canonical Issue #23 (P1_visible_crash).  This probe verifies the patch holds under the same Associations matrix fixture used by PR #112 and looks for any new blocker behind the JET 3061 layer.

## Setup

- **Fixture:** `assoc_437_unfiltered` (matrix `_make_assoc_fixtures` first fixture; picker_ids = [437]; controls = {'FrameFilterYears': 1})
- **Chain:** `CmdQuery,CmdNeo4j` via Form.Tag, directory mode
- **Watchdog:** records and dismisses any MsgBox the driver missed; each entry is a runtime signal we surface (not silently swallow).
- **click_via_timer cap:** 180 s  ·  **outer cap:** 300 s

## Raw observed facts

- **chain_elapsed_sec:** 10.04
- **file_count:** 8
- **chain_observed_done:** True
- **click_via_timer_returned:** 11867
- **msgbox_observed_via_watchdog_count:** 5
- **total_wall_elapsed_sec:** 19.03

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

| +t (s) | msg_text |
|---:|---|
| 9.53 | `Kinship code records = 1` |
| 9.94 | `Literary genre code records = 0` |
| 10.36 | `Institution code records = 0` |
| 10.78 | `Occasion code records = 6` |
| 11.19 | `Topic code records = 2` |

## Q1-Q4 answers

**Q1 — JET 3061 'c_index_addr_type_code' :ERR disappeared?**

- has_jet_3061_marker: **False** (expected False)
- all_err_markers_observed: `[]`

If has_jet_3061_marker is False, the patch successfully prevented the JET 3061 'unknown field name: c_index_addr_type_code' that PR #112's probe observed.  If True, the rewrite did not take effect (anchor mismatch / scoping bug / etc.) and the patch needs investigation.

**Q2 — chain wrote files (was 0 in PR #112)?**

- file_count = 8
- chain_observed_done = True
- chain_elapsed_sec = 10.04

PR #112 observed 0 files; this probe expects file_count >= 1 if the rewrite took effect AND the rest of the chain runs cleanly.

**Q3 — file shape:**

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

**Q4 — next blocker exposed?**

- outcome_bucket = `patch_verified_chain_clean`
- other_err_markers = `[]`
- msgbox_observed_count = 5

If outcome == patch_verified_chain_clean: no next blocker on this fixture.  If outcome == patch_partial_jet_3061_still_observed: rewrite did not take effect.  If outcome == next_blocker_exposed_different_err: rewrite worked but a different runtime error appeared.  If outcome == patch_resolved_jet_3061_but_zero_files: rewrite worked but no file got written for some other reason.

## Verdict: `patch_verified_chain_clean`

Patch verified.  Chain produced 8 files in 10.04s with no :ERR markers in ZZ_TEST_DEBUG (specifically no JET 3061 'c_index_addr_type_code' marker).  This is the strongest available signal that the target-column rewrite is the only blocker on this fixture for LookAtAssociations × CmdNeo4j.  Note: this probe does NOT open a coverage PR; the next step is a separate coverage PR brief.

## Markers (timeline, 14 entries)

  - `+  0.00s` constructing_session
  - `+  5.44s` session_opened_attempt_1
  - `+  5.79s` filedialog_patched
  - `+  6.53s` form_opened
  - `+  6.54s` picker_seeded_1_codes
  - `+  6.54s` fixture_seeded
  - `+  6.54s` form_tag_set_chain_CmdQuery_CmdNeo4j
  - `+  6.54s` chain_fire_t_start
  - `+ 11.58s` click_via_timer_returned_11867
  - `+ 16.58s` chain_quiescent_files_8_stable_for_5s
  - `+ 16.58s` chain_elapsed_10.04s
  - `+ 16.60s` files_inventoried_8
  - `+ 16.60s` row_counts_captured
  - `+ 16.60s` zz_test_debug_captured

## Constraints honoured per brief

- ✅ Did NOT open a coverage PR
- ✅ Driver patch is narrow-scoped: only `Form_LookAtAssociations`, only `CmdNeo4j_Click`, only the literal `c_index_addr_type_code, c_female` -> `c_addr_type, c_female` rewrite
- ✅ Did NOT introduce a generic SQL text rewriting policy
- ✅ Did NOT touch other forms or the SELECT-side qualified `BIOG_MAIN.c_index_addr_type_code` reference
- ✅ Did NOT change canonical reports / issue severity / README / triage docs
- ✅ Did NOT change tests/