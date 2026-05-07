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

PR #112 observed 0 files; this probe expects file_count >= 1 if the rewrite took effect AND the rest of the chain runs to its natural end.

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

- outcome_bucket = `patch_resolved_issue23_but_exposed_msgbox_blocker`
- other_err_markers = `[]`
- msgbox_watchdog_count = 5  (REAL pop-up dialogs the driver's literal-only generic neutralizer could not catch; each one would block unattended coverage)
- zz_test_debug_msgbox_marker_count = 1  (generic-neutralizer footprints for literal `MsgBox "<lit>"` lines; NOT pop-up dialogs)

**The two MsgBox signal layers are distinct:**

- *Watchdog-dismissed dialogs:* REAL pop-up dialogs that the driver's generic literal-only `MsgBox "<lit>"` neutralizer could not catch — typically concat-form `MsgBox "<lit>" + Trim(Str(...))` calls.  Watchdog dismisses them so the probe does not hang, but for unattended coverage a dialog is a hard blocker.

- *ZZ_TEST_DEBUG :MSGBOX markers:* INSERT rows the driver's generic literal neutralizer wrote BEFORE the runtime would see the original `MsgBox "<lit>"` line as a dialog.  These do NOT correspond to dialogs that ever popped — they are the neutralizer's footprint, not a blocker.  Distinct from the line-1035 early-bail marker (which would only appear if the RecordCount=0 bail fired; on this fixture it does not, so the :MSGBOX marker present here is the terminal `MsgBox "Finished saving to Neo4j"` neutralizer footprint, indicating the chain reached its end).

Bucket meanings (mutually exclusive, first match wins):
  - patch_verified_chain_clean: file_count >= 1 AND no :ERR AND zero watchdog dialogs.  Strict.
  - patch_resolved_issue23_but_exposed_msgbox_blocker: file_count >= 1 AND no :ERR BUT >= 1 watchdog dialog dismissed.  Issue #23 JET 3061 is gone, but a downstream debug-MsgBox layer is exposed and would block unattended coverage.
  - patch_partial_jet_3061_still_observed: JET 3061 :ERR still present.
  - next_blocker_exposed_different_err: a :ERR from a different class.
  - patch_resolved_jet_3061_but_zero_files: no :ERR but no files either.

## Verdict: `patch_resolved_issue23_but_exposed_msgbox_blocker`

**Two-layer outcome.**

Layer 1 (resolved): the Issue #23 JET 3061 target-column mismatch ('unknown field name: c_index_addr_type_code') has been neutralized by the narrow per-form rewrite.  Direct evidence:
  - 0 :ERR markers in ZZ_TEST_DEBUG (no `LookAtAssociations:ERR ... c_index_addr_type_code` row).
  - file_count = 8 (was 0 in PR #112's pre-patch probe).
  - ZZ_SCRATCH_PEOPLE populated by the INSERT that previously failed.
  - Chain reached its natural end (chain_observed_done = True).

Layer 2 (newly exposed): a downstream debug-MsgBox layer is now live and would block unattended coverage.  Direct evidence:
  - watchdog dismissed 5 dialog(s) during the chain run.  Sample: ['Kinship code records = 1', 'Literary genre code records = 0', 'Institution code records = 0', 'Occasion code records = 6', 'Topic code records = 2']
  - These are concat-form `MsgBox "<lit>" + Trim(Str(...))` calls inside CmdNeo4j_Click that the driver's literal-only generic neutralizer cannot match.  Same shape as the AssocPairs CmdNeo4j debug-MsgBox layer that PR #109 suppressed via `_suppress_assocpairs_cmdneo4j_debug_msgbox` — but in a different form, so PR #109's per-form patch does NOT cover it.

This is a NEW per-host observation that PR #112's pre-patch probe could not surface, because that probe's chain bailed at the JET 3061 BEFORE reaching the debug-MsgBox lines.

**Implication for next step:** coverage PR is NOT directly unblocked by this rewrite alone.  For unattended coverage, one of two paths is required (each its own brief, NOT this PR):
  (a) extend `_PER_FORM_CMDGIS_PATCHES` with a `Form_LookAtAssociations.CmdNeo4j_Click` debug-MsgBox suppress (mirror of PR #109's AssocPairs-side patch); OR
  (b) the cross-form coverage test inlines a watchdog dismisser (this probe demonstrates the watchdog approach works in principle).

Note on the single ZZ_TEST_DEBUG `:MSGBOX` marker present here: that row was written by the driver's generic literal-only MsgBox neutralizer when it rewrote the terminal `MsgBox "Finished saving to Neo4j"` line (literal form, neutralizable) BEFORE the runtime would see it.  It is NOT the line-1035 RecordCount=0 early-bail marker (which would only appear if the bail fired; on this fixture it does not — ZZ_SCRATCH_P_ASSOC has 8087 rows).  The :MSGBOX marker is a clean-end signal, NOT a blocker — distinct layer from the watchdog dismissals above.

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