# AssociationPairs × CmdNeo4j verification probe — post MsgBox suppress patch

**Date:** 2026-05-07  ·  **Branch:** `driver/assocpairs-cmdneo4j-msgbox-suppress`

Verification probe for the narrow scoped driver patch `_suppress_assocpairs_cmdneo4j_debug_msgbox` in `tests/cbdb_driver/vba_session.py`.  The patch comments out the 6 unconditional debug MsgBox calls in `Form_LookAtAssociationPairs.CmdNeo4j_Click` (lines 1069 / 1151 / 1234 / 1317 / 1400 / 1470) — the confirmed blocker identified by PR AX's probe.  This probe verifies the patch holds under the same 1×3 fixture and looks for any new blocker that may have been hidden behind the MsgBox layer.

## Setup

- **Fixture:** 1×3 known-edged (TxtID1=1, TxtID2=3, FrameFilterYears=1, ChkKinship=0, Chk2Nodes=0)
- **Chain:** `CmdQuery,CmdNeo4j` via Form.Tag, directory mode
- **Watchdog:** records (and dismisses to keep the probe moving) any MsgBox the patch missed.  Each observation is a patch-correctness failure.
- **click_via_timer cap:** 120 s  ·  **outer cap:** 240 s

## Raw observed facts

- **chain_elapsed_sec:** 5.53
- **file_count:** 6
- **chain_observed_done:** True
- **msgboxes_observed_count:** 0  (expected: 0)
- **click_via_timer_returned:** 5
- **total_wall_elapsed_sec:** 15.46

## Q1-Q3 answers

**Q1 — Is the terminal Finished-saving path now fully observable?**

- msgboxes_observed_count = 0 (expected 0)
- chain_observed_done = True
- chain_elapsed_sec = 5.53
With all 6 unconditional MsgBoxes commented out, the terminal MsgBox is no longer the observable signal — chain quiescence (file count stable >= 5s) AND zero MsgBox dialogs observed AND no runtime exception are the new completeness criteria.  All three present means the terminal path is fully reached.

**Q2 — Is the file set stable across the full chain?**

- file_count = 6
- names = ['f12.out.csv', 'f15.out.csv', 'f18.out.csv', 'f3.out.csv', 'f6.out.csv', 'f9.out.csv']
Probe expects file_count >= 6 (matching the previous probe's observed set); observed = 6.

**Q3 — Does removing MsgBox layer expose a next blocker?**

- outcome bucket = `patch_verified_chain_clean`
If outcome == patch_verified_chain_clean, no next blocker exposed by this fixture.  If outcome == patch_partial_msgbox_still_observed, the scoped patch missed one of the 6 — see msgboxes_observed list.  If outcome == next_blocker_exposed_no_msgbox_but_chain_unhealthy, MsgBox layer is gone but a NEW blocker shows up (empty file set, slow chain, exception).

## Verdict: `patch_verified_chain_clean`

Patch verified.  Chain produced 6 files in 5.53s with zero MsgBox dialogs observed.  This is the strongest available signal that the 6 unconditional debug MsgBox calls are the only blocker for unattended coverage on this fixture.  Note: this probe does NOT open a coverage PR; the next step is a separate coverage PR brief.

## MsgBoxes observed (must be empty for clean verify)

(none observed — scoped patch held)

## Per-file detail

| name | size | n_cols | first_col | data_rows | header_preview |
|---|---:|---:|---|---:|---|
| f12.out.csv | 490 | 13 | `Person1_ID` | 5 | `Person1_ID,Person2_ID,Association_Code,Kin_ID,Kin_Code,AssocKin_ID,AssocKin_Code` |
| f15.out.csv | 186 | 3 | `AssociationCode` | 5 | `AssociationCode,AssociationTrans,AssociationHZ` |
| f18.out.csv | 1322 | 3 | `KinshipCode` | 96 | `KinshipCode,KinshipTrans,KinshipHZ` |
| f3.out.csv | 164 | 6 | `nameID` | 4 | `nameID,nameHZ,namePY,indexyear,dynasty,sex` |
| f6.out.csv | 180 | 5 | `placeID` | 3 | `placeID,placePY,placeHZ,placeX,placeY` |
| f9.out.csv | 220 | 5 | `nameID` | 3 | `nameID,placeID,personPlaceCode,personPlaceTrans,personPlaceHZ` |

## Scratch row counts (post-chain)

- `ZZ_SOCIAL_NETWORK`: 5
- `ZZ_SCRATCH_PEOPLE`: 3
- `ZZ_SCRATCH_P_ASSOC`: 4
- `ZZ_KIN_LIST_TMP`: 132
- `ZZ_TEST_DEBUG`: 2

## ZZ_TEST_DEBUG content

- `LookAtAssociationPairs:ENTER`
- `LookAtAssociationPairs:DONE`

## Markers (timeline, 13 entries)

  - `+  0.00s` constructing_session
  - `+  6.26s` session_opened_attempt_1
  - `+  6.63s` filedialog_patched
  - `+  7.30s` form_opened
  - `+  7.37s` fixture_controls_set
  - `+  7.38s` form_tag_set_chain_CmdQuery_CmdNeo4j
  - `+  7.38s` chain_fire_t_start
  - `+  7.91s` click_via_timer_returned_5
  - `+ 12.91s` chain_quiescent_files_6_stable_for_5s
  - `+ 12.91s` chain_elapsed_5.53s
  - `+ 12.91s` files_inventoried_6
  - `+ 12.91s` row_counts_captured
  - `+ 12.91s` zz_test_debug_captured

## Constraints honoured

- ✅ Did NOT open a coverage PR
- ✅ Driver patch is narrow-scoped: only `Form_LookAtAssociationPairs.CmdNeo4j_Click`, only the 6 unconditional debug MsgBox lines
- ✅ Did NOT introduce a generic auto-dismiss policy
- ✅ Did NOT touch unrelated forms or other MsgBox calls in CmdNeo4j_Click
- ✅ Did NOT change README / canonical reports / issue severity
- ✅ Did NOT change tests/