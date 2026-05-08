# LookAtPlace × CmdNeo4j verification probe — post tRstPeople SELECT projection patch

**Date:** 2026-05-08  ·  **Branch:** `driver/place-cmdneo4j-trstpeople-projection-rewrite`

Verification probe for the narrow scoped driver patch `_rewrite_place_cmdneo4j_trstpeople_projection` in `tests/cbdb_driver/vba_session.py`.  The patch extends the `tRstPeople` SELECT projection inside `Form_LookAtPlace.CmdNeo4j_Click` from 4 cols (only `ZZ_SCRATCH_P_TEXT` cols) to 7 cols (also `DYNASTIES.c_dynasty`, `DYNASTIES.c_dynasty_chn`, `BIOG_MAIN.c_female`).  Anchor:

```
old: "ZZ_SCRATCH_P_TEXT.c_index_year "  (with trailing space)
new: "ZZ_SCRATCH_P_TEXT.c_index_year, DYNASTIES.c_dynasty, DYNASTIES.c_dynasty_chn, BIOG_MAIN.c_female "
```

Verifies under the same matrix Place fixture used by PR #120, looking for whether the JET 3265 disappears and whether the chain now writes files.

## Setup

- **Fixture:** `place_addr_7213` (matrix `_make_place_fixtures` first fixture; same as PR #120)
- **Pre-fixture step:** `set_control("LookAtPlace", "TabPlaces", 0)` (mirrors the cross-form CmdNeo4j test's special-case for Place)
- **Chain:** `CmdQuery,CmdNeo4j` via Form.Tag, directory mode
- **Watchdog:** records and dismisses any MsgBox the driver missed; each entry is a runtime signal we surface (not silently swallow).
- **click_via_timer cap:** 180 s  ·  **outer cap:** 300 s

## Raw observed facts

- **chain_elapsed_sec:** 11.04
- **file_count:** 6
- **chain_observed_done:** True
- **click_via_timer_returned:** 5962
- **msgbox_watchdog_count:** 0
- **total_wall_elapsed_sec:** 20.17

### Scratch row counts (post-chain)

- `ZZ_SCRATCH_STATUS`: 17
- `ZZ_SCRATCH_PLACE_PEOPLE`: 5764
- `ZZ_SCRATCH_PLACE_AGG`: 5764
- `ZZ_SCRATCH_PEOPLE`: 5764
- `ZZ_SCRATCH_ADDR`: 1
- `ZZ_TEST_DEBUG`: 3

### ZZ_TEST_DEBUG content

- `LookAtPlace:ENTER`
- `LookAtPlace:MSGBOX`
- `LookAtPlace:DONE`

### Watchdog MsgBox observations

(none observed)

## Q1-Q4 answers

**Q1 — JET 3265 'Item not found in this collection.' :ERR disappeared?**

- has_jet_3265_marker: **False** (expected False)
- all_err_markers_observed: `[]`

If has_jet_3265_marker is False, the patch successfully prevented the JET 3265 'Item not found in this collection.' that PR #120's probe observed.  If True, the rewrite did not take effect.

**Q2 — chain wrote files (was 0 in PR #120)?**

- file_count = 6
- chain_observed_done = True
- chain_elapsed_sec = 11.04

PR #120 observed 0 files; this probe expects file_count >= 1 if the rewrite took effect AND the rest of the chain runs to its natural end.

**Q3 — file shape:**

| name | size | n_cols | first_col | data_rows |
|---|---:|---:|---|---:|
| f12.out.csv | 409721 | 9 | `PersonID` | 5961 |
| f15.out.csv | 4383 | 4 | `PersonPlaceRelCode` | 103 |
| f18.out.csv | 199 | 3 | `IndexAddrTypeCode` | 4 |
| f3.out.csv | 233329 | 6 | `NameID` | 5764 |
| f6.out.csv | 84255 | 3 | `NameID` | 5642 |
| f9.out.csv | 14476 | 5 | `PlaceID` | 286 |

**Q4 — next blocker exposed?**

- outcome_bucket = `patch_verified_chain_clean`
- other_err_markers = `[]`
- msgbox_watchdog_count = 0

If outcome == patch_verified_chain_clean: no next blocker on this fixture.  If outcome == patch_partial_jet_3265_still_observed: rewrite did not take effect.  If outcome == next_blocker_exposed_runtime_err: rewrite worked but a different runtime error class appeared.  If outcome == patch_resolved_issue24_but_exposed_msgbox_blocker: rewrite worked AND chain wrote files BUT a watchdog-dismissed dialog appeared (analog of PR #116's MsgBox-blocker layer for Associations CmdNeo4j).  If outcome == patch_resolved_jet_3265_but_zero_files: rewrite worked AND no :ERR but no files either.

**Pre-chain observations (preserved, not silenced):**

- set_control failures during fixture seeding (2 entries):
    - `set_control_ChkAssoc_fail: com_error(-2147352567, 'Exception occurred.', (0, None, "Welcome to CBDB! can't find the field 'ChkAssoc' referred to in your expression.", 'acmain11.chm', 11730, -2146825823), None)`
    - `set_control_ChkPosting_fail: com_error(-2147352567, 'Exception occurred.', (0, None, "Welcome to CBDB! can't find the field 'ChkPosting' referred to in your expression.", 'acmain11.chm', 11730, -2146825823), None)`

These are pre-chain failures from the matrix _make_place_fixtures controls dict (ChkAssoc / ChkPosting NOT controls on Form_LookAtPlace). Preserved per brief — NOT silently swallowed. They do NOT cause the JET 3265 (which fires later inside CmdNeo4j_Click body) and do NOT cause any post-patch failure.  Fixing the fixture is out-of-scope for this probe.

## Verdict: `patch_verified_chain_clean`

**Strict clean.**  Chain produced 6 files in 11.04s with 0 :ERR markers AND 0 watchdog-dismissed dialogs.  This is the strongest available signal that the SELECT projection rewrite is the only blocker on this fixture for LookAtPlace × CmdNeo4j.  Per the brief, do NOT open a coverage PR off this verdict alone — the coverage PR is a separate brief.

## Markers (timeline, 17 entries)

  - `+  0.00s` constructing_session
  - `+  5.43s` session_opened_attempt_1
  - `+  5.69s` filedialog_patched
  - `+  6.50s` form_opened
  - `+  6.51s` tab_places_set_to_0
  - `+  6.53s` set_control_ChkAssoc_fail: com_error(-2147352567, 'Exception occurred.', (0, None, "Welcome to CBDB! can't find the field 'ChkAssoc' referred to in your expression.", 'acmain11.chm', 11730, -2146825823), None)
  - `+  6.53s` set_control_ChkPosting_fail: com_error(-2147352567, 'Exception occurred.', (0, None, "Welcome to CBDB! can't find the field 'ChkPosting' referred to in your expression.", 'acmain11.chm', 11730, -2146825823), None)
  - `+  6.54s` addr_picker_seeded_1_codes
  - `+  6.54s` fixture_seeded
  - `+  6.54s` form_tag_set_chain_CmdQuery_CmdNeo4j
  - `+  6.54s` chain_fire_t_start
  - `+ 12.58s` click_via_timer_returned_5962
  - `+ 17.58s` chain_quiescent_files_6_stable_for_5s
  - `+ 17.58s` chain_elapsed_11.04s
  - `+ 17.58s` files_inventoried_6
  - `+ 17.59s` row_counts_captured
  - `+ 17.59s` zz_test_debug_captured

## Constraints honoured per brief

- ✅ Did NOT open a coverage PR
- ✅ Driver patch is narrow-scoped: only `Form_LookAtPlace`, only `CmdNeo4j_Click`, only the `tRstPeople` SELECT projection literal
- ✅ Did NOT introduce a generic SQL rewrite policy
- ✅ Did NOT touch other forms or other recordsets in the same sub
- ✅ Did NOT change canonical reports / issue severity / README / triage docs
- ✅ Did NOT change tests/
- ✅ Pre-chain `ChkAssoc / ChkPosting` `set_control` failures preserved as observations (per brief), NOT silenced