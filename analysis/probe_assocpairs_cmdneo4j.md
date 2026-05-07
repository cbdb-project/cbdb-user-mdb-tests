# AssociationPairs × CmdNeo4j probe (probe-first investigation)

**Date:** 2026-05-06  ·  **Branch:** `probe/assocpairs-cmdneo4j` (off main `bd6337f`)

Per the export-gap triage refresh (`analysis/export_gap_triage_plan.md`, Refresh 2026-05-06 second), AssociationPairs × CmdNeo4j is Rank 1 probe-first candidate after the SetFocus driver patch landed (commits `3bb69ef` + `0c0eaf1`).  This probe answers Q1-Q5 from the brief BEFORE any coverage PR is opened.

## Static pre-analysis

Before running, static read of `analysis/dump/vba/Form_LookAtAssociationPairs.vb` found **6 unconditional blocking `MsgBox` calls** in `CmdNeo4j_Click`:

| Line | Message |
|---:|---|
| 1069 | `MsgBox "Kinship code records = ..."` |
| 1151 | `MsgBox "Literary genre code records = ..."` |
| 1234 | `MsgBox "Institution code records = ..."` |
| 1317 | `MsgBox "Occasion code records = ..."` |
| 1400 | `MsgBox "Topic code records = ..."` |
| 1470 | `MsgBox "Finished saving to Neo4j"` |

These are debug artifacts left in the production VBA — not behind any `If` conditional.  The probe auto-dismisses them via pywinauto to allow the full chain to complete.

## Setup

- **Fixture:** 1×3 known-edged pair (TxtID1=1, TxtID2=3, FrameFilterYears=1, ChkKinship=0, Chk2Nodes=0).  Same pair as `test_vba_pajek_gephi_cross_form.py` `_assocpairs_1x3_fixture`; verified ZZ_SCRATCH_PEOPLE=2, ZZ_SOCIAL_NETWORK>0.
- **Chain:** `CmdQuery,CmdNeo4j` via Form.Tag, directory mode (trailing backslash → `f<n>.out.csv` per Show call)
- **MsgBox dismisser:** background thread (pywinauto win32 backend) auto-clicks OK on Access `#32770` class dialogs
- **click_via_timer cap:** 120 s  ·  **outer cap:** 300 s
- **Promote threshold (strict):** chain_elapsed ≤ 120 s **AND** file_count > 0 **AND** `Finished saving to Neo4j` MsgBox observed and dismissed. All three required; none may be substituted.

## Raw observed facts

(These are the unprocessed facts captured during the probe run.  Classification is derived from them in the next section.)

- **chain_elapsed_sec:** 12.54
- **file_count:** 6
- **chain_observed_done:** True
- **finished_msgbox_seen:** False
- **msgbox_dismissed_count:** 5
- **static_unconditional_msgboxes_in_vba:** 6
- **click_via_timer_returned:** 5
- **total_wall_elapsed_sec:** 21.37

## Classification

Strict gate evaluation against the promote threshold:

| Gate | Required | Observed | Pass |
|---|---|---|---|
| chain_elapsed ≤ 120 s | True | 12.54 s | ✅ |
| file_count > 0 | True | 6 | ✅ |
| finished_msgbox_seen | True | False | ❌ |

**Per-probe outcome:** `confirmed_blocking_msgbox_layer_needs_decision_before_coverage`

## Answers to brief Q1-Q5

- **Q1** — CmdQuery + CmdNeo4j completed in same session?  **True**
- **Q2** — Chain elapsed ≤ 120 s?  **True** (12.54 s)
- **Q3** — Files produced: **6**.  Headers: `['Person1_ID', 'AssociationCode', 'KinshipCode', 'nameID', 'placeID', 'nameID']`
- **Q4** — Blocking symptoms: MsgBox auto-dismissed count = **5**; Finished-Neo4j MsgBox seen = **False**; Static VBA MsgBox count = **6**
- **Q5** — Same family as LookAtAssociations 0-file mode?  Three observations, kept separate because the probe does not have evidence to collapse them into a single exclusive failure class:
  (a) different from LookAtAssociations × CmdNeo4j 0-file mode — LookAtAssociations produces 0 files in directory mode (likely bails before any SaveAs), whereas AssocPairs CmdNeo4j uses ZZ_SOCIAL_NETWORK (populated by CmdQuery on the 1x3 fixture) and produced 6 files in this run.
  (b) confirmed blocking MsgBox layer exists in CmdNeo4j_Click — static analysis lists 6 unconditional debug MsgBox calls, and 5 of them were observed and dismissed by the probe's pywinauto auto-dismisser at runtime. Unattended coverage would be blocked by this layer.
  (c) post-MsgBox terminal behavior remains only partially observed because the final 'Finished saving to Neo4j' MsgBox (line 1470) was not seen within the polling window. The probe therefore cannot claim that 'blocking_debug_msgbox' is the exclusive failure class after MsgBox removal.

## Verdict: `confirmed_blocking_msgbox_layer_needs_decision_before_coverage`

Chain produced 6 files in 12.54s and 5 of the 6 unconditional debug MsgBox calls were observed and dismissed by the probe's auto-dismisser. The terminal 'Finished saving to Neo4j' MsgBox (line 1470) was NOT observed within the polling window, so post-MsgBox terminal behavior is only partially observed.

What this confirms:
  - a blocking debug MsgBox layer exists in CmdNeo4j_Click and would prevent unattended coverage.
  - this is NOT the same as LookAtAssociations × CmdNeo4j 0-file mode.

What this does NOT confirm:
  - that the debug MsgBox layer is the only failure class after removal.
  - the full terminal completion path of CmdNeo4j on this fixture.

Decision required from maintainer/reviewer before any coverage PR is opened: (a) remove the 6 unconditional MsgBox calls from upstream VBA, or (b) handle them in the test driver. NOT a coverage candidate yet.

## Reviewer / maintainer decision points

1. **Unconditional debug MsgBox calls in CmdNeo4j_Click (confirmed runtime blocker for unattended coverage).**  Static analysis identifies 6 unconditional `MsgBox` calls in `Form_LookAtAssociationPairs.vb` at lines 1069, 1151, 1234, 1317, 1400, 1470.  The probe observed and dismissed 5 of them at runtime.  Decision required before any coverage PR is opened: (a) remove these calls from upstream VBA, or (b) handle them in the test driver.  This probe does not presume which option is correct.

## Future coverage-implementation tasks (not blockers)

The following items would only matter if and when a coverage PR is opened.  They are NOT blockers to this probe and do NOT block reviewer decision on the MsgBox question above.

- New file-shape headers observed in this run that are not yet in `_NEO4J_SHAPES` / `_NEO4J_SHAPES_BY_TWO_COLS` in `tests/test_vba_cmdneo4j_cross_form.py`: `Person1_ID` (PeopleAssociations, 13-col), `AssociationCode` (3-col), `KinshipCode` (3-col).  Adding these is a coverage-implementation task, not a probe finding.

## Observations / possible follow-up questions (not blockers)

These are noted only for completeness.  This probe does NOT have evidence sufficient to elevate them to confirmed blockers.

- `ZZ_KIN_LIST_TMP` post-chain row count = `132` with ChkKinship=0 in the fixture.  The DELETE at Form_LookAtAssociationPairs.vb line 900 is gated by `ChkKinship.Value`, so a non-zero count here is consistent with carry-over from prior sessions, but the probe did NOT verify whether the KinshipCodes file contents (96 rows in this run) reflect stale data vs. the current fixture's expected output.  Treat as an observation, not a confirmed blocker, until either (i) a future probe seeds and verifies the table state, or (ii) a fixture isolation gap is independently demonstrated.

## MsgBox dismissal log

| +t (s) | msg_text |
|---:|---|
| 7.08 | `Kinship code records = 4` |
| 7.5 | `Literary genre code records = 0` |
| 7.91 | `Institution code records = 0` |
| 8.33 | `Occasion code records = 0` |
| 8.75 | `Topic code records = 0` |

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
- `ZZ_TEST_DEBUG`: 3

## Markers (timeline, 12 entries)

  - `+  0.00s` constructing_session
  - `+  5.40s` session_opened_attempt_1
  - `+  5.77s` filedialog_patched
  - `+  6.43s` form_opened
  - `+  6.50s` fixture_controls_set
  - `+  6.50s` form_tag_set_chain_CmdQuery_CmdNeo4j
  - `+  6.50s` chain_fire_t_start
  - `+  9.03s` click_via_timer_returned_5
  - `+ 19.05s` chain_quiescent_files_6
  - `+ 19.05s` chain_elapsed_12.54s
  - `+ 19.05s` files_inventoried_6
  - `+ 19.05s` row_counts_captured

## Constraints honoured per brief

- ✅ Investigation artifacts only — no `tests/` changes
- ✅ Did NOT touch README / canonical reports / issue severity / driver
- ✅ Did NOT open a coverage PR
- ✅ Reused 1×3 known-edged fixture from `test_vba_pajek_gephi_cross_form.py` — no new fixture design
- ✅ Used Access COM via `VbaSession.make_fixture`
- ✅ LookAtAssociations × CmdNeo4j noted as companion observation only (not separately probed)