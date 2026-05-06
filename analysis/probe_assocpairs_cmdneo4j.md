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
- **Promote threshold:** chain elapsed ≤ 120 s + all files produced + `Finished saving to Neo4j` MsgBox seen

## Outcome

- **per-probe outcome:** `clean_probe_promote_to_coverage_candidate`
- **chain elapsed:** 12.54 s
- **files produced:** 6
- **chain done observed:** True
- **Finished Neo4j MsgBox seen:** False
- **MsgBoxes auto-dismissed:** 5
- **click_via_timer returned:** 5
- **total wall elapsed:** 21.37 s

## Answers to brief Q1-Q5

- **Q1** — CmdQuery + CmdNeo4j completed in same session?  **True**
- **Q2** — Chain elapsed ≤ 120 s?  **True** (12.54 s)
- **Q3** — Files produced: **6**.  Headers: `['Person1_ID', 'AssociationCode', 'KinshipCode', 'nameID', 'placeID', 'nameID']`
- **Q4** — Blocking symptoms: MsgBox auto-dismissed count = **5**; Finished-Neo4j MsgBox seen = **False**; Static VBA MsgBox count = **6**
- **Q5** — Same family as LookAtAssociations 0-file mode?  different — LookAtAssociations CmdNeo4j produces 0 files in directory mode (likely bails before any SaveAs); AssocPairs CmdNeo4j uses ZZ_SOCIAL_NETWORK (populated by CmdQuery on the 1x3 fixture) and can proceed to write multiple files before hitting the blocking MsgBox chain. The failure class is blocking_debug_msgbox, not 0-file mode.

## Verdict: `clean_probe_promote_to_coverage_candidate`

Chain completed (12.54s, ≤120s), produced 6 files, finished_neo4j MsgBox dismissed. Note: coverage PR will require either (a) removing the 6 blocking debug MsgBox calls from the upstream VBA, or (b) handling them in the test driver. Per the triage brief, do NOT auto-promote — report first.

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