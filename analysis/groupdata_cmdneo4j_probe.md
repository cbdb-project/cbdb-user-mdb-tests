# GroupData × CmdNeo4j probe (probe-first investigation)

**Date:** 2026-05-05  ·  **Branch:** `probe/groupdata-cmdneo4j` (off main `47e506d`)

Per the export-gap triage refresh (`analysis/export_gap_triage_plan.md`, refresh section 2026-05-05), GroupData × CmdNeo4j is the only remaining cheapest-next-cell after applying the brief's exclusions (no AssociationPairs · no driver/meta-PR · no CmdUCINet new family).  This probe answers Q1-Q5 from that brief BEFORE any coverage PR is opened.

## Setup

- Fixture: person_1 (matches `matrix_hard_forms`'s `groupdata_person_1_small`)
- Pattern: all 11 Chk* reset to False, then enable 6 clean-branch boxes (Status/Office/Addr + GIS sisters).  Same as the just-merged GroupData × CmdGIS test.
- Chain: `CmdRun,CmdNeo4j` via Form.Tag, directory mode (trailing backslash → `f<n>.out` per Show)
- click_via_timer cap: 180 s  ·  per-probe outer cap: 240 s
- Promote threshold per brief: chain elapsed ≤ 120 s, ≥ 1 file, no `LookAtGroupData:ERR`

## Outcome

- **per-probe outcome:** `needs_investigation_mid_chain_err_with_files`
- **chain elapsed:** 2.53 s
- **files produced:** 8
- **chain done observed:** True
- **click_via_timer returned:** 2
- **total wall elapsed:** 11.68 s

## Answers to brief Q1-Q5

- **Q1** — CmdRun + CmdNeo4j completed in same session?  **True**
- **Q2** — Total chain elapsed ≤ 120 s?  **True** (2.53 s)
- **Q3** — Files produced: **8**.  Names: `['f12.out.csv', 'f15.out.csv', 'f18.out.csv', 'f21.out.csv', 'f24.out.csv', 'f3.out.csv', 'f6.out.csv', 'f9.out.csv']`
- **Q4** — `ZZ_TEST_DEBUG` has `LookAtGroupData:ERR`?  **True**
- **Q5** — Branch-specific known issues?  `DAO_3021_no_current_record (distinct from Issue #6; suggests unguarded recordset access, e.g. .Fields read without .EOF check)`

## Verdict: `do_not_open_coverage_pr_mid_chain_err`

ZZ_TEST_DEBUG contains LookAtGroupData:ERR.  Per the refresh brief, do NOT open a coverage PR — open an investigation PR instead.

Error classification: DAO_3021_no_current_record (distinct from Issue #6; suggests unguarded recordset access, e.g. .Fields read without .EOF check)

Notes for the next maintainer reading this report:
- The chain DID complete in 2.53 s (well under the 120 s promote threshold) and DID produce 8 well-formed CSVs.  See the per-file table for the shapes that landed.
- The error fired mid-chain but the form-level error handler caught it; the chain still emitted its DONE marker.  This is similar in *shape* to the GroupData × CmdGIS investigation finding (an Entry-related branch failing while sibling branches complete cleanly), but the underlying error class (see classification above) determines whether this is the SAME bug as Issue #6 or a NEW bug worth filing separately.
- Recommended investigation: per-Chk* isolation probe (like analysis/probe_groupdata_cmdgis_subcalls.py from PR `investigate/groupdata-cmdgis-subcall-isolation`) to localise which CmdNeo4j section raises the error.  CmdNeo4j has no `If ChkX.Value` gating in its own body; the isolation probe needs to vary CmdRun's Chk* state to vary which ZZ_SCRATCH_* tables are populated, then re-run the chain and see which Chk*-state combo eliminates the error.

## Findings supplement (interpretation)

- Expected dlgSaveAs.Show blocks in CmdNeo4j: **11** (People, Places, PeoplePlaces, PeoplePlacesCodes, PeopleStatus, StatusCode, PeopleOffice, OfficeCodes, PeopleEntry, EntryCode, InstitutionCodes).  Last (InstitutionCodes) is conditional on `If tRecDeleted > 0`.
- Actual files produced: **8**.
- **3 blocks did not produce a file.** Most likely the chain bailed via `GoTo Exit_CmdNeo4j_Click` after the error surfaced.  Looking at the file-name sequence (patch_filedialog increments `f<n>.out` per Show call), the gap localises which block raised the error: the chain stopped at f<n>=24 (every Show call increments the sequence by ~3 due to driver-side bookkeeping), suggesting the error fired in or just before the `PeopleEntry` block.
- File first-column shapes seen: `['NameID', 'OfficeCode', 'PlaceID', 'StatusCode', 'personPlaceCode']`.
- `ZZ_SCRATCH_ENTRY` row count is 0 (because we excluded ChkEntry per the all-Chk*-reset + Status/Office/Addr-only enable pattern, matching the GroupData × CmdGIS coverage scope).  CmdNeo4j's PeopleEntry / EntryCode blocks (#9, #10) read from this table.  If the error is `No current record.`, that's an unguarded recordset access in those blocks (DAO 3021 family — `.Fields(...)` called without `.EOF` check), which is a *different* bug pattern than Issue #6's JET 3061 column-typo.
- Suggested next-step probe: re-run with `ChkEntry = True` enabled (and continue to exclude `ChkGisEntry` since this is Neo4j not GIS).  If the `:ERR No current record.` disappears when ZZ_SCRATCH_ENTRY is non-empty, that confirms the bug is "PeopleEntry/EntryCode block doesn't guard against empty source recordset" and is worth filing as a NEW issue distinct from Issue #6.
- Caveat: enabling ChkEntry will re-trigger Issue #6 (queryEntry's typo'd column ref) during CmdRun.  ZZ_SCRATCH_ENTRY may end up non-empty due to the prior INSERT executing before the JET 3061 raises, OR may stay 0 depending on transactional behaviour — the next-step probe needs to capture ZZ_SCRATCH_ENTRY count both with and without ChkEntry to disambiguate.

## Per-file detail

| name | size | n_cols | first_col | header_preview |
|---|---:|---:|---|---|
| f12.out.csv | 93 | 3 | `personPlaceCode` | `personPlaceCode,personPlaceTrans,personPlaceHZ` |
| f15.out.csv | 61 | 4 | `NameID` | `NameID,StatusCode,FirstYear,LastYear` |
| f18.out.csv | 108 | 3 | `StatusCode` | `StatusCode,StatusDesc,StatusDescHZ` |
| f21.out.csv | 379 | 8 | `NameID` | `NameID,OfficeCode,OfficeAddrID,SocialInstID,SocialInstID,PostingFirstYear,Postin` |
| f24.out.csv | 711 | 4 | `OfficeCode` | `OfficeCode,OfficeTrans,OfficePinyin,OfficeHZ` |
| f3.out.csv | 286 | 9 | `NameID` | `NameID,NameHZ,NamePY,IndexYear,IndexYearTypeCode,IndexYearTypeDesc,IndexYearType` |
| f6.out.csv | 342 | 5 | `PlaceID` | `PlaceID,PlacePY,PlaceHZ,PlaceX,PlaceY` |
| f9.out.csv | 47 | 3 | `NameID` | `NameID,PlaceID,PersonPlaceCode` |

## Scratch row counts (post-chain)

- `ZZ_SCRATCH_STATUS`: 2
- `ZZ_SCRATCH_OFFICE`: 12
- `ZZ_SCRATCH_ENTRY`: 0
- `ZZ_SCRATCH_BIOG_ADDR_DATA`: 1
- `ZZ_SCRATCH_BIOG_TEXT_DATA`: 0
- `ZZ_SCRATCH_P_TEXT`: 3
- `ZZ_SCRATCH_IMPORT_PEOPLE`: 1
- `ZZ_ADDRESSES`: 8
- `ZZ_PLACE`: 2

## ZZ_TEST_DEBUG transcript (3 entries)

  - `   1`: `LookAtGroupData:ENTER`
  - `   2`: `LookAtGroupData:ERR No current record.`
  - `   3`: `LookAtGroupData:DONE`

## Markers (timeline)

  - `+  0.00s` constructing_session
  - `+  5.24s` session_opened_attempt_1
  - `+  5.87s` filedialog_patched
  - `+  6.75s` form_opened
  - `+  6.75s` picker_seeded_pid_1
  - `+  6.82s` all_chk_reset_to_False
  - `+  6.85s` clean_branches_enabled
  - `+  6.85s` form_tag_set_chain_CmdRun_CmdNeo4j
  - `+  6.85s` chain_fire_t_start
  - `+  7.38s` click_via_timer_returned_2
  - `+  9.39s` chain_done_observed_files_8
  - `+  9.39s` chain_elapsed_2.53s
  - `+  9.39s` files_inventoried_8
  - `+  9.39s` debug_captured

## Constraints honoured per brief

- ✅ Investigation artifacts only — no `tests/` changes
- ✅ Did NOT touch README / reports* / issues / driver  (*reports/groupdata_cmdneo4j_probe.json IS an output of this probe per the brief; no other reports edited)
- ✅ Did NOT design a new fixture — reused person_1 inline (matches matrix_hard_forms's `groupdata_person_1_small`)
- ✅ Used Access COM via `VbaSession.make_fixture`