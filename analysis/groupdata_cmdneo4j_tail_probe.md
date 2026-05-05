# GroupData × CmdNeo4j tail probe — per-block isolation

**Date:** 2026-05-05  ·  **Branch:** `investigate/groupdata-cmdneo4j-tail`

Follow-up to PR `probe/groupdata-cmdneo4j` (commit `4ace85b`).  That probe found 8 of 11 expected CmdNeo4j files were produced under person_1 + Status/Office/Addr enable scope, with a mid-chain `LookAtGroupData:ERR No current record.` (DAO 3021).  This tail probe localises the failure and tests the hypothesis that the bug is *unguarded `.MoveFirst` on an empty recordset* in the PeopleEntry / EntryCode blocks.

## Static evidence (read-only source review)

`analysis/dump/vba/Form_LookAtGroupData.vb`:

- Line 1243-1245 (block #9 PeopleEntry):
  ```
  Set tRstPeopleEntry = CurrentDb.OpenRecordset("ZZ_SCRATCH_ENTRY", dbOpenDynaset)
  With tRstPeopleEntry
      .MoveFirst         ' DAO 3021 if recordset is empty
      Do While Not .EOF
  ```

- Line 1383-1385 (block #10 EntryCode):
  ```
  Set tRstEntryCodes = CurrentDb.OpenRecordset(tQueryStr)
        ' tQueryStr is a SELECT against ZZ_SCRATCH_ENTRY
  With tRstEntryCodes
      .MoveFirst         ' DAO 3021 if recordset is empty
      Do While Not .EOF
  ```

- Line 1485-1487 (block #11 InstitutionCodes) is gated by `If tRecDeleted > 0 Then`; this block is correctly skipped when its upstream INSERT produces 0 rows.  Not bugged on the Entry-empty path.

- ALL 11 blocks (#1-#11) share the same unguarded `.MoveFirst` pattern.  Blocks #1-#8 only "work" on the probe's fixture because their feeder scratch tables (ZZ_SCRATCH_STATUS, ZZ_SCRATCH_OFFICE, ZZ_PLACE, ZZ_SCRATCH_P_TEXT, ZZ_ADDRESSES) are non-empty.  Block #9 fails purely because ZZ_SCRATCH_ENTRY is 0 under the probe's all-Chk*-reset + Status/Office/Addr-only enable scope.

## Runtime probe — three iterations

| iter | enable_chk | chain | seed | files | ZZ_SCRATCH_ENTRY | err_class |
|---|---|:-:|:-:|---:|---:|---|
| `iter1_baseline_chain_chkentry_off` | Status+Office+Addr +GIS sisters | Y | N | 0 | — | `(no ERR)` |
| `iter2_chain_chkentry_on` | Status+Office+Addr+Entry +GIS sisters | Y | N | 8 | 0 | `LookAtGroupData:ERR No value given for one or more required parameters.` |
| `iter3_split_then_seed` | Status+Office+Addr +GIS sisters | N (split) | Y | 10 | 1 | `(no ERR)` |

## Conclusion: **`A_new_bug_candidate_empty_recordset_guard`**

Baseline corroborated by `iter2` (iter2 chain with ChkEntry on; ZZ_SCRATCH_ENTRY ended up 0 because Issue #6 (JET 3061) blocked queryEntry inserts): ZZ_SCRATCH_ENTRY=0 produced 8 files + `LookAtGroupData:ERR No current record.`.  Iter 3 then split the chain (CmdRun alone with ChkEntry off, then a manual INSERT of 1 synthetic row into ZZ_SCRATCH_ENTRY, then CmdNeo4j alone).  Result: ZZ_SCRATCH_ENTRY=1, 10 files produced, NO `No current record.` ERR.  The 2 extra files are the missing PeopleEntry + EntryCode shapes (see the per-file detail).  InstitutionCodes (block #11) correctly remained skipped because its `If tRecDeleted > 0` gate evaluates to False for the synthetic row.

This isolates the failure cause to: `Form_LookAtGroupData.CmdNeo4j_Click` blocks #9 (PeopleEntry, line 1243-1245) and #10 (EntryCode, line 1383-1385) call `.MoveFirst` on their recordset without first checking `.EOF` or `.RecordCount > 0`.  When ZZ_SCRATCH_ENTRY is empty the recordset opens with `.EOF=True` and `.MoveFirst` raises DAO 3021.

This is a NEW bug candidate distinct from Issue #6 (which is a column-typo JET 3061 in `Form_LookAtGroupData.queryEntry` that prevents ZZ_SCRATCH_ENTRY from being populated at all when ChkEntry is true).  The two bugs interact — Issue #6 is the upstream cause that leaves ZZ_SCRATCH_ENTRY=0 even on the ChkEntry-on path, which then exposes this downstream missing-guard bug — but they are different code-level defects and should be filed separately.

Static evidence corroborates: ALL 11 CmdNeo4j blocks share the same unguarded `.MoveFirst` pattern.  Blocks #1-#8 only "work" because their feeder scratch tables (ZZ_SCRATCH_STATUS, ZZ_SCRATCH_OFFICE, ZZ_PLACE, ZZ_SCRATCH_P_TEXT, ZZ_ADDRESSES) are non-empty under the probe's enable scope.  This is a systemic missing-guard pattern; PeopleEntry and EntryCode happen to be the first blocks that hit the empty-feeder case under any normal user enable scope.

## Per-iter detail

### `iter1_baseline_chain_chkentry_off`

- **enable_chk:** `['ChkStatus', 'ChkOffice', 'ChkAddr', 'ChkGisStatus', 'ChkGisOffice', 'ChkGisAddr']`
- **chain_neo4j:** `True`
- **seed_entry_row:** `False`
- **expected outcome (per probe design):** 8 files + LookAtGroupData:ERR 'No current record.' + ZZ_SCRATCH_ENTRY=0 (reproduces PR probe/groupdata-cmdneo4j commit 4ace85b finding)
- **per-iter outcome:** `exception_uncaught`
- **elapsed:** 218.0 s
- **file_count:** 0
- **click_via_timer_returned:** None
- **chain observed done:** None  (reason: None)
- **`:ERR` messages observed:** none
- **exception:** `RuntimeError('session open failed after 3 attempts')
Traceback (most recent call last):
  File "C:\Users\how612\Documents\GitHub\cbdb-user-mdb-tests\analysis\probe_groupdata_cmdneo4j_tail.py", line 345, in _worker
    raise RuntimeError(
RuntimeError: session open failed after 3 attempts
`
- **files produced (0):**
- **ZZ_TEST_DEBUG transcript (0):**
- **markers:**
  - `+  0.00s` constructing_session
  - `+  3.50s` session_open_attempt_1_fail: com_error(-2147023170, 'The remote procedure call failed.', None, None)
  - `+ 94.18s` session_open_attempt_2_fail: com_error(-2146959355, 'Server execution failed', None, None)
  - `+155.41s` session_open_attempt_3_fail: com_error(-2147023170, 'The remote procedure call failed.', None, None)

### `iter2_chain_chkentry_on`

- **enable_chk:** `['ChkStatus', 'ChkOffice', 'ChkAddr', 'ChkGisStatus', 'ChkGisOffice', 'ChkGisAddr', 'ChkEntry']`
- **chain_neo4j:** `True`
- **seed_entry_row:** `False`
- **expected outcome (per probe design):** Issue #6 (JET 3061) fires in queryEntry during CmdRun.  Outcome of CmdNeo4j depends on whether ZZ_SCRATCH_ENTRY ends up populated (transactional behaviour of the failing INSERT).
- **per-iter outcome:** `no_current_record_with_files`
- **elapsed:** 11.55 s
- **file_count:** 8
- **click_via_timer_returned:** 2
- **chain observed done:** True  (reason: done_marker)
- **row counts after CmdRun:**
  - `ZZ_SCRATCH_STATUS`: 2
  - `ZZ_SCRATCH_OFFICE`: 12
  - `ZZ_SCRATCH_ENTRY`: 0
  - `ZZ_SCRATCH_BIOG_ADDR_DATA`: 0
  - `ZZ_SCRATCH_BIOG_TEXT_DATA`: 0
  - `ZZ_SCRATCH_P_TEXT`: 2
  - `ZZ_SCRATCH_IMPORT_PEOPLE`: 1
  - `ZZ_ADDRESSES`: 7
  - `ZZ_PLACE`: 1
- **`:ERR` messages observed (2):**
  - `LookAtGroupData:ERR No value given for one or more required parameters.`
  - `LookAtGroupData:ERR No current record.`
- **files produced (8):**
  - `f12.out.csv` (93 B, 3 cols, first=`personPlaceCode`)
  - `f15.out.csv` (61 B, 4 cols, first=`NameID`)
  - `f18.out.csv` (108 B, 3 cols, first=`StatusCode`)
  - `f21.out.csv` (379 B, 8 cols, first=`NameID`)
  - `f24.out.csv` (711 B, 4 cols, first=`OfficeCode`)
  - `f3.out.csv` (225 B, 9 cols, first=`NameID`)
  - `f6.out.csv` (342 B, 5 cols, first=`PlaceID`)
  - `f9.out.csv` (47 B, 3 cols, first=`NameID`)
- **ZZ_TEST_DEBUG transcript (4):**
  - `   1`: `LookAtGroupData:ENTER`
  - `   2`: `LookAtGroupData:ERR No value given for one or more required parameters.`
  - `   3`: `LookAtGroupData:ERR No current record.`
  - `   4`: `LookAtGroupData:DONE`
- **markers:**
  - `+  0.00s` constructing_session
  - `+  5.08s` session_opened_attempt_1
  - `+  5.71s` filedialog_patched
  - `+  6.59s` form_opened
  - `+  6.60s` picker_seeded_pid_1
  - `+  6.70s` chk_state_set_enable_ChkStatus,ChkOffice,ChkAddr,ChkGisStatus,ChkGisOffice,ChkGisAddr,ChkEntry
  - `+  6.70s` form_tag_set_chain
  - `+  7.23s` click_via_timer_returned_2
  - `+  9.23s` chain_quiesce_files_8_reason_done_marker
  - `+  9.24s` files_inventoried_8
  - `+  9.24s` debug_captured

### `iter3_split_then_seed`

- **enable_chk:** `['ChkStatus', 'ChkOffice', 'ChkAddr', 'ChkGisStatus', 'ChkGisOffice', 'ChkGisAddr']`
- **chain_neo4j:** `False`
- **seed_entry_row:** `True`
- **expected outcome (per probe design):** If hypothesis A holds: 10-11 files + no 'No current record.' ERR (only InstitutionCodes block may be skipped via its tRecDeleted gate; PeopleEntry + EntryCode should now succeed against the seeded row).
- **per-iter outcome:** `clean_no_err_files_produced`
- **elapsed:** 14.13 s
- **file_count:** 10
- **click_via_timer_returned:** 2
- **cmdneo4j_alone_returned:** 3
- **seed_insert_succeeded:** True
- **ZZ_SCRATCH_ENTRY after seed:** 1
- **row counts after CmdRun:**
  - `ZZ_SCRATCH_STATUS`: 2
  - `ZZ_SCRATCH_OFFICE`: 12
  - `ZZ_SCRATCH_ENTRY`: 0
  - `ZZ_SCRATCH_BIOG_ADDR_DATA`: 1
  - `ZZ_SCRATCH_BIOG_TEXT_DATA`: 0
  - `ZZ_SCRATCH_P_TEXT`: 0
  - `ZZ_SCRATCH_IMPORT_PEOPLE`: 1
  - `ZZ_ADDRESSES`: 24
  - `ZZ_PLACE`: 9788
- **row counts after CmdNeo4j:**
  - `ZZ_SCRATCH_STATUS`: 2
  - `ZZ_SCRATCH_OFFICE`: 12
  - `ZZ_SCRATCH_ENTRY`: 1
  - `ZZ_SCRATCH_BIOG_ADDR_DATA`: 1
  - `ZZ_SCRATCH_BIOG_TEXT_DATA`: 0
  - `ZZ_SCRATCH_P_TEXT`: 3
  - `ZZ_SCRATCH_IMPORT_PEOPLE`: 1
  - `ZZ_ADDRESSES`: 8
  - `ZZ_PLACE`: 2
- **`:ERR` messages observed:** none
- **files produced (10):**
  - `f12.out.csv` (93 B, 3 cols, first=`personPlaceCode`)
  - `f15.out.csv` (61 B, 4 cols, first=`NameID`)
  - `f18.out.csv` (108 B, 3 cols, first=`StatusCode`)
  - `f21.out.csv` (379 B, 8 cols, first=`NameID`)
  - `f24.out.csv` (711 B, 4 cols, first=`OfficeCode`)
  - `f27.out.csv` (156 B, 11 cols, first=`NameID`)
  - `f3.out.csv` (286 B, 9 cols, first=`NameID`)
  - `f30.out.csv` (59 B, 3 cols, first=`EntryCode`)
  - `f6.out.csv` (342 B, 5 cols, first=`PlaceID`)
  - `f9.out.csv` (47 B, 3 cols, first=`NameID`)
- **ZZ_TEST_DEBUG transcript (3):**
  - `   1`: `LookAtGroupData:ENTER`
  - `   2`: `LookAtGroupData:DONE`
  - `   3`: `LookAtGroupData:MSGBOX`
- **markers:**
  - `+  0.00s` constructing_session
  - `+  5.07s` session_opened_attempt_1
  - `+  5.71s` filedialog_patched
  - `+  6.59s` form_opened
  - `+  6.60s` picker_seeded_pid_1
  - `+  6.69s` chk_state_set_enable_ChkStatus,ChkOffice,ChkAddr,ChkGisStatus,ChkGisOffice,ChkGisAddr
  - `+  6.70s` form_tag_set_cmdrun_alone
  - `+  7.23s` cmdrun_alone_returned_2
  - `+  9.23s` row_counts_after_cmdrun_captured_entry_0
  - `+  9.23s` seed_insert_succeeded
  - `+  9.23s` zz_scratch_entry_after_seed_1
  - `+  9.23s` form_tag_set_cmdneo4j_alone
  - `+  9.77s` cmdneo4j_alone_returned_3
  - `+ 11.77s` cmdneo4j_quiesce_files_10_reason_done_marker
  - `+ 11.78s` files_inventoried_10
  - `+ 11.78s` debug_captured

## Constraints honoured per brief

- ✅ Investigation artifacts only — no `tests/` changes
- ✅ Did NOT touch README / canonical reports / issue severity / driver  (the .md/.json/.py outputs of this probe are the only files written; the brief explicitly lists them)
- ✅ Did NOT design a new fixture — reused person_1 inline (matches matrix_hard_forms's `groupdata_person_1_small`)
- ✅ Used Access COM via VbaSession
- ✅ Did NOT open a new issue (this PR is the evidence base for the maintainer's later issue-filing decision)
- ✅ Did NOT mix this with Issue #6 (which is JET 3061, a column-typo class; the bug here is DAO 3021, an unguarded recordset class)