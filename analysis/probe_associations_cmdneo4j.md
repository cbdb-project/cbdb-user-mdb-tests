# LookAtAssociations × CmdNeo4j probe (probe-first investigation)

**Date:** 2026-05-07  ·  **Branch:** `probe/associations-cmdneo4j` (off main `ed61cb6`)

Per `analysis/export_gap_triage_plan.md` § Refresh 2026-05-07, this is the rank-1 cheapest unfinished local PR.  The cell has been skipped in `tests/test_vba_cmdneo4j_cross_form.py::_spec_skip_marks` with reason "produces 0 files in directory mode — needs investigation alongside Place"; this probe characterises *why*.

## Static pre-analysis

`Form_LookAtAssociations.vb::CmdNeo4j_Click` (lines 480-1565) contains an early-bail at lines 517-520:

```vb
If Me.ZZ_SCRATCH_P_ASSOC.Form.Recordset.RecordCount = 0 Then
    MsgBox "There are no records to save."  ' line 518
    GoTo Exit_CmdNeo4j_Click                 ' line 519
End If
```

`Dim dlgSaveAs As FileDialog` is declared at line 524 — i.e. **AFTER** the bail.  A hit on the bail therefore produces 0 files (no SaveAs dialog ever opens).

Driver context: `LookAtAssociations` is **NOT** in `_SUBFORMS_TO_REQUERY` (see `tests/cbdb_driver/vba_session.py` line 603).  Sibling forms `Place` and `Kinship` are in that dict because their subforms cache a saved-query recordset that stays stale after CmdQuery's INSERTs into the underlying table.  Candidate hypothesis (this probe tests it): Associations has the same staleness on `ZZ_SCRATCH_P_ASSOC.Form.Recordset` — CmdGIS / CmdPajek / CmdGephi don't trip it because they read different scratch tables.

The driver's generic `MsgBox "<lit>"` neutralizer rewrites the bail-MsgBox at line 518 into `CurrentDb.Execute INSERT INTO ZZ_TEST_DEBUG VALUES ('LookAtAssociations:MSGBOX')`, so a hit on the bail leaves a direct `LookAtAssociations:MSGBOX` row in `ZZ_TEST_DEBUG` — that is the direct evidence chain for the 0-file mode.

## Setup

- **Form:** `LookAtAssociations`
- **Fixture:** `assoc_437_unfiltered` (reused from matrix `_make_assoc_fixtures`; picker_ids = [437], controls = {'FrameFilterYears': 1})
- **Chain:** `CmdQuery,CmdNeo4j` via Form.Tag, directory mode (trailing backslash → `f<n>.out.csv` per `dlgSaveAs.Show` call)
- **Watchdog:** records (and dismisses to keep the probe moving) any MsgBox not caught by the driver's generic literal-neutralizer.
- **click_via_timer cap:** 180 s  ·  **outer cap:** 300 s
- **Promote threshold (strict):** chain_elapsed ≤ 120 s + file_count >= 1 + no `:ERR` markers in ZZ_TEST_DEBUG.

## Raw observed facts

- **chain_elapsed_sec:** 11.56
- **file_count:** 0
- **chain_observed_done:** True
- **click_via_timer_returned:** 11867
- **msgbox_observed_via_watchdog_count:** 0
- **total_wall_elapsed_sec:** 20.62

### Scratch row counts (post-chain)

- `ZZ_SCRATCH_ASSOC`: 11867
- `ZZ_SCRATCH_P_ASSOC`: 8087
- `ZZ_SCRATCH_PEOPLE`: 0
- `ZZ_TEST_DEBUG`: 3

### ZZ_TEST_DEBUG content

- `LookAtAssociations:ENTER`
- `LookAtAssociations:ERR The INSERT INTO statement contains the following unknown field name: 'c_index_addr_type_code'. Make sure you have typed the name correctly, and try the operation again.`
- `LookAtAssociations:DONE`

### Watchdog MsgBox observations

(none observed via watchdog — driver's generic literal-neutralizer caught everything)

## Classification

Strict gate evaluation (the four buckets are mutually exclusive; the first matching one wins):

| Bucket | Required | Observed | Match |
|---|---|---|---|
| blocked_exception | exception AND file_count==0 | exception=no, file_count=0 | — |
| probe_found_new_runtime_bug_candidate | any :ERR | :ERR present=True | ✅ |
| probe_hit_existing_known_failure_family | file_count==0 AND ZZ_TEST_DEBUG has LookAtAssociations:MSGBOX | file_count=0, marker=False | — |
| clean_probe_promote_to_coverage_candidate | file_count>=1 AND no :ERR AND chain quiesced AND elapsed<=120s | file_count=0, no_err=False, done=True, elapsed_ok=True | — |

**Per-probe outcome:** `probe_found_new_runtime_bug_candidate`

## Brief Q1-Q4 answers

**Q1 — Chain outcome label:** `runtime_ERR_zero_files`

**Q2 — 0-file mode evidence chain:**

- file_count = 0
- **zero_file_path_classification:** `runtime_ERR_after_first_SaveAs_show — chain advanced PAST the line-554 dlgSaveAs.Show (SaveAs captured a filename via FILEDIALOG_PATCH v8) and INTO the True branch at lines 555-1559. The :ERR marker means the JET / VBA runtime error fired BEFORE gStream.WriteText flushed any data to the captured filename. SaveAs dialog 'fired' logically but no disk file resulted. Direct evidence chain: ZZ_TEST_DEBUG :ERR marker -> error trap (Err_CmdNeo4j_Click) -> Exit_CmdNeo4j_Click -> 0 files on disk.`
- **bailed_before_any_saveas_filedialog_stage:** `False`
- ZZ_TEST_DEBUG contains `LookAtAssociations:MSGBOX` (the line-518 bail-MsgBox marker): **False**
- ZZ_TEST_DEBUG contains `:ERR` marker: **True**
- `:ERR` marker text:
    - `LookAtAssociations:ERR The INSERT INTO statement contains the following unknown field name: 'c_index_addr_type_code'. Make sure you have typed the name correctly, and try the operation again.`
- scratch tables with query output (rows > 0): `{'ZZ_SCRATCH_ASSOC': 11867, 'ZZ_SCRATCH_P_ASSOC': 8087}`
- watchdog MsgBox observation count: 0
- ZZ_TEST_DEBUG full: `['LookAtAssociations:ENTER', "LookAtAssociations:ERR The INSERT INTO statement contains the following unknown field name: 'c_index_addr_type_code'. Make sure you have typed the name correctly, and try the operation again.", 'LookAtAssociations:DONE']`

**Q3 — vs AssociationPairs × CmdNeo4j failure class:**

- Associations observed path: `runtime_err`
- AssocPairs known path: files_written_THEN_blocking_debug_msgbox_layer (per PR AX probe; the MsgBox layer was suppressed by PR #109's driver patch)
- **Are failure classes distinct?** `True`

AssocPairs CmdNeo4j writes >=1 files BEFORE its blocker (debug-MsgBox layer) fires; Associations CmdNeo4j is observed here to either bail at RecordCount=0 (0 files, MsgBox marker) or hit a different runtime path.  Either way, the failure classes are not the same.

**Q4 — Outcome bucket:** `probe_found_new_runtime_bug_candidate`

## Verdict: `probe_found_new_runtime_bug_candidate`

ZZ_TEST_DEBUG contains :ERR marker(s): ["LookAtAssociations:ERR The INSERT INTO statement contains the following unknown field name: 'c_index_addr_type_code'. Make sure you have typed the name correctly, and try the operation again."].  file_count = 0.

What this confirms (direct from runtime + static):
  - CmdQuery completed cleanly (click_via_timer returned 11867); scratch tables ZZ_SCRATCH_ASSOC / ZZ_SCRATCH_P_ASSOC are populated. So this is NOT the static-suspected line-517 RecordCount=0 bail (which would have left ZZ_SCRATCH_P_ASSOC empty AND a 'LookAtAssociations:MSGBOX' marker).
  - The :ERR fires INSIDE CmdNeo4j_Click body, AFTER the line-554 dlgSaveAs.Show True branch is entered. The error message ('unknown field name "c_index_addr_type_code"') traces to the INSERT at Form_LookAtAssociations.vb:644-647, which references 'BIOG_MAIN.c_index_addr_type_code'.
  - This is a JET 3061 column-not-found family (same shape as Issue #6 LookAtGroupData CmdGIS queryEntry typo, but on a different form / different target column).

What this does NOT yet confirm:
  - whether 'c_index_addr_type_code' is missing on BIOG_MAIN (source side) or on ZZ_SCRATCH_PEOPLE (target side) — needs a schema check that this probe intentionally does not perform (out of scope).
  - whether the bug existed in older dumps or is introduced by current dump's schema drift.

Failure class (Q3): DISTINCT from AssociationPairs × CmdNeo4j. AssocPairs writes >=1 files, then hits a debug-MsgBox layer (now suppressed by PR #109's driver patch). Associations writes 0 files because of a JET column-not-found error fired BEFORE any gStream.WriteText. The two failures are in different VBA error families and at different chain depths.

Recommend opening a new investigation line via a separate maintainer brief: (a) confirm whether the missing column is BIOG_MAIN side (schema drift) or ZZ_SCRATCH_PEOPLE side (table-schema typo in CBDB), (b) decide whether to file a new canonical Issue, (c) decide whether a driver-side per-form patch (à la _PER_FORM_CMDGIS_PATCHES Issue #4 / #5 fixes) is appropriate. NOT a coverage candidate; NOT yet a canonical issue.

## Per-file detail

(no files produced)

## Markers (timeline, 14 entries)

  - `+  0.00s` constructing_session
  - `+  5.22s` session_opened_attempt_1
  - `+  5.57s` filedialog_patched
  - `+  6.32s` form_opened
  - `+  6.33s` picker_seeded_1_codes
  - `+  6.33s` fixture_seeded
  - `+  6.33s` form_tag_set_chain_CmdQuery_CmdNeo4j
  - `+  6.33s` chain_fire_t_start
  - `+  9.89s` click_via_timer_returned_11867
  - `+ 17.89s` chain_quiescent_zero_files_stable_for_8s
  - `+ 17.89s` chain_elapsed_11.56s
  - `+ 17.89s` files_inventoried_0
  - `+ 17.90s` row_counts_captured
  - `+ 17.90s` zz_test_debug_captured

## Constraints honoured per brief

- ✅ Investigation artifacts only — no `tests/` changes
- ✅ Did NOT touch driver, README, canonical reports, issue severity, or triage docs
- ✅ Did NOT open a coverage PR
- ✅ Reused matrix `_make_assoc_fixtures` first fixture — no new long-term fixture design
- ✅ Did NOT pre-assume same failure class as AssociationPairs × CmdNeo4j
- ✅ Used Access COM via `VbaSession.make_fixture`