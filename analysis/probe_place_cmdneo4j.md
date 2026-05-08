# LookAtPlace × CmdNeo4j probe (probe-first investigation)

**Date:** 2026-05-08  ·  **Branch:** `probe/place-cmdneo4j` (off main `d288d0c`)

Per `analysis/export_gap_triage_plan.md` § Refresh 2026-05-07 (later), this is the rank-1 cheapest unfinished local PR.  The cell has been skipped in `tests/test_vba_cmdneo4j_cross_form.py::_spec_skip_marks` with reason "fires `Item not found in this collection.` mid-body — looks like a real CBDB bug (SQL or recordset field reference against a renamed/missing column)"; this probe characterises *why* and decides whether the methodology that worked end-to-end for Associations × CmdNeo4j (#112 → #118) transfers.

## Static pre-analysis

`Form_LookAtPlace.vb::CmdNeo4j_Click` (lines 435-1778 in the current dump, 6810 total lines, verified via Python `splitlines()` with cp1252):

- **NO** `RecordCount = 0` early-bail check (unlike Associations line 1033 / AssocPairs line 363)
- **NO** `MsgBox "There are no records to save."` literal
- **NO** literal `"Item not found in this collection."` string in the VBA source — that text is **JET / DAO error 3265** raised at runtime when a `Recordset!fieldname` reference (or `Recordset.Fields("name")` lookup) fails because the field isn't in the recordset's runtime field collection.
- 7 `Set tRst* = CurrentDb.OpenRecordset(...)` calls (lines 531 / 651 / 881 / 1073 / 1259 / 1479 / 1671)
- 6 `If dlgSaveAs.Show = -1 Then` blocks (lines 545 / 827 / 959 / 1205 / 1427 / 1591)
- 54 `Recordset!c_<col>` field references — many candidates for the JET 3265 trigger
- `Dim dlgSaveAs As FileDialog` at line 457; `Exit_CmdNeo4j_Click:` at line 1761; `Err_CmdNeo4j_Click:` at line 1767

**Family hypothesis (pre-runtime):** the JET 3265 skip-reason text would put this cell in a **different family** from Issue #23 (JET 3061 unknown field in INSERT) and from Issue #21 (DAO 3021 No current record on empty recordset .MoveFirst).  Same SURFACE SYMPTOM (CBDB-side renamed/missing column) but DIFFERENT trigger surface (DAO field lookup vs SQL parser vs unguarded MoveFirst).

## Setup

- **Form:** `LookAtPlace`
- **Fixture:** `place_addr_7213` (reused from matrix `_make_place_fixtures`; addr_ids = [7213], controls = {'ChkIndividual': True, 'ChkOffice': False, 'ChkAssoc': False, 'ChkPosting': False, 'ChkEntry': False})
- **Pre-fixture step:** `set_control("LookAtPlace", "TabPlaces", 0)` (mirrors the cross-form CmdNeo4j test's special-case handling for Place)
- **Chain:** `CmdQuery,CmdNeo4j` via Form.Tag, directory mode (trailing backslash → `f<n>.out.csv` per `dlgSaveAs.Show` call)
- **Watchdog:** records (and dismisses to keep the probe moving) any MsgBox not caught by the driver's generic literal-neutralizer.
- **click_via_timer cap:** 180 s  ·  **outer cap:** 300 s

## Raw observed facts

- **chain_elapsed_sec:** 17.05
- **file_count:** 0
- **chain_observed_done:** True
- **click_via_timer_returned:** 5962
- **msgbox_observed_via_watchdog_count:** 0
- **total_wall_elapsed_sec:** 26.51

### Scratch row counts (post-chain)

- `ZZ_SCRATCH_STATUS`: 17
- `ZZ_SCRATCH_PLACE_PEOPLE`: 5764
- `ZZ_SCRATCH_PLACE_AGG`: 5764
- `ZZ_SCRATCH_PEOPLE`: 5764
- `ZZ_SCRATCH_ADDR`: 1
- `ZZ_TEST_DEBUG`: 3

### ZZ_TEST_DEBUG content

- `LookAtPlace:ENTER`
- `LookAtPlace:ERR Item not found in this collection.`
- `LookAtPlace:DONE`

### Watchdog MsgBox observations

(none observed)

## Q1-Q6 answers

**Q1 — Chain outcome label:** `runtime_ERR_zero_files`

**Q2 — "Item not found in this collection." evidence chain:**

- skip_reason_phrase searched for: `"Item not found in this collection"`
- appears in ZZ_TEST_DEBUG :ERR row(s)? **True**
- matching :ERR markers:
    - `LookAtPlace:ERR Item not found in this collection.`
- appears in watchdog-dismissed dialogs? `False`
- file_count at failure: `0` (0 = before any disk write; >0 = at least one SaveAs block completed before the error fired)

If appears_in_zz_test_debug is True, the documented skip-reason error reproduced — JET 3265 fired mid-body and the driver's generic Err.Description neutralizer captured it as a ZZ_TEST_DEBUG :ERR row.  The chain stage (before/after any SaveAs) is inferred from file_count: 0 means before any disk write; > 0 means at least one SaveAs block completed before the error fired.

**Q3 — ZZ_TEST_DEBUG markers:** see Raw observed facts → ZZ_TEST_DEBUG content section above.

**Q4 — ZZ_SCRATCH_* row counts at failure:** see Raw observed facts → Scratch row counts section above.

**Q5 — same family as Issue #23?**

- verdict: **`DIFFERENT_FAMILY_from_Issue_23 — JET_3265_recordset_field_lookup`**
- rationale:

Skip reason reproduces.  JET 3265 fires when a `Recordset!field` (or `Recordset.Fields(name)`) lookup fails at the VBA / DAO layer because the field isn't in the recordset's runtime field collection.  Issue #23 (JET 3061) fires from the SQL parser when an INSERT/SELECT/UPDATE field name doesn't exist on the named target/source table.  Same SURFACE SYMPTOM (CBDB-side missing/renamed column) but DIFFERENT TRIGGER SURFACE (DAO field lookup vs SQL parser).  Per-form workaround would also differ: this would rewrite a `!c_<col>` identifier, not an INSERT target column literal.

- comparison:
    - `issue_23_associations_x_cmdneo4j`: JET 3061 'unknown field name in INSERT': INSERT INTO ZZ_SCRATCH_PEOPLE references non-existent target column c_index_addr_type_code
    - `this_probe_place_x_cmdneo4j`: JET 3265 'Item not found in this collection.': a Recordset!c_<col> reference in CmdNeo4j_Click body fails to find the field on the open recordset

**Q6 — Outcome bucket:** `probe_hit_existing_known_failure_family`

## Verdict: `probe_hit_existing_known_failure_family`

**Documented skip reason reproduced.**  ZZ_TEST_DEBUG contains the JET 3265 "Item not found in this collection." :ERR row(s): ['LookAtPlace:ERR Item not found in this collection.'].  file_count = 0.

This is the **JET 3265 family** — a `Recordset!field` or `Recordset.Fields("name")` reference at the VBA layer fails because the field isn't in the recordset's runtime field collection.  DIFFERENT family from Issue #23 (JET 3061 unknown field name in INSERT statement) and from Issue #21 (DAO 3021 'No current record' on empty recordset .MoveFirst).  Same surface symptom (missing/renamed column) but different trigger surface (DAO field lookup vs SQL parser).

Recommended next step (separate brief, NOT this PR): static investigation analogous to PR #114 — locate the specific `!c_<col>` reference inside CmdNeo4j_Click that fails (54 candidates per the static pre-analysis), determine whether the source recordset's column has been renamed or removed, then file as a new canonical Issue (analogous to Issue #23 filing in PR #115).  The driver-side workaround would mirror PR #116's `.replace()` shape but on the `!c_<col>` identifier rather than the INSERT target column.

## Markers (timeline, 17 entries)

  - `+  0.00s` constructing_session
  - `+  5.81s` session_opened_attempt_1
  - `+  6.07s` filedialog_patched
  - `+  6.94s` form_opened
  - `+  6.95s` tab_places_set_to_0
  - `+  6.97s` set_control_ChkAssoc_fail: com_error(-2147352567, 'Exception occurred.', (0, None, "Welcome to CBDB! can't find the field 'ChkAssoc' referred to in your expression.", 'acmain11.chm', 11730, -2146825823), None)
  - `+  6.98s` set_control_ChkPosting_fail: com_error(-2147352567, 'Exception occurred.', (0, None, "Welcome to CBDB! can't find the field 'ChkPosting' referred to in your expression.", 'acmain11.chm', 11730, -2146825823), None)
  - `+  6.99s` addr_picker_seeded_1_codes
  - `+  6.99s` fixture_seeded
  - `+  6.99s` form_tag_set_chain_CmdQuery_CmdNeo4j
  - `+  6.99s` chain_fire_t_start
  - `+ 16.04s` click_via_timer_returned_5962
  - `+ 24.04s` chain_quiescent_zero_files_stable_for_8s
  - `+ 24.04s` chain_elapsed_17.05s
  - `+ 24.04s` files_inventoried_0
  - `+ 24.05s` row_counts_captured
  - `+ 24.05s` zz_test_debug_captured

## Constraints honoured per brief

- ✅ Investigation artifacts only — no `tests/` changes
- ✅ Did NOT touch driver, README, canonical reports, issue severity, or triage docs
- ✅ Did NOT open a coverage PR
- ✅ Reused matrix `_make_place_fixtures` first fixture — no new fixture design
- ✅ Did NOT pre-assume same family as Issue #23 — Q5 explicitly distinguishes JET 3265 (DAO field lookup) from JET 3061 (SQL parser)
- ✅ Used Access COM via `VbaSession.make_fixture`
- ✅ Raw facts and conclusion separated