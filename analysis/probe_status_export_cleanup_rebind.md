# LookAtStatus export-cleanup-rebind family probe (driver/meta investigation)

**Date:** 2026-05-08  ·  **Branch:** `investigate/status-cleanup-rebind-family` (off main `089b2f9`)

Per `analysis/export_gap_triage_plan.md` § Refresh 2026-05-08 (Rank 2 — driver/meta investigation), this probe characterises the Status × {CmdNeo4j, CmdPajek, CmdGephi} family blocker.  The existing skip-reason text claims all three share a 'CmdQuery cleanup-rebind' root cause; this probe verifies whether that's true OR whether the three are sibling / distinct failures.  **Read-only investigation; no driver, test, README, or triage changes.**

## Static pre-analysis (pre-runtime evidence)

`Form_LookAtStatus.vb` (3304 lines, current dump):

- `CmdQuery_Click` (line 1156) drops both subform recordsets via dummy-rebind (lines 1174 / 1186), runs INSERTs into ZZ_SCRATCH_STATUS / ZZ_SCRATCH_P_STATUS, then cleanup section `Exit_Run_Query:` (line 1452) **rebinds** both subforms via `Set ZZ_SCRATCH_STATUS.Form.Recordset = CurrentDb.OpenRecordset(...)` (lines 1457 + 1460).
- `CmdPajek_Click` (line 2133) AND `CmdGephi_Click` (line 18) **both** check subform recordset RecordCount upfront: `If ZZ_SCRATCH_STATUS.Form.Recordset.RecordCount = 0 Then MsgBox 'There are no records to save.'` (Pajek lines 2156-2161; Gephi lines 45-50).  Then both also use `Set tRstEdge = ZZ_SCRATCH_STATUS.Form.Recordset` to read data.  Same structural pattern.
- `CmdNeo4j_Click` (line 479) does **NOT** check subform RecordCount upfront.  Opens fresh dynasets directly on the underlying scratch TABLES: `Set tRstPeopleStatus = CurrentDb.OpenRecordset("ZZ_SCRATCH_STATUS", dbOpenDynaset)` (line 527) and `Set tRstPeople = CurrentDb.OpenRecordset("ZZ_SCRATCH_P_STATUS", dbOpenDynaset)` (line 528).  Structurally **bypasses** the subform recordset.

**Pre-runtime hypothesis:** CmdPajek + CmdGephi share the cleanup-rebind blocker (subform RecordCount=0); CmdNeo4j does NOT — it must fail elsewhere or possibly succeed.

## Setup

- **Form:** `LookAtStatus`
- **Fixture:** matrix `_make_status_fixtures` first fixture (`status_<top_code>_unfiltered`)
- **Phases:** 3 sequential, one per export button (`CmdPajek, CmdGephi, CmdNeo4j`).  Each phase = own fresh MDB copy + own VbaSession + own out_dir → fully isolated evidence per button.
- **Per-phase chain:** `CmdQuery,<button>` via Form.Tag, directory mode
- **Watchdog:** records (and dismisses to keep the probe moving) any MsgBox not caught by the driver's generic literal-neutralizer.  Watchdog dialog texts are surfaced as observations, NOT silently swallowed.
- **click_via_timer cap:** 180 s  ·  **per-phase outer cap:** 240 s
- **Total wall elapsed:** 62.13 s

## Raw observed facts (per phase)

### Phase: `CmdPajek`

- **chain_elapsed_sec:** 9.54
- **file_count:** 0
- **chain_observed_done:** True
- **click_via_timer_returned:** 17023
- **msgbox_watchdog_count:** 0
- **per_phase_wall_elapsed_sec:** 18.48

**Scratch row counts (post-chain):**
- `ZZ_SCRATCH_STATUS`: 17023
- `ZZ_SCRATCH_P_STATUS`: 17022
- `ZZ_TEST_DEBUG`: 3

**ZZ_TEST_DEBUG content:**
- `LookAtStatus:ENTER`
- `LookAtStatus:ERR Object required`
- `LookAtStatus:DONE`

**Watchdog MsgBox observations:**
(none observed)

**Files produced:** 0 (per-file shape detail in JSON).

**Phase signature category:** `runtime_err_zero_files_other`

### Phase: `CmdGephi`

- **chain_elapsed_sec:** 9.54
- **file_count:** 0
- **chain_observed_done:** True
- **click_via_timer_returned:** 17023
- **msgbox_watchdog_count:** 0
- **per_phase_wall_elapsed_sec:** 18.12

**Scratch row counts (post-chain):**
- `ZZ_SCRATCH_STATUS`: 17023
- `ZZ_SCRATCH_P_STATUS`: 17022
- `ZZ_TEST_DEBUG`: 3

**ZZ_TEST_DEBUG content:**
- `LookAtStatus:ENTER`
- `LookAtStatus:ERR Object required`
- `LookAtStatus:DONE`

**Watchdog MsgBox observations:**
(none observed)

**Files produced:** 0 (per-file shape detail in JSON).

**Phase signature category:** `runtime_err_zero_files_other`

### Phase: `CmdNeo4j`

- **chain_elapsed_sec:** 7.04
- **file_count:** 6
- **chain_observed_done:** True
- **click_via_timer_returned:** 17023
- **msgbox_watchdog_count:** 0
- **per_phase_wall_elapsed_sec:** 15.76

**Scratch row counts (post-chain):**
- `ZZ_SCRATCH_STATUS`: 17023
- `ZZ_SCRATCH_P_STATUS`: 17022
- `ZZ_TEST_DEBUG`: 3

**ZZ_TEST_DEBUG content:**
- `LookAtStatus:ENTER`
- `LookAtStatus:MSGBOX`
- `LookAtStatus:DONE`

**Watchdog MsgBox observations:**
(none observed)

**Files produced:** 6 (per-file shape detail in JSON).

**Phase signature category:** `files_produced_clean`

## Cross-phase signatures

| Button | Category | Files | :ERR | Watchdog | ZZ_SCRATCH_STATUS | ZZ_SCRATCH_P_STATUS | no-records dialog? |
|---|---|---:|---:|---:|---:|---:|---|
| `CmdPajek` | `runtime_err_zero_files_other` | 0 | 1 | 0 | 17023 | 17022 | False |
| `CmdGephi` | `runtime_err_zero_files_other` | 0 | 1 | 0 | 17023 | 17022 | False |
| `CmdNeo4j` | `files_produced_clean` | 6 | 0 | 0 | 17023 | 17022 | False |

## Q1-Q5 answers

**Q1 — same root cause? cleanup-rebind invalidates subform recordset?** verdict: `shared_root_for_pajek_gephi_only_neo4j_runs_clean_unaffected`

Per-button :ERR text (the :ERR text after the form prefix, e.g. `Object required`):
  - `CmdPajek`: `Object required`
  - `CmdGephi`: `Object required`
  - `CmdNeo4j`: `None`

Per-button files-produced-clean (no :ERR, files written):
  - `CmdPajek`: False
  - `CmdGephi`: False
  - `CmdNeo4j`: True

CmdPajek + CmdGephi share identical :ERR text? **True**

**Q2 — chain stage of failure per button:**
  - `CmdPajek`: `before_file_write_object_required_from_subform_recordset_access`
  - `CmdGephi`: `before_file_write_object_required_from_subform_recordset_access`
  - `CmdNeo4j`: `completed_export_files_written`

**Q3 — same / sibling / 3 separate?** verdict: `two_sibling_same_err_text_third_runs_clean`

**Q4 — evidence chain per button:** see Cross-phase signatures table above + per-phase Raw observed facts.

**Q5 — outcome bucket:** `export_specific_sibling_blockers`

**Minimum intervention surface (only meaningful when bucket = shared / sibling):**

Two scopes, NOT three.  Probe found that CmdPajek + CmdGephi share an identical :ERR text (`Object required`, VBA 424) — both bail at the same structural point: their pre-export `subform.Form.Recordset.RecordCount` access (Pajek lines 2156+2161, Gephi lines 45+50) raises 424 because the cleanup rebind (Exit_Run_Query, lines 1457+1460) leaves the subform recordset in an Object-Required state.  **CmdNeo4j runs CLEAN on the same fixture** — produces files, no :ERR, no dialogs.  This REFUTES the existing CmdNeo4j skip-reason claim of 'same root family'.  CmdNeo4j is structurally different (opens fresh dbOpenDynaset on the underlying scratch TABLES directly, lines 527+528; bypasses subform recordset entirely).

Minimum intervention surfaces:
  (1) **Driver-side cleanup-rebind fix on Form_LookAtStatus.vb:1457+1460** — narrow per-form patch in `_PER_FORM_CMDGIS_PATCHES` style; candidates: replace `Set X.Form.Recordset = CurrentDb.OpenRecordset(...)` with `X.Form.Requery`, OR add `MoveFirst` after Set, OR drop the rebind block.  Unblocks Pajek + Gephi.  Separate driver PR (NOT this one).
  (2) **CmdNeo4j unskip** — the existing `_spec_skip_marks` entry in `tests/test_vba_cmdneo4j_cross_form.py` is a false positive that copy-pasted the Pajek/Gephi reason without verifying.  Direct unskip (no driver patch needed) → coverage cell.  Separate coverage PR (NOT this one).

## Verdict: `export_specific_sibling_blockers`

**Sibling-not-shared.**  Two of the three Status export buttons (CmdPajek + CmdGephi) share an identical :ERR text and bail at the same structural point.  The third — CmdNeo4j — **runs cleanly** on the same fixture and produces files.  This **refutes** the existing CmdNeo4j skip-reason claim of 'same root family as Pajek/Gephi Status skip' — that skip is a false positive that was copy-pasted from the Pajek/Gephi skip without verification.

Two minimum intervention surfaces (NOT one): see `minimum_intervention_surface_note`.  Recommended next steps (each a separate PR):
  (1) Driver/meta PR — narrow per-form patch to fix the cleanup-rebind 'Object required' failure for Pajek + Gephi.
  (2) Coverage PR — direct unskip of CmdNeo4j in `_spec_skip_marks` (no driver change needed; the cell already runs clean), with per-shape pinning analogous to other covered Neo4j cells.
  (3) Triage refresh — split the existing Status × CmdNeo4j skipped entry from the Status × CmdPajek + CmdGephi skipped entries; they are separate blockers (or, in CmdNeo4j's case, no blocker).

The probe REFUTES the rank-2 framing in refresh_2026_05_08 that '1 fix unblocks 3 cells'.  It's '1 fix unblocks 2 cells; 1 cell is already unblocked'.  Total leverage is still high (3 cells movable) but the shape is split, not bundled.

## Constraints honoured per brief

- ✅ Investigation artifacts only — no `tests/`, `tests/cbdb_driver/`, `README.md`, `analysis/export_gap_triage_plan.md`, `reports/generate_report.py`, or canonical issue content changed
- ✅ No coverage PR opened; no driver patch landed; no canonical issue filed
- ✅ Did NOT pre-assume per-form literal-rewrite would solve this — explicit driver/meta scope
- ✅ Used Access COM via `VbaSession.make_fixture` across 3 isolated phases
- ✅ Raw facts (per-phase observations) and classification (cross-phase signatures + verdict) separated into different sections
- ✅ `--reclassify-from-json` supported for re-running verdict logic without COM