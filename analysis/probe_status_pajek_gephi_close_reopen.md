# LookAtStatus × {CmdPajek, CmdGephi}: close + reopen form between Phase A and Phase B

**Date:** 2026-05-09  ·  **Branch:** `investigate/status-close-reopen` (off main `6b06d6a`)

Last non-UI local feasibility check before fallback to UI direct simulation. Tests close+reopen of the form instance between Phase A (CmdQuery) and Phase B (export button) as a structurally distinct mechanism vs the already-exhausted PR #129/#132/#133/#134/#135/#136/#137/#141 families (all of which kept the same form-class instance).

## Experiment design

Per button, in a fresh single-session MDB copy:

1. Open form; seed fixture (controls + picker).
2. **Phase A**: trigger `CmdQuery` via `click_via_timer`; wait DONE; capture scratch row counts (must be PR #127 baseline 17023 / 17022).
3. **Snapshot**: `SELECT * INTO ZZ_SNAPSHOT_*_CR FROM ZZ_SCRATCH_*` for both scratch tables.
4. **Close form**: `DoCmd.Close acForm, "LookAtStatus", acSaveNo`.
5. **Reopen form**: `DoCmd.OpenForm`. This triggers `Form_Open()` which destructively `DELETE *`s `ZZ_SCRATCH_STATUS` + `ZZ_SCRATCH_P_STATUS` at `Form_LookAtStatus.vb:2090` / `:2103`. We capture the post-reopen row counts (expected: 0, 0) to confirm Form_Open ran.
6. **Restore**: `INSERT INTO ZZ_SCRATCH_* SELECT * FROM ZZ_SNAPSHOT_*_CR`. Capture post-restore row counts (must be back to 17023 / 17022).
7. **Re-seed**: re-apply `set_control` for fixture controls + `set_picker_codes` for picker (runtime form properties don't survive close).
8. **Phase B**: trigger `Cmd<button>` via `click_via_timer` on the new form instance. This is a FIRST OnTimer use on this instance — fresh event-binding cache, fresh subform Recordset bindings (via Form_Open's own logic).
9. Capture: file_count, ZZ_TEST_DEBUG, scratch counts, watchdog dialogs.

## State restoration: what we restored and why

**Restored (with rationale):**

- **`scratch_tables_via_sql_insert`**: Form_Open destructively DELETEs ZZ_SCRATCH_STATUS and ZZ_SCRATCH_P_STATUS at lines 2090/2103 of Form_LookAtStatus.vb. Without restoration, Phase B has no data to export. Snapshot via SELECT INTO before close; INSERT INTO ... SELECT * FROM snapshot after reopen. Restores the data precondition only — does not alter the subform Recordset binding (Form_Open's own logic re-binds the subform Recordsets, which is the exact pathology we are testing close+reopen against).
- **`form_tag_for_export_path`**: Set Form.Tag to '<button>|<output_dir>' for the autodetect to pick up output dir. Runtime form property; doesn't survive close. Re-set after reopen. This is a test-scaffolding precondition, NOT a state-restoration that would smuggle the experiment into a different workflow.
- **`form_controls_via_setcontrol`**: Re-set the same picker / control values used in Phase A (picker_codes, etc.). This matches the input state the user would have set before clicking Cmd<X> manually in the same session; the control values are runtime form properties that don't survive close. Required so Cmd<X>_Click reads the same inputs it would in real use.

**NOT restored (with rationale):**

- **`subform_recordset_binding`**: We DO NOT restore or rebind the subform Recordset (ZZ_SCRATCH_STATUS / ZZ_SCRATCH_P_STATUS subforms). Form_Open of the new instance handles that itself. If the new instance's subform binding is healthy, the close+reopen workaround works; if it isn't, the workaround doesn't work. This is exactly what we are testing — restoring it would beg the question.
- **`any_in_process_recordset_variables`**: Module-level / Dim'd-local Recordset variables from Phase A's CmdQuery_Click are NOT restored. They cannot be — they don't survive form close, by VBA's scoping rules. This is a key part of why close+reopen is a structurally distinct mechanism vs PR #129-#137/#141 in-process workarounds.
- **`onTimer_state_or_eventBinding_cache`**: We DO NOT touch OnTimer / event-binding cache state. The new form instance gets fresh cache, fresh OnTimer binding. PR #137 pinned the second-Form_Timer-call failure at the form-class-instance event-binding cache for the SAME instance. New instance = fresh cache by construction.

The restoration set is the minimum needed to keep Phase B testable. Critically, we do NOT restore the subform Recordset binding or any in-process Recordset variable — those are exactly the runtime state that close+reopen is meant to reset, so restoring them would beg the question.

**Total wall elapsed:** 299.16 s  ·  **buttons probed:** 2

## Raw facts (per button)

### CmdPajek

- **phase_a_click_via_timer_returned:** `17023`
- **phase_a_row_counts:** `{'ZZ_SCRATCH_STATUS': 17023, 'ZZ_SCRATCH_P_STATUS': 17022}`
- **snapshot_counts:** `{'ZZ_SNAPSHOT_STATUS_CR': 17023, 'ZZ_SNAPSHOT_P_STATUS_CR': 17022}`
- **close_ok:** `True` (err: `None`)
- **reopen_ok:** `True` (err: `None`)
- **row_counts_after_form_open_delete:** `{'ZZ_SCRATCH_STATUS': 0, 'ZZ_SCRATCH_P_STATUS': 0}`
- **row_counts_after_restore:** `{'ZZ_SCRATCH_STATUS': 17023, 'ZZ_SCRATCH_P_STATUS': 17022}`
- **phase_b_click_via_timer_returned:** `17023`
- **phase_b_zz_test_debug_msgs:** `[]`
- **phase_b_row_counts:** `{'ZZ_SCRATCH_STATUS': 17023, 'ZZ_SCRATCH_P_STATUS': 17022}`
- **file_count:** 0
- **msgbox_observed:** 1 dialogs

### CmdGephi

- **phase_a_click_via_timer_returned:** `17023`
- **phase_a_row_counts:** `{'ZZ_SCRATCH_STATUS': 17023, 'ZZ_SCRATCH_P_STATUS': 17022}`
- **snapshot_counts:** `{'ZZ_SNAPSHOT_STATUS_CR': 17023, 'ZZ_SNAPSHOT_P_STATUS_CR': 17022}`
- **close_ok:** `True` (err: `None`)
- **reopen_ok:** `True` (err: `None`)
- **row_counts_after_form_open_delete:** `{'ZZ_SCRATCH_STATUS': 0, 'ZZ_SCRATCH_P_STATUS': 0}`
- **row_counts_after_restore:** `{'ZZ_SCRATCH_STATUS': 17023, 'ZZ_SCRATCH_P_STATUS': 17022}`
- **phase_b_click_via_timer_returned:** `17023`
- **phase_b_zz_test_debug_msgs:** `[]`
- **phase_b_row_counts:** `{'ZZ_SCRATCH_STATUS': 17023, 'ZZ_SCRATCH_P_STATUS': 17022}`
- **file_count:** 0
- **msgbox_observed:** 1 dialogs

## Interpretation (per button)

| Button | Outcome | q1 state restorable | q2/q3 truly fired | q4 Object required disappeared | q5 file_count >= 1 |
|---|---|---|---|---|---|
| **CmdPajek** | `sub_did_not_fire` | True | False | False | False |
| **CmdGephi** | `sub_did_not_fire` | True | False | False | False |

Per-button raw signals:

- **CmdPajek**: `{'snapshot_counts_ok': True, 'restored_counts_ok': True, 'phase_b_enter_seen': False, 'phase_b_done_seen': False, 'phase_b_other_marker': False, 'phase_b_object_required': False, 'phase_b_err_texts': [], 'phase_b_file_count': 0}`
- **CmdGephi**: `{'snapshot_counts_ok': True, 'restored_counts_ok': True, 'phase_b_enter_seen': False, 'phase_b_done_seen': False, 'phase_b_other_marker': False, 'phase_b_object_required': False, 'phase_b_err_texts': [], 'phase_b_file_count': 0}`

## 5 required answers (overall)

1. **Was state restorable after close+reopen?** True. Snapshot counts and post-restore counts both match PR #127 baseline (17023 / 17022) for each button's session — see raw facts above.
2. **Did `CmdPajek` truly fire?** False.
3. **Did `CmdGephi` truly fire?** False.
4. **Did `Object required` disappear?** False.
5. **Did file_count go from 0 to >= 1?** False.

## Verdict

- **bucket:** `close_reopen_does_not_unblock`
- **failure shapes per button:** `{'CmdPajek': 'sub_did_not_fire', 'CmdGephi': 'sub_did_not_fire'}`
- **recommendation:** non-UI local workaround line is now exhausted; next step should be UI direct simulation fallback (pywinauto on the live Access UI). Do NOT carve more non-UI sub-variants — they would re-test structurally already-failed mechanisms.
- **next_step:** `ui_direct_simulation_fallback`

## Self-review checklist (programmer-self-review-template.md)

**A. Branch shape**
- [x] Branch cut clean from current `main` (`6b06d6a`).
- [x] `git diff --name-only main..HEAD` contains only the 3 permitted artifact files (probe py + md + json).
- [x] `git diff --stat main..HEAD` is additive-only.

**B. Source-of-truth sync**
- [x] Paired MD + JSON updated together.
- [x] No canonical-issue / triage / inventory drift (this PR doesn't touch those surfaces).
- N/A — bilingual: probe artifact PR; no EN/ZH tier summaries to sync.

**C. Evidence vs claim**
- [x] Raw facts (per-button raw_signals + ZZ_TEST_DEBUG transcripts + row counts + file lists) recorded separately from interpretation/classification.
- [x] Verdict bucket follows mechanically from raw facts via `_classify_button` + `_verdict`; no interpretation smuggled into raw fields.
- [x] No extrapolation: this probe tests close+reopen for Status × CmdPajek/Gephi only; no claims about other forms / buttons.
- [x] No runtime behavioural pin missing — close+reopen is a runtime test and we ran it.

**D. Residual risk**
- [x] What we did NOT verify: any close+reopen variant that ALSO modifies subform binding manually. We chose minimum restoration to keep the experiment clean; stronger restoration shapes weren't tried per brief constraint ("不要顺手测试别的 close+reopen 变体").
- [x] Next step that should NOT be autopiloted: the UI direct simulation fallback (pywinauto) is a separately-scoped maintainer brief; this PR recommends but does not begin it.
- [x] One **new** finding worth recording (does not change verdict): Phase B's `phase_b_zz_test_debug_msgs = []` (no `:ENTER`, no `:DONE`) means `Cmd<X>_Click` literally never executed on the new form instance. Since the autodetect always emits `:ENTER` at the top of every patched sub, an empty Phase B transcript proves the Form_Timer dispatch itself never reached the patched sub. This is a notable extension of PR #137's pinned cache — the timer-binding pathology may NOT be strictly per-form-class-instance; it may persist across instances within the same Access process. This is a *new piece of evidence* for the maintainer-line — the local non-UI workaround surface is even thinner than PR #141 implied. Recorded here; not promoted to a separate canonical issue (still part of the same maintainer-line concern packaged in PR #139 / PR #140).
- [x] No downstream-work pre-claim: this PR does NOT claim close+reopen would fix any other unblocked cell or any other form/button beyond Status × CmdPajek/Gephi.

**Pytest scope actually run**: `pytest tests/test_schema.py tests/test_saved_views.py -W ignore -q` → 34 passed in 25.43s. The full Access COM matrix (`pytest tests/`, ~138 tests) was NOT re-run because (i) this PR adds 3 artifact files only — no changes to `tests/`, `tests/cbdb_driver/`, or any production code; (ii) `pytest --collect-only` succeeded over all 138 tests, confirming imports clean; (iii) the probe imports `make_fixture` + `LOOKATSTATUS` + `_all_fixtures` and ran them successfully, exercising the same module surface as the matrix tests; (iv) full COM matrix takes hours, and the artifacts-only diff cannot regress matrix behaviour by construction.