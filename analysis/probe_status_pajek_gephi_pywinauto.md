# LookAtStatus × {CmdPajek, CmdGephi}: UI direct simulation (pywinauto) feasibility probe

**Date:** 2026-05-09  ·  **Branch:** `investigate/status-pywinauto` (off main `6b06d6a`)

Last local feasibility check before pure maintainer-line. Tests whether direct UI driving via pywinauto on the live Access UI can trigger the export Cmd<X>_Click handlers, bypassing the Form_Timer / chained-dispatch / COM-side invoke paths that have all failed in PR #129/#132/#133/#134/#135/#136/#137/#141/#142.

## Experiment design

Per button, fresh single-session MDB copy:

1. Open form; seed fixture (controls + picker).
2. **Phase A**: trigger `CmdQuery` via `click_via_timer` (uses existing Form_Timer machinery — but only for CmdQuery, which is not the path we are testing). Wait DONE; capture scratch baseline.
3. **Diagnostic**: enumerate the LookAtStatus form's controls recursively via COM (Forms / Controls / subform Form); enumerate UIA descendants of the Access app window; identify button candidates that pywinauto would click.
4. **Phase B**: attempt pywinauto direct UI invocation in this order, each attempt logged independently:
    1. UIA backend: find by `auto_id == 'Cmd<X>'`
    2. UIA backend: find by `name == 'Cmd<X>'`
    3. Win32 backend: enumerate child windows of the Access main window
    4. COM `Forms('LookAtStatus').Controls('Cmd<X>').SetFocus` + pywinauto.keyboard ENTER
5. After all attempts, wait up to 120s for any ZZ_TEST_DEBUG marker (`:ENTER`/`:DONE`/`:ERR`/`:MSGBOX`); empty transcript = sub never fired.
6. Capture file_count, scratch counts, watchdog dialogs, all per-attempt diagnostics.

Per-button outer timeout: 480 s.

**Total wall elapsed:** 287.30 s  ·  **buttons probed:** 2

## Raw facts (per button)

### CmdPajek

- **phase_a_click_via_timer_returned:** `17023`
- **phase_a_row_counts:** `{'ZZ_SCRATCH_STATUS': 17023, 'ZZ_SCRATCH_P_STATUS': 17022}`
- **diag_form_controls_via_com:**
    - total_controls_recursive: 136
    - subform_controls: `['ZZ_SCRATCH_STATUS', 'ZZ_SCRATCH_P_STATUS']`
    - all_cmd_controls (count: 16): `['CmdQuery', 'CmdPickStatus', 'CmdGIS', 'CmdFanti', 'CmdJianti', 'CmdSelectPlace', 'CmdImportPlaces', 'CmdAllPlaces', 'CmdHelp', 'CmdStoreID', 'CmdFromDynasty', 'CmdToDynasty', 'CmdAllDynasties', 'CmdNeo4j', 'CmdImportStatusCodes', 'CmdSaveStatusCodes']`
    - controls_named_pajek_or_gephi: `[]`
- **diag_uia_descendants:**
    - access_window_hwnd: 30738458, title='Welcome to CBDB!'
    - uia_descendants_total: 1187
    - control_type_histogram: `{'Pane': 62, 'Window': 11, 'ScrollBar': 20, 'Button': 428, 'Thumb': 21, 'TitleBar': 13, 'MenuBar': 12, 'MenuItem': 17, 'ToolBar': 1, 'SplitButton': 2, 'Tab': 1, 'TabItem': 5, 'Edit': 2, 'StatusBar': 1, 'Text': 293, 'Custom': 1, 'Group': 5, 'Image': 292}`
    - all_form_windows (UIA OForm): `['Look at Status 查詢社會區分', 'NAVIGATION_PANE', 'LookAtEntry', 'LookAtOffice', 'LookAtTexts', 'LookAtAssociations', 'LookAtPlace', 'LookAtAssociationPairs', 'LookAtKinship', 'LookAtNetworks', 'LookAtGroupData']`
    - lookatstatus_uia_visible_as_form_window: `False`
    - buttons_named_with_cmd_or_pajek_or_gephi: `[]`
- **phase_b_pywinauto_attempts:**
    - strategy=`uia_auto_id_lookup` ok=False error=`TypeError("IUIA.build_condition() got an unexpected keyword argument 'auto_id'")` detail=`None`
    - strategy=`uia_name_lookup` ok=False error=`no UIA Button descendant with name='CmdPajek'` detail=`None`
    - strategy=`win32_enum_child` ok=False error=`no Win32 child window with title matching 'CmdPajek'` detail=`None`
    - strategy=`com_setfocus_sendkeys_enter` ok=False error=`COM Controls('CmdPajek').SetFocus failed: com_error(-2147352567, 'Exception occurred.', (0, None, "Welcome to CBDB! can't find the field 'CmdPajek' referred to in your expression.", 'acmain11.chm', 11730, -2146825823), None)` detail=`None`
- **phase_b_zz_test_debug_msgs:** `[]`
- **phase_b_row_counts:** `{'ZZ_SCRATCH_STATUS': 17023, 'ZZ_SCRATCH_P_STATUS': 17022}`
- **file_count:** 0
- **msgbox_observed:** 0 dialogs

### CmdGephi

- **phase_a_click_via_timer_returned:** `17023`
- **phase_a_row_counts:** `{'ZZ_SCRATCH_STATUS': 17023, 'ZZ_SCRATCH_P_STATUS': 17022}`
- **diag_form_controls_via_com:**
    - total_controls_recursive: 136
    - subform_controls: `['ZZ_SCRATCH_STATUS', 'ZZ_SCRATCH_P_STATUS']`
    - all_cmd_controls (count: 16): `['CmdQuery', 'CmdPickStatus', 'CmdGIS', 'CmdFanti', 'CmdJianti', 'CmdSelectPlace', 'CmdImportPlaces', 'CmdAllPlaces', 'CmdHelp', 'CmdStoreID', 'CmdFromDynasty', 'CmdToDynasty', 'CmdAllDynasties', 'CmdNeo4j', 'CmdImportStatusCodes', 'CmdSaveStatusCodes']`
    - controls_named_pajek_or_gephi: `[]`
- **diag_uia_descendants:**
    - access_window_hwnd: 31262448, title='Welcome to CBDB!'
    - uia_descendants_total: 1187
    - control_type_histogram: `{'Pane': 62, 'Window': 11, 'ScrollBar': 20, 'Button': 428, 'Thumb': 21, 'TitleBar': 13, 'MenuBar': 12, 'MenuItem': 17, 'ToolBar': 1, 'SplitButton': 2, 'Tab': 1, 'TabItem': 5, 'Edit': 2, 'StatusBar': 1, 'Text': 293, 'Custom': 1, 'Group': 5, 'Image': 292}`
    - all_form_windows (UIA OForm): `['Look at Status 查詢社會區分', 'NAVIGATION_PANE', 'LookAtEntry', 'LookAtOffice', 'LookAtTexts', 'LookAtAssociations', 'LookAtPlace', 'LookAtAssociationPairs', 'LookAtKinship', 'LookAtNetworks', 'LookAtGroupData']`
    - lookatstatus_uia_visible_as_form_window: `False`
    - buttons_named_with_cmd_or_pajek_or_gephi: `[]`
- **phase_b_pywinauto_attempts:**
    - strategy=`uia_auto_id_lookup` ok=False error=`TypeError("IUIA.build_condition() got an unexpected keyword argument 'auto_id'")` detail=`None`
    - strategy=`uia_name_lookup` ok=False error=`no UIA Button descendant with name='CmdGephi'` detail=`None`
    - strategy=`win32_enum_child` ok=False error=`no Win32 child window with title matching 'CmdGephi'` detail=`None`
    - strategy=`com_setfocus_sendkeys_enter` ok=False error=`COM Controls('CmdGephi').SetFocus failed: com_error(-2147352567, 'Exception occurred.', (0, None, "Welcome to CBDB! can't find the field 'CmdGephi' referred to in your expression.", 'acmain11.chm', 11730, -2146825823), None)` detail=`None`
- **phase_b_zz_test_debug_msgs:** `[]`
- **phase_b_row_counts:** `{'ZZ_SCRATCH_STATUS': 17023, 'ZZ_SCRATCH_P_STATUS': 17022}`
- **file_count:** 0
- **msgbox_observed:** 0 dialogs

## Interpretation (per button)

| Button | Outcome | q1 truly fired | q2 marker | q3 Object required | q4 file>=1 | q5 scratch baseline |
|---|---|---|---|---|---|---|
| **CmdPajek** | `ui_path_infeasible_no_button_on_form` | False | False | False | False | True |
| **CmdGephi** | `ui_path_infeasible_no_button_on_form` | False | False | False | False | True |

Per-button raw signals:

- **CmdPajek**: `{'phase_b_enter_seen': False, 'phase_b_done_seen': False, 'phase_b_msgbox_seen': False, 'phase_b_err_seen': False, 'phase_b_object_required': False, 'phase_b_file_count': 0, 'button_exists_as_com_control': False, 'button_exists_as_uia_button': False, 'any_pywinauto_attempt_ok': False, 'n_attempts': 4}`
- **CmdGephi**: `{'phase_b_enter_seen': False, 'phase_b_done_seen': False, 'phase_b_msgbox_seen': False, 'phase_b_err_seen': False, 'phase_b_object_required': False, 'phase_b_file_count': 0, 'button_exists_as_com_control': False, 'button_exists_as_uia_button': False, 'any_pywinauto_attempt_ok': False, 'n_attempts': 4}`

## Verdict

- **bucket:** `ui_direct_simulation_infeasible_no_ui_buttons`
- **next_step:** `pure_maintainer_line_only`
- **recommendation:** UI direct simulation is mechanically impossible for these buttons on the LookAtStatus form: the form's design has no UI controls named CmdPajek / CmdGephi (verified by recursive COM Controls enumeration AND UIA descendants scan). This probe runtime-corroborates the existing canonical missing-UI issues #16 (LookAtStatus is missing its CmdPajek button) and #17 (LookAtStatus is missing its CmdGephi button); the same shape is also already filed as canonical Issue #18 for CmdUCINet. This PR does NOT expand canonical scope, does NOT open a new issue candidate, and does NOT add the missing-UI fix to the existing maintainer-line. It supplies runtime corroboration for canonical Issues #16/#17 and closes the pywinauto fallback line. Next step remains the existing pure maintainer-line.

## Self-review (per docs/skills/programmer-self-review-template.md)

**A. Branch shape**
- [x] Branch cut clean from current `main` (`6b06d6a`).
- [x] `git diff --name-only main..HEAD` contains only the 3 permitted artifact files (probe py + md + json).
- [x] `git diff --stat main..HEAD` is additive-only.

**B. Source-of-truth sync**
- [x] Paired MD + JSON updated together (`--reclassify-from-json` byte-identical roundtrip verified).
- [x] No canonical-issue / triage / inventory drift (this PR doesn't touch those surfaces).
- N/A — bilingual: probe artifact PR; no EN/ZH tier summaries to sync.

**C. Evidence vs claim**
- [x] Raw facts (per-button raw_signals + per-attempt trail + diagnostic enumerations + ZZ_TEST_DEBUG transcripts) recorded separately from interpretation/classification.
- [x] Verdict bucket follows mechanically from raw facts via `_classify_button` + `_verdict`; no interpretation smuggled into raw fields.
- [x] No extrapolation: this probe tests UI direct simulation for Status × CmdPajek / CmdGephi only; no claims about other forms / buttons.
- [x] No runtime behavioural pin missing — UI direct simulation is a runtime test and we ran it.

**D. Residual risk**
- [x] What we did NOT verify: hardware-level mouse-event simulation (e.g. SendInput) at OS level below pywinauto's abstractions. If pywinauto's UIA + Win32 backends + COM SetFocus can't see the controls, lower-level simulation has no remaining surface to click — but that gap is acknowledged.
- [x] Runtime corroboration of existing canonical missing-UI issues: this probe's diagnostic enumerations (recursive COM `Forms('LookAtStatus').Controls` finding 16 Cmd* controls but no CmdPajek/CmdGephi; UIA descendants scan finding zero matching buttons; grep of `analysis/dump/vba/Form_LookAtStatus.vb` finding zero `Me.CmdPajek.*` / `Me.CmdGephi.*` / `Me.CmdUCINet.*` references) all runtime-confirm the immediate UI blocker already captured by canonical Issues #16 (`LookAtStatus is missing its CmdPajek button`) and #17 (`LookAtStatus is missing its CmdGephi button`); the same shape is also already filed for CmdUCINet as canonical Issue #18. This PR does NOT discover a new finding and does NOT open a candidate for canonicalization — it provides runtime evidence for issues that are already canonical.
- [x] No downstream-work pre-claim: this PR does NOT claim UI fallback would unblock any other form/button.

**Pytest scope actually run**: artifacts-only diff; no changes to `tests/`, `tests/cbdb_driver/`, or any production code. `pytest --collect-only` succeeds; fast non-COM subset (`tests/test_schema.py` + `tests/test_saved_views.py`) was the run. Full COM matrix (~138 tests, hours) NOT re-run because the diff cannot regress matrix behaviour by construction.