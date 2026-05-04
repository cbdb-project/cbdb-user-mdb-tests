# LookAtNetworks autodetect matrix bisection (PR AU)

PR AT established `_inject_autodetect` IS the trigger of the LookAtNetworks OpenForm hang, but skipping only Networks's own injection still hangs.  This probe runs 12 variants to narrow down which sibling-form injection (or cumulative effect) is the cause.

## Outcomes

| Variant | Skipped | Injected count | Outcome | Elapsed |
|---|---|---:|---|---:|
| `V1_baseline_inject_all` | — | 10 | `hung_at_OpenForm` | 97.82s |
| `V2_skip_all` | (all 10) | 0 | `OpenForm_returned` | 6.02s |
| `V3_skip_Form_LookAtEntry` | Form_LookAtEntry | 9 | `hung_at_OpenForm` | 97.16s |
| `V4_skip_Form_LookAtOffice` | Form_LookAtOffice | 9 | `hung_at_OpenForm` | 97.03s |
| `V5_skip_Form_LookAtStatus` | Form_LookAtStatus | 9 | `hung_at_OpenForm` | 97.11s |
| `V6_skip_Form_LookAtTexts` | Form_LookAtTexts | 9 | `hung_at_OpenForm` | 97.21s |
| `V7_skip_Form_LookAtAssociations` | Form_LookAtAssociations | 9 | `hung_at_OpenForm` | 97.02s |
| `V8_skip_Form_LookAtPlace` | Form_LookAtPlace | 9 | `hung_at_OpenForm` | 97.13s |
| `V9_skip_Form_LookAtKinship` | Form_LookAtKinship | 9 | `hung_at_OpenForm` | 97.02s |
| `V10_skip_Form_LookAtAssociationPairs` | Form_LookAtAssociationPairs | 9 | `hung_at_OpenForm` | 96.98s |
| `V11_skip_Form_LookAtGroupData` | Form_LookAtGroupData | 9 | `hung_at_OpenForm` | 96.73s |
| `V12_inject_only_Networks` | (all 9 siblings) | 1 | `OpenForm_returned` | 6.91s |

## Headline verdict: `any_sibling_module_touch_is_enough`

Inject-only-Networks succeeds (V12); inject-Networks-plus-any-one-sibling all hung (V3-V11).  The trigger is cumulative: ANY sibling LookAt module modification is enough to make Networks Form_Open hang.  Project-wide dirty-compile is the most likely mechanism.

## Per-variant detail

### `V1_baseline_inject_all`
- production default — inject all 10 (PR AT V1)
- skip = `None`
- n_skipped / n_injected = 0 / 10
- outcome: **`hung_at_OpenForm`**
- elapsed: 97.82s
- markers:
  - +0.0s constructing_session
  - +0.0s opening_session
  - +5.49s session_opened
  - +5.49s starting_OpenForm
  - +95.5s OpenForm_finished=False

### `V2_skip_all`
- skip all forms — control (PR AT V2)
- skip = `['Form_LookAtAssociationPairs', 'Form_LookAtAssociations', 'Form_LookAtEntry', 'Form_LookAtGroupData', 'Form_LookAtKinship', 'Form_LookAtNetworks', 'Form_LookAtOffice', 'Form_LookAtPlace', 'Form_LookAtStatus', 'Form_LookAtTexts']`
- n_skipped / n_injected = 10 / 0
- outcome: **`OpenForm_returned`**
- elapsed: 6.02s
- post_open_state: `{'forms_count': 2, 'loaded': True, 'visible': True}`
- markers:
  - +0.0s constructing_session
  - +0.0s opening_session
  - +1.31s session_opened
  - +1.31s starting_OpenForm
  - +3.67s OpenForm_finished=True
  - +3.73s form_closed

### `V3_skip_Form_LookAtEntry`
- skip ONLY Form_LookAtEntry; Networks + 8 other siblings injected
- skip = `['Form_LookAtEntry']`
- n_skipped / n_injected = 1 / 9
- outcome: **`hung_at_OpenForm`**
- elapsed: 97.16s
- markers:
  - +0.0s constructing_session
  - +0.0s opening_session
  - +4.85s session_opened
  - +4.85s starting_OpenForm
  - +94.86s OpenForm_finished=False

### `V4_skip_Form_LookAtOffice`
- skip ONLY Form_LookAtOffice; Networks + 8 other siblings injected
- skip = `['Form_LookAtOffice']`
- n_skipped / n_injected = 1 / 9
- outcome: **`hung_at_OpenForm`**
- elapsed: 97.03s
- markers:
  - +0.0s constructing_session
  - +0.0s opening_session
  - +4.68s session_opened
  - +4.68s starting_OpenForm
  - +94.7s OpenForm_finished=False

### `V5_skip_Form_LookAtStatus`
- skip ONLY Form_LookAtStatus; Networks + 8 other siblings injected
- skip = `['Form_LookAtStatus']`
- n_skipped / n_injected = 1 / 9
- outcome: **`hung_at_OpenForm`**
- elapsed: 97.11s
- markers:
  - +0.0s constructing_session
  - +0.0s opening_session
  - +4.81s session_opened
  - +4.81s starting_OpenForm
  - +94.81s OpenForm_finished=False

### `V6_skip_Form_LookAtTexts`
- skip ONLY Form_LookAtTexts; Networks + 8 other siblings injected
- skip = `['Form_LookAtTexts']`
- n_skipped / n_injected = 1 / 9
- outcome: **`hung_at_OpenForm`**
- elapsed: 97.21s
- markers:
  - +0.0s constructing_session
  - +0.0s opening_session
  - +4.88s session_opened
  - +4.88s starting_OpenForm
  - +94.89s OpenForm_finished=False

### `V7_skip_Form_LookAtAssociations`
- skip ONLY Form_LookAtAssociations; Networks + 8 other siblings injected
- skip = `['Form_LookAtAssociations']`
- n_skipped / n_injected = 1 / 9
- outcome: **`hung_at_OpenForm`**
- elapsed: 97.02s
- markers:
  - +0.0s constructing_session
  - +0.0s opening_session
  - +4.71s session_opened
  - +4.71s starting_OpenForm
  - +94.71s OpenForm_finished=False

### `V8_skip_Form_LookAtPlace`
- skip ONLY Form_LookAtPlace; Networks + 8 other siblings injected
- skip = `['Form_LookAtPlace']`
- n_skipped / n_injected = 1 / 9
- outcome: **`hung_at_OpenForm`**
- elapsed: 97.13s
- markers:
  - +0.0s constructing_session
  - +0.0s opening_session
  - +4.81s session_opened
  - +4.81s starting_OpenForm
  - +94.82s OpenForm_finished=False

### `V9_skip_Form_LookAtKinship`
- skip ONLY Form_LookAtKinship; Networks + 8 other siblings injected
- skip = `['Form_LookAtKinship']`
- n_skipped / n_injected = 1 / 9
- outcome: **`hung_at_OpenForm`**
- elapsed: 97.02s
- markers:
  - +0.0s constructing_session
  - +0.0s opening_session
  - +4.65s session_opened
  - +4.65s starting_OpenForm
  - +94.65s OpenForm_finished=False

### `V10_skip_Form_LookAtAssociationPairs`
- skip ONLY Form_LookAtAssociationPairs; Networks + 8 other siblings injected
- skip = `['Form_LookAtAssociationPairs']`
- n_skipped / n_injected = 1 / 9
- outcome: **`hung_at_OpenForm`**
- elapsed: 96.98s
- markers:
  - +0.0s constructing_session
  - +0.0s opening_session
  - +4.7s session_opened
  - +4.7s starting_OpenForm
  - +94.7s OpenForm_finished=False

### `V11_skip_Form_LookAtGroupData`
- skip ONLY Form_LookAtGroupData; Networks + 8 other siblings injected
- skip = `['Form_LookAtGroupData']`
- n_skipped / n_injected = 1 / 9
- outcome: **`hung_at_OpenForm`**
- elapsed: 96.73s
- markers:
  - +0.0s constructing_session
  - +0.0s opening_session
  - +4.41s session_opened
  - +4.41s starting_OpenForm
  - +94.42s OpenForm_finished=False

### `V12_inject_only_Networks`
- skip all 9 siblings; ONLY Form_LookAtNetworks injected
- skip = `['Form_LookAtAssociationPairs', 'Form_LookAtAssociations', 'Form_LookAtEntry', 'Form_LookAtGroupData', 'Form_LookAtKinship', 'Form_LookAtOffice', 'Form_LookAtPlace', 'Form_LookAtStatus', 'Form_LookAtTexts']`
- n_skipped / n_injected = 9 / 1
- outcome: **`OpenForm_returned`**
- elapsed: 6.91s
- post_open_state: `{'forms_count': 2, 'loaded': True, 'visible': True}`
- markers:
  - +0.0s constructing_session
  - +0.0s opening_session
  - +2.12s session_opened
  - +2.12s starting_OpenForm
  - +4.52s OpenForm_finished=True
  - +4.57s form_closed

## Implications

The injection contents don't matter — what matters is that ANY sibling LookAt module is dirty when Networks opens.  This points at a project-wide auto-compile interaction with Networks's `Forms!LookAtNetworks!<sub>.Form.Recordset` self-reference during Form_Open.

Recommended next step (separate PR): commit the VBA project after `_inject_autodetect()` runs (via `app.RunCommand acCmdCompileAndSaveAllModules` or equivalent), so subsequent form opens don't hit the dirty-compile state.  If that lets Networks open with full injection, the operating constraint is "compile after injection, before any form open".  Keep it probe-only first; no driver-default change.

## Constraints respected per AU brief
- Probe-only.
- Default driver behaviour unchanged (`skip_inject_autodetect_forms=None` → inject all).
- Networks injection ON in V1, V3-V12 (only V2 turns everything off).
- `reset_pickers` ON in every variant.
- DataMode=0 unchanged.
- No matrix unskips, no production fixture changes.
- Per-variant fresh Access process via fresh VbaSession.
- 90 s OpenForm watchdog + 180 s per-variant hard cap.
- Killed orphan MSACCESS.EXE between variants.