# LookAtNetworks warm-open-before-inject probe (PR AW)

PR AV falsified the dirty-compile theory.  PR AU's remaining "any sibling module touch is enough" verdict suggests the trigger is Access's per-form compiled-class cache being invalidated by sibling module modifications.  This probe tests whether opening Networks BEFORE any sibling injection runs establishes form state that survives later modifications.

## Outcomes

| Variant | Outcome | Elapsed |
|---|---|---:|
| `W1_baseline_inject_then_open` | `hung_at_OpenForm` | 95.21s |
| `W2_warm_close_inject_reopen` | `hung_at_second_OpenForm` | 95.77s |
| `W3_warm_keep_loaded_inject` | `OpenForm_returned` | 5.84s |
| `W4_warm_close_inject_siblings_only_reopen` | `hung_at_second_OpenForm` | 96.66s |

## Headline verdict: `warmup_works_W3_only`

Keep-form-loaded-during-inject works, but close+reopen (W2) does NOT.  More restrictive: production test must keep Networks loaded across sibling injection.  Riskier.

## Per-variant detail

### `W1_baseline_inject_then_open`
- full inject first, then open Networks (expected hang; PR AT V1)
- outcome: **`hung_at_OpenForm`**
- elapsed: 95.21s
- markers:
  - +0.0s constructing_session_default
  - +0.0s opening_session
  - +5.21s session_opened_with_full_inject
  - +5.21s starting_OpenForm_after_inject
  - +95.21s OpenForm_finished=False

### `W2_warm_close_inject_reopen`
- open Networks → close → inject all 10 → reopen Networks
- outcome: **`hung_at_second_OpenForm`**
- elapsed: 95.77s
- markers:
  - +0.0s constructing_session_skip_all_inject
  - +0.0s opening_session
  - +1.29s session_opened_no_inject
  - +1.29s starting_first_OpenForm
  - +1.92s first_OpenForm_finished=True
  - +1.92s first_OpenForm_returned
  - +1.97s first_form_closed
  - +1.97s starting_inject_post_warmup
  - +5.76s inject_finished=True
  - +5.76s starting_second_OpenForm
  - +95.77s second_OpenForm_finished=False

### `W3_warm_keep_loaded_inject`
- open Networks → keep loaded → inject all 10 → check usable
- outcome: **`OpenForm_returned`**
- elapsed: 5.84s
- post_open_state: `{'forms_count': 11, 'loaded': True, 'visible': True}`
- markers:
  - +0.0s constructing_session_skip_all_inject
  - +0.0s opening_session
  - +1.3s session_opened_no_inject
  - +1.3s starting_first_OpenForm
  - +1.92s first_OpenForm_finished=True
  - +1.92s first_OpenForm_returned_keeping_loaded
  - +1.92s starting_inject_with_form_loaded
  - +5.82s inject_finished=True
  - +5.82s checking_form_state_post_inject
  - +5.84s form_closed_after_inject

### `W4_warm_close_inject_siblings_only_reopen`
- open Networks → close → inject 9 siblings (skip Networks) → reopen
- outcome: **`hung_at_second_OpenForm`**
- elapsed: 96.66s
- markers:
  - +0.0s constructing_session_skip_all_inject
  - +0.0s opening_session
  - +1.29s session_opened_no_inject
  - +1.29s starting_first_OpenForm
  - +3.64s first_OpenForm_finished=True
  - +3.64s first_OpenForm_returned
  - +3.69s first_form_closed
  - +3.69s starting_inject_skip_only_Networks
  - +6.65s inject_finished=True
  - +6.65s starting_second_OpenForm
  - +96.66s second_OpenForm_finished=False

## Implications

More restrictive working recipe.  The form must stay loaded across the injection.  Probably a harder ergonomic for the existing matrix harness (open_form is a one-shot helper, not open-and-keep).  Recommend trying W2-equivalent alternatives first (e.g. opening + closing more than once) before committing to keep-loaded.

## Constraints respected per AW brief
- Probe-only.  No matrix unskips.
- Default driver behaviour unchanged (probe uses the PR AT `skip_inject_autodetect_forms` kwarg + direct private-attr toggling between phases; no new driver code added).
- Networks injection ON in W1 / W2 / W3 (W4 skips by design per the brief).
- `reset_pickers` ON (runs inside `sess.open()` before all warm opens).
- DataMode=0 unchanged.
- Per-variant fresh Access process.
- 90 s OpenForm watchdog × 2 + 60 s inject watchdog + 360 s per-variant outer cap.
- Killed orphan MSACCESS.EXE between variants.
- Did NOT attempt CmdRun (deferred to a follow-up PR).