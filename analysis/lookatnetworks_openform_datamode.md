# LookAtNetworks OpenForm DataMode bisection (PR AS)

PR AR found that `VbaSession.open_form("LookAtNetworks")` hangs while PR AA's stripped-down `app.DoCmd.OpenForm(..., DataMode=2)` returned in ~2 s. Three possible triggers: DataMode value, autodetect injection, reset_pickers.  This probe holds the latter two constant (full `VbaSession.open()`) and varies just the DataMode arg passed to OpenForm.

## Outcomes

| Variant | DataMode | Outcome | Elapsed | OpenForm marker |
|---|---:|---|---:|---|
| `D0_acFormPropertySettings` | 0 | `hung_at_OpenForm` | 97.59s | `OpenForm_finished=False` |
| `D1_acFormAdd` | 1 | `hung_at_OpenForm` | 97.37s | `OpenForm_finished=False` |
| `D2_acFormEdit` | 2 | `hung_at_OpenForm` | 97.38s | `OpenForm_finished=False` |

## Per-variant detail

### `D0_acFormPropertySettings` (DataMode=0)
- VbaSession default (PR AR's hang)
- outcome: **`hung_at_OpenForm`**
- elapsed: 97.59s
- markers:
  - +0.0s opening_session
  - +5.25s session_opened
  - +5.25s starting_OpenForm
  - +95.26s OpenForm_finished=False

### `D1_acFormAdd` (DataMode=1)
- acFormAdd
- outcome: **`hung_at_OpenForm`**
- elapsed: 97.37s
- markers:
  - +0.0s opening_session
  - +5.05s session_opened
  - +5.05s starting_OpenForm
  - +95.07s OpenForm_finished=False

### `D2_acFormEdit` (DataMode=2)
- PR AA's stripped-probe success
- outcome: **`hung_at_OpenForm`**
- elapsed: 97.38s
- markers:
  - +0.0s opening_session
  - +5.07s session_opened
  - +5.07s starting_OpenForm
  - +95.07s OpenForm_finished=False

## Verdict

**All variants hung** — DataMode is not the trigger.  The hang must come from autodetect injection or reset_pickers.  Recommended next step: bisect axis (a) `_inject_autodetect` for Networks (open the form before the injection runs, or skip injection entirely).

## Constraints respected
- Per-variant fresh Access process via fresh VbaSession.
- 90 s OpenForm watchdog + 180 s per-variation hard cap.
- Did NOT remove _AUTODETECT[Form_LookAtNetworks].
- Did NOT skip reset_pickers.
- No matrix unskips, no production fixture changes.
- Fast suite still 111 passed, 9 skipped (pre-probe check).