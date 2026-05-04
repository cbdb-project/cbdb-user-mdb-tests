# LookAtNetworks _inject_autodetect bisection (PR AT)

PR AS falsified DataMode.  PR AR's remaining hypothesis (a) — `_inject_autodetect` modifies CmdRun_Click in the VBA module; auto-compile during Form_Open may be the trigger — is what this probe tests.

Held constant: `reset_pickers` ON, DataMode=0, full VbaSession setup otherwise.  Only the new `skip_inject_autodetect_forms` constructor kwarg varies between variants.

## Outcomes

| Variant | Skip | Outcome | Elapsed |
|---|---|---|---:|
| `V1_baseline_inject_all` | — | `hung_at_OpenForm` | 97.57s |
| `V2_skip_all_inject` | all | `OpenForm_returned` | 4.33s |
| `V3_skip_only_networks` | Form_LookAtNetworks | `hung_at_OpenForm` | 96.54s |

**Headline verdict: `_inject_autodetect_IS_the_trigger`**

## Per-variant detail

### `V1_baseline_inject_all`
- production default — autodetect injected for all forms
- skip_inject_autodetect_forms = `None`
- outcome: **`hung_at_OpenForm`**
- elapsed: 97.57s
- markers:
  - +0.0s constructing_session
  - +0.0s opening_session
  - +5.23s session_opened
  - +5.23s starting_OpenForm
  - +95.25s OpenForm_finished=False

### `V2_skip_all_inject`
- skip _inject_autodetect for ALL forms
- skip_inject_autodetect_forms = `[]`
- outcome: **`OpenForm_returned`**
- elapsed: 4.33s
- post_open_state: `{'forms_count': 2, 'loaded': True, 'visible': True}`
- markers:
  - +0.0s constructing_session
  - +0.0s opening_session
  - +1.3s session_opened
  - +1.3s starting_OpenForm
  - +1.92s OpenForm_finished=True
  - +1.98s form_closed

### `V3_skip_only_networks`
- skip _inject_autodetect for Form_LookAtNetworks only
- skip_inject_autodetect_forms = `['Form_LookAtNetworks']`
- outcome: **`hung_at_OpenForm`**
- elapsed: 96.54s
- markers:
  - +0.0s constructing_session
  - +0.0s opening_session
  - +4.2s session_opened
  - +4.2s starting_OpenForm
  - +94.21s OpenForm_finished=False

## Implications — and the V3 surprise

`_inject_autodetect` **IS** the trigger of the OpenForm hang
(V2 succeeded, V1 hung — clean signal).  But the per-form
detail rules out a single-module culprit:

  - V3 skipped autodetect injection for **Form_LookAtNetworks
    only** (other 9 forms still injected) → **still hung**.

So the trigger is **NOT** Networks's own injected
CmdRun_Click body.  It must be one (or the cumulative effect)
of the OTHER 9 forms' injections — i.e. modifying any LookAt
module triggers a project-level state that makes Networks's
Form_Open hang.

Plausible mechanisms (testable in a follow-up bisection):

  1. **VBA-project-wide auto-compile** — when a form opens,
     Access verifies the VBE project; if multiple sibling
     modules have uncompiled changes pending, the verify
     pass touches all of them.  Networks's Form_Open
     specifically does `Forms!LookAtNetworks!<sub>.Form
     .Recordset` (self-reference during Open, the AGENTS
     landmine), and the compile-while-open interaction may
     deadlock specifically on Networks because of that
     pattern.
  2. **Module-by-module bisection** — toggle each of the 9
     other forms' autodetect entries individually with
     V3-style probes to find which one(s) trigger the hang.
     If just one form does, that form's injection BODY can
     be bisected next.  If all 9 contribute, mechanism #1
     is the more likely root cause.

Recommended next step (separate PR): a 9-variant probe that
skips one OTHER form's autodetect at a time, with Networks's
own injection always ON.  Whichever form's skip lets Networks
open is the trigger; then bisect that form's injection body
or accept "any LookAt module touch breaks Networks" as the
operating constraint.

Do NOT fix in this PR.  The Networks autodetect entry is real
test infrastructure (gUsePersonID + gUseADDRID need to be set
for CmdRun); a fix that "skip Networks injection if any other
LookAt is opened first" is not yet justified.

### Workaround usable today (without a fix)

The new `skip_inject_autodetect_forms=set()` constructor
kwarg gives a clean path for a focused experiment:

```python
sess = VbaSession(USER_MDB, WORK,
                  skip_inject_autodetect_forms=set())
sess.open()
sess.app.DoCmd.OpenForm("LookAtNetworks", 0, "", "", 0, 0)
# Form opens in ~2s.  But: gUsePersonID auto-detect is OFF for
# Networks (and every other form), so CmdRun's gating MsgBox
# will fire unless gUsePersonID is set through some other
# mechanism (manual VBA call, picker UI, etc.).  Fine for a
# Form_Open-only diagnostic; not enough for an unattended
# CmdRun test.
```

## Constraints respected
- Probe-only.
- Driver opt-out is a constructor kwarg with default `None` → production behaviour preserved.  Production tests do NOT pass this kwarg.
- `reset_pickers` unchanged.
- DataMode=0 unchanged (PR AS-confirmed not the cause).
- Networks autodetect entry NOT removed.
- No matrix unskips.  No production fixture changes.
- Per-variant fresh Access process via fresh VbaSession.
- 90 s OpenForm watchdog + 180 s per-variant hard cap.
- Fast suite still 111 passed, 9 skipped (pre-probe).