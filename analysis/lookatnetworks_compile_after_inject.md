# LookAtNetworks compile-after-inject mitigation probe (PR AV)

PR AU established that ANY sibling LookAt module modification dirties the VBA project, deadlocking Networks's Form_Open on the project-wide auto-compile.  This probe tests the direct mitigation: explicitly compile + save all modules after `_inject_autodetect()` runs and before any form opens.

## Outcomes

| Variant | Compile method | Outcome | Compile elapsed | OpenForm elapsed | Total |
|---|---|---|---:|---:|---:|
| `V1_baseline_no_compile` | `no_compile` | `hung_at_OpenForm` | — | `OpenForm_finished=False` | 97.56s |
| `V2_RunCommand_126` | `RunCommand_126` | `hung_at_OpenForm` | 0.34s | `OpenForm_finished=False` | 97.76s |
| `V3_VBE_touch` | `VBE_touch_components` | `hung_at_OpenForm` | 0.16s | `OpenForm_finished=False` | 97.46s |

## Headline verdict: `no_compile_method_helped`

Neither RunCommand nor VBE-touch unblocked Networks.  The dirty-project-compile theory may need refinement; next mitigation candidate is to open LookAtNetworks BEFORE the sibling injections run (open-first ordering).

## Per-variant detail

### `V1_baseline_no_compile`
- full inject, no compile (PR AT V1 reproduction)
- compile method: `no_compile`
- outcome: **`hung_at_OpenForm`**
- elapsed: 97.56s
- markers:
  - +0.0s constructing_session
  - +0.0s opening_session
  - +5.23s session_opened
  - +5.23s starting_OpenForm
  - +95.25s OpenForm_finished=False

### `V2_RunCommand_126`
- full inject + acCmdCompileAndSaveAllModules
- compile method: `RunCommand_126`
- outcome: **`hung_at_OpenForm`**
- elapsed: 97.76s
- compile detail: `{'method': 'RunCommand_126', 'elapsed_sec': 0.34, 'outcome': 'compile_returned', 'exception': None}`
- markers:
  - +0.0s constructing_session
  - +0.0s opening_session
  - +5.01s session_opened
  - +5.01s compile_start_RunCommand_126
  - +5.34s compile_returned_in_0.34s
  - +5.34s starting_OpenForm
  - +95.36s OpenForm_finished=False

### `V3_VBE_touch`
- full inject + VBE.VBComponents touch (fallback)
- compile method: `VBE_touch_components`
- outcome: **`hung_at_OpenForm`**
- elapsed: 97.46s
- compile detail: `{'method': 'VBE_dummy_write', 'elapsed_sec': 0.16, 'outcome': 'vbe_touch_complete', 'exception': None, 'proj_mode_before': 2, 'n_components_touched': 69}`
- markers:
  - +0.0s constructing_session
  - +0.0s opening_session
  - +5.02s session_opened
  - +5.18s vbe_touch_done_69_comps_in_0.16s
  - +5.18s starting_OpenForm
  - +95.18s OpenForm_finished=False

## Implications

Compile-after-injection didn't help.  The dirty-project-compile theory needs refinement.  Next candidate (separate PR): open LookAtNetworks BEFORE `_inject_autodetect()` runs, then close it; subsequent re-opens may not hit the deadlock since Networks's form-cache state is already established.

## Constraints respected per AV brief
- Probe-only.  No matrix unskips.
- Default driver behaviour unchanged (no `compile_after_inject` option added; the test would be premature without first verifying mitigation works).
- Full injection ON in every variant (per brief).
- Networks injection ON.
- `reset_pickers` unchanged.
- DataMode=0 unchanged.
- Per-variant fresh Access process via fresh VbaSession.
- 90 s OpenForm watchdog + 60 s compile timeout + 240 s per-variant hard cap.
- Killed orphan MSACCESS.EXE between variants.
- Did NOT attempt CmdRun (deferred to a follow-up PR).