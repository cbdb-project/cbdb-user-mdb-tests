# click_via_timer second-dispatch infra investigation (diagnostic-only; no public driver edit)

**Date:** 2026-05-08  ·  **Branch:** `investigate/click-via-timer-second-dispatch` (off main `948978d`)

Diagnoses PR #135's 'second `click_via_timer` call doesn't dispatch' failure by bypassing `click_via_timer` and using raw Access COM with an instrumented Form_Timer that writes `TIMER_FIRED` and `TIMER_RETURNED` markers bracketing `Call <ctl>_Click`.

## Probe shape

Per phase (one per export button):

1. Open VbaSession + seed fixture.
2. **Phase A** — inject probe-instrumented Form_Timer for `CmdQuery`, arm via OnTimer + TimerInterval=100, wait for `TIMER_FIRED CmdQuery` and `TIMER_RETURNED CmdQuery` markers.  Capture form module state at 3 checkpoints (after inject, after arm, after completion).
3. Clear `ZZ_TEST_DEBUG`; set Form.Tag for export path; Python sleep 1.5 s.
4. **Phase B** — inject probe-instrumented Form_Timer for `Cmd<button>`, arm, wait up to 60 s for `TIMER_FIRED <button>`, `TIMER_FIRED CmdQuery` (would indicate stale binding), or `TIMER_RETURNED <button>`.  Capture state.
5. Capture Phase B `ZZ_TEST_DEBUG`, files, scratch counts.

Bypasses `click_via_timer` entirely.  All instrumentation is in the probe script; **no public driver edit**.

**Total wall:** 155.97 s.

## Raw observed facts (per phase)

### Phase: `CmdPajek`

- **per-phase outcome:** `second_dispatch_timer_did_not_fire`
- Phase A inject: `True` — instrumented Form_Timer for ctl=CmdQuery; force-compile OK
- Phase A arm: `True` — armed (OnTimer=[Event Procedure], TimerInterval=100)
- Phase A `TIMER_FIRED CmdQuery` seen: `True` (after 0.5 s)
- Phase A `TIMER_RETURNED CmdQuery` seen: `True` (after 1.0 s)
- Phase A scratch: {'ZZ_SCRATCH_STATUS': 17023, 'ZZ_SCRATCH_P_STATUS': 17022}
- Phase A ZZ_TEST_DEBUG: ['TIMER_FIRED CmdQuery', 'LookAtStatus:ENTER', 'LookAtStatus:DONE', 'TIMER_RETURNED CmdQuery']

- Phase B inject: `True` — instrumented Form_Timer for ctl=CmdPajek; force-compile OK
- Phase B arm: `True` — armed (OnTimer=[Event Procedure], TimerInterval=100)
- Phase B `TIMER_FIRED CmdPajek` seen: `False`
- Phase B `TIMER_FIRED CmdQuery` seen (would indicate stale binding): `False`
- Phase B `TIMER_RETURNED CmdPajek` seen: `False`
- Phase B file_count: 0
- Phase B scratch: {'ZZ_SCRATCH_STATUS': 17023, 'ZZ_SCRATCH_P_STATUS': 17022}
- Phase B ZZ_TEST_DEBUG: []
- Phase B OnTimer property after arm: `[Event Procedure]`
- Phase B TimerInterval after arm: `100`
- Phase B TimerInterval after wait: `0`
- per-phase elapsed: 74.73 s

**Phase A Form_Timer source (after inject):**
```vba
' PROBE_INSTRUMENTED_FORM_TIMER ctl=CmdQuery
Private Sub Form_Timer()
    On Error Resume Next
    CurrentDb.Execute "INSERT INTO ZZ_TEST_DEBUG (msg) VALUES ('TIMER_FIRED CmdQuery')"
    Me.TimerInterval = 0
    Call CmdQuery_Click
    CurrentDb.Execute "INSERT INTO ZZ_TEST_DEBUG (msg) VALUES ('TIMER_RETURNED CmdQuery')"
End Sub
```

**Phase B Form_Timer source (after inject):**
```vba
' PROBE_INSTRUMENTED_FORM_TIMER ctl=CmdPajek
Private Sub Form_Timer()
    On Error Resume Next
    CurrentDb.Execute "INSERT INTO ZZ_TEST_DEBUG (msg) VALUES ('TIMER_FIRED CmdPajek')"
    Me.TimerInterval = 0
    Call CmdPajek_Click
    CurrentDb.Execute "INSERT INTO ZZ_TEST_DEBUG (msg) VALUES ('TIMER_RETURNED CmdPajek')"
End Sub
```

**Phase B Form_Timer source (after wait):**
```vba
' PROBE_INSTRUMENTED_FORM_TIMER ctl=CmdPajek
Private Sub Form_Timer()
    On Error Resume Next
    CurrentDb.Execute "INSERT INTO ZZ_TEST_DEBUG (msg) VALUES ('TIMER_FIRED CmdPajek')"
    Me.TimerInterval = 0
    Call CmdPajek_Click
    CurrentDb.Execute "INSERT INTO ZZ_TEST_DEBUG (msg) VALUES ('TIMER_RETURNED CmdPajek')"
End Sub
```

### Phase: `CmdGephi`

- **per-phase outcome:** `second_dispatch_timer_did_not_fire`
- Phase A inject: `True` — instrumented Form_Timer for ctl=CmdQuery; force-compile OK
- Phase A arm: `True` — armed (OnTimer=[Event Procedure], TimerInterval=100)
- Phase A `TIMER_FIRED CmdQuery` seen: `True` (after 0.5 s)
- Phase A `TIMER_RETURNED CmdQuery` seen: `True` (after 1.0 s)
- Phase A scratch: {'ZZ_SCRATCH_STATUS': 17023, 'ZZ_SCRATCH_P_STATUS': 17022}
- Phase A ZZ_TEST_DEBUG: ['TIMER_FIRED CmdQuery', 'LookAtStatus:ENTER', 'LookAtStatus:DONE', 'TIMER_RETURNED CmdQuery']

- Phase B inject: `True` — instrumented Form_Timer for ctl=CmdGephi; force-compile OK
- Phase B arm: `True` — armed (OnTimer=[Event Procedure], TimerInterval=100)
- Phase B `TIMER_FIRED CmdGephi` seen: `False`
- Phase B `TIMER_FIRED CmdQuery` seen (would indicate stale binding): `False`
- Phase B `TIMER_RETURNED CmdGephi` seen: `False`
- Phase B file_count: 0
- Phase B scratch: {'ZZ_SCRATCH_STATUS': 17023, 'ZZ_SCRATCH_P_STATUS': 17022}
- Phase B ZZ_TEST_DEBUG: []
- Phase B OnTimer property after arm: `[Event Procedure]`
- Phase B TimerInterval after arm: `100`
- Phase B TimerInterval after wait: `0`
- per-phase elapsed: 74.45 s

**Phase A Form_Timer source (after inject):**
```vba
' PROBE_INSTRUMENTED_FORM_TIMER ctl=CmdQuery
Private Sub Form_Timer()
    On Error Resume Next
    CurrentDb.Execute "INSERT INTO ZZ_TEST_DEBUG (msg) VALUES ('TIMER_FIRED CmdQuery')"
    Me.TimerInterval = 0
    Call CmdQuery_Click
    CurrentDb.Execute "INSERT INTO ZZ_TEST_DEBUG (msg) VALUES ('TIMER_RETURNED CmdQuery')"
End Sub
```

**Phase B Form_Timer source (after inject):**
```vba
' PROBE_INSTRUMENTED_FORM_TIMER ctl=CmdGephi
Private Sub Form_Timer()
    On Error Resume Next
    CurrentDb.Execute "INSERT INTO ZZ_TEST_DEBUG (msg) VALUES ('TIMER_FIRED CmdGephi')"
    Me.TimerInterval = 0
    Call CmdGephi_Click
    CurrentDb.Execute "INSERT INTO ZZ_TEST_DEBUG (msg) VALUES ('TIMER_RETURNED CmdGephi')"
End Sub
```

**Phase B Form_Timer source (after wait):**
```vba
' PROBE_INSTRUMENTED_FORM_TIMER ctl=CmdGephi
Private Sub Form_Timer()
    On Error Resume Next
    CurrentDb.Execute "INSERT INTO ZZ_TEST_DEBUG (msg) VALUES ('TIMER_FIRED CmdGephi')"
    Me.TimerInterval = 0
    Call CmdGephi_Click
    CurrentDb.Execute "INSERT INTO ZZ_TEST_DEBUG (msg) VALUES ('TIMER_RETURNED CmdGephi')"
End Sub
```

## Layer diagnosis

MIXED — see per-phase signatures for detail.

## Verdict: `second_dispatch_timer_did_not_fire_both`

**Layer diagnosis:** MIXED — see per-phase signatures for detail.

**Strong evidence collected:**
  - Phase A: timer fires correctly (TIMER_FIRED +     TIMER_RETURNED both observed at 0.5 s)
  - Phase B: TimerInterval transitions 100 → 0     (Access timer subsystem ticks)
  - Phase B: Form_Timer SOURCE in the module is     the new ctl=Cmd<button> body (verified via     direct CodeModule.Lines read)
  - Phase B: OnTimer property =     `[Event Procedure]` (verified before AND after     arm)
  - Phase B: `Application.RunCommand(126)`     (acCmdCompileAndSaveAllModules) succeeded     after AddFromString — so the new body IS     compiled
  - Phase B: NO marker fires (neither for the     new ctl nor for the old) — body never     executes

**Conclusion:** Access's compiled timer-event binding does NOT refresh to the newly-AddFromString'd Form_Timer body, even after explicit `acCmdCompileAndSaveAllModules` AND OnTimer rebind.  The timer-event dispatcher appears to have a per-form-instance cache that survives module recompilation; only a fresh form instance (close + reopen) would refresh it.

**Force-compile rejected as a narrow fix.**  Did not unblock either button.  Both phases still show timer_did_not_fire after the compile fix was applied.

**Candidate fixes RANKED for next brief:**
  1. **close + reopen the form between dispatches**      — heaviest but most likely to work.  Drops      the form's class instance entirely, forcing      Access to re-resolve event handlers from the      freshly-compiled module on Form_Open.       Probably needs the test fixture to reseed      scratch tables.
  2. **inject a Form_Timer in a fresh standard      module + redirect the form's OnTimer to call      into it** — moves the dispatch out of the      form's class module, may bypass the per-     instance cache.
  3. **pywinauto button click** — UI-driven,      completely different surface.  Listed as      fallback in PR #136.
  4. **maintainer-line / canonical Issue** —      leave Status × CmdPajek + CmdGephi skipped;      the underlying CBDB Set-rebind pattern is      the root cause regardless of what      test-driver path we pick.

Per the brief: this PR ships diagnostic + the rejected force-compile candidate.  No public driver edit is appropriate yet (the close+reopen candidate is significant enough to warrant a separate brief and verification probe).

## Constraints honoured per brief

- ✅ Investigation artifacts only — public driver (`tests/cbdb_driver/vba_session.py`) NOT modified by this first-pass diagnostic.  All instrumentation in the probe script.
- ✅ Minimum instrumentation only — TIMER_FIRED + TIMER_RETURNED markers around the Call; module state snapshots at 3 checkpoints.  No deeper VBA changes.
- ✅ Both buttons covered (CmdPajek + CmdGephi).
- ✅ First vs second dispatch comparison explicit (Phase A vs Phase B per session).
- ✅ No `tests/test_*` changed; no README, triage, canonical reports / issue severity touched.
- ✅ CmdNeo4j NOT touched.
- ✅ `--reclassify-from-json` supported.
- ✅ `analysis/report_screenshot_audit.md` drift left alone.