# LookAtStatus × {CmdPajek, CmdGephi} direct-invocation probe (Application.Run via COM)

**Date:** 2026-05-08  ·  **Branch:** `investigate/status-direct-invoke` (off main `5f100b4`; rebased to current main)

Bypasses the `Form_Timer` second-dispatch limitation PR #135 surfaced.  Tests whether `Cmd<X>_Click` actually runs post-CmdQuery when invoked via Access COM `Application.Run` through a Public-wrapper sub.  Isolates whether the blocker is the export sub itself or the `click_via_timer` / Form.Timer re-dispatch infrastructure.

## Probe shape

Per phase (one per export button), in a fresh VbaSession with its own MDB copy:

1. Open form, seed fixture (no Form.Tag chain).
2. **Inject a probe-script-only Public wrapper** into the Form_LookAtStatus VBA module:
   ```vba
   Public Sub RunExport<Pajek/Gephi>()
       Call Cmd<Pajek/Gephi>_Click
   End Sub
   ```
   The wrapper is needed because `Cmd<X>_Click` is `Private Sub` — `Application.Run` from external COM can't reach Private subs, but a Public wrapper inside the same form module can call them.  Wrappers exist only in the working MDB copy (regenerated per phase) — NEVER committed to `tests/cbdb_driver/vba_session.py`.
3. **Phase A** — fire CmdQuery via standard `click_via_timer`.  Wait for `:DONE`.  Snapshot scratch counts.
4. Clear `ZZ_TEST_DEBUG`.  Set Form.Tag for export path.
5. **Python-side `time.sleep(1.5)`** — matches PR #131's positive signal.
6. **Phase B (DIRECT INVOKE)** — `app.Run('Form_LookAtStatus.RunExport<X>')`.  Synchronous COM call; returns after the wrapper's `Call Cmd<X>_Click` finishes.  Bypasses Form_Timer entirely.  Capture exception text if Run rejects.
7. Capture Phase B `ZZ_TEST_DEBUG`, files, scratch counts, watchdog dialogs.

**Total wall:** 44.68 s.

## Raw observed facts (per phase)

### Phase: `CmdPajek`

- **outcome category:** `direct_invoke_app_run_rejected`
- **wrapper inject:** `True` — wrapper RunExportPajek injected OK
- **Phase A click_via_timer returned:** 17023
- **Phase A elapsed:** 2.03 s
- **COM sleep:** 1.5 s
- **Phase B Application.Run returned OK:** False
- **Phase B Application.Run error (if any):** `attempt 1 'Form_LookAtStatus.RunExportPajek': com_error(-2147352567, 'Exception occurred.', (0, None, "Welcome to CBDB! cannot find the procedure 'Form_LookAtStatus.RunExportPajek.'", None, -1, -2146825771), None) | attempt 2 'RunExportPajek' (unqualified): com_error(-2147352567, 'Exception occurred.', (0, None, "Welcome to CBDB! cannot find the procedure 'RunExportPajek.'", None, -1, -2146825771), None)`
- **Phase B invoke elapsed:** 0.01 s
- **Phase B file_count:** 0
- **Phase B chain_observed_done:** True
- **msgbox_watchdog_count:** 0
- **per-phase wall:** 19.25 s

**Phase A scratch counts (post-CmdQuery):**
- `ZZ_SCRATCH_STATUS`: 17023
- `ZZ_SCRATCH_P_STATUS`: 17022
- `ZZ_TEST_DEBUG`: 2

**Phase A `ZZ_TEST_DEBUG`:**
- `LookAtStatus:ENTER`
- `LookAtStatus:DONE`

**Phase B scratch counts (post-direct-invoke):**
- `ZZ_SCRATCH_STATUS`: 17023
- `ZZ_SCRATCH_P_STATUS`: 17022
- `ZZ_TEST_DEBUG`: 0

**Phase B `ZZ_TEST_DEBUG` (cleared between phases):**
(empty)

**Watchdog MsgBox observations:**
(none observed)

### Phase: `CmdGephi`

- **outcome category:** `direct_invoke_app_run_rejected`
- **wrapper inject:** `True` — wrapper RunExportGephi injected OK
- **Phase A click_via_timer returned:** 17023
- **Phase A elapsed:** 1.53 s
- **COM sleep:** 1.5 s
- **Phase B Application.Run returned OK:** False
- **Phase B Application.Run error (if any):** `attempt 1 'Form_LookAtStatus.RunExportGephi': com_error(-2147352567, 'Exception occurred.', (0, None, "Welcome to CBDB! cannot find the procedure 'Form_LookAtStatus.RunExportGephi.'", None, -1, -2146825771), None) | attempt 2 'RunExportGephi' (unqualified): com_error(-2147352567, 'Exception occurred.', (0, None, "Welcome to CBDB! cannot find the procedure 'RunExportGephi.'", None, -1, -2146825771), None)`
- **Phase B invoke elapsed:** 0.01 s
- **Phase B file_count:** 0
- **Phase B chain_observed_done:** True
- **msgbox_watchdog_count:** 0
- **per-phase wall:** 18.63 s

**Phase A scratch counts (post-CmdQuery):**
- `ZZ_SCRATCH_STATUS`: 17023
- `ZZ_SCRATCH_P_STATUS`: 17022
- `ZZ_TEST_DEBUG`: 2

**Phase A `ZZ_TEST_DEBUG`:**
- `LookAtStatus:ENTER`
- `LookAtStatus:DONE`

**Phase B scratch counts (post-direct-invoke):**
- `ZZ_SCRATCH_STATUS`: 17023
- `ZZ_SCRATCH_P_STATUS`: 17022
- `ZZ_TEST_DEBUG`: 0

**Phase B `ZZ_TEST_DEBUG` (cleared between phases):**
(empty)

**Watchdog MsgBox observations:**
(none observed)

## Q-A summary

**Q (key isolation question) — Did the export step actually trigger via direct invoke?**
- CmdPajek: **False** · CmdGephi: **False** · both: **False**

**Q1 — `Object required` :ERR remains 0 per button (Phase B)?**
- CmdPajek: **True** · CmdGephi: **True** · both: **True**

**Q2 — Both buttons start writing files?**
- CmdPajek: **False** (0 files) · CmdGephi: **False** (0 files) · both: **False**

**Q3 — Watchdog dialogs zero per button?**
- CmdPajek: **True** · CmdGephi: **True** · both: **True**

**Q4 — Phase B `ZZ_TEST_DEBUG` content per button:**
- CmdPajek msgs: `[]`
- CmdGephi msgs: `[]`

**Q5 — Scratch counts (post-Phase B) match PR #127 baseline?**
- ZZ_SCRATCH_STATUS · CmdPajek=17023 · CmdGephi=17023 · baseline=17023
- ZZ_SCRATCH_P_STATUS · CmdPajek=17022 · CmdGephi=17022 · baseline=17022
- **both_match_baseline:** **True**

## Verdict: `direct_invoke_sub_did_not_run`

**Direct-invoke via `Application.Run` is technically not feasible in this environment for form-module subs.**  Wrapper injection succeeded — Public Sub `RunExport<X>` was added to `Form_LookAtStatus` and the AddFromString call returned without error.  But `Application.Run` rejected BOTH name forms tried per phase:
  - `Form_LookAtStatus.RunExport<X>` (qualified):     "cannot find the procedure ..."
  - `RunExport<X>` (unqualified): same error.

**Implication: this is the well-known Access limitation that `Application.Run` only resolves Public subs in STANDARD modules, not class (form/report) modules.**  The wrapper exists but is not addressable via the Application.Run name-resolution path.  The brief's question — "would the export sub run cleanly post-CmdQuery if invoked outside the Form_Timer dispatch path?" — **cannot be answered via this mechanism**.

Phase A still verified the prior infrastructure is intact: scratch counts match PR #127 baseline on both phases (17023 / 17022); CmdQuery body was non-destructive.

**Per brief: do NOT silently fallback to a different mechanism.**  Documented honestly as a probe finding.  Next brief candidate space:
  - **wrapper in a STANDARD module** — inject the Public wrapper into a new Module (not the form's class module).  The wrapper would call into the form via `[Forms]![LookAtStatus].SetFocus` then `SendKeys` or `DoCmd`, OR by re-using `_inject_timer_trigger` from a standard-module context.  Adds complexity but keeps invocation off Form_Timer's re-dispatch path.
  - **pywinauto button click** — UI-driven; depends on form/button visibility + focus.  Brief explicitly listed this as a fallback option ("focused non-chain UI trigger path if that is the narrowest viable equivalent").
  - **driver-infra fix to `click_via_timer`** — investigate why PR #135's second-call dispatch didn't fire.  Possibly close+reopen form between dispatches.
  - **maintainer-line / canonical Issue** — file the upstream CBDB pattern (Dim'd-local Set rebind in CmdQuery cleanup) as a defect; leave Status × CmdPajek + CmdGephi skipped pending upstream fix.

## Constraints honoured per brief

- ✅ Investigation artifacts only — public driver (`tests/cbdb_driver/vba_session.py`) NOT modified.  Probe-script-internal Public wrapper is added to the session-local MDB copy only.
- ✅ Form.Tag chain dispatcher NOT used for Phase B; `Application.Run` via COM is the sole invocation mechanism.
- ✅ Both buttons covered (CmdPajek + CmdGephi).
- ✅ All 5 brief gates explicitly recorded.
- ✅ Failure modes documented; no silent fallback to a different invocation mechanism.
- ✅ No `tests/test_*` changed; no README, triage, canonical reports / issue severity touched.
- ✅ CmdNeo4j NOT touched (covered as PR #128).
- ✅ `--reclassify-from-json` supported.
- ✅ `analysis/report_screenshot_audit.md` drift left alone (standing rule).