# LookAtStatus × {CmdPajek, CmdGephi} bounded exploratory sweep

**Date:** 2026-05-09  ·  **Branch:** `investigate/status-bounded-sweep` (off main `6b06d6a`)

Per maintainer-authorized one-shot exploratory sweep before any UI fallback or maintainer-line.  Bounded sweep across distinct mechanism families; stops at first candidate that crosses the viability threshold (both buttons fire AND Object required disappears AND file_count >= 1 for both).  If none cross, recommends pywinauto fallback or maintainer-line.

## Candidate families tested in this sweep

| ID | Mechanism | External evidence | Implementation cost |
|---|---|---|---|
| **A** | Chain-dispatch with `RecordSource = RecordSource` self-rebinding (PR #140 F4) | Microsoft Learn `Form.Recordset` doc explicitly recommends this pattern; never tested in PR #129-#137 | low (1-line VBA inject before Select Case dispatch) |
| **E** | Standard-module `Form_Timer` dispatch via `OnTimer = "=Func()"` form | Distinct binding mechanism; bypasses form-class-instance event-binding cache (PR #137 pinned layer) | medium (new standard module + Public wrapper in form module) |

## Candidates DELIBERATELY EXCLUDED

| ID | Reason for exclusion |
|---|---|
| **B** (RecordSource reset to literal + Requery) | small variant of Family A; if A fails because Access doesn't honor the rebind, B has no independent reason to succeed |
| **C** (close + reopen form between phases) | `Form_LookAtStatus.Form_Open()` at lines 2090/2103 explicitly DELETEs ZZ_SCRATCH_STATUS and ZZ_SCRATCH_P_STATUS on every open; re-opening between Phase A and Phase B would wipe the data we just populated; testing close+reopen would require invasive workarounds that confound the mechanism |
| **D** (global/module-level recordset ownership) | PR #137 pinned the failure layer at form-class-instance event-binding cache, NOT recordset ownership; insufficient mechanism distinction to justify another test |

**Total wall elapsed:** 447.04 s  ·  **families executed:** 2  ·  **stopped early on success:** False

## Per-family results

### Family A

- **summary:** `family=A CmdPajek=sub_fired_object_required CmdGephi=sub_fired_object_required crossed=False scratch_ok=True`
- **crosses viability threshold:** **False**
- **scratch baseline preserved:** True
- **CmdPajek:** `sub_fired_object_required`
    - file_count: 0
    - scratch: ZZ_SCRATCH_STATUS=17023, ZZ_SCRATCH_P_STATUS=17022
    - ZZ_TEST_DEBUG: ['LookAtStatus:ENTER', 'LookAtStatus:ERR Object required', 'LookAtStatus:DONE']
    - watchdog dialogs: 0
- **CmdGephi:** `sub_fired_object_required`
    - file_count: 0
    - scratch: ZZ_SCRATCH_STATUS=17023, ZZ_SCRATCH_P_STATUS=17022
    - ZZ_TEST_DEBUG: ['LookAtStatus:ENTER', 'LookAtStatus:ERR Object required', 'LookAtStatus:DONE']
    - watchdog dialogs: 0

### Family E

- **summary:** `family=E CmdPajek=sub_did_not_fire CmdGephi=sub_did_not_fire crossed=False scratch_ok=True`
- **crosses viability threshold:** **False**
- **scratch baseline preserved:** True
- **CmdPajek:** `sub_did_not_fire`
    - file_count: 0
    - scratch: ZZ_SCRATCH_STATUS=17023, ZZ_SCRATCH_P_STATUS=17022
    - ZZ_TEST_DEBUG: []
    - watchdog dialogs: 0
- **CmdGephi:** `sub_did_not_fire`
    - file_count: 0
    - scratch: ZZ_SCRATCH_STATUS=17023, ZZ_SCRATCH_P_STATUS=17022
    - ZZ_TEST_DEBUG: []
    - watchdog dialogs: 0

## Key new evidence from this sweep

Family E's negative result is the most substantive finding. Family E injected a **Public Function `DispatchToCmdPajek()` in a standard module** and set `OnTimer = "=DispatchToCmdPajek()"` (Access expression-service binding form), then armed `TimerInterval = 100`. The dispatch function never executed — `ZZ_TEST_DEBUG` contains no `FAMILY_E_DISPATCH_FIRED` marker; `dispatch_fired_seen=False` after a 90-second wait.

This **extends PR #137's pinned failure layer**. PR #137 pinned the second-Form_Timer-call failure at the form-class-instance event-binding cache for `OnTimer = "[Event Procedure]"` form, even after `AddFromString` + force-compile of a fresh `Form_Timer` sub. Family E shows the cache resists the **expression-service binding form too** — `OnTimer = "=StandardModuleFunc()"` is a structurally distinct dispatch path (no class-module event lookup; goes through Access's expression service to a standard module) yet it fires no more reliably than `[Event Procedure]` after the form's first OnTimer use.

Practical consequence: any in-process workaround that relies on the form's `OnTimer` after `CmdQuery_Click` has already used the timer once is structurally blocked, regardless of which `OnTimer` binding form is used. The cache is at the **TimerInterval-arming-event level**, not the binding-resolution level.

Family A's negative result is consistent with PR #129/#132/#133/#134's finding that no in-CmdQuery-stack VBA-side intervention restores the subform Recordsets in time for the chained Cmd<X>_Click. F4's `RecordSource = RecordSource` is structurally distinct from Recordset-level interventions (it triggers Access's RecordSource setter), but in this fixture the subforms remain unbound for the chained dispatch — the cleanup-rebind defect (`Form_LookAtStatus.vb:1457-1460`) leaves them in a state F4 alone cannot recover. The maintainer-line (apply the cleanup-rebind fix upstream) remains the correct and only known full unblock.

## Verdict

- **bucket:** `no_viable_local_candidate`
- **viable_family:** `None`
- **recommendation (concrete):** stay with the **pure maintainer-line** posture for Status × CmdPajek / CmdGephi. Reasons:
    - The maintainer-line is already filed end-to-end (Issue #X — Status cleanup-rebind defect; PR #139 handoff memo; PR #140 external-evidence corroboration).
    - The maintainer-line is one localized VBA fix in production (`Form_LookAtStatus.vb:1457` and `:1460`, swap the Dim'd-local pattern for the `gRstPeople`-style global ownership shape already proven on line 1184).
    - A pywinauto UI fallback would add a ~50-second per-session wall cost AND a new dependency on real-time Windows-message timing, AND would still only verify what the maintainer-line shape is supposed to fix — net negative cost/benefit.
    - This sweep added one new piece of evidence (Family E above) that **strengthens** the maintainer-line case rather than weakens it: the timer cache is structural; no in-process workaround on this form-instance is going to restore Form_Timer dispatch after the first use, period.
- **driver state:** `tests/cbdb_driver/vba_session.py` is byte-identical to `main` (verified via `git diff main`). No revert needed; no probe-side driver edits were ever staged.

## Constraints honoured per brief

- ✅ Bounded sweep across distinct mechanism families (2 candidate families A + E; 3 explicitly excluded with rationale: B/C/D)
- ✅ Non-UI paths first; pywinauto UI fallback NOT attempted in this sweep (deferred per brief)
- ✅ Stop-at-first-success early-termination logic
- ✅ No tests / README / triage / canonical reports / issue severity changes
- ✅ Driver edits (Family A's autodetect chain block extension) applied per-session via probe-side textual patching of CodeModule, NOT via public driver edit; if no candidate crosses threshold, public driver remains byte-for-byte identical to main
- ✅ Family E's standard-module + form Public wrapper are per-session VBA injection only (working MDB copy regenerated per phase)
- ✅ `analysis/report_screenshot_audit.md` drift left alone (standing rule)