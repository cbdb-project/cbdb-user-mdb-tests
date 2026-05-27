# LookAtStatus × {CmdPajek, CmdGephi} verification probe — explicit-Requery variant

**Date:** 2026-05-08  ·  **Branch:** `investigate/status-explicit-requery-variant` (off main `5f100b4`; rebased to current main for clean PR diff)

Verifies the EXPLICIT-REQUERY variant in `tests/cbdb_driver/vba_session.py::_inject_autodetect`'s chain-dispatch loop: a per-form per-step Requery shim that calls `<subform>.Form.Requery` for `ZZ_SCRATCH_STATUS` and `ZZ_SCRATCH_P_STATUS` INSIDE the For loop body (BEFORE the Select Case Call), followed by DoEvents.  Minimal settle: DoEvents only — no sleep.  Targets the mechanism PR #131 confirmed works (Form.Requery + DoEvents recovers subform recordset).

## Driver edit under test (REVERTED if did_not_unblock)

```vba
' New per-form per-step requery shim — Status only,
' via _PER_FORM_PER_STEP_REQUERY_SUBFORMS dict.
For chnI = 1 To UBound(chnParts)
    ' Per-step explicit Requery (PR #133 variant):
    On Error Resume Next
    ZZ_SCRATCH_STATUS.Form.Requery
    ZZ_SCRATCH_P_STATUS.Form.Requery
    DoEvents
    On Error GoTo 0
    Select Case Trim(chnParts(chnI))
        Case ...: Call Cmd<X>_Click
    End Select
Next chnI
```

**Scope:** only `Form_LookAtStatus` (per the `_PER_FORM_PER_STEP_REQUERY_SUBFORMS` dict — only one entry).  Other forms get the dispatcher unchanged.

## Setup

- **Form:** `LookAtStatus`
- **Fixture:** matrix `_make_status_fixtures` first fixture (`status_<top_code>_unfiltered`).
- **Phases:** 2 sequential, one per export button (`CmdPajek, CmdGephi`).
- **Per-phase chain:** `CmdQuery,<button>` via Form.Tag.
- **CmdQuery cleanup-intent gate:** scratch counts must match PR #127 baseline on BOTH phases (17023 / 17022).
- **Total wall elapsed:** 43.32 s

## Raw observed facts (per phase)

### Phase: `CmdPajek`

- **outcome category:** `variant_partial_object_required_still_observed`
- **chain_elapsed_sec:** 9.54
- **file_count:** 0
- **click_via_timer_returned:** 17023
- **chain_observed_done:** True
- **msgbox_watchdog_count:** 0
- **per-phase wall elapsed:** 18.42 s

**Scratch row counts (post-chain):**
- `ZZ_SCRATCH_STATUS`: 17023
- `ZZ_SCRATCH_P_STATUS`: 17022
- `ZZ_TEST_DEBUG`: 3

**ZZ_TEST_DEBUG content:**
- `LookAtStatus:ENTER`
- `LookAtStatus:ERR Object required`
- `LookAtStatus:DONE`

**Watchdog MsgBox observations:**
(none observed)

### Phase: `CmdGephi`

- **outcome category:** `variant_partial_object_required_still_observed`
- **chain_elapsed_sec:** 9.54
- **file_count:** 0
- **click_via_timer_returned:** 17023
- **chain_observed_done:** True
- **msgbox_watchdog_count:** 0
- **per-phase wall elapsed:** 18.04 s

**Scratch row counts (post-chain):**
- `ZZ_SCRATCH_STATUS`: 17023
- `ZZ_SCRATCH_P_STATUS`: 17022
- `ZZ_TEST_DEBUG`: 3

**ZZ_TEST_DEBUG content:**
- `LookAtStatus:ENTER`
- `LookAtStatus:ERR Object required`
- `LookAtStatus:DONE`

**Watchdog MsgBox observations:**
(none observed)

## Q1-Q5 answers

**Q1 — `Object required` :ERR remains 0 per button?**
- CmdPajek: **False** · CmdGephi: **False** · both: **False**

**Q2 — Both buttons start writing files?**
- CmdPajek: **False** (0 files) · CmdGephi: **False** (0 files) · both: **False**

**Q3 — Watchdog dialogs zero per button?**
- CmdPajek: **True** · CmdGephi: **True** · both: **True**

**Q4 — `ZZ_TEST_DEBUG` is `ENTER / MSGBOX? / DONE` with no `:ERR`?**
- CmdPajek: **False** · CmdGephi: **False** · both: **False**
- CmdPajek msgs: ['LookAtStatus:ENTER', 'LookAtStatus:ERR Object required', 'LookAtStatus:DONE']
- CmdGephi msgs: ['LookAtStatus:ENTER', 'LookAtStatus:ERR Object required', 'LookAtStatus:DONE']

**Q5 — Scratch counts match PR #127 baseline?**
- ZZ_SCRATCH_STATUS · CmdPajek=17023 · CmdGephi=17023 · baseline=17023
- ZZ_SCRATCH_P_STATUS · CmdPajek=17022 · CmdGephi=17022 · baseline=17022
- **both_match_baseline:** **True**

## Verdict: `variant_did_not_unblock`

**Useful negative evidence — explicit-Requery variant tested, NOT being landed.**  Driver edit was added to `tests/cbdb_driver/vba_session.py` ONCE on this branch to capture the evidence below, then REVERTED before this PR was opened for review.  Probe artifacts ship; driver code does not.

**What the variant achieved (1 of 5 gates):**
  - Q5 cleanup-intent gate PASSED — scratch counts match PR #127 baseline.  Edit was non-destructive.

**What the variant did NOT achieve (4 of 5):**
  - Q1 `Object required` :ERR REAPPEARED on both phases — same as PR #127 / PR #132.
  - Q2 file_count = 0 on both phases.
  - Q3 watchdog = 0 (incidental).
  - Q4 ZZ_TEST_DEBUG has `:ERR` row.

**Mechanism implication.**  Even calling `<subform>.Form.Requery` directly inside the dispatcher (with DoEvents, no sleep) is INSUFFICIENT — the explicit Requery does not synchronously establish a usable Recordset before the next `Cmd<X>_Click` reads it.  PR #131's success at returning the expected RecordCount required `time.sleep(1.5)` after the Requery.  The Access UI thread takes non-trivial time to commit Form.Requery's side-effects regardless of whether the Requery is called from VBA inline or from external COM.

**Per the brief: NOT silently switching to a second variant or longer settle.**  Failed experiment; next brief decides.  Candidate space:
  - longer-settle bisect — explicit Requery + small explicit sleep (e.g. 100ms / 250ms / 500ms) inside the dispatcher loop, find the minimum sufficient delay
  - explicit Requery + different placement — Requery BEFORE the For loop (after CmdQuery returns, before any chained step), with a longer settle then; only one settle window regardless of chain length
  - deeper Access/dispatcher investigation — is there a way to *synchronously* wait for Form.Requery to commit (e.g. `MoveLast` after Requery, or `Application.RefreshDatabaseWindow`)?

## Constraints honoured per brief

- ✅ Investigation artifacts only — `tests/cbdb_driver/vba_session.py` was modified ONCE on this branch to capture the evidence below, then REVERTED before this PR was opened for review (Q1+Q2+Q4 failed; per brief's exception clause, driver edit is retained ONLY if both cells unblocked, which they did not).  Final diff = 3 probe artifacts only.
- ✅ Variant was more mechanism-aligned than PR #132 — explicit `Form.Requery` (PR #131's positive signal), not generic settle.  Result: insufficient on its own; needs settle TIME after the Requery.
- ✅ Settle was minimal — DoEvents only, no sleep.
- ✅ Verification probe covered BOTH buttons (CmdPajek + CmdGephi).
- ✅ No `tests/test_*` changed; no README, triage, canonical reports / issue severity touched; **no driver workaround landed**.
- ✅ CmdNeo4j NOT touched (already covered as PR #128).
- ✅ Did NOT silently try a second variant after this one failed.
- ✅ `--reclassify-from-json` supported (and used by this respin to regenerate MD/JSON without re-running COM after the driver revert).
- ✅ `analysis/report_screenshot_audit.md` drift left alone (standing rule).