# LookAtStatus × {CmdPajek, CmdGephi} COM-side split-dispatch probe

**Date:** 2026-05-08  ·  **Branch:** `investigate/status-com-split-dispatch` (off main `5f100b4`; rebased to current main)

Tests whether splitting the dispatch into two SEPARATE COM-driven clicks — with a Python-side `time.sleep(1.5)` between them — replicates PR #131's positive signal AND unblocks both Status export cells.  Structurally-different intervention from PR #132/#133/#134 (which all relied on a VBA-side DoEvents settle, ruled out by PR #134's mechanism boundary finding).  No driver edit; all logic in the probe script.

## Probe shape

Per phase (one per export button):

1. Open VbaSession + seed fixture.
2. **Phase A** — fire `CmdQuery` ALONE via `click_via_timer` with `Form.Tag = "CmdQuery"` (NOT `"CmdQuery,Cmd<X>"`).  Wait for `LookAtStatus:DONE`.  Snapshot scratch counts.
3. **Clear `ZZ_TEST_DEBUG`** so Phase B's `_wait_for_done` does not short-circuit on Phase A's stale `:DONE` marker.
4. **Python-side `time.sleep(1.5)`** — fully releases the COM thread.
5. **Phase B** — fire `Cmd<X>` ALONE via `click_via_timer` with `Form.Tag = "Cmd<X>"`.  Wait for fresh `LookAtStatus:DONE` (or quiescence on file count).
6. Capture Phase B `ZZ_TEST_DEBUG`, files, scratch counts.

**Sleep value:** 1.5 s (matches PR #131's positive signal; per brief, only one value tested).  **Total wall:** 166.56 s.

## Raw observed facts (per phase)

### Phase: `CmdPajek`

- **outcome category:** `split_dispatch_blocked_zero_files_no_err`
- **Phase A `click_via_timer` returned:** 17023
- **Phase A elapsed:** 1.53 s
- **COM-side sleep:** 1.5 s
- **Phase B `click_via_timer` returned:** 17023
- **Phase B elapsed:** 68.16 s
- **Phase B chain_observed_done:** True
- **Phase B file_count:** 0
- **msgbox_watchdog_count:** 0
- **Per-phase wall elapsed:** 79.91 s

**Phase A scratch row counts (post-CmdQuery):**
- `ZZ_SCRATCH_STATUS`: 17023
- `ZZ_SCRATCH_P_STATUS`: 17022
- `ZZ_TEST_DEBUG`: 2

**Phase A `ZZ_TEST_DEBUG`:**
- `LookAtStatus:ENTER`
- `LookAtStatus:DONE`

**Phase B scratch row counts (post-Cmd<X>):**
- `ZZ_SCRATCH_STATUS`: 17023
- `ZZ_SCRATCH_P_STATUS`: 17022
- `ZZ_TEST_DEBUG`: 0

**Phase B `ZZ_TEST_DEBUG` (cleared between phases, so only Phase B msgs):**
(empty)

**Watchdog MsgBox observations:**
(none observed)

### Phase: `CmdGephi`

- **outcome category:** `split_dispatch_blocked_zero_files_no_err`
- **Phase A `click_via_timer` returned:** 17023
- **Phase A elapsed:** 1.53 s
- **COM-side sleep:** 1.5 s
- **Phase B `click_via_timer` returned:** 17023
- **Phase B elapsed:** 68.16 s
- **Phase B chain_observed_done:** True
- **Phase B file_count:** 0
- **msgbox_watchdog_count:** 0
- **Per-phase wall elapsed:** 79.84 s

**Phase A scratch row counts (post-CmdQuery):**
- `ZZ_SCRATCH_STATUS`: 17023
- `ZZ_SCRATCH_P_STATUS`: 17022
- `ZZ_TEST_DEBUG`: 2

**Phase A `ZZ_TEST_DEBUG`:**
- `LookAtStatus:ENTER`
- `LookAtStatus:DONE`

**Phase B scratch row counts (post-Cmd<X>):**
- `ZZ_SCRATCH_STATUS`: 17023
- `ZZ_SCRATCH_P_STATUS`: 17022
- `ZZ_TEST_DEBUG`: 0

**Phase B `ZZ_TEST_DEBUG` (cleared between phases, so only Phase B msgs):**
(empty)

**Watchdog MsgBox observations:**
(none observed)

## Q1-Q5 answers

**Q1 — `Object required` :ERR remains 0 (Phase B) per button?**
- CmdPajek: **True** · CmdGephi: **True** · both: **True**

**Q2 — Both buttons start writing files?**
- CmdPajek: **False** (0 files) · CmdGephi: **False** (0 files) · both: **False**

**Q3 — Watchdog dialogs zero per button?**
- CmdPajek: **True** · CmdGephi: **True** · both: **True**

**Q4 — Phase B `ZZ_TEST_DEBUG` is `ENTER / MSGBOX? / DONE` with no `:ERR`?**
- CmdPajek well-formed: **False** · CmdGephi well-formed: **False** · both: **False**

**Q5 — Scratch counts (post-Phase B) match PR #127 baseline?**
- ZZ_SCRATCH_STATUS · CmdPajek=17023 · CmdGephi=17023 · baseline=17023
- ZZ_SCRATCH_P_STATUS · CmdPajek=17022 · CmdGephi=17022 · baseline=17022
- **both_match_baseline:** **True**

## Verdict: `split_dispatch_did_not_unblock`

**Split dispatch did NOT execute Phase B's `Cmd<X>_Click` at all.**  Phase B's `ZZ_TEST_DEBUG` is empty (we cleared it between phases; nothing was added).  No `:ENTER`, no `:MSGBOX`, no `:ERR`.  Phase B's `click_via_timer` returned 17023 (the result_table count, unchanged from Phase A) but timed out waiting for `:DONE`.  file_count = 0.

**Infrastructure finding (the value of this PR):** the existing `click_via_timer` + `_inject_timer_trigger` mechanism does NOT reliably re-dispatch to a DIFFERENT sub after a prior call in the same session.  When Phase A fired CmdQuery via Form_Timer and TimerInterval got reset to 0, the path to re-fire Form_Timer with a different target (CmdPajek) — even after `_inject_timer_trigger`'s delete-and-readd-Form_Timer-sub logic + OnTimer rebind — empirically does not invoke the new target.  Phase B's `Form_Timer` event either did not fire, or fired but the dispatched `Cmd<X>_Click` did not actually run.

**This is an INFRASTRUCTURE limitation, not a CBDB-runtime question.**  The brief's hypothesis (COM-side sleep between two click_via_timer calls unblocks the cells) could not be tested because the second click doesn't reliably execute its sub.  The intervention shape requires either (a) a different invocation mechanism on the COM side, OR (b) infrastructure work on the existing click_via_timer to support sequential different-target dispatches.

**Per the brief: do NOT silently switch candidate.**  Documented honestly.  Next brief candidate space:
  - **direct method invocation via COM** — try `app.Forms('LookAtStatus').Controls('CmdPajek').SetFocus` then a click via pywinauto, OR call the underlying VBA function directly via `Application.Run('Form_LookAtStatus.CmdPajek_Click')`.  Bypasses Form_Timer entirely.
  - **driver infra fix** — investigate why the second `click_via_timer` with a different ctl doesn't dispatch.  If this is a Form_Timer binding issue (likely), there may be a fix (e.g. close+reopen the form between dispatches, or use a different timer-arming sequence).
  - **maintainer-line / canonical Issue** — the underlying CBDB pattern (Dim'd-local Set rebind in CmdQuery cleanup) is fragile against any sequential test driver pattern; upstream fix would address it once for any future test infrastructure.

## Constraints honoured per brief

- ✅ Investigation artifacts only — NO driver edit made (all split-dispatch logic in the probe script).
- ✅ Form.Tag chain `CmdQuery,<button>` NOT used; two SEPARATE `set_form_tag` + `click_via_timer` calls with COM-side sleep between.
- ✅ One sleep value tested (1.5 s); not extended to a grid sweep.
- ✅ Both buttons covered (CmdPajek + CmdGephi).
- ✅ Did NOT silently switch candidate after this run.
- ✅ No `tests/test_*` changed; no README, triage, canonical reports / issue severity touched.
- ✅ CmdNeo4j NOT touched (covered as PR #128).
- ✅ `--reclassify-from-json` supported.
- ✅ `analysis/report_screenshot_audit.md` drift left alone (standing rule).