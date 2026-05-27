# LookAtStatus × {CmdPajek, CmdGephi} verification probe — post chain-dispatcher 250ms settle tweak

**Date:** 2026-05-08  ·  **Branch:** `driver/chain-dispatcher-settle` (off main `5f100b4`)

Verifies the narrow scoped driver edit in `tests/cbdb_driver/vba_session.py::_inject_autodetect`'s chain-dispatch loop: a 250 ms `DoEvents` loop inserted AT THE TOP of each chained-step iteration in the `For chnI = 1 To UBound(chnParts)` loop.  Compares 2 phases (CmdPajek, CmdGephi) against PR #127's pre-patch baseline on the same fixture.  CmdNeo4j is NOT re-probed — already covered as PR #128.

## Driver edit under test (one VBA template change)

Inside `_inject_autodetect`'s `done_insert` block — at the TOP of the For loop body, BEFORE the `Select Case` switch, INSIDE the loop:

```vba
Dim chnSettleT As Double  ' added to existing Dim block
...
For chnI = 1 To UBound(chnParts)
    ' 250 ms DoEvents loop drains the Access UI
    ' message queue so subform Set/Requery side-
    ' effects from the previous step commit before
    ' the next Cmd<X>_Click reads the recordset.
    chnSettleT = Timer
    Do While (Timer - chnSettleT) < 0.25
        DoEvents
    Loop
    Select Case Trim(chnParts(chnI))
        Case ...: Call Cmd<X>_Click
    End Select
Next chnI
```

**Scope of edit:**
- only chain dispatcher / Form.Tag multi-step path
- settle is at step boundary (before each Cmd<X>_Click call), not global
- first step (entry sub itself) runs without an introductory settle (no preceding state to settle)
- 250 ms is deliberately a smallest-reasonable lower bound vs PR #131's 1.5 s empirical baseline; anything shorter would be unverified
- `Timer`-based loop keeps DoEvents firing throughout the window, draining the message queue continuously

## Setup

- **Form:** `LookAtStatus`
- **Fixture:** matrix `_make_status_fixtures` first fixture (`status_<top_code>_unfiltered`), same as PR #127 / PR #129 baseline.
- **Phases:** 2 sequential, one per export button (`CmdPajek, CmdGephi`).  Each = own MDB copy + own VbaSession + own out_dir.
- **Per-phase chain:** `CmdQuery,<button>` via Form.Tag, directory mode.
- **CmdQuery cleanup-intent gate:** scratch counts must match PR #127 baseline on BOTH phases (ZZ_SCRATCH_STATUS = 17023, ZZ_SCRATCH_P_STATUS = 17022).
- **click_via_timer cap:** 180 s · **per-phase outer cap:** 240 s
- **Total wall elapsed:** 43.85 s

## Raw observed facts (per phase)

### Phase: `CmdPajek`

- **outcome category:** `settle_partial_object_required_still_observed`
- **chain_elapsed_sec:** 9.53
- **file_count:** 0
- **click_via_timer_returned:** 17023
- **chain_observed_done:** True
- **msgbox_watchdog_count:** 0
- **per-phase wall elapsed:** 18.37 s

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

- **outcome category:** `settle_partial_object_required_still_observed`
- **chain_elapsed_sec:** 10.04
- **file_count:** 0
- **click_via_timer_returned:** 17023
- **chain_observed_done:** True
- **msgbox_watchdog_count:** 0
- **per-phase wall elapsed:** 18.66 s

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
- CmdPajek: **False**  ·  CmdGephi: **False**  ·  both: **False**

**Q2 — Both buttons start writing files?**
- CmdPajek: **False** (0 files)  ·  CmdGephi: **False** (0 files)  ·  both started writing: **False**

**Q3 — Watchdog dialogs zero per button?**
- CmdPajek: **True**  ·  CmdGephi: **True**  ·  both: **True**

**Q4 — `ZZ_TEST_DEBUG` is `ENTER / MSGBOX? / DONE` with no `:ERR`?**
- CmdPajek well-formed: **False**  ·  CmdGephi well-formed: **False**  ·  both: **False**
- CmdPajek msgs: ['LookAtStatus:ENTER', 'LookAtStatus:ERR Object required', 'LookAtStatus:DONE']
- CmdGephi msgs: ['LookAtStatus:ENTER', 'LookAtStatus:ERR Object required', 'LookAtStatus:DONE']

**Q5 — Scratch counts match PR #127 baseline?**
- ZZ_SCRATCH_STATUS  · CmdPajek=17023 · CmdGephi=17023 · baseline=17023
- ZZ_SCRATCH_P_STATUS · CmdPajek=17022 · CmdGephi=17022 · baseline=17022
- **both_match_baseline:** **True**

## Verdict: `settle_did_not_unblock`

**Useful negative evidence — 250 ms settle tested, NOT being landed.**  The DoEvents loop was added to `_inject_autodetect`'s dispatch block ONCE on this branch to capture the evidence below, then REVERTED before this PR was opened for review.  Probe artifacts ship; driver code does not.

**What 250 ms settle achieved (1 of 5 gates):**
  - Q5 cleanup-intent gate PASSED — scratch counts match PR #127 baseline (17023 / 17022) on both phases.  The DoEvents loop did NOT regress CmdQuery body's INSERT outcome; the edit was at least non-destructive.

**What 250 ms settle did NOT achieve (4 of 5 gates):**
  - Q1 `Object required` :ERR REAPPEARED on both phases (same shape as PR #127's pre-patch baseline; identical to the failure mode the settle was supposed to fix).
  - Q2 file_count = 0 on both phases.
  - Q3 watchdog dialogs = 0 (incidental — the literal-only neutralizer caught the bail MsgBox; not a Q3 win in any meaningful sense).
  - Q4 ZZ_TEST_DEBUG = [ENTER, :ERR Object required, DONE] — has a `:ERR` row.

**Why 250 ms is below threshold but 1.5 s worked.**  PR #131's micro-check used 1.5 s between Form.Requery and the RecordCount read; that succeeded.  This PR's 250 ms in the chain dispatcher was insufficient.  The actual threshold for Access UI message-pump completion of the subform Set/Requery side-effects on this fixture is **somewhere between 250 ms and 1.5 s** — narrowing it requires more probes.

**Per the brief: NOT silently extending the window in this PR.**  Documenting the below-threshold failure honestly.  The next brief decides whether to:
  - try a longer settle (e.g. 500 ms, 750 ms, 1000 ms) — bisect the threshold to find the smallest that works
  - try a different placement (e.g. settle AFTER the chain dispatch's requery_lines and BEFORE the For loop, instead of inside the For-loop body)
  - reconsider the intervention shape entirely (e.g. PR #131's discovery that explicit .Form.Requery + DoEvents is sufficient suggests an alternative: have the dispatcher itself call `<subform>.Form.Requery` for the Status subforms before each chained step, rather than relying on CmdQuery's cleanup-rebind to commit via DoEvents)

**Bottom line.**  250 ms is empirically insufficient; the threshold is between 250 ms and 1.5 s.  This evidence narrows the next-brief candidate space — that's the value this PR delivers.

## Constraints honoured per brief

- ✅ Investigation artifacts only — `tests/cbdb_driver/vba_session.py` was modified ONCE on this branch to capture the evidence below, then REVERTED before this PR was opened for review.  Final diff vs `main` contains only this probe script + paired MD + paired JSON.
- ✅ Driver edit (when active during the COM run) was narrow-scoped: only chain dispatcher / Form.Tag multi-step path; settle at step boundary only; not a global sleep policy.
- ✅ Settle was minimal — 250 ms (smallest reasonable lower bound vs PR #131's 1.5 s empirical baseline).
- ✅ Verification probe covers BOTH buttons (CmdPajek + CmdGephi).
- ✅ No `tests/test_*` changed; no README, triage, canonical reports / issue severity touched; **no driver workaround landed**.
- ✅ CmdNeo4j NOT re-probed (already covered as PR #128).
- ✅ No coverage PR opened.
- ✅ Failure shape documented honestly; did NOT silently extend the settle window or switch to a different intervention.
- ✅ `--reclassify-from-json` supported (and used by this respin to regenerate MD/JSON without re-running COM after the driver revert).
- ✅ `analysis/report_screenshot_audit.md` drift left alone (standing rule).