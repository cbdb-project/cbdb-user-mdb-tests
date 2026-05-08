# LookAtStatus × {CmdPajek, CmdGephi} verification probe — post Set→Requery cleanup-rebind patch

**Date:** 2026-05-08  ·  **Branch:** `driver/status-cleanup-rebind-requery` (off main `22667eb`)

Verifies the narrow scoped driver patch `_rewrite_status_cmdquery_cleanup_rebind_to_requery` (candidate (a) per the PR brief: replace each `Set <subform>.Form.Recordset = <local-var>` in `Form_LookAtStatus.CmdQuery_Click` cleanup section with `<subform>.Form.Requery`).  Compares 2 phases (CmdPajek, CmdGephi) against PR #127's pre-patch baseline on the same fixture.  CmdNeo4j is NOT re-probed — it was already covered as PR #128 (false-positive skip; bypasses the subform recordset).

## Setup

- **Form:** `LookAtStatus`
- **Fixture:** matrix `_make_status_fixtures` first fixture (`status_<top_code>_unfiltered`), same as PR #127's baseline.
- **Patch under test:** candidate (a) — replace `Set <subform>.Form.Recordset = <var>` with `<subform>.Form.Requery` at `Form_LookAtStatus.vb:1457+1460` (CmdQuery_Click Exit_Run_Query section).
- **Phases:** 2 sequential, one per export button (`CmdPajek, CmdGephi`).  Each = own MDB copy + own VbaSession + own out_dir.
- **Per-phase chain:** `CmdQuery,<button>` via Form.Tag, directory mode
- **CmdQuery cleanup-intent gate:** scratch counts must match PR #127 baseline on BOTH phases (ZZ_SCRATCH_STATUS = 17023, ZZ_SCRATCH_P_STATUS = 17022).
- **click_via_timer cap:** 180 s · **per-phase outer cap:** 240 s
- **Total wall elapsed:** 43.23 s

## Why candidate (a) over (b) / (c)

PR #127's diagnosis pinned the failure to a local-var lifetime issue: `tRstStatus` (line 1160) is a Dim'd local; after Exit Sub it dies, and `<subform>.Form.Recordset` reads as Nothing in the next button → VBA 424 'Object required'.  `gRstPeople` (global) sometimes survives but the Pajek/Gephi reads fire on `ZZ_SCRATCH_STATUS` first, so the local-var death is the deciding factor.

- **(a) Set→Requery** — Subform's Recordset becomes owned by the form (re-derived from its design-time RecordSource).  No local-var lifetime issue.  Most idiomatic Access pattern for this case.  *Chosen.*
- **(b) Keep Set + add MoveFirst** — Does NOT fix the lifetime issue.  After Exit Sub the local var dies regardless.  Would not address the symptom.
- **(c) Drop rebind** — Subform stays bound to `tRstDummy` (line 1176 → Z_SCRATCH_DUMMY_SC).  Subsequent reads see dummy data, not the freshly-populated ZZ_SCRATCH_STATUS data.  Wrong content even if the chain runs.

## Raw observed facts (per phase)

### Phase: `CmdPajek`

- **outcome category:** `patch_blocked_zero_files_no_err`
- **chain_elapsed_sec:** 9.54
- **file_count:** 0
- **click_via_timer_returned:** 17023
- **chain_observed_done:** True
- **msgbox_watchdog_count:** 0
- **per-phase wall elapsed:** 18.38 s

**Scratch row counts (post-chain):**
- `ZZ_SCRATCH_STATUS`: 17023
- `ZZ_SCRATCH_P_STATUS`: 17022
- `ZZ_TEST_DEBUG`: 3

**ZZ_TEST_DEBUG content:**
- `LookAtStatus:ENTER`
- `LookAtStatus:MSGBOX`
- `LookAtStatus:DONE`

**Watchdog MsgBox observations:**
(none observed)


### Phase: `CmdGephi`

- **outcome category:** `patch_blocked_zero_files_no_err`
- **chain_elapsed_sec:** 9.53
- **file_count:** 0
- **click_via_timer_returned:** 17023
- **chain_observed_done:** True
- **msgbox_watchdog_count:** 0
- **per-phase wall elapsed:** 18.05 s

**Scratch row counts (post-chain):**
- `ZZ_SCRATCH_STATUS`: 17023
- `ZZ_SCRATCH_P_STATUS`: 17022
- `ZZ_TEST_DEBUG`: 3

**ZZ_TEST_DEBUG content:**
- `LookAtStatus:ENTER`
- `LookAtStatus:MSGBOX`
- `LookAtStatus:DONE`

**Watchdog MsgBox observations:**
(none observed)


## Q1-Q5 answers

**Q1 — `Object required` disappeared per button?**
- CmdPajek: **True**  ·  CmdGephi: **True**  ·  both: **True**

**Q2 — Both buttons start writing files?**
- CmdPajek: **False** (0 files)  ·  CmdGephi: **False** (0 files)  ·  both started writing: **False**

**Q3 — Watchdog dialogs zero per button?**
- CmdPajek: **True**  ·  CmdGephi: **True**  ·  both: **True**

**Q4 — `ZZ_TEST_DEBUG` is `ENTER / MSGBOX? / DONE` with no `:ERR`?**
- CmdPajek well-formed: **True**  ·  CmdGephi well-formed: **True**  ·  both: **True**
- CmdPajek msgs: ['LookAtStatus:ENTER', 'LookAtStatus:MSGBOX', 'LookAtStatus:DONE']
- CmdGephi msgs: ['LookAtStatus:ENTER', 'LookAtStatus:MSGBOX', 'LookAtStatus:DONE']

**Q5 — CmdQuery cleanup-intent preserved? (scratch counts match PR #127 baseline)**
- ZZ_SCRATCH_STATUS  · CmdPajek=17023 · CmdGephi=17023 · PR127 baseline=17023
- ZZ_SCRATCH_P_STATUS · CmdPajek=17022 · CmdGephi=17022 · PR127 baseline=17022
- **both_match_baseline:** **True**

If both_match_baseline is True, the Set→Requery patch did not regress CmdQuery body's INSERT outcome — same scratch population as the PR #127 unpatched baseline.  This is the primary 'no broken cleanup intent' check.

## Verdict: `patch_did_not_unblock`

**Useful negative evidence — candidate (a) tested, NOT being landed as a driver workaround.**  The Set→Requery helper was added to `tests/cbdb_driver/vba_session.py` ONCE on this investigation branch to capture the evidence below, then REVERTED before this PR was opened for review.  The probe artifacts (this script + paired MD + paired JSON) are what ships; the driver code is not.

**What candidate (a) achieved (4 of 5 gates):**
  - `Object required` :ERR is GONE on both phases (the local-var lifetime issue PR #127 pinned IS resolved by Set→Requery).
  - Watchdog dialogs = 0 on both phases (the driver's literal-only neutralizer caught the bail `MsgBox`).
  - `ZZ_TEST_DEBUG` = [ENTER, MSGBOX, DONE] with no `:ERR` rows on both phases (the `:MSGBOX` is the neutralized `MsgBox "There are no records to save."` from the bail path).
  - `ZZ_SCRATCH_STATUS / ZZ_SCRATCH_P_STATUS` row counts MATCH PR #127's pre-patch baseline (17023 / 17022) on both phases — CmdQuery cleanup intent PRESERVED.

**What candidate (a) did NOT achieve (1 of 5 gates):**
  - `file_count = 0` on both phases — the buttons bailed at their `If <subform>.Form.Recordset.RecordCount = 0 Then` check.  The failure shape *shifted* (Object required → legitimately-zero RecordCount) but did NOT disappear.

**Why this means candidate (a) does NOT belong in `vba_session.py`.**  Per repo convention, driver-side workarounds are landed only when they unblock the cell they target — partial fixes that resolve a symptom without enabling file production accumulate as dead weight.  Candidate (a) eliminates PR #127's Object-required symptom AND preserves CmdQuery cleanup intent, but neither CmdPajek nor CmdGephi produces files post-patch.  The patch is insufficient.

**Interpretation of the file-write blocker.**  Set→Requery successfully removed the Nothing-recordset symptom, but Requery re-executes the subform's *design-time RecordSource* — which apparently does NOT see the populated ZZ_SCRATCH_STATUS / ZZ_SCRATCH_P_STATUS data after CmdQuery body finished its INSERTs.  Most likely the design-time RecordSource is an empty saved query, OR a stale binding to the original (now-replaced) recordset.  Confirming this would need a static read of the form's design-time RecordSource setting (out of scope for this PR).

**Per the brief: NOT silently switching to candidate (b) or (c).**  Documenting the shifted-symptom shape.  The next brief decides whether to:
  - try candidate (b) (Set + MoveFirst): structurally less promising now that (a) confirmed the local-var lifetime issue is real and is resolvable — (b) preserves the same Set + dynaset shape that (a) replaced.
  - try candidate (c) (drop rebind): leaves the subform bound to `tRstDummy` (Z_SCRATCH_DUMMY_SC, not ZZ_SCRATCH_STATUS); fails the cleanup-intent gate by definition.
  - try a candidate (d) NEW (surfaced by this probe): combine Requery with an EXPLICIT `Set <subform>.Form.RecordSource = "ZZ_SCRATCH_STATUS"` BEFORE the Requery, forcing the design-time RecordSource to point at the populated table.  Closer to the actual fix surface.
  - escalate to canonical issue filing: the subform design-time RecordSource not seeing the populated scratch data may be a CBDB source-level binding bug; warrants its own static investigation.

**Bottom line.**  Candidate (a) is verifiably INSUFFICIENT: it removes one symptom layer (local-var lifetime) but exposes a second (design-time RecordSource binding).  The next attempt needs to address the second layer, not the first.  This evidence narrows the next-brief candidate space — that's the value this PR delivers.

## Constraints honoured per brief

- ✅ Investigation artifacts only — `tests/cbdb_driver/vba_session.py` was modified ONCE on this branch to capture the evidence below, then REVERTED before this PR was opened for review.  Final diff vs `main` contains only this probe script + paired MD + paired JSON.
- ✅ No `tests/test_*` changed; no README, triage, or canonical reports / issue severity touched.
- ✅ No coverage PR opened; no canonical issue filed; **no driver workaround landed**.
- ✅ Only candidate (a) implemented and verified — did NOT silently switch to (b) or (c).  Failure shape (symptom shifted, file-write still blocked) documented honestly.
- ✅ CmdNeo4j NOT re-probed (already covered as PR #128).
- ✅ `--reclassify-from-json` supported for re-running verdict logic without COM (and used by this respin to regenerate MD/JSON without re-running COM after the driver revert).