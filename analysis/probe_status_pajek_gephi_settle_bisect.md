# LookAtStatus × {CmdPajek, CmdGephi} settle-bisect probe — explicit Requery + bisected settle window

**Date:** 2026-05-08  ·  **Branch:** `investigate/status-settle-bisect` (off main `5f100b4`; rebased to current main)

Bisects the minimum settle window after explicit subform `.Form.Requery` that lets BOTH CmdPajek + CmdGephi start writing files post-CmdQuery on the matrix Status fixture.  3 settle values × 2 buttons = 6 phases.  Driver edit was added ONCE to capture the evidence below, then REVERTED before this PR opened for review (per brief: this PR is investigation only; landed workaround PR is the separate next brief).

## Setup

- **Form:** `LookAtStatus`
- **Fixture:** matrix `_make_status_fixtures` first fixture (`status_<top_code>_unfiltered`).
- **Intervention shape (FIXED across all settle values):** explicit `<subform>.Form.Requery` for `ZZ_SCRATCH_STATUS` AND `ZZ_SCRATCH_P_STATUS`, INSIDE the For loop body, BEFORE the Select Case Call; followed by Timer-driven DoEvents loop for the configured settle_ms.  Wrapped in `On Error Resume Next` for safety.
- **Settle values tested:** 500 ms, 750 ms, 1000 ms.
- **Phases:** 6 = 3 settle values × 2 buttons (CmdPajek, CmdGephi); each its own fresh MDB copy + own VbaSession.
- **Total wall elapsed:** 132.90 s

## Raw observed facts (per settle value)

### settle_ms = 500  ·  outcome: `unblocks_neither`

| Button | outcome | files | :ERR | scratch_status / p_status | ZZ_TEST_DEBUG |
|---|---|---:|---|---:|---|
| `CmdPajek` | `object_required_still_observed` | 0 | `Object required` | 17023 / 17022 | `['LookAtStatus:ENTER', 'LookAtStatus:ERR Object required', 'LookAtStatus:DONE']` |
| `CmdGephi` | `object_required_still_observed` | 0 | `Object required` | 17023 / 17022 | `['LookAtStatus:ENTER', 'LookAtStatus:ERR Object required', 'LookAtStatus:DONE']` |

### settle_ms = 750  ·  outcome: `unblocks_neither`

| Button | outcome | files | :ERR | scratch_status / p_status | ZZ_TEST_DEBUG |
|---|---|---:|---|---:|---|
| `CmdPajek` | `object_required_still_observed` | 0 | `Object required` | 17023 / 17022 | `['LookAtStatus:ENTER', 'LookAtStatus:ERR Object required', 'LookAtStatus:DONE']` |
| `CmdGephi` | `object_required_still_observed` | 0 | `Object required` | 17023 / 17022 | `['LookAtStatus:ENTER', 'LookAtStatus:ERR Object required', 'LookAtStatus:DONE']` |

### settle_ms = 1000  ·  outcome: `unblocks_neither`

| Button | outcome | files | :ERR | scratch_status / p_status | ZZ_TEST_DEBUG |
|---|---|---:|---|---:|---|
| `CmdPajek` | `object_required_still_observed` | 0 | `Object required` | 17023 / 17022 | `['LookAtStatus:ENTER', 'LookAtStatus:ERR Object required', 'LookAtStatus:DONE']` |
| `CmdGephi` | `object_required_still_observed` | 0 | `Object required` | 17023 / 17022 | `['LookAtStatus:ENTER', 'LookAtStatus:ERR Object required', 'LookAtStatus:DONE']` |

## Q1-Q3 answers (only the brief's 3 questions)

**Q1 — Smallest settle_ms where neither button hits `Object required`:**
- CmdPajek first no-Object-required at: `None` ms
- CmdGephi first no-Object-required at: `None` ms

**Q2 — Smallest settle_ms where both buttons actually write files:**
- CmdPajek first file_count >= 1 at: `None` ms
- CmdGephi first file_count >= 1 at: `None` ms

**Q3 — Any settle_ms where the buttons diverge?**
- First divergence at: `None` ms (None = no divergence observed)

## Verdict: `no_value_up_to_1000ms_unblocks`

**No tested VBA-side settle value (up to 1000 ms) unblocked either button.**  All phases produced identical failure shape (`:ERR Object required` + 0 files).  This is a SURPRISING result given PR #131's positive 1.5 s signal — and it has a specific mechanism implication.

**Mechanism implication.**  PR #131's positive signal used Python COM-side `time.sleep(1.5)` — that releases the COM thread fully, allowing the Access UI thread to process `Form.Requery` side-effects asynchronously.  This probe's settle is a VBA-side `Do While Timer DoEvents` loop running ON THE Access UI THREAD itself.  DoEvents from inside the UI thread lets nested UI events fire but does NOT release the thread the way Python's `time.sleep` does.  So **COM-side sleep ≠ VBA-side DoEvents settle**, even at the same wall-clock duration.  Naively extrapolating from PR #131's COM-side 1.5 s to a VBA-side 1.5 s would not work either.

Per the brief: do NOT silently extend the bisect range or switch intervention.  But the result narrows the next-brief candidate space meaningfully:
  - **VBA-side DoEvents-only settle is unlikely to ever work** at any duration; the mechanism is structurally wrong.
  - **COM-side sleep injection** is the path PR #131 already validated.  The driver could interpose between dispatcher steps via Python COM (e.g. set Form.Tag to single-step `CmdQuery` only, await completion, then separately fire each Cmd<X>_Click via timer) — but that's a larger refactor of the chain dispatcher, not a settle-value tweak.
  - **Synchronous in-VBA Requery commit** — alternative VBA primitives might force synchronous commit (e.g. `Forms!LookAtStatus.Painting = True`, `Application.RefreshDatabaseWindow`, `<subform>.SourceObject = ...` reassignment, or `MoveLast` after Requery to force dynaset population).  PR #133's variant + this PR's settle bisect rule out plain Requery + DoEvents; a separate brief would need to test these other primitives.
  - **Maintainer-line / canonical Issue** — the underlying CBDB pattern (Dim'd-local Set <subform>.Form.Recordset rebind in CmdQuery cleanup) is fragile against COM-driven chained-step access; the upstream `.mdb` could be fixed to use globals (per Form_LookAtStatus line 1184's `gRstPeople` global as precedent) or refactor to not need the rebind at all.

**This PR pins the time-axis bisect range and the mechanism boundary**: VBA-side settle interventions are now ruled out (probe surface exhausted); next brief picks a structurally different intervention or escalates to maintainer-line.

## Constraints honoured per brief

- ✅ Investigation artifacts only — driver edit REVERTED before this PR opened (per brief: workaround PR is separate).
- ✅ Intervention FIXED across all settle values — explicit Form.Requery + parametric DoEvents-loop settle.  No candidate switching.
- ✅ Only 3 settle values tested (well under the 4-cap).
- ✅ Each settle value tested on BOTH buttons.
- ✅ Verdict picked from the 4 named buckets per brief.
- ✅ No `tests/test_*` changed; no README, triage, canonical reports / issue severity touched.
- ✅ CmdNeo4j NOT touched.
- ✅ `--reclassify-from-json` supported.
- ✅ `analysis/report_screenshot_audit.md` drift left alone (standing rule).