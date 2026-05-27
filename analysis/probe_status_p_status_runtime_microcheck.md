# LookAtStatus × {CmdPajek, CmdGephi} runtime micro-check — 3 reads

**Date:** 2026-05-08  ·  **Branch:** `investigate/status-p-status-runtime-microcheck` (off main `a783b72`, post PR #129 + PR #130 merge)

Single-purpose micro-probe per PR #130's `minimum_next_confirmation` block.  Answers exactly 3 reads after `CmdQuery_Click` on the matrix Status fixture; classifies into one of 4 outcome buckets; ships investigation artifacts only.  Does NOT test candidate (b)/(c)/(d), does NOT propose driver workarounds.

## Setup

- **Form:** `LookAtStatus`
- **Fixture:** matrix `_make_status_fixtures` first fixture (`status_40_unfiltered`)
- **Chain:** CmdQuery alone (no chained button) — we only need the post-CmdQuery form state to ask the 3 questions.
- **click_via_timer cap:** 180 s  ·  **outer cap:** 240 s  ·  **total elapsed:** 11.65 s

## Raw observed facts

- **click_via_timer_returned:** `17023`  (matches PR #127 / PR #129 baseline = 17023)
- **scratch_status_count (independent COM read):** `17023`
- **scratch_p_status_count (independent COM read):** `17022`

### Q1 — `ZZ_SCRATCH_P_STATUS` c_dynasty='unknown' counts

| Predicate | COUNT(*) |
|---|---:|
| `c_dynasty = 'unknown'` | `7` |
| `c_dynasty IS NULL OR c_dynasty <> 'unknown'` | `17015` |
| total | `17022` |

### Q2 — `Form.FilterOn` runtime state

| Subform | FilterOn |
|---|---|
| `ZZ_SCRATCH_P_STATUS` (brief's primary Q2) | `False` |
| `ZZ_SCRATCH_STATUS` (supplementary cross-check) | `False` |

### Q3 — RecordCount after explicit Requery + brief settle

| Subform | RecordCount | PR #127 baseline | match? |
|---|---:|---:|---|
| `ZZ_SCRATCH_STATUS` | `17023` | 17023 | True |
| `ZZ_SCRATCH_P_STATUS` | `17022` | 17022 | True |

## Interpretation (separated from raw facts)

**H_chain_timing supported.**  After explicit `<subform>.Form.Requery` + brief settle, BOTH subform recordsets return expected RecordCounts (STATUS=17023, P_STATUS=17022).  This means Requery DOES rebind to the populated table when given breathing room; the chain dispatcher's compressed timeline (CmdQuery → Cmd<X> in one Form_Timer cycle) is what prevents the rebind from completing in PR #129's post-(a) state.

Decision space: driver-side dispatcher tweak — insert a DoEvents (or short sleep) between chain steps in `_inject_autodetect`'s dispatch loop (`tests/cbdb_driver/vba_session.py`).  This is test-infrastructure work, NOT a CBDB defect.  Separate brief; this micro-probe does NOT implement it.

## Outcome: `H_chain_timing_supported`

Per the brief, this is a single-purpose micro-probe.  It does NOT implement the next intervention (whichever outcome bucket it lands in).  The next brief picks the intervention based on this outcome.

## Markers (timeline, 15 entries)

  - `+  0.00s` constructing_session
  - `+  5.46s` session_opened_attempt_1
  - `+  6.19s` form_opened
  - `+  6.21s` picker_seeded_1_codes
  - `+  6.21s` fixture_seeded
  - `+  6.21s` form_tag_set_cmdquery_only
  - `+  7.74s` click_via_timer_returned_17023
  - `+  7.74s` scratch_counts_captured
  - `+  7.75s` q1_counts_captured
  - `+  7.76s` q2_filter_on_pstatus=False_status=False
  - `+  7.78s` q3_status_requery_called
  - `+  7.79s` q3_p_status_requery_called
  - `+  7.79s` q3_doevents_skipped: AttributeError('Access.Application.DoEvents')
  - `+  9.29s` q3_settle_slept_1500ms
  - `+  9.29s` q3_recordcounts_status=17023_pstatus=17022

## Constraints honoured per brief

- ✅ Investigation artifacts only — probe + paired MD + paired JSON; no driver / test / README / triage / canonical reports / issue severity changed.
- ✅ Probe asks exactly the 3 brief Q's — Q1 c_dynasty counts, Q2 P_STATUS FilterOn, Q3 RecordCount post-Requery+settle.  No additional reads beyond a single supplementary STATUS FilterOn (kept in raw facts only, not used in inference).
- ✅ Did NOT test candidate (b)/(c)/(d) — those belong to a separate brief AFTER this micro-probe's outcome selects the next decision.
- ✅ Raw facts and interpretation separated into different sections (raw under '## Raw observed facts'; inference under '## Interpretation').
- ✅ `--reclassify-from-json` supported.
- ✅ `analysis/report_screenshot_audit.md` drift left alone (standing rule).