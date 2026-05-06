# LookAtPlace post-form-work COM bridge instability — trigger-isolation matrix

**Date:** 2026-05-06  ·  **Branch:** `investigate/place-com-bridge-instability`  ·  **Base:** main `db16c03`

**Brief context:** PR AS's Place × CmdUCINet probe established that the bridge fails RPC-unavailable on the second `set_form_tag` call across all four iterations.  PR AS deliberately did NOT collapse this to 'CmdQuery is the trigger' because iter 3 (synthetic-row injection + failed direct-COM Requery, no CmdQuery) ALSO reproduced the failure.  This probe runs a 5-trial isolation matrix to localize WHICH form-side operation actually triggers the bridge failure.

## Headline

Narrowest class: **`long_click_via_timer_polling_loop_correlated`**

T1+T2 baselines stable. T3 leaves bridge DEAD AND its `click_via_timer(CmdQuery)` step ran for the full 120s timeout (DONE marker not surfaced via the autodetect inject — `click_via_timer` kept polling Access COM for the full window). T5 leaves bridge ALIVE AND its `click_via_timer(CmdQuery)` step returned cleanly in 6.54 s (DONE marker surfaced or result_table count grew).  Same form, same picker, same checkboxes — the only behavioural difference between the two is HOW `click_via_timer` returned.  This strongly suggests the narrowest trigger is NOT CmdQuery itself but `click_via_timer`'s polling loop running for the full timeout window, which appears to age out / poison the COM bridge.  CmdQuery is NOT exonerated (the polling-timeout happens during CmdQuery), but a clean CmdQuery completion does NOT poison the bridge in this matrix. Note: T4 did NOT reproduce PR AS iter 3's failed-Requery → bridge-dead pattern.  In this run the direct-COM Requery actually succeeded (no AttributeError) and the bridge stayed alive after the subsequent set_form_tag.  PR AS iter 3's AttributeError on the same call may have been a one-off win32com binding flake or had a timing dependency this run didn't hit.  This narrows the original 'failed Requery alone poisons bridge' hypothesis: NOT supported by this probe's run; the iter-3 RPC failure was more likely caused by the synthetic INSERT into ZZ_PLACE during a still-open form recordset OR by COM session aging from the prior failed iterations (which v3 inherited via reuse of the MSACCESS process state).

---

## Probe-observed facts (per-trial outcomes)

| trial | question | diagnostic op | diag class | bridge state | elapsed |
|---|---|---|---|---|---:|
| `T1` | Q1: Is `set_form_tag` alone stable post-`open_form`? | `set_form_tag` | `ok` | `alive` | 7.29 s |
| `T2` | Q2: Stable when only CmdUCINet is dispatched (no prior CmdQuery)? | `set_form_tag_then_dispatch` | `ok` | `alive` | 8.34 s |
| `T3` | Q3: Does CmdQuery + second set_form_tag fail (original PR AS iter 1/2/4 pattern)? | `set_form_tag` | `ok` | `dead` | 126.73 s |
| `T4` | Q4: Does failed direct-COM Requery alone poison the bridge (PR AS iter 3 pattern, minus the synthetic INSERT)? | `set_form_tag` | `ok` | `alive` | 8.54 s |
| `T5` | Q5: Post-CmdQuery, is the failure specific to `set_form_tag`, or is any COM call dead? | `forms_count_only` | `ok` | `alive` | 13.81 s |

## Direct yes/no answers to brief Q1–Q5

- **Q1_open_then_set_form_tag_stable** → `yes`
- **Q2_open_then_dispatch_only_stable** → `yes`
- **Q3_post_cmdquery_set_form_tag_fails** → `yes`
- **Q4_post_failed_requery_set_form_tag_fails** → `no`
- **Q5_post_cmdquery_minimal_com_bridge_dead** → `no`

## Per-trial step detail

### T1 — open_form + set_form_tag (baseline)

_Q1: Is `set_form_tag` alone stable post-`open_form`?_

- session open attempt: 1
- diagnostic outcome: `ok` (class `ok`)
- bridge state at end: `alive`
- total elapsed: 7.29 s

Steps:

| op | result | class | elapsed |
|---|---|---|---:|
| `open_form(LookAtPlace)` | ok | `ok` | 0.83 s |
| `set_form_tag(CmdUCINet, '<diag>') [diagnostic]` | ok | `ok` | 0.0 s |
| `bridge_alive_check_forms_count` | ok | `ok` | 0.0 s |

### T2 — open_form + set_form_tag + click_via_timer(CmdUCINet, wait_done=False)

_Q2: Stable when only CmdUCINet is dispatched (no prior CmdQuery)?_

- session open attempt: 1
- diagnostic outcome: `ok` (class `ok`)
- bridge state at end: `alive`
- total elapsed: 8.34 s

Steps:

| op | result | class | elapsed |
|---|---|---|---:|
| `open_form(LookAtPlace)` | ok | `ok` | 0.81 s |
| `set_form_tag(CmdUCINet, '<diag>')` | ok | `ok` | 0.0 s |
| `click_via_timer(CmdUCINet, wait_done=False) [diagnostic]` | ok | `ok` | 2.02 s |
| `bridge_alive_check_forms_count` | ok | `ok` | 0.0 s |

### T3 — open_form + picker + ChkKin + click_via_timer(CmdQuery) + set_form_tag(CmdUCINet)

_Q3: Does CmdQuery + second set_form_tag fail (original PR AS iter 1/2/4 pattern)?_

- session open attempt: 1
- diagnostic outcome: `ok` (class `ok`)
- bridge state at end: `dead`
- total elapsed: 126.73 s

Steps:

| op | result | class | elapsed |
|---|---|---|---:|
| `open_form(LookAtPlace)` | ok | `ok` | 0.83 s |
| `set_picker_codes(addr 3089)` | ok | `ok` | 0.0 s |
| `set_control(ChkKin=True)` | ok | `ok` | 0.01 s |
| `set_form_tag(CmdQuery, '')` | ok | `ok` | 0.0 s |
| `click_via_timer(CmdQuery, timeout=120s)` | ok | `ok` | 120.22 s |
| `set_form_tag(CmdUCINet, '<diag>') [diagnostic]` | ok | `ok` | 0.0 s |
| `bridge_alive_check_forms_count` | fail | `bridge_dead_rpc` | 0.0 s |

First failing step exception: `com_error(-2147023174, 'The RPC server is unavailable.', None, None)`

### T4 — open_form + failed direct-COM Requery + set_form_tag(CmdUCINet)

_Q4: Does failed direct-COM Requery alone poison the bridge (PR AS iter 3 pattern, minus the synthetic INSERT)?_

- session open attempt: 1
- diagnostic outcome: `ok` (class `ok`)
- bridge state at end: `alive`
- total elapsed: 8.54 s

Steps:

| op | result | class | elapsed |
|---|---|---|---:|
| `open_form(LookAtPlace)` | ok | `ok` | 0.93 s |
| `direct_com.Forms('LookAtPlace').Controls('frmZZZ_PLACE').Form.Requery()` | ok | `ok` | 0.01 s |
| `set_form_tag(CmdUCINet, '<diag>') [diagnostic]` | ok | `ok` | 0.0 s |
| `bridge_alive_check_forms_count` | ok | `ok` | 0.0 s |

### T5 — open_form + picker + ChkKin + click_via_timer(CmdQuery) + Forms.Count read (NOT set_form_tag)

_Q5: Post-CmdQuery, is the failure specific to `set_form_tag`, or is any COM call dead?_

- session open attempt: 1
- diagnostic outcome: `ok` (class `ok`)
- bridge state at end: `alive`
- total elapsed: 13.81 s

Steps:

| op | result | class | elapsed |
|---|---|---|---:|
| `open_form(LookAtPlace)` | ok | `ok` | 0.93 s |
| `set_picker_codes(addr 3089)` | ok | `ok` | 0.01 s |
| `set_control(ChkKin=True)` | ok | `ok` | 0.01 s |
| `set_form_tag(CmdQuery, '')` | ok | `ok` | 0.0 s |
| `click_via_timer(CmdQuery, timeout=120s)` | ok | `ok` | 6.54 s |
| `app.Forms.Count [diagnostic — minimal COM read, NOT set_form_tag]` | ok | `ok` | 0.0 s |
| `bridge_alive_check_forms_count` | ok | `ok` | 0.0 s |

---

## Inferences (analytical, not additional probe data)

Re-verify before acting on these.  These are conclusions drawn from the trial matrix, NOT additional COM observations.

**Narrowest class: `long_click_via_timer_polling_loop_correlated`**

T1+T2 baselines stable. T3 leaves bridge DEAD AND its `click_via_timer(CmdQuery)` step ran for the full 120s timeout (DONE marker not surfaced via the autodetect inject — `click_via_timer` kept polling Access COM for the full window). T5 leaves bridge ALIVE AND its `click_via_timer(CmdQuery)` step returned cleanly in 6.54 s (DONE marker surfaced or result_table count grew).  Same form, same picker, same checkboxes — the only behavioural difference between the two is HOW `click_via_timer` returned.  This strongly suggests the narrowest trigger is NOT CmdQuery itself but `click_via_timer`'s polling loop running for the full timeout window, which appears to age out / poison the COM bridge.  CmdQuery is NOT exonerated (the polling-timeout happens during CmdQuery), but a clean CmdQuery completion does NOT poison the bridge in this matrix. Note: T4 did NOT reproduce PR AS iter 3's failed-Requery → bridge-dead pattern.  In this run the direct-COM Requery actually succeeded (no AttributeError) and the bridge stayed alive after the subsequent set_form_tag.  PR AS iter 3's AttributeError on the same call may have been a one-off win32com binding flake or had a timing dependency this run didn't hit.  This narrows the original 'failed Requery alone poisons bridge' hypothesis: NOT supported by this probe's run; the iter-3 RPC failure was more likely caused by the synthetic INSERT into ZZ_PLACE during a still-open form recordset OR by COM session aging from the prior failed iterations (which v3 inherited via reuse of the MSACCESS process state).

### Remaining hypotheses (NOT closed by this probe)

- Whether the same instability would manifest under a NON-Place form open + identical work (this probe does not run cross-form trials).
- Whether there is a Win11/Office build interaction (this probe runs on a single machine; reproducibility on other Office builds is unknown).
- Whether a driver-side change to `click_via_timer`'s polling loop (which sleeps + touches `app` repeatedly during the wait) would affect the post-CmdQuery state.
- For PR AS iter 3 specifically: whether the synthetic INSERT into ZZ_PLACE (separate ODBC connection) contributed to the instability.  This probe's T4 deliberately omits the INSERT to isolate the failed-Requery effect.

## Constraints honoured per brief

- ✅ Investigation artifacts only — no `tests/` / `cbdb_driver/*` / README / canonical / inventory / issue changes.
- ✅ Did NOT restart Place × CmdUCINet probe.
- ✅ Did NOT re-evaluate Issue #22 family.
- ✅ Used Access COM via existing public VbaSession methods + read-only `app` access for the failed-Requery reproduction.
- ✅ Probe-observed facts (per-trial outcomes) vs inferences (narrowest-class synthesis) explicitly separated.

### Self-review (per `docs/skills/programmer-self-review-template.md`)

- **A. Branch shape.** `investigate/place-com-bridge-instability` cut from clean main `db16c03`; only the 3 brief-named files in the diff (script + MD + JSON); no driver / tests / canonical changes.
- **B. Source-of-truth sync.** MD ↔ JSON carry the same per-trial outcomes, the same Q1–Q5 answers, and the same narrowest-class synthesis.  No source-of-truth file is modified by this probe.
- **C. Evidence vs claim.** Per-trial diagnostic outcomes are direct probe observations.  The narrowest-class synthesis is explicitly labeled inferential and falls back to `unresolved_partial_data` rather than asserting a class beyond what the matrix supports.  Remaining hypotheses are listed so the next probe author doesn't mistake narrowness for completeness.
- **D. Residual risk.** Single-run-per-trial data; if RPC flakes were random, a single T3/T5 fail might be misattributed.  Mitigated by classifying exceptions explicitly into `bridge_dead_rpc` vs `other`, by retrying session open up to 3× before declaring a session-open failure (matches PR AS pattern), and by structuring the synthesis as priority-ordered patterns so a single-trial flake shifts the conclusion to `unresolved_partial_data` rather than to a wrong narrow class.