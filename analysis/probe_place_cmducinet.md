# LookAtPlace × CmdUCINet — runtime characterization probe

**Date:** 2026-05-06  ·  **Branch:** `investigate/place-cmducinet-shape`  ·  **Base:** main `cdfca69`

**Brief context:** PR AR's queue refresh ranked Place × CmdUCINet as the cheapest next probe in the CmdUCINet family because Place's writer is *structurally* different from the FSO ANSI path that drives Issue #22 (ADO Stream + UTF-8 vs FSO CreateTextFile + cp1252). This probe answers — at runtime — whether Place CmdUCINet produces a clean parseable file, whether the ADO Stream UTF-8 path actually writes UTF-8, and whether any new blocker exists. **It is NOT an Issue #22 reproduction**; Place is not to be folded into Issue #22's family on the strength of this probe.

## Headline

Conclusion: **`still_needs_better_fixture`**

**Place CmdUCINet runtime behaviour remains unobserved.** Every probe iteration on this dump ended with the second `set_form_tag` call (for CmdUCINet) raising `com_error('The RPC server is unavailable.')`, before any CmdUCINet code path could execute.

Honest framing of the failure: this is **post-form-work COM bridge instability in the current Place probe shape**.  The narrower claim 'CmdQuery on the Place form is the trigger' is NOT supported by the evidence: iteration 3 deliberately bypassed CmdQuery (synthetic-row injection via pyodbc) and the next `set_form_tag` STILL failed RPC-unavailable after a separately-failed direct-COM Requery attempt.  So whether CmdQuery is necessary to trigger the instability is unresolved; what we CAN say is that under the current probe shape, the COM bridge becomes unavailable after the form-side work that precedes the CmdUCINet fire (CmdQuery in iters 1/2/4, a failed Requery in iter 3).

**Bucket label caveat: `still_needs_better_fixture` is an imperfect fit.**  The brief's fixed bucket vocabulary forces this label, but the actual blocker is NOT fixture insufficiency — it is unresolved probe-shape / COM instability.  Picker size, checkbox state, and synthetic-row injection were all tried; none of them is the root cause.  Reading this bucket as 'try a different fixture and Place will be characterized' would be a misread.

What this PR's evidence DOES support:
  - Place CmdUCINet runtime still unobserved.
  - The current probe shape hits unresolved COM/RPC instability after form-side work.
  - The static prediction (ADO Stream + UTF-8, 3-section file shape, FSO path commented out) remains the only available characterization of Place CmdUCINet's writer.

What a follow-up would need (out of scope for this PR's brief boundary): either (a) a driver-side change targeted at the post-form-work instability, after a separate investigation localizes which form-side operation actually triggers it, OR (b) a different probe shape that fires CmdUCINet via a path that doesn't go through `click_via_timer` (e.g. pywinauto UI click on the Access window — but `VbaSession.click_button` currently expects a visible form, and headless Access may not support it).

---

## Probe-observed facts

- Outcome: `no_file_no_err_no_guard`
- Elapsed: 189.81 s
- Picker addr_id: `3089` (Chenliu (陳留)) — matrix-fixture address; CJK-rich network by construction (no new fixture introduced)
- ZZ_PLACE row count after CmdQuery: `3`

### ZZ_PLACE.c_rel_type distribution after CmdQuery

| c_rel_type | count |
|---|---:|
| `Biography` | 3 |

Network-eligible rows (c_rel_type IN ('Kinship', 'Associate Place')): **0**. These are the only rows CmdUCINet's first INSERT into ZZ_SOCIAL_NETWORK selects from.

### Output file

- Path: `None`
- Size: `None` bytes
- First 16 bytes hex: `None`
- BOM detected: `None`
- Decoded as: `None`

### File shape

(no sections parsed — file did not appear or was empty)

### Static prediction vs runtime

- Predicted sections (static): `['*node data', '*node properties', '*tie data']`
- Actual sections (runtime):   `[]`
- Section order match: **False**
- Predicted encoding (static): UTF-8 (default branch `tStream.Charset = "utf-8"`)
- Actual encoding (runtime):   `None`
- Encoding match: **False**
- Predicted BOM (static): unspecified (ADO Stream `SaveToFile adSaveCreateOverWrite` may or may not emit BOM depending on Charset)
- Actual BOM (runtime):   `None`

### Row-count parity

| metric | file | scratch-table | parity |
|---|---:|---:|---|
| node-data rows | None | ZZ_SCRATCH_PEOPLE = 0 | `None` |
| node-properties rows | None | (should equal node-data rows) | `None` |
| tie-data rows | None | ZZ_SOCIAL_NETWORK = 0 | `None` |

### ZZ_TEST_DEBUG transcript

- `  1`: `LookAtPlace:ENTER`

### Markers timeline

- `+  0.00s` constructing_session
- `+  6.05s` session_opened_attempt_1
- `+  6.33s` filedialog_patched
- `+  7.22s` picker_seeded_addr_3089
- `+  7.23s` checkbox_overrides_{'ChkKin': {'requested': True, 'actual': True}}
- `+127.44s` cmdquery_returned_3
- `+127.44s` zz_place_count_3_rel_types_1
- `+127.44s` cmducinet_fire_exc: com_error(-2147023174, 'The RPC server is unavailable.', None, None)
- `+187.47s` file_appeared_False
- `+187.48s` scratch_tbl_counts_{'ZZ_SOCIAL_NETWORK': 0, 'ZZ_SCRATCH_PEOPLE': 0}
- `+187.48s` debug_captured
- `+187.48s` parity_checks_done

---

## Inferences for canonicalization / coverage follow-up

These are CONCLUSIONS drawn from the probe + static evidence, NOT additional probe observations. Re-verify before relying.

**Conclusion bucket: `still_needs_better_fixture`**

**Place CmdUCINet runtime behaviour remains unobserved.** Every probe iteration on this dump ended with the second `set_form_tag` call (for CmdUCINet) raising `com_error('The RPC server is unavailable.')`, before any CmdUCINet code path could execute.

Honest framing of the failure: this is **post-form-work COM bridge instability in the current Place probe shape**.  The narrower claim 'CmdQuery on the Place form is the trigger' is NOT supported by the evidence: iteration 3 deliberately bypassed CmdQuery (synthetic-row injection via pyodbc) and the next `set_form_tag` STILL failed RPC-unavailable after a separately-failed direct-COM Requery attempt.  So whether CmdQuery is necessary to trigger the instability is unresolved; what we CAN say is that under the current probe shape, the COM bridge becomes unavailable after the form-side work that precedes the CmdUCINet fire (CmdQuery in iters 1/2/4, a failed Requery in iter 3).

**Bucket label caveat: `still_needs_better_fixture` is an imperfect fit.**  The brief's fixed bucket vocabulary forces this label, but the actual blocker is NOT fixture insufficiency — it is unresolved probe-shape / COM instability.  Picker size, checkbox state, and synthetic-row injection were all tried; none of them is the root cause.  Reading this bucket as 'try a different fixture and Place will be characterized' would be a misread.

What this PR's evidence DOES support:
  - Place CmdUCINet runtime still unobserved.
  - The current probe shape hits unresolved COM/RPC instability after form-side work.
  - The static prediction (ADO Stream + UTF-8, 3-section file shape, FSO path commented out) remains the only available characterization of Place CmdUCINet's writer.

What a follow-up would need (out of scope for this PR's brief boundary): either (a) a driver-side change targeted at the post-form-work instability, after a separate investigation localizes which form-side operation actually triggers it, OR (b) a different probe shape that fires CmdUCINet via a path that doesn't go through `click_via_timer` (e.g. pywinauto UI click on the Access window — but `VbaSession.click_button` currently expects a visible form, and headless Access may not support it).

### Iteration history (for the probe author's future self)

This probe was run in four iterations before the current configuration. None of them produced a runtime characterization of CmdUCINet's writer; all of them showed the SAME terminal failure (`com_error('The RPC server is unavailable.')` on the second `set_form_tag` call, the one that fires CmdUCINet). The form-side work that preceded that failure differed across iterations — CmdQuery in iters 1/2/4, a failed direct-COM Requery in iter 3 — so CmdQuery is *not* a proven common cause; see closing paragraph for the narrowest claim the evidence supports. Iterations recorded so future probe authors don't re-run the same dead ends:

1. **Kaifeng (100658), default checkboxes.** CmdQuery returned 2208 rows in 187s — all `Biography` / `Office Place` / `Entry`, zero `Kinship` / `Associate Place`. CmdUCINet would have bailed at its 'no networks' guard even if the COM bridge had stayed alive (which it did not).
2. **Kaifeng + ChkKin + ChkAssocPlace enabled.** CmdQuery returned 8370 rows in 240s (incl. 6159 `Kinship`); ChkAssocPlace contributed 0 rows on Kaifeng, suggesting that source needs more than just the checkbox toggle. COM bridge died immediately after CmdQuery.
3. **Synthetic-row injection bypassing CmdQuery** (open form, INSERT 4 Kinship rows directly via pyodbc, then attempt subform Requery via direct COM). CmdQuery was NOT fired in this iteration. The subform Requery via `sess.app.Forms('LookAtPlace').Controls('frmZZZ_PLACE').Form.Requery()` raised `AttributeError('Access.Application.Forms')`; the next set_form_tag still failed with RPC unavailable. NB: this iteration shows that CmdQuery is NOT a necessary precondition for the post-form-work COM failure — a failed direct-COM Requery is also enough to leave the bridge unavailable.
4. **Chenliu (3089) + ChkKin (current run).** Smallest realistic picker (3 BIOG_ADDR_DATA rows, 3 KIN_DATA links per pyodbc scan of `data/CBDB_*_DATA.mdb`). CmdQuery still timed out at 120s and returned 3 rows (all `Biography` — ChkKin's filter wants people whose *kin's* address is in the picker, not people whose own address is). COM bridge died immediately after.

Common factor: every iteration that reached the second `set_form_tag` call (for CmdUCINet) failed the same way. The instability is NOT picker-size, NOT checkbox-state, NOT row-count dependent. The narrowest claim the evidence supports is **post-form-work COM bridge instability in the current Place probe shape** — the COM bridge becomes unavailable after the form-side work that precedes the CmdUCINet fire (CmdQuery in iters 1/2/4, a failed Requery in iter 3). Whether CmdQuery specifically is necessary to trigger the instability is unresolved.

### Why Place should NOT be folded into Issue #22 regardless of this probe's outcome


(this paragraph is unchanged regardless of iteration outcome — it is a static-evidence argument)

- Place's writer uses ADO Stream (`tStream.WriteText ... adWriteLine`) with explicit `tStream.Charset = "utf-8"`; Issue #22 is canonically about the FSO `CreateTextFile(tFileName, True)` 2-arg call whose ANSI cp1252 path cannot encode CJK Han ideographs.
- The FSO path in Place's source is COMMENTED OUT (`'Set tFileSystem = CreateObject(...)`, `'Set tVNA = tFileSystem.CreateTextFile(...)`).
- Even if this probe surfaces a different Place-CmdUCINet runtime issue, that issue is a *different bug class* and warrants its own canonical entry, NOT a sibling-form note under Issue #22.

## Constraints honoured per brief

- ✅ Investigation artifacts only — no `tests/` or `cbdb_driver/*` changes
- ✅ Did NOT modify README / canonical reports / issue severity
- ✅ Did NOT do an upstream fix
- ✅ Did NOT fold Place into Issue #22's family
- ✅ Did NOT open a coverage PR
- ✅ Used Access COM via `VbaSession.make_fixture`
- ✅ Reused existing matrix-fixture address (Kaifeng 100658) — no new long-term fixture
- ✅ Probe-observed facts vs inferences explicitly separated