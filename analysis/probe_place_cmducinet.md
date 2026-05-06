# LookAtPlace × CmdUCINet — runtime characterization probe

**Date:** 2026-05-06  ·  **Branch:** `investigate/place-cmducinet-shape`  ·  **Base:** main `cdfca69`

**Brief context:** PR AR's queue refresh ranked Place × CmdUCINet as the cheapest next probe in the CmdUCINet family because Place's writer is *structurally* different from the FSO ANSI path that drives Issue #22 (ADO Stream + UTF-8 vs FSO CreateTextFile + cp1252). This probe answers — at runtime — whether Place CmdUCINet produces a clean parseable file, whether the ADO Stream UTF-8 path actually writes UTF-8, and whether any new blocker exists. **It is NOT an Issue #22 reproduction**; Place is not to be folded into Issue #22's family on the strength of this probe.

## Headline

Conclusion: **`still_needs_better_fixture`**

CmdUCINet was never able to fire: every probe iteration on this dump shows that the second COM call (set_form_tag for CmdUCINet) raises `com_error('The RPC server is unavailable.')` AFTER the CmdQuery timer either completes or times out. This is a *driver-CmdQuery interaction issue on the Place form*, not a Place CmdUCINet bug — Place CmdUCINet's runtime behaviour is still unobserved.

Picker size / checkbox state / synthetic-row injection are NOT the root cause: the same failure mode reproduced across (a) Kaifeng addr 100658 with default checkbox state, (b) Kaifeng with ChkKin+ChkAssocPlace enabled, (c) synthetic-row inject bypassing CmdQuery, (d) Chenliu addr 3089 (3 BIOG_ADDR_DATA rows). The common factor is that CmdQuery on the Place form holds Access in a state where the next set_form_tag call fails.

The brief explicitly forbids driver changes in this PR, so this probe cannot complete CmdUCINet's runtime characterization. A follow-up needs either (a) a driver-side change so CmdQuery's completion releases the COM bridge cleanly (separate maintainer brief), OR (b) a different probe shape that fires CmdUCINet via a path that doesn't go through click_via_timer (e.g. pywinauto UI click on the Access window — but VbaSession.click_button currently expects a visible form, and headless Access may not support it). The static prediction (ADO Stream + UTF-8, 3-section file shape, FSO path commented out) remains the only available characterization of Place CmdUCINet's writer.

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

CmdUCINet was never able to fire: every probe iteration on this dump shows that the second COM call (set_form_tag for CmdUCINet) raises `com_error('The RPC server is unavailable.')` AFTER the CmdQuery timer either completes or times out. This is a *driver-CmdQuery interaction issue on the Place form*, not a Place CmdUCINet bug — Place CmdUCINet's runtime behaviour is still unobserved.

Picker size / checkbox state / synthetic-row injection are NOT the root cause: the same failure mode reproduced across (a) Kaifeng addr 100658 with default checkbox state, (b) Kaifeng with ChkKin+ChkAssocPlace enabled, (c) synthetic-row inject bypassing CmdQuery, (d) Chenliu addr 3089 (3 BIOG_ADDR_DATA rows). The common factor is that CmdQuery on the Place form holds Access in a state where the next set_form_tag call fails.

The brief explicitly forbids driver changes in this PR, so this probe cannot complete CmdUCINet's runtime characterization. A follow-up needs either (a) a driver-side change so CmdQuery's completion releases the COM bridge cleanly (separate maintainer brief), OR (b) a different probe shape that fires CmdUCINet via a path that doesn't go through click_via_timer (e.g. pywinauto UI click on the Access window — but VbaSession.click_button currently expects a visible form, and headless Access may not support it). The static prediction (ADO Stream + UTF-8, 3-section file shape, FSO path commented out) remains the only available characterization of Place CmdUCINet's writer.

### Iteration history (for the probe author's future self)

This probe was run in four iterations before the current configuration. None of them produced a runtime characterization of CmdUCINet's writer; all of them showed the SAME failure mode (`com_error('The RPC server is unavailable.')` on the second set_form_tag call after CmdQuery). Iterations recorded so future probe authors don't re-run the same dead ends:

1. **Kaifeng (100658), default checkboxes.** CmdQuery returned 2208 rows in 187s — all `Biography` / `Office Place` / `Entry`, zero `Kinship` / `Associate Place`. CmdUCINet would have bailed at its 'no networks' guard even if the COM bridge had stayed alive (which it did not).
2. **Kaifeng + ChkKin + ChkAssocPlace enabled.** CmdQuery returned 8370 rows in 240s (incl. 6159 `Kinship`); ChkAssocPlace contributed 0 rows on Kaifeng, suggesting that source needs more than just the checkbox toggle. COM bridge died immediately after CmdQuery.
3. **Synthetic-row injection bypassing CmdQuery** (open form, INSERT 4 Kinship rows directly, Requery the subform). The subform Requery via `sess.app.Forms('LookAtPlace').Controls('frmZZZ_PLACE').Form.Requery()` raised `AttributeError('Access.Application.Forms')`; the next set_form_tag failed with RPC unavailable.
4. **Chenliu (3089) + ChkKin (current run).** Smallest realistic picker (3 BIOG_ADDR_DATA rows, 3 KIN_DATA links per pyodbc scan of `data/CBDB_*_DATA.mdb`). CmdQuery still timed out at 120s and returned 3 rows (all `Biography` — ChkKin's filter wants people whose *kin's* address is in the picker, not people whose own address is). COM bridge died immediately after.

Common factor: every iteration that reached the second set_form_tag call (for CmdUCINet) failed the same way. The instability is NOT picker-size, NOT checkbox-state, NOT row-count dependent — it is a *driver-CmdQuery interaction on the Place form*.

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