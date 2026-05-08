# Maintainer handoff — CmdNeo4j family + LookAtStatus export buttons

**To:** CBDB upstream maintainer
**From:** cbdb-user-mdb-tests test framework maintainers
**Date:** 2026-05-08
**Subject:** Four upstream-fix concerns from the cross-form export test surface; consolidated for one round of maintainer review.

---

## Executive summary

Four concerns across the `Form_LookAt*.Cmd*_Click` export-button surface have reached a state where the upstream `.mdb` is the natural next intervention point. Three are already filed as canonical P1 issues with anchored test markers (#21, #23, #24); one (Status × CmdPajek/Gephi cleanup-rebind) is brought to your attention here as a maintainer-review concern, not yet filed. All four share a "cross-form export-button blocker rooted in a CBDB source pattern" theme; bundling them in one round of maintainer review is more efficient than four sequential interactions.

For three of the four (Issues #23, #24, plus the Status concern), the test framework has either landed a repo-local workaround (#23, #24) or exhausted the local-workaround surface across 11 PRs (Status). For the fourth (Issue #21), no local workaround was attempted because the failure is genuinely a runtime bug, not a test-driver fragility.

The fix surfaces are small in every case (10–20 lines of VBA per concern). We are NOT asking for architectural changes or schema changes.

---

## Concern 1 — Issue #21: `LookAtGroupData.CmdNeo4j` 'No current record' on empty sections

| Field | Value |
|---|---|
| **Affected form / button** | `Form_LookAtGroupData` · `CmdNeo4j_Click` |
| **Observed runtime symptom** | DAO error 3021 'No current record.' fires mid-chain; partial files written before the crash; unguarded `.MoveFirst` on an empty recordset in the People-Entry export block |
| **Specific source location** | `Form_LookAtGroupData.vb` lines around 1243–1245 (the `Set tRstPeopleEntry = CurrentDb.OpenRecordset(...)` followed by `.MoveFirst` without an empty-recordset guard); also at the analogous tRstPeoplePlace export block (similar pattern, ~10 lines later) |
| **Current repo disposition** | Filed as canonical P1 (`reports/generate_report.py::ISSUES` id=21); 4-5 anchored static markers in `tests/test_known_bugs.py::test_bug21_*` will fire automatically when upstream ships the fix; cell `LookAtGroupData × CmdNeo4j` remains skipped in cross-form Neo4j coverage |
| **Smallest known upstream fix surface** | Add an `If tRstPeopleEntry.RecordCount > 0 Then ... End If` guard around the `.MoveFirst`+iteration block; same for tRstPeoplePlace. ~6–8 added lines per block. |
| **Repo-local workaround exists?** | No. Static markers + skip; no driver-side rewrite was viable because the bug is structural (unguarded MoveFirst) not symbolic. |

---

## Concern 2 — Issue #23: `LookAtAssociations.CmdNeo4j` INSERT target column typo

| Field | Value |
|---|---|
| **Affected form / button** | `Form_LookAtAssociations` · `CmdNeo4j_Click` |
| **Observed runtime symptom** | JET 3061 'INSERT INTO statement contains the following unknown field name: c_index_addr_type_code'; INSERT into `ZZ_SCRATCH_PEOPLE` fails immediately; 0 files written |
| **Specific source location** | The INSERT into `ZZ_SCRATCH_PEOPLE` references target column `c_index_addr_type_code` which does NOT exist on that table; the table DOES have `c_addr_type` (per the schema); the surrounding UPDATE statements correctly use `c_addr_type`. Looks like a single-column typo in the INSERT statement. |
| **Current repo disposition** | Filed as canonical P1 (id=23); 4 anchored static markers in `tests/test_known_bugs.py::test_bug23_*`; `LookAtAssociations × CmdNeo4j` IS covered locally via two driver workarounds (PR #116 + #117); coverage state and canonical issue state are independent assertions on the same source defect |
| **Smallest known upstream fix surface** | Rename the target column in the INSERT statement: `c_index_addr_type_code` → `c_addr_type`. Single token change. |
| **Repo-local workaround exists?** | Yes — `_rewrite_associations_cmdneo4j_target_column` in `tests/cbdb_driver/vba_session.py` rewrites the in-memory VBA at test time. Workaround does NOT change the on-disk `.mdb`; user-visible defect remains. |

---

## Concern 3 — Issue #24: `LookAtPlace.CmdNeo4j` `tRstPeople` SELECT projection missing columns

| Field | Value |
|---|---|
| **Affected form / button** | `Form_LookAtPlace` · `CmdNeo4j_Click` |
| **Observed runtime symptom** | JET 3265 'Item not found in this collection.' fires mid-body when the loop reads `!c_dynasty` / `!c_dynasty_chn` / `!c_female` from a recordset whose SELECT projected only 4 ZZ_SCRATCH_P_TEXT columns; downstream field reads can't find these columns; 0 files written |
| **Specific source location** | The `tRstPeople` recordset opens with a SELECT that lists only 4 ZZ_SCRATCH_P_TEXT columns, but the loop body reads 3 additional fields via `!c_dynasty` / `!c_dynasty_chn` / `!c_female`. Those 3 fields exist on the JOINed source tables (DYNASTIES + BIOG_MAIN) but were omitted from the SELECT projection. |
| **Current repo disposition** | Filed as canonical P1 (id=24); 5 anchored static markers in `tests/test_known_bugs.py::test_bug24_*`; `LookAtPlace × CmdNeo4j` IS covered locally via one driver workaround (PR #123); same independence rule as #23 |
| **Smallest known upstream fix surface** | Extend the `tRstPeople` SELECT projection to add `DYNASTIES.c_dynasty, DYNASTIES.c_dynasty_chn, BIOG_MAIN.c_female`. ~3 added column references in the SELECT clause; no JOIN changes needed. |
| **Repo-local workaround exists?** | Yes — `_rewrite_place_cmdneo4j_trstpeople_projection` in `tests/cbdb_driver/vba_session.py` rewrites the SELECT projection at test time. |

---

## Concern 4 (NEW; not yet canonicalized) — `LookAtStatus.CmdPajek` + `LookAtStatus.CmdGephi` cleanup-rebind root cause

This is a new maintainer-review concern that emerged from a recent investigation chain. It is **NOT** filed as a canonical Issue; whether to canonicalize it is your call.

| Field | Value |
|---|---|
| **Affected form / button** | `Form_LookAtStatus` · `CmdPajek_Click` AND `CmdGephi_Click` (sibling cells; same root cause) |
| **Observed runtime symptom** | VBA error 424 'Object required' when the export sub reads `<subform>.Form.Recordset.RecordCount` immediately post-CmdQuery (under any sequential test-driver dispatch). Both sibling buttons (`CmdPajek`, `CmdGephi`) bail at the same structural point; both produce 0 files. |
| **Specific source location** | `Form_LookAtStatus.vb` cleanup section `Exit_Run_Query` (around lines 1457 + 1460): `Set ZZ_SCRATCH_STATUS.Form.Recordset = tRstStatus` and `Set ZZ_SCRATCH_P_STATUS.Form.Recordset = gRstPeople` where `tRstStatus` is a `Dim`'d-local recordset variable. After the CmdQuery sub returns, the local-variable's recordset reference dies; the subform's `.Form.Recordset` reads as Nothing in any subsequent sub that touches it. |
| **Current repo disposition** | Both cells remain skipped on the cross-form Pajek/Gephi test (`tests/test_vba_pajek_gephi_cross_form.py::_case_skip_marks`). The CmdNeo4j sibling on the same form is covered (PR #128) because `Form_LookAtStatus.CmdNeo4j_Click` opens fresh `dbOpenDynaset` directly on the underlying scratch tables (lines 527+528) and bypasses the subform recordset entirely. |
| **Smallest known upstream fix surface** | Switch the cleanup-rebind from `Dim`'d-local to a globally-owned recordset variable, OR refactor away the rebind entirely. The same form already uses a global `gRstPeople` for the People scratch (line 1184 precedent); applying the same pattern to `tRstStatus` (rename to `gRstStatus`, declare in a Module instead of Dim'd-locally) is the smallest viable change. ~3–5 added lines + 1 line modified. Alternative: re-open the recordset at the top of each consuming sub (CmdPajek, CmdGephi) instead of relying on the cleanup binding. |
| **Repo-local workaround exists?** | **No, and the local-workaround surface has been demonstrably exhausted.** 7 distinct mechanism layers were tested across PR #129, #132, #133, #134, #135, #136, #137: per-form VBA literal rewrite (Set→Requery); VBA-side DoEvents settle window (250 / 500 / 750 / 1000 ms); COM-side sleep + sequential dispatch; direct invoke via `Application.Run`; raw-COM `Form_Timer` re-injection + force-compile. All 7 attempts failed; final probe pinned the layer at Access form-class-instance event-binding cache (PR #137). The mechanism boundary is at Access internals, not in our test driver — meaning further test-driver-side carving has steeply diminishing returns. |

**Why this is brought to maintainer review now (instead of remaining in local investigation):**

The 7-PR exhaustion record is the trigger. Each layer of test-driver-side intervention was tried and failed for an independent reason (mechanism boundary; event-binding cache). The remaining theoretical local candidates (close+reopen form between dispatches; standard-module Form_Timer redirect; pywinauto button click) each carry significant implementation cost AND uncertain outcome AND would not address the root cause.

A small upstream change to the cleanup section would address it once for any future test infrastructure AND remove a fragility in the production code path itself (the `Dim`'d-local Set-rebind pattern is fragile against any caller that doesn't re-open the recordset before reading).

---

## What is already covered locally vs what is still broken upstream

| Cell | Test framework state | Upstream defect state |
|---|---|---|
| `LookAtGroupData × CmdNeo4j` | skipped on cross-form CmdNeo4j test | Issue #21 P1 — open upstream; fix landed → static markers fire → skip can be removed |
| `LookAtAssociations × CmdNeo4j` | covered (via PR #116 + #117 driver workarounds) | Issue #23 P1 — **still open upstream**; covered in tests does NOT mean fixed in `.mdb`; fix landed → static markers fire → workaround can be removed |
| `LookAtPlace × CmdNeo4j` | covered (via PR #123 driver workaround) | Issue #24 P1 — **still open upstream**; same independence rule as #23 |
| `LookAtStatus × CmdNeo4j` | covered (no workaround needed; prior skip was a false-positive) | NO defect — CmdNeo4j on this form bypasses the subform recordset; runs cleanly |
| `LookAtStatus × CmdPajek` | skipped on cross-form Pajek/Gephi test | NEW concern (Concern 4 above) — not yet canonicalized; awaiting maintainer review |
| `LookAtStatus × CmdGephi` | skipped on cross-form Pajek/Gephi test | NEW concern (Concern 4 above) — sibling of CmdPajek; same root cause |

**Headline:** **CmdNeo4j family is 8 covered / 0 skipped** in the repo's cross-form Neo4j test (all 8 hosts running). Issues #23 and #24 have repo-local workarounds that make the cells *testable* on the existing source; these workarounds are NOT upstream fixes — the underlying `.mdb` defects remain user-visible. **Status × CmdPajek + CmdGephi remain skipped** with no local workaround landed; they're parked pending this maintainer-review outcome.

---

## What this handoff is NOT asking for

- It is NOT asking you to merge any of the test-driver-side workarounds into the `.mdb`. Those workarounds exist only in the test repo; they should never reach production.
- It is NOT asking for a schema change. All four fix surfaces are VBA source edits.
- It is NOT asking you to canonicalize the Status × CmdPajek/Gephi concern as Issue #25 (or whichever next ID is free). Whether to file it formally is your decision; this memo presents it as a maintainer-review concern with the same evidence rigor as the filed issues.
- It is NOT escalating Issue #22 (`LookAtAssociations × CmdUCINet` FSO Unicode). That's a separate channel — different export family, different defect class. Not bundled here.

---

## Verifiable signals after each upstream fix lands

For each of the three filed issues (#21 / #23 / #24), the test framework already has anchored static markers in `tests/test_known_bugs.py::test_bug{21,23,24}_*` that read the on-disk VBA dump. Once the upstream fix is in the `.mdb` we use as source-of-truth (`data/CBDB_BJ_User.mdb` or whatever the next dump captures), the static markers will fire automatically. That's the agreed signal for marking each canonical issue resolved AND for removing the corresponding driver workaround (where one exists).

For Concern 4 (Status cleanup-rebind): if it's filed as a new canonical issue at your discretion, an analogous static-marker test could be added in a follow-up PR (anchoring on the `Set ZZ_SCRATCH_STATUS.Form.Recordset = tRstStatus` literal at lines 1457+1460). That detection would automatically fire once the upstream fix replaces the pattern.

---

*End of handoff memo.*

---

## Constraints honoured per brief

- ✅ Read-only analysis; no Access COM run; no tests / driver / README / canonical reports / issue severity changes
- ✅ Only the two handoff artifacts touched (this MD + paired JSON)
- ✅ Did NOT decide whether to canonicalize Concern 4 — explicitly written as a maintainer-review concern; the formal filing decision is left to the maintainer
- ✅ Each of 4 concerns includes affected form/button + observed symptom + current repo disposition + smallest known upstream fix surface + workaround-exists-or-not
- ✅ Status × CmdPajek/Gephi explicitly grounded in 7-PR exhaustion (PR #129/#132/#133/#134/#135/#136/#137) with the layer-by-layer attempt record
- ✅ "What's covered vs still broken upstream" section is explicit per-cell
- ✅ Memo style throughout; not a triage log; speaks to a single audience (CBDB upstream maintainer)
- ✅ Issue #22 explicitly excluded from bundle (separate channel)
- ✅ `analysis/report_screenshot_audit.md` drift left alone (standing rule)

## Self-review (per `docs/skills/programmer-self-review-template.md`)

**A. Branch shape.** Branch `chore/maintainer-handoff-cmdneo4j-and-status` cut from current `main = 43600fa` (post PR #138 merge). Only the two handoff files touched (this MD + paired JSON). No tests, driver, README, canonical reports, issue severity, or other artifacts changed. Pre-existing `analysis/report_screenshot_audit.md` drift left alone per standing instruction.

**B. Source-of-truth sync.** All four concerns' canonical-issue facts (titles, IDs, P1 status, fix recommendations) are quoted from the live `reports/generate_report.py::ISSUES` entries OR from the recent triage refresh `refresh_2026_05_08_later`'s `next_work_items_ranked` rank-1 bundle. The cell-state grid in the "What is already covered locally vs what is still broken upstream" section reflects the actual current state of `tests/test_vba_cmdneo4j_cross_form.py::_spec_skip_marks` (post-PR #128: empty body) and `tests/test_vba_pajek_gephi_cross_form.py::_case_skip_marks` (Status entries unchanged). No source-of-truth file is being modified by this handoff.

**C. Evidence vs claim.** Concern 4's "local-workaround surface exhausted" claim is grounded in the explicit 7-PR enumeration (#129 / #132 / #133 / #134 / #135 / #136 / #137); the mechanism boundary at Access form-class-instance event-binding cache is grounded in PR #137's verdict. The "smallest known upstream fix surface" for Concern 4 cites the in-source `gRstPeople` precedent at line 1184 — verifiable in the dump. No claim of "guaranteed fix" anywhere; all four concerns are framed as maintainer-review with concrete suggested directions, not prescriptions. The "covered locally != fixed upstream" rule is restated identically to how prior triage refreshes have framed it.

**D. Residual risk.** Handoff memo, not implementation; residual risk is purely communication-style: (1) the maintainer may prefer a different bundling — e.g. handle each concern as its own message; the memo doesn't prevent that. (2) The "smallest known upstream fix surface" for Concern 4 lists two candidates (global recordset variable OR re-open at consumer); the maintainer may choose neither and prefer a third approach — that's fine, the memo presents directions, not requirements. (3) The decision whether to canonicalize Concern 4 as Issue #25 is explicitly left open. (4) No code path or test altered, so no runtime regression risk.
