# Export coverage gap triage plan

**Date:** 2026-05-04 · **Refreshed:** 2026-05-08 (later — post Status × CmdPajek/Gephi local-workaround line exhaustion: PR #127 + #129 → #137; CmdNeo4j family fully closed at 8/0)
**Branch:** `plan/export-gap-triage` (off main `434168a`); refresh on `refresh/export-gap-triage-after-groupdata`
**Source data (read-only):**
- `reports/export_coverage_inventory.json` — **12 `gap` cells** (was 13; GroupData × CmdGIS landed in PR `cover/groupdata-cmdgis-clean-branches` commit `294cbda`)
- `analysis/dump/vba/Form_LookAt*.vb` — handler presence + chain shape
- `analysis/dump/control_inventory.json` — button presence
- `tests/test_vba_matrix_all_forms.py::_xfail_marks` — matrix CmdQuery/CmdRun blockers
- `tests/test_vba_matrix_hard_forms.py` — small-fixture pattern proven for AssociationPairs (4×5) and GroupData (person_id=1)
- `tests/test_vba_networks_small_fixture.py` — minimal-injection pattern proven for Networks (Cao Zhi, c_personid=30270)
- Per-test skip-reason strings in the four cross-form export tests

**No Access COM run.  No tests changed.  No reports/issues/README touched.**

This document does NOT prescribe writing tests.  It triages the 13
gaps so the next implementation PR has a defensible scope and the
high-time-cost areas (Networks/AssociationPairs full-injection +
matrix CmdQuery family) are NOT autopilot-implemented.

---

## Refresh 2026-05-05 — what changed

After PR `cover/groupdata-cmdgis-clean-branches` (commit `294cbda`)
landed, inventory dropped from 13 → 12 gaps and the only
bucket-A cell (GroupData × CmdGIS) became `covered`.

**Bucket distribution of remaining 12 gaps:**

| Bucket | Count | Cells |
|---|---:|---|
| A small_candidate | **0** (was 1) | — |
| B blocked_by_known_driver_issue | 5 | AssociationPairs CmdGIS / CmdPajek / CmdGephi / CmdNeo4j; Networks CmdGIS / CmdPajek (split below) |
| C blocked_by_form_query_timeout | 1 (was 2) | GroupData × CmdNeo4j |
| D new_export_family_needs_design | 3 | Associations / Place / Kinship × CmdUCINet |
| D + B stacked | 2 | AssociationPairs × CmdUCINet; Networks × CmdUCINet |

(B-bucket count is 5 across two driver causes: 4 cells gated on
the AssociationPairs SetFocus driver patch, 2 cells gated on
Networks Form_Open landmine #3.5.  CmdNeo4j AssociationPairs sits
in B; the 4 vs 2 above sums to 6 because Networks CmdUCINet is
counted under "D + B stacked", not in the pure-B row.)

**Next cheapest 1-3 cells, after applying the refresh-brief
exclusions** (no AssociationPairs · no driver/meta-PR-needed · no
CmdUCINet new family):

| Rank | Cell | Why it is cheaper than every other remaining cell |
|---|---|---|
| 1 | ~~**LookAtGroupData × CmdNeo4j**~~ — **SUPERSEDED** by probe results, see "Refresh 2026-05-05 (later)" below | (original rationale, kept for history): The just-merged PR proved the person_1 small fixture drives GroupData CmdRun + a multi-file export chain end-to-end inside the standard 180s watcher.  That dispels the "matrix CmdRun timeout" half of this cell's bucket-C blocker.  The all-`Chk*`-reset pattern from the GroupData CmdGIS test directly transfers (Issue #6 avoidance).  Only one unknown remains: whether the CmdNeo4j chain's own SaveAs count fits the 180s watcher on person_1 (cheaper to settle than any other remaining cell, all of which need either a driver patch or a new export-family design pass first). |

> **Update (post-probe, 2026-05-05 later):** the rank-1
> candidate above was probed (`probe/groupdata-cmdneo4j` and
> `investigate/groupdata-cmdneo4j-tail`, both merged to main).
> The probe found mid-chain `LookAtGroupData:ERR No current
> record.` — confirmed bug-candidate class (unguarded
> `.MoveFirst` on empty recordset; see the tail probe for
> the per-block isolation evidence).  Filed as **canonical
> Issue #21 (P1)** in `reports/generate_report.py::ISSUES`
> via PR `chore/file-issue-21-v2` (commit `bc85092`, merged
> to main).  GroupData × CmdNeo4j is now firmly on the
> issue / investigation line, NOT a coverage candidate.
> See the "Refresh 2026-05-05 (later)" section below for
> the current ranking (which is empty).

There is no rank 2 or rank 3 under the brief's exclusions —
**every other gap requires a driver patch, the Networks Form_Open
scaffold, or a CmdUCINet family-level design + probe pass**, none of
which is cheap.  The refresh deliberately reports a single
candidate rather than padding the list.

**Required prerequisite before opening rank-1 PR:** a probe
(read-only, similar shape to `analysis/probe_groupdata_cmdgis.py`)
that drives CmdRun + CmdNeo4j on person_1, times the chain, and
counts files.  If chain runtime ≤ 120s and ≥ 1 file is produced,
promote GroupData × CmdNeo4j from bucket C → bucket A and open the
coverage PR.  If chain trips the watcher OR mid-chain `:ERR`
appears, do NOT open the coverage PR — instead investigate first
(this is the AssociationPairs lesson: don't trust passing
infrastructure-adjacent tests as evidence the export chain works).

**Cells the brief explicitly excludes from "next cheapest"
ranking** (do not touch in the next 1-3 PRs without a fresh
maintainer brief):

| Cell | Exclusion reason |
|---|---|
| LookAtAssociationPairs × CmdGIS / CmdNeo4j / CmdPajek / CmdGephi | AssociationPairs (per refresh brief); also gated on CmdQuery SetFocus driver patch |
| LookAtAssociationPairs × CmdUCINet | AssociationPairs + CmdUCINet new family (per refresh brief) |
| LookAtNetworks × CmdGIS / CmdPajek | Driver/meta-PR (Form_Open landmine #3.5) — per refresh brief |
| LookAtNetworks × CmdUCINet | Driver/meta-PR + CmdUCINet new family — per refresh brief |
| LookAtAssociations × CmdUCINet | CmdUCINet new family — per refresh brief |
| LookAtPlace × CmdUCINet | CmdUCINet new family — per refresh brief |
| LookAtKinship × CmdUCINet | CmdUCINet new family — per refresh brief |

11 of the 12 remaining gaps fall into the do-not-touch table; only
GroupData × CmdNeo4j is open for consideration, and only after a
read-only probe.

---

## Refresh 2026-05-05 (later) — AssociationPairs × CmdGIS confirmed NOT cheap

Sequel mini-refresh after the SetFocus driver patch
(`feat/driver-patch-associationpairs-setfocus`, commit `89b46a9`)
and the AssociationPairs × CmdPajek + CmdGephi coverage PR
(`cover/assocpairs-pajek-gephi-1x3`, commit `4b8a927`) both
landed.  The natural next-cheapest cell appeared to be
`LookAtAssociationPairs × CmdGIS`, but a focused attempt
(branch `cover/assocpairs-cmdgis-1x3`, deleted; no commits)
confirmed the cell is **NOT** in the same equivalence class as
LookAtPlace / LookAtKinship for the existing
`_SUBFORMS_TO_REQUERY` driver mechanism.

**Negative finding** documented in
[`analysis/assocpairs_cmdgis_note.md`](./assocpairs_cmdgis_note.md):
the 1-line `_SUBFORMS_TO_REQUERY` extension fails because
AssociationPairs's CmdQuery cleanup (line 2000) opens a fresh
`dbOpenDynaset` whose `RecordCount` returns 0 until visited,
and the existing dict-mechanism's `.Form.Requery` after that
rebind likely *invalidates* the fresh recordset (per the
pre-existing Status warning comment in the dict).  A more
aggressive shim (e.g. per-form `.MoveLast` action) was outside
the brief's authorized scope; reviewer chose option (c) accept
& document.

**Triage adjustment:**

- AssociationPairs × CmdGIS stays in bucket B
  (`blocked_by_known_driver_issue`) — already classified there
  per the original 2026-05-04 SetFocus reclassification.  This
  mini-refresh just confirms that the SetFocus driver patch
  alone (now merged) does NOT lift this cell out of bucket B;
  it has a second, independent stale-subform-RecordCount
  blocker that needs its own driver/meta investigation.
- **Removed from "cheapest next cell" candidates.**  After the
  AssocPairs SetFocus driver patch + Pajek/Gephi coverage,
  this cell briefly looked like the next cheap win.  It isn't.
  Future work needs either (a) per-form-action expansion of
  `_SUBFORMS_TO_REQUERY` (e.g. `MoveLast` after `Requery` for
  AssocPairs only) or (b) a deeper investigation into why
  AssocPairs's `Set ... = OpenRecordset(...)` rebind doesn't
  populate the subform's RecordCount the way Place / Kinship's
  do.  Neither is justified on cost-benefit today.

**Cheapest-next ranking after this refresh: EMPTY.**

The earlier 2026-05-05 refresh (above) named LookAtGroupData ×
CmdNeo4j as the rank-1 cheapest-next candidate.  That is now
**SUPERSEDED** by the actual probe results that have since
landed on main:

- `probe/groupdata-cmdneo4j` (commit `4ace85b`) — chain runs
  in 2.5 s and produces 8 of 11 expected files, but emits
  mid-chain `LookAtGroupData:ERR No current record.` (DAO
  3021) before the Entry-related tail blocks fire.  Verdict:
  **`do_not_open_coverage_pr_mid_chain_err`**.
- `investigate/groupdata-cmdneo4j-tail` (commit `3bfcba8`) —
  per-block isolation probe confirmed the bug class:
  **`A_new_bug_candidate_empty_recordset_guard`** (PeopleEntry
  / EntryCode blocks unguarded `.MoveFirst` on empty
  `ZZ_SCRATCH_ENTRY`; distinct from Issue #6's column-typo
  family).  Filed as **canonical Issue #21 (P1)** in
  `reports/generate_report.py::ISSUES` via PR
  `chore/file-issue-21-v2` (commit `bc85092`, merged to
  main).  Both source-side static marker
  (`tests/test_known_bugs.py::test_bug21_groupdata_cmdneo4j
  _missing_eof_guard`) and runtime behavioural pin
  (`tests/test_vba_bug_behaviors.py::test_bug21_lookat
  _groupdata_cmdneo4j_fires_no_current_record`) are in
  place.

So GroupData × CmdNeo4j has moved from "probe-first cheapest-
next coverage candidate" to **"investigation-first / bug-
candidate confirmed; not a coverage candidate"**.  It belongs
in the same not-cheap class as AssociationPairs × CmdGIS — both
have a confirmed downstream blocker that a coverage PR alone
won't address.

**Therefore the cheapest-next list is now empty.**  Every
remaining gap is one of:

- AssociationPairs × CmdGIS — bucket B + this refresh's not-
  cheap negative finding
- AssociationPairs × CmdNeo4j / CmdUCINet — bucket B + brief
  exclusions
- Networks × {CmdGIS, CmdPajek, CmdUCINet} — driver/meta-PR
- {Associations, Place, Kinship} × CmdUCINet — CmdUCINet new
  family
- GroupData × CmdNeo4j — investigation-first per the tail
  probe's bug-candidate finding (this refresh)

None of these is autopilot-safe under the standing brief
(no AssocPairs, no Networks driver-meta, no CmdUCINet new
family, no investigation-first cells without a fresh brief).

**Successor recommendation (revised):** rather than picking
a "next cell" by ranking, the natural next step is one of:

1. **GroupData CmdNeo4j tail / empty-recordset-guard
   follow-up** — Issue #21 is now canonical
   (`reports/generate_report.py::ISSUES`, merged via
   `chore/file-issue-21-v2` commit `bc85092`).  Next step
   is either coordinating an upstream CBDB fix (per the
   canonical issue's `fix_en` recommendation: guard the
   `.MoveFirst` in blocks #9 and #10), OR writing a
   per-block bugfix verification probe that flips the
   existing test_bug21 markers when the upstream fix
   lands.
2. **A fresh whole-triage refresh** — re-baseline the export-
   gap queue from scratch given that two cells (AssocPairs ×
   CmdGIS and GroupData × CmdNeo4j) have moved out of the
   cheap-next zone since the original 2026-05-04 plan.

What this refresh deliberately does NOT recommend:
"continue closing the next cheapest cell" — there is no next
cheapest cell on this dump, under the standing brief, today.

**AssociationPairs line of work — final state per this
refresh:**

| PR | Status | Outcome |
|---|---|---|
| `feat/driver-patch-associationpairs-setfocus` (`89b46a9`) | merged | SetFocus driver patch unblocks CmdQuery body |
| `cover/assocpairs-pajek-gephi-1x3` (`4b8a927`) | merged | 2 cells closed: CmdPajek + CmdGephi |
| `cover/assocpairs-cmdgis-1x3` | abandoned (no commits) | CmdGIS proven NOT cheap; see `analysis/assocpairs_cmdgis_note.md` |
| this refresh | in flight | reclassifies CmdGIS as not-cheap; removes from candidate list |

The line has produced its high-value PRs.  Successor
recommendation: change direction (a fresh export-gap triage
refresh after this lands), not deeper into AssociationPairs.

---

## Bucket taxonomy

| Bucket | Meaning |
|---|---|
| **A. small_candidate** | Existing small-fixture pattern (matrix_hard_forms / networks_small_fixture) covers this form's CmdQuery, AND the export's per-block logic is well-understood enough that a new cross-form test slice should pass with low driver risk.  Recommended for next 1-3 PRs. |
| **B. blocked_by_known_driver_issue** | Form_Open / VBE-injection blocker (Networks AGENTS landmine #3.5) — CmdQuery itself works only via the `skip_inject_autodetect_forms` minimal-injection scaffold.  Standard cross-form test infra (full injection) cannot host it.  Needs a separate test file or a driver-level option pass-through. |
| **C. blocked_by_form_query_timeout** | Matrix CmdQuery/CmdRun on the default heavy fixture times out within `click_via_timer`'s 180s watcher.  Could in principle become **A. small_candidate** by reusing the small fixture pattern (4×5 / person_1) — but the export chain's own runtime might still blow past the timeout.  Investigate before promising. |
| **D. new_export_family_needs_design** | Handler family has zero existing tests; classifier knows none of its file shapes; export semantics need a probe pass before any assertion strictness can be defined.  CmdUCINet is the only family in this bucket today. |
| **E. low_value_or_duplicate** | Cell is meaningfully covered by another form/family already, and the marginal value of adding it is low.  None today. |

---

## Per-cell triage (13 cells)

### LookAtAssociations × CmdUCINet — bucket D

- **handler exists?** yes
- **button exists?** yes
- **current blocker:** CmdUCINet handler family has zero existing tests anywhere; the depth-check classifier has zero `_NEO4J_SHAPES` entries that match a UCINet output (UCINet `.dl` format is unrelated to People/Place/EntryCode shapes the existing classifier knows).
- **nearest existing test pattern:** `tests/test_vba_pajek_gephi_cross_form.py` is the closest (different file format per export, single-file output assertion) — but it doesn't help with .dl-specific structure (UCINet has its own file syntax).
- **recommended next action:** **NOT this PR.**  Needs a separate "design + probe" PR first: dump 1-2 sample CmdUCINet outputs by driving the handler against an Associations small fixture, identify the output file format (`.dl` vs another), define what "passes" means (well-formed `.dl` header? row count > 0?  shape-specific anchors?), THEN write the test.
- **risk:** **medium** — handler logic is standard CBDB style (likely reuses ZZ_SCRATCH_*); the unknown is the file-format depth check.

### LookAtPlace × CmdUCINet — bucket D

- **handler exists?** yes
- **button exists?** yes
- **current blocker:** same family as Associations.CmdUCINet — no test infrastructure for CmdUCINet anywhere.
- **nearest existing test pattern:** none for CmdUCINet; closest non-UCINet equivalent is `tests/test_vba_pajek_gephi_cross_form.py` for LookAtPlace × CmdPajek.
- **recommended next action:** **NOT this PR.**  Same as Associations.CmdUCINet — needs the family-level design pass first.  Once CmdUCINet test infra exists, Place is one of three forms a cross-form CmdUCINet test would cover.
- **risk:** **medium**.

### LookAtKinship × CmdUCINet — bucket D

- **handler exists?** yes
- **button exists?** yes
- **current blocker:** same CmdUCINet family blocker.
- **nearest existing test pattern:** none for CmdUCINet; closest for Kinship is `tests/test_vba_cmdguess_cross_form.py` (CmdGUESS) and `tests/test_vba_pajek_gephi_cross_form.py` (CmdPajek).
- **recommended next action:** **NOT this PR.**  Same CmdUCINet design dependency.
- **risk:** **medium**.

### LookAtAssociationPairs × CmdGIS — **RE-CLASSIFIED 2026-05-04 → bucket B (blocked_by_known_driver_issue)**

**Original triage:** bucket A (small_candidate, low risk).
**Reason for re-classification:** same CmdQuery SetFocus blocker
that was surfaced for AssociationPairs CmdPajek/Gephi (see those
entries below).  AssociationPairs CmdGIS is downstream of the
same CmdQuery (which seeds `ZZ_SOCIAL_NETWORK`); if CmdQuery's
INSERTs never run, CmdGIS has nothing to export.

(Investigation didn't directly probe CmdGIS, but the blocker is
upstream of the export — there's no path that bypasses CmdQuery's
SetFocus error.)

- **handler exists?** yes
- **button exists?** yes
- **actual blocker:** upstream CmdQuery_Click `Me.CmdQuery.SetFocus`
  fails under Form_Timer dispatch (see CmdPajek/Gephi entries
  below for details).  CmdGIS export has no source data.
- **nearest existing test pattern:** `tests/test_vba_cmdgis_other_forms.py` (5 forms × CmdGIS, single-file output, structural assertion).  Once the driver-side CmdQuery patch lands, this cell drops back to bucket A.
- **recommended next action:** wait for the AssociationPairs CmdQuery driver patch (same blocker as Pajek/Gephi).
- **risk:** **medium** — driver-level fix needed.

### LookAtAssociationPairs × CmdNeo4j — bucket C (blocked_by_form_query_timeout)

- **handler exists?** yes
- **button exists?** yes
- **current blocker:** declared in `tests/test_vba_cmdneo4j_cross_form.py::_spec_skip_marks`-style intent but actually missing from `_SPECS`; the underlying matrix-CmdQuery timeout would block it just like it does CmdGIS, but Neo4j's multi-file chain (6-10 SaveAs blocks) on a 4×5 fixture might still complete OR might blow past 180s — this is the unknown.
- **nearest existing test pattern:** `tests/test_vba_cmdneo4j_cross_form.py` is the cross-form host; LookAtAssociations is the closest sibling and is currently *skipped* with `"produces 0 files in directory mode — needs investigation alongside Place"` (NOT a timeout — a different kind of failure).
- **recommended next action:** **NOT in the next 1-3 PRs.**  Two unknowns stack here (matrix timeout + 0-file Associations sibling failure).  Probe first to confirm small-fixture chain runtime AND whether the AssociationPairs-specific Neo4j has the same 0-file mode as Associations.  If both probes are clean, promote to bucket A.  If not, open a separate investigation PR.
- **risk:** **medium-high** — cumulative unknowns.

### LookAtAssociationPairs × CmdPajek — **RE-CLASSIFIED 2026-05-04 → bucket B (blocked_by_known_driver_issue)**

**Original triage (PR plan/export-gap-triage):** bucket A (small_candidate, low risk).
**Investigation outcome (PR cover/assocpairs-pajek-gephi probe):** the
small-fixture wiring works mechanically, BUT
`Form_LookAtAssociationPairs.CmdQuery_Click` line 1635 calls
`Me.CmdQuery.SetFocus`, which fails under Form_Timer dispatch with
`Welcome to CBDB! can't move the focus to the control CmdQuery.`
(VBA error 2110).  The active form during timer dispatch isn't
necessarily LookAtAssociationPairs (it's whatever was last
foregrounded — typically the welcome / NAVIGATION_PANE form),
so SetFocus to a control on a non-active form raises.  CmdQuery
exits via the error handler BEFORE its INSERT statements run,
leaving `ZZ_SCRATCH_PEOPLE = 0` and `ZZ_SOCIAL_NETWORK = 0`.
CmdPajek then bails on `RecordCount = 0` (per
`Form_LookAtAssociationPairs.vb:363-364`).

This is **not just the matrix CmdQuery timeout** that the original
`_xfail_marks` docstring documents — it's a deeper Form_Timer-
dispatch-vs-Access-focus-model interaction that no other LookAt
form's CmdQuery hits because no other form's CmdQuery_Click body
calls `Me.<button>.SetFocus`.

**Why matrix_hard_forms's `assocpair_4x5_small` PASSES today:** that
test's `_check_assoc_pairs` deliberately doesn't assert on row
count (`# Don't assert >= 1 — small fixtures may legitimately have
0 network edges`), so the SetFocus failure is silently swallowed.
matrix_hard_forms reports `ZZ_SOCIAL_NETWORK rows: 0` and passes —
which the gap-triage incorrectly read as "small fixture works".

- **handler exists?** yes
- **button exists?** yes
- **actual blocker:** `Me.CmdQuery.SetFocus` inside CmdQuery_Click
  fails under Form_Timer dispatch.  Needs driver-level inline VBA
  patch (à la `_PER_FORM_CMDGIS_PATCHES`) to comment out the
  SetFocus line, OR a different dispatch mechanism that ensures
  the form has focus.  Both are out of scope for "no driver
  changes".
- **nearest existing test pattern:** none, given the SetFocus
  bug.  (`tests/test_vba_pajek_gephi_cross_form.py::Case
  ("LookAtAssociations", "CmdPajek", ...)` works because
  Associations' CmdQuery_Click does not call `Me.CmdQuery.SetFocus`.)
- **recommended next action:** **NOT a small-PR candidate.**
  Open a driver-side investigation PR to either (a) add a
  per-form CmdQuery patch that strips `Me.<button>.SetFocus`
  lines from AssociationPairs, OR (b) extend the Form_Timer
  dispatcher to first activate the form via DoCmd / similar.
  Once that lands, this cell drops back to bucket A and is
  trivial to close.
- **risk:** **medium** — the underlying bug is well-understood;
  fix is mechanical at the driver level but isn't appropriate
  for autopilot.

### LookAtAssociationPairs × CmdGephi — **RE-CLASSIFIED 2026-05-04 → bucket B (blocked_by_known_driver_issue)**

- **handler exists?** yes
- **button exists?** yes
- **actual blocker:** identical to CmdPajek above — same
  CmdQuery SetFocus bug blocks any export downstream of it.
  `Form_LookAtAssociationPairs.vb:113-119` shows CmdGephi has
  the same `RecordCount = 0` early-bail.
- **recommended next action:** same as CmdPajek — wait for
  the driver-side fix, then both cells close together.
- **risk:** **medium**.

### LookAtAssociationPairs × CmdUCINet — bucket D (CmdUCINet family) + C (matrix-blocked)

- **handler exists?** yes
- **button exists?** yes
- **current blocker:** stacks both CmdUCINet family blocker AND matrix-CmdQuery timeout.
- **recommended next action:** **NOT this PR.**  Two stacked blockers.  Resolve CmdUCINet family first, AssociationPairs is then the highest-risk member of that family (matrix-blocker on top).
- **risk:** **high** — combined blockers.

### LookAtNetworks × CmdGIS — bucket B (blocked_by_known_driver_issue)

- **handler exists?** yes
- **button exists?** yes
- **current blocker:** Networks `Form_Open` deadlock under default full injection (AGENTS landmine #3.5).  Cross-form tests use full-injection `VbaSession` from `tests/test_vba_cmdneo4j_cross_form.py`-style infrastructure; Networks needs the `skip_inject_autodetect_forms=SKIP_SIBLINGS` minimal-injection mode used by `tests/test_vba_networks_small_fixture.py`.  AND the matrix CmdRun on Zhu Xi (default heavy fixture) times out — needs Cao Zhi small fixture.
- **nearest existing test pattern:** `tests/test_vba_networks_small_fixture.py` uses minimal injection successfully on Cao Zhi for `CmdRun` only — has not extended to CmdGIS/CmdPajek.
- **recommended next action:** **NOT in the next 1-3 PRs.**  Needs either (a) extending the existing `tests/test_vba_networks_small_fixture.py` with a CmdGIS slice, OR (b) refactoring the cross-form CmdGIS test infrastructure to thread `skip_inject_autodetect_forms` through.  Both are scope-defining design work; neither is mechanical.
- **risk:** **medium** — Networks driver risk is well-understood; just costly per landmine #3.5.

### LookAtNetworks × CmdPajek — bucket B

- **handler exists?** yes
- **button exists?** yes
- **current blocker:** same Networks Form_Open blocker as CmdGIS.  Note: CmdPajek hosts list at `test_vba_pajek_gephi_cross_form.py:9` includes Networks, but it's NOT in `_CASES` (the test author omitted it precisely because of the Form_Open blocker).
- **recommended next action:** **NOT in the next 1-3 PRs.**  Same Networks driver-pattern dependency.
- **risk:** **medium**.

### LookAtNetworks × CmdUCINet — bucket B + D

- **handler exists?** yes
- **button exists?** yes
- **current blocker:** Networks Form_Open AND CmdUCINet family.  Stacked.
- **recommended next action:** **NOT in the next 1-3 PRs.**  Resolve CmdUCINet design AND Networks driver pattern first.
- **risk:** **high** — combined.

### LookAtGroupData × CmdGIS — **COVERED 2026-05-05** (was bucket A small_candidate)

- **status:** **covered** by `tests/test_vba_cmdgis_other_forms.py::test_cmd_gis_groupdata_clean_branches` (PR `cover/groupdata-cmdgis-clean-branches`, commit `294cbda`).  Inventory `gap: 13 → 12`, `real_vba_covered: 16 → 17`.
- **scope:** clean branches only — Status / Office / Addr.  Entry deliberately excluded (Issue #6 separately bug-pinned in `tests/test_vba_bug_behaviors.py::test_bug6_lookat_groupdata_query_entry_fires_no_such_field`).  Text excluded (benign 0-files on person_1).
- **lesson for refresh:** the all-`Chk*`-reset pattern (reset all 11 checkboxes to False BEFORE setting target ones) is now baked in — first-attempt failure surfaced because Form_Open defaults left Issue #6's branch enabled.  This pattern is directly transferable to a future GroupData × CmdNeo4j probe.

### LookAtGroupData × CmdNeo4j — bucket C (blocked_by_form_query_timeout)

- **handler exists?** yes
- **button exists?** yes
- **current blocker:** matrix CmdRun timeout AND multi-file Neo4j chain heaviness on top.  CmdNeo4j cross-form test's hardcoded 180s watcher timeout (already on the edge for Office's 37k rows) might still trip on GroupData even with person_1 small fixture if the chain has many SaveAs blocks.
- **nearest existing test pattern:** `tests/test_vba_cmdneo4j_cross_form.py` cross-form host (currently has GroupData absent from `_SPECS`).
- **recommended next action:** **NOT in the next 1-3 PRs** unless GroupData CmdGIS (bucket A above) lands first AND a probe confirms the Neo4j chain completes on person_1 within the watcher timeout.  Sequence-dependent.
- **risk:** **medium** — depends on probe outcome.

---

## Summary

**Updated 2026-05-04** after PR cover/assocpairs-pajek-gephi probe
re-classified 3 AssociationPairs cells from bucket A to bucket B
(driver-level CmdQuery SetFocus blocker — see those entries above
for full context).

| Cell | Bucket | Risk |
|---|---|---|
| LookAtAssociationPairs × CmdGIS | ~~A small_candidate~~ → **B blocked_by_known_driver_issue** | medium |
| LookAtAssociationPairs × CmdPajek | ~~A small_candidate~~ → **B blocked_by_known_driver_issue** | medium |
| LookAtAssociationPairs × CmdGephi | ~~A small_candidate~~ → **B blocked_by_known_driver_issue** | medium |
| LookAtGroupData × CmdGIS | ~~A small_candidate~~ → **COVERED 2026-05-05** | — |
| LookAtAssociationPairs × CmdNeo4j | C blocked_by_form_query_timeout | medium-high |
| LookAtGroupData × CmdNeo4j | C blocked_by_form_query_timeout | medium |
| LookAtNetworks × CmdGIS | B blocked_by_known_driver_issue | medium |
| LookAtNetworks × CmdPajek | B blocked_by_known_driver_issue | medium |
| LookAtAssociations × CmdUCINet | D new_export_family_needs_design | medium |
| LookAtPlace × CmdUCINet | D new_export_family_needs_design | medium |
| LookAtKinship × CmdUCINet | D new_export_family_needs_design | medium |
| LookAtAssociationPairs × CmdUCINet | D + B stacked | high |
| LookAtNetworks × CmdUCINet | D + B stacked | high |

| Bucket (12 remaining gaps after 2026-05-05) | Count |
|---|---:|
| A small_candidate | **0** (was 1 on 2026-05-04; GroupData × CmdGIS now covered) |
| B blocked_by_known_driver_issue | 5 |
| C blocked_by_form_query_timeout | 2 |
| D new_export_family_needs_design | 3 |
| D + B stacked | 2 |

(Counts above use the post-2026-05-04 reclassifications.  See the
JSON for canonical bucket-per-cell assignments — the per-cell rows
in this MD's table above predate the JSON's CmdNeo4j AssocPairs
re-class and are kept as historical record.)

(LookAtAssociationPairs × CmdUCINet was previously stacked
D + C; the underlying matrix-CmdQuery blocker is actually the
same SetFocus driver issue as the Pajek/Gephi cells, so its
inner blocker also moves from C → B.  Stack is now D + B.)

---

## Recommended next 1-3 PRs (ranked by risk + ROI)

### ~~PR 1 — AssociationPairs × CmdPajek + CmdGephi via small 4×5 fixture~~ — **WITHDRAWN 2026-05-04**

**Original:** bucket A (×2 cells closed), risk low.
**Withdrawn because:** PR cover/assocpairs-pajek-gephi probed
this and surfaced a deeper blocker — `Form_LookAt
AssociationPairs.CmdQuery_Click:1635` calls
`Me.CmdQuery.SetFocus`, which fails under Form_Timer dispatch
because the active form is the welcome / NAVIGATION_PANE form,
not LookAtAssociationPairs.  CmdQuery exits via the error
handler before its INSERT statements run, so
`ZZ_SOCIAL_NETWORK = 0`, and CmdPajek/CmdGephi bail on
RecordCount=0.  matrix_hard_forms's existing
`assocpair_4x5_small` test passes only because its
`_check_assoc_pairs` doesn't assert row count — it silently
swallows the SetFocus failure.  The triage plan was based on
that misleading "passing" test.

**Re-classification:** all three AssociationPairs export cells
(CmdGIS, CmdPajek, CmdGephi) are now bucket B (driver-level
blocker).  See per-cell entries above.

### ~~PR 2 — AssociationPairs × CmdGIS via the same 4×5 fixture~~ — **WITHDRAWN 2026-05-04**

Same blocker as PR 1 — withdrawn for the same reason.

### ~~PR 3 (now PR 1) — GroupData × CmdGIS via person_1 small fixture~~ — **LANDED 2026-05-05**

Implemented in PR `cover/groupdata-cmdgis-clean-branches`, commit
`294cbda`.  The probe-first caveat from this PR's recommendation
proved correct: a probe (`analysis/probe_groupdata_cmdgis.py` and
the sub-call follow-up) found mid-chain `:ERR` from Issue #6's
queryEntry sub.  The shipped coverage PR therefore covered Status
/ Office / Addr only and explicitly excluded Entry (separately
bug-pinned) and Text (benign 0-files on person_1).  Inventory
delta: `gap: 13 → 12`, `real_vba_covered: 16 → 17`.

### PR 1 (post-refresh) — GroupData × CmdNeo4j probe-first

- **Bucket:** C → maybe A after probe
- **Risk:** medium until probed
- **Why this is the cheapest remaining cell:** see the "Refresh
  2026-05-05" section near the top of this document.  Person_1
  fixture is now end-to-end-proven for GroupData CmdRun + a multi-
  file export chain, dispelling the matrix-CmdRun-timeout half of
  the original blocker.  All other remaining gaps require either a
  driver patch, the Networks Form_Open scaffold, or a CmdUCINet
  family design pass first.
- **Required prerequisite:** read-only probe — drive CmdRun +
  CmdNeo4j on person_1, time the chain, count files, snapshot
  ZZ_TEST_DEBUG.  Reuse the all-`Chk*`-reset pattern baked into
  `test_cmd_gis_groupdata_clean_branches`.
- **Promote-to-coverage condition:** chain runtime ≤ 120s AND ≥ 1
  file produced AND no `LookAtGroupData:ERR` mid-chain.
- **Reject-to-investigation condition:** chain trips watcher OR
  mid-chain `:ERR` appears.  In that case, do NOT open the
  coverage PR — open an investigation PR instead (the
  AssociationPairs lesson).

### Driver-side meta-PR (would unblock 3-5 cells at once)

A driver-side PR adding a per-form CmdQuery patch (à la
`_PER_FORM_CMDGIS_PATCHES`) that strips `Me.<button>.SetFocus`
lines from `Form_LookAtAssociationPairs` would unblock all 3
AssociationPairs export cells (CmdGIS / CmdPajek / CmdGephi)
at once, AND make the bucket-D-stacked AssociationPairs
CmdUCINet a normal D-only blocker.  Net unblock: 3-4 cells.
Risk: medium (driver patch surface area; need to confirm no
other forms regress).  Out of scope for the current PR
sequence but a high-leverage future direction.

---

## Explicitly NOT recommended for autopilot implementation

The following cells should **NOT** be picked up by an implementer without first opening a scope-defining brief from the maintainer:

| Cell | Why NOT autopilot |
|---|---|
| LookAtAssociationPairs × CmdGIS / CmdPajek / CmdGephi | Driver-side CmdQuery SetFocus blocker (re-classified 2026-05-04 from bucket A → B).  Needs the meta-PR above. |
| LookAtAssociationPairs × CmdNeo4j | Same SetFocus blocker upstream + multi-file chain risk if/when CmdQuery is unblocked. |
| LookAtGroupData × CmdNeo4j | Multi-file Neo4j chain on heavy form; watcher-timeout risk.  Sequence after GroupData CmdGIS. |
| LookAtNetworks × CmdGIS / CmdPajek | Form_Open landmine #3.5 — needs minimal-injection scaffolding either via extending `tests/test_vba_networks_small_fixture.py` or refactoring cross-form test driver.  Either path is design work, not mechanical. |
| LookAtAssociations × CmdUCINet | New export family — needs design + probe pass to identify file format and assertion strictness BEFORE writing the test. |
| LookAtPlace × CmdUCINet | Same family blocker. |
| LookAtKinship × CmdUCINet | Same family blocker. |
| LookAtAssociationPairs × CmdUCINet | Stacked blockers (CmdUCINet + the SetFocus driver issue). |
| LookAtNetworks × CmdUCINet | Stacked blockers (CmdUCINet + Networks Form_Open). |

---

## What this PR's investigation actually delivered

This PR (`cover/assocpairs-pajek-gephi`) was opened as a
mechanical implementation of the triage's PR 1.  When the wiring
was done and the test was run, it surfaced a previously-hidden
driver-level blocker that affects ALL AssociationPairs export
cells.  The fix to the test/inventory was reverted; the value
delivered is the **re-triage**: 3 cells move from bucket A to
bucket B, the recommended-PR list shrinks from 3 to 1 (and that
1 needs a probe), and a meta-PR direction is identified that
would unblock 3-4 cells in a single driver-side change.

This is a useful negative result — better that the gap-triage
plan be honest about the actual blockers than that a
later implementer re-discover the SetFocus issue after spending
hours wiring tests.

The 4 bucket-A small_candidates above are the only cells worth promoting in narrow PRs without first investing in design/probe work.  The rest will close eventually but each requires a PR brief from the maintainer that names the specific blocker being lifted.

---

## Refresh 2026-05-06 — post Issue #22 + Kinship sibling-form alignment

Triggering event: `chore/align-issue22-kinship-sibling-confirmed`
merged into main (current `14099a7`). The CmdUCINet family's
3-cell truth has stabilized; this refresh re-baselines the queue
on top of that truth. **Read-only analysis. No COM. No tests /
driver / README / canonical-report changes in this PR.**

### CmdUCINet family — current state on main (14099a7)

| Cell | Status | Truth label (consistent across canonical + inventory + README + tests) |
|---|---|---|
| `LookAtAssociations × CmdUCINet` | **gap** | canonical **Issue #22** (P1_visible_crash); FSO ANSI cp1252 cannot encode CJK — runtime-pinned, static-marked. NOT a coverage candidate. |
| `LookAtKinship × CmdUCINet` | **covered (fixture-fragile)** | runtime-confirmed sibling form of Issue #22 (probe `154bb4b`, picker pid 152930 → kin pid 140733 He Mou 取 U+53D6); test passes only on the matrix `kinship_person_3211` fixture. NOT downgraded to gap; the fragility IS the truth. |
| `LookAtPlace × CmdUCINet` | **gap (uninvestigated, different mechanism)** | not yet probed; uses ADO Stream not FSO (see static signal below). Explicitly outside Issue #22's scope. |

### Place × CmdUCINet — read-only static signal (this refresh's analysis)

Static read of `analysis/dump/vba/Form_LookAtPlace.vb`
inside `Sub CmdUCINet_Click`:

```
tStream.Charset = "utf-8"   ' default; alt branches: "big5", "gb18030"
tStream.Type    = adTypeText
tStream.Open
'Set tVNA = tFileSystem.CreateTextFile(tFileName, True)   ' ← FSO path COMMENTED OUT
```

Implication: **Place's CmdUCINet write path is Unicode-by-design
via ADO Stream**; the FSO `CreateTextFile` path that drives Issue
#22 in Associations + Kinship is commented out in Place. So
Place CmdUCINet does NOT structurally share Issue #22's bug class.
A probe is therefore expected to either (a) succeed cleanly on a
Han-name fixture, or (b) reveal a *different* failure mode
(charset encoder mismatch, recordset-empty guard, or ADO-specific
issue).

### Next 1–3 work items, ranked — by category

#### Rank 1 — probe-first investigation: `LookAtPlace × CmdUCINet`

- **Category:** probe-first investigation (NOT coverage PR, NOT
  canonical issue).
- **Why now:** It is the cheapest next probe in the family. The
  static signal above is favorable enough that the probe is mostly
  about *characterizing the file shape* (sections, columns, row
  counts, encoding/BOM), not about reproducing Issue #22. Closing
  this probe gives the queue a clean fork: either a coverage PR
  candidate, a new canonical issue, or a deferred gap with a
  recorded reason.
- **Shape:** one tiny fixture (a place network whose member
  recordset includes ≥1 CJK person), a Form_Open + CmdRun + CmdUCINet
  split-fire, a file-poll for the `.vna` output, and a strict
  file-shape classifier. Probe MD + JSON only — no `test_*`
  additions.
- **Risk:** low; Place's writer is structurally distinct from the
  failing FSO writer.

#### Rank 2 — canonical issue / maintainer-line: Issue #22 upstream-fix coordination

- **Category:** maintainer line (outside this repo's PR surface).
- **Why second:** Issue #22 is filed P1 and has a recommended fix
  (`CreateTextFile(..., True, True)` for the 3rd Unicode arg, with
  the downstream-acceptance hedge). The next forward step is
  reaching the CBDB maintainer to land that fix on the upstream
  `.mdb`, not more local activity here.
- **Shape:** out-of-band; not a PR action.

#### Rank 3 — deferred coverage hardening: Kinship CmdUCINet runtime pin

- **Category:** coverage hardening (NOT new coverage).
- **Why third:** the previous Issue #22 alignment PR deliberately
  deferred adding a runtime `:ERR Invalid procedure call or
  argument` pin in `tests/test_vba_bug_behaviors.py` for Kinship
  (the analog of the Associations one). It would tighten evidence
  but does not change any cell's coverage state.
- **Shape:** small; mirrors the existing Associations runtime pin.

### Place × CmdUCINet — direct answers to the brief

**Q1: Is `LookAtPlace × CmdUCINet` the cheapest next probe?**

**Yes**, conditional on staying on the export-CmdUCINet line.
Reasons:

1. *Adjacency:* it is the last unprobed cell in a family whose
   other two cells just stabilized — context cost is at its
   minimum.
2. *Static prep already in hand:* the write mechanism is
   characterized (ADO Stream, UTF-8 default), so the probe
   does not need a separate static-analysis pass first.
3. *Bounded outcome space:* probable result is "Place is robust
   against Issue #22's bug class; characterize sections/columns";
   alternative result is a *different* failure mode that itself
   would be a clean separate finding. Either branch produces a
   small, self-contained probe artifact.
4. *No driver / fixture-design dependency:* unlike the AssocPairs
   driver-meta cells, Place CmdUCINet does not require any
   driver-side change first.

**Q2: Is anything else cheaper / better-sequenced than Place
CmdUCINet right now?**

**Realistically, no — within this repo's PR surface.** The
remaining queue is thin:

- AssocPairs × CmdGIS / CmdNeo4j / CmdUCINet — all in bucket B/C
  (driver-meta or query-timeout); not cheap.
- Networks × CmdGIS / CmdPajek / CmdUCINet — bucket B (driver-meta).
- GroupData × CmdNeo4j — Issue #21 maintainer line, same shape as
  Issue #22 maintainer line (outside this repo's PR surface).
- Associations × CmdUCINet — explicitly NOT coverage line.

If you want to *leave the CmdUCINet line entirely*, the strongest
candidate is the Issue #21 upstream-fix maintainer line (already
listed in `refresh_2026_05_05_later` as "next reasonable forward
move"), but that is also outside this repo's PR surface — not a
local-implementation alternative.

So the practical choice is: do the Place probe (Rank 1), or pause
implementation and pursue maintainer-line work (Rank 2 / Issue #21).
There is no third "cheaper local PR" that out-ranks Place.

### Explicitly NOT to touch (this refresh + onward, until briefed otherwise)

- ❌ Do NOT pursue Associations × CmdUCINet **coverage** — canonical
  Issue #22 IS the truth label for this cell; trying to wrap
  coverage around a known crashing form would either (a) require
  asserting `:ERR` (which we already do as a runtime pin in
  `test_vba_bug_behaviors.py`) or (b) silently mask the bug.
- ❌ Do NOT downgrade Kinship × CmdUCINet from `covered` back to
  `gap` — it is covered with a fixture-fragile caveat documented
  in 4 places (canonical Issue #22, manifest, README, test
  docstring). The fragility is the truth, not absence of coverage.
- ❌ Do NOT reopen Issue #22 wording — current canonical (Associations
  directly + Kinship sibling-runtime-confirmed via probe `154bb4b`)
  is the agreed truth. Issue #23 was deliberately NOT filed (sibling
  pattern was the chosen shape).
- ❌ Do NOT return to AssociationPairs × CmdGIS — canonical not-cheap
  per `analysis/assocpairs_cmdgis_note.md` and the 2026-05-04 refresh.
- ❌ Do NOT promote Place × CmdUCINet to coverage PR before the
  Rank-1 probe lands — the static UTF-8 signal is favorable but
  not a substitute for runtime evidence.

### Self-review (per `docs/skills/programmer-self-review-template.md`)

**A. Branch shape.** Branch
`refresh/queue-after-cmducinet-alignment` cut from clean main
`14099a7`. Only two files touched: this MD and the paired JSON.
No tests / driver / README / canonical-report changes — matches
brief boundary exactly.

**B. Source-of-truth sync.** MD section ↔ JSON
`refresh_2026_05_06` block carry the same 3-cell state, the same
ranked list, the same Place static finding, and the same
do-not-touch list. The 3-cell state itself was checked against
canonical truth on `main`: Issue #22 entry in
`reports/generate_report.py`, the manifest entry in
`analysis/inventory_export_coverage.py`, the static + runtime
markers in `tests/test_known_bugs.py` and
`tests/test_vba_bug_behaviors.py`, the README coverage table +
roadmap-8 lines, and the Kinship coverage test docstring. No
source-of-truth file is being changed by this refresh; the refresh
only reflects what is already canonical.

**C. Evidence vs claim.** Place's UTF-8 signal is grounded in a
verifiable static read (`grep` on `Form_LookAtPlace.vb` reproduced
in this PR's commit message) — not an inferred property. The
"probable outcome" language for the proposed Place probe is
explicitly hedged ("expected to either … or …"); no claim about
Place's runtime behavior is made beyond what static analysis
supports. The Issue #22 / Kinship / Associations state is reported
as-of `14099a7` and cited to canonical files, not paraphrased.

**D. Residual risk.** This is a triage-document refresh, not an
implementation; the residual risk is purely *advisory error*:
ranking Place above maintainer-line work could be wrong if the
user's priority is upstream coordination. Mitigated by: ranking
maintainer-line work explicitly at Rank 2 and by stating the
"leave CmdUCINet line entirely" alternative in the Q2 answer
above. No code path or test is altered, so no runtime regression
risk.

---

## Refresh 2026-05-07 — post AssociationPairs × CmdNeo4j coverage merge

Triggering events:

- PR #109 (`driver/assocpairs-cmdneo4j-msgbox-suppress`, merged
  as commit `b695a92`) — narrow scoped driver patch
  `_suppress_assocpairs_cmdneo4j_debug_msgbox` comments out the
  6 unconditional debug `MsgBox` calls in
  `Form_LookAtAssociationPairs.CmdNeo4j_Click` (lines 1069 /
  1151 / 1234 / 1317 / 1400 / 1470).
- PR #110 (`test/assocpairs-cmdneo4j-coverage`, merged as commit
  `f0e7594`) — coverage test for `LookAtAssociationPairs ×
  CmdNeo4j` on the 1×3 known-edged fixture; 3 new shape entries
  in `_NEO4J_SHAPES` (`Person1_ID` / `AssociationCode` /
  `KinshipCode`); 6-file set pinned exactly.

Read-only refresh on top of `main = f0e7594`. No COM, no tests /
driver / canonical reports / issue severity changes.

### Current truth deltas (vs `refresh_2026_05_06`)

| Cell | Before this refresh | After this refresh |
|---|---|---|
| `LookAtAssociationPairs × CmdNeo4j` | gap (bucket B; SetFocus + 6-MsgBox blockers) | **covered** — 1×3 fixture, exact 6-file set, ENTER+DONE markers, no `:ERR` (PR #110) |
| `LookAtAssociationPairs × CmdPajek` | gap (bucket B in cells-array; covered in fact) | **covered** (acknowledged — `cover/assocpairs-pajek-gephi-1x3` already on main since `4b8a927`) |
| `LookAtAssociationPairs × CmdGephi` | gap (bucket B in cells-array; covered in fact) | **covered** (acknowledged — same PR) |
| `LookAtAssociationPairs × CmdUCINet` | D + B stacked | **D-only** (CmdUCINet new family) — the inner B (SetFocus / debug-MsgBox) is now removed by the AssocPairs SetFocus patch + PR #109 driver patch; only the CmdUCINet family blocker remains |
| `LookAtPlace × CmdUCINet` | D, rank-1 probe-first candidate per `refresh_2026_05_06` | **D, paused** — investigation paused on a timeout-path-correlated COM bridge instability (separate driver/meta concern; see "do-not-touch" below) |
| `LookAtAssociations × CmdNeo4j` | bucket-? skipped sibling (open question per AssocPairs CmdNeo4j entry) | **still separate skipped / investigation line** — confirmed by PR AX Q5: 0-file mode, structurally different from AssocPairs CmdNeo4j's blocking-MsgBox failure class; not unlocked by PR #109's driver patch |

### Bucket distribution after this refresh

| Bucket | Count | Cells |
|---|---:|---|
| A small_candidate | 0 | — |
| B blocked_by_known_driver_issue | **3** *(was 5)* | AssocPairs CmdGIS; Networks CmdGIS / CmdPajek |
| C blocked_by_form_query_timeout | 1 | GroupData × CmdNeo4j *(now investigation line, Issue #21 canonical)* |
| D new_export_family_needs_design | **4** *(was 3)* | Associations / Place / Kinship / **AssociationPairs** × CmdUCINet |
| D + B stacked | **1** *(was 2)* | Networks × CmdUCINet |

**Remaining gaps: 9** *(was 12 on 2026-05-05; was 10 implied by earlier refreshes' running tallies after the Pajek/Gephi pair landed)*. The 3-cell drop since `refresh_2026_05_06` is: AssocPairs × CmdNeo4j (this refresh, just merged) plus implicit acknowledgement of AssocPairs × CmdPajek and CmdGephi as covered (no longer in any rank).

### Already-landed context observed during this refresh (NOT next work)

Two prior-refresh "future" items were checked and found already on `main` as of `f0e7594` — they are **NOT** in the ranked list below:

- **Kinship × CmdUCINet runtime `:ERR` pin** — already landed at `tests/test_vba_bug_behaviors.py::test_bug22_kinship_cmducinet_sibling_form_fires_invalid_procedure_call` (line 631). Pins `LookAtKinship:ERR Invalid procedure call or argument` on the He-Mou-取 (U+53D6) fixture; documents the sibling-form pattern under Issue #22; states inventory status unchanged. The `refresh_2026_05_06` "Rank 3 — coverage hardening" entry is therefore **completed**, not deferred.

This refresh is therefore restricted to ranking *only* genuinely unfinished work.

### Next work, ranked (max 3)

#### Rank 1 — probe-first investigation: `LookAtAssociations × CmdNeo4j`

- **Category:** probe-first investigation (NOT coverage PR, NOT canonical issue yet).
- **Why:** PR AX's Q5 confirmed this is a *different failure class* from AssocPairs × CmdNeo4j (0-file mode, likely bails before any SaveAs — vs AssocPairs which writes files then hits MsgBox layer). PR #109's driver patch does NOT address it. The cell remains skipped in `tests/test_vba_cmdneo4j_cross_form.py::_spec_skip_marks` with reason "produces 0 files in directory mode — needs investigation alongside Place." This refresh proposes the probe finally happen, scoped strictly to characterizing why the 0-file mode triggers (static read of `Form_LookAtAssociations.vb::CmdNeo4j_Click` + one COM run on the same kind of small fixture used by Associations CmdGIS / CmdPajek tests).
- **Shape:** read-only probe MD + JSON only; same shape as PR AX (`probe/assocpairs-cmdneo4j` artifacts). NOT a coverage PR; NOT a driver PR.
- **Risk:** low; outcome is one of (a) clean probe → coverage candidate, (b) confirmed blocker → new investigation line, (c) different bug class → canonical issue candidate.

#### Rank 2 — maintainer-line: Issue #22 upstream-fix coordination

- **Category:** maintainer-line (out-of-band; NOT a PR action in this repo).
- **Why:** Issue #22 is filed P1 and has a recommended fix (`CreateTextFile(..., True, True)` for the 3rd Unicode arg, with the downstream-acceptance hedge). Reaching the CBDB maintainer to land the fix on the upstream `.mdb` is the next forward step that would meaningfully change the cell-state of `LookAtAssociations × CmdUCINet` and the Kinship sibling fragility. Not a local PR.
- **Shape:** out-of-band coordination; no repo file touched.

#### Rank 3 — maintainer-line: Issue #21 upstream-fix coordination

- **Category:** maintainer-line (out-of-band; NOT a PR action in this repo).
- **Why:** Issue #21 is the canonical P1 for `LookAtGroupData × CmdNeo4j`'s mid-chain `:ERR No current record.` (DAO 3021, unguarded `.MoveFirst` on empty recordset in blocks #9 and #10). Same shape as Rank 2 — a forward step that lives on the CBDB upstream side rather than this repo. Listed here because both Issue #21 and Issue #22 maintainer-lines are genuinely unfinished and either could be the next forward step depending on the maintainer's priority.
- **Shape:** out-of-band coordination; no repo file touched.

### Direct answers to the brief

**Q: After AssocPairs × CmdNeo4j merge, has rank-1 changed?**

**Yes — rank-1 has changed.** `refresh_2026_05_06`'s rank-1 was `LookAtPlace × CmdUCINet` probe-first investigation. That candidate is **paused** as of this refresh on a timeout-path-correlated COM bridge instability that is structurally a separate driver/meta concern, not an export-coverage concern. The new rank-1 is the **`LookAtAssociations × CmdNeo4j` probe-first investigation**, chosen because:

1. It is the cheapest genuinely-unfinished local PR available today.
2. PR AX explicitly declined to bundle it with the AssocPairs probe (different failure class per Q5); now is the natural moment to take the sibling probe given the AssocPairs CmdNeo4j coverage just merged.
3. It has no driver dependency and no new fixture design — same shape as PR AX, scoped strictly to characterizing the 0-file mode.
4. The two maintainer-line items below it (Issue #22, Issue #21) are out-of-band; the Associations CmdNeo4j probe is the only ranked item that is a local PR.

(An earlier draft of this refresh proposed `Kinship × CmdUCINet runtime :ERR pin` as rank-1. That was a stale-premise error: the pin is already on `main` at `tests/test_vba_bug_behaviors.py::test_bug22_kinship_cmducinet_sibling_form_fires_invalid_procedure_call` line 631, completing what `refresh_2026_05_06` Rank 3 had proposed. The corrected ranking only contains genuinely unfinished work.)

### Explicitly NOT to do (this refresh + onward, until briefed otherwise)

- ❌ Do NOT reopen `LookAtAssociationPairs × CmdNeo4j` — covered as of PR #110, scope locked to the 1×3 fixture; broader coverage of this cell is a separate brief.
- ❌ Do NOT mix `LookAtAssociations × CmdNeo4j` with `LookAtAssociationPairs × CmdNeo4j` — different failure classes per PR AX Q5; merging them into one investigation conflates the 0-file-mode root cause with the (now-fixed) MsgBox layer.
- ❌ Do NOT continue dissecting the `LookAtPlace × CmdUCINet` COM bridge instability inside this triage line; that needs its own driver/meta brief (timeout-path-correlated COM bridge issues are a different repo concern from export coverage).
- ❌ Do NOT promote `LookAtAssociationPairs × CmdUCINet` to coverage now that it is D-only — CmdUCINet family infra still does not exist; the inner B blocker being resolved doesn't make the family blocker go away.
- ❌ Do NOT re-promote `LookAtGroupData × CmdNeo4j` — Issue #21 canonical, on the maintainer line.
- ❌ Do NOT assert AssocPairs × CmdNeo4j coverage on fixtures other than the 1×3 known-edged pair — PR #110 deliberately scoped to that fixture and the per-shape pin will fail if a new fixture is added without a fresh brief.

### Self-review (per `docs/skills/programmer-self-review-template.md`)

**A. Branch shape.** Branch `refresh/queue-after-assocpairs-cmdneo4j-covered` cut from clean `main = f0e7594`. Only the two triage files touched (this MD + paired JSON). No tests, driver, README, canonical reports, issue severity, or other artifacts changed — matches the brief's `read-only analysis only` boundary exactly.

**B. Source-of-truth sync.** This MD section ↔ the new JSON `refresh_2026_05_07` block carry the same truth deltas, the same bucket distribution, the same ranked list, the same brief Q-A, and the same do-not-touch list. The five truth-delta items were checked against canonical truth on `main`: PR #109 + PR #110 are merged (verified via `git log` `b695a92` and `f0e7594`); the AssocPairs CmdPajek/Gephi covered state is in the README cross-form table at line 238 and in `tests/test_vba_pajek_gephi_cross_form.py::_assocpairs_1x3_fixture`; `LookAtAssociations × CmdNeo4j` skip reason is in `tests/test_vba_cmdneo4j_cross_form.py::_spec_skip_marks` and PR AX's Q5 in `reports/probe_assocpairs_cmdneo4j.json`. No source-of-truth file is being changed by this refresh.

**C. Evidence vs claim.** Truth deltas cite specific commit SHAs (`b695a92`, `f0e7594`, `4b8a927`) and PR numbers (#109, #110). The "Place paused on COM bridge instability" item is reported as the user's standing instruction, not as a finding this refresh produced — no claim about Place's runtime state is added. Rank-1 (`LookAtAssociations × CmdNeo4j` probe-first) is hedged with three explicit outcome branches, none of which is pre-claimed. The Kinship-pin-already-landed observation is grounded in a verifiable file read (`tests/test_vba_bug_behaviors.py` line 631) and recorded under the "Already-landed context observed during this refresh" section, not silently dropped.

**D. Residual risk.** Triage-document refresh, not an implementation; residual risk is purely advisory. Specifically: (1) Place pause assumes the COM bridge instability is a current-session standing constraint — if that's resolved, Place returns to rank-1 candidate via a separate refresh, not silently. (2) The corrected ranking ranks two maintainer-line items at Rank 2 and Rank 3; if the maintainer priority is "ship more local coverage first", the Associations CmdNeo4j probe at Rank 1 is the only local PR available — there is no honest Rank-2 local PR today (every other unfinished cell needs driver/meta or family-design work first). No code path or test altered, so no runtime regression risk.

**Correction tracking.** An earlier draft of this refresh proposed `LookAtKinship × CmdUCINet runtime :ERR pin` as Rank 1. Reviewer correctly flagged it as a stale-premise error: the pin is already on `main` at `tests/test_vba_bug_behaviors.py::test_bug22_kinship_cmducinet_sibling_form_fires_invalid_procedure_call` line 631 — completing what `refresh_2026_05_06` Rank 3 had proposed. The corrected ranking removes the Kinship pin from `next_work_items_ranked`, records it under `already_landed_observed_during_refresh` for honesty, and re-ranks only genuinely unfinished work. Same correction applies to the MD `### Already-landed context observed during this refresh (NOT next work)` section above and the JSON `already_landed_observed_during_refresh` block.

---

## Refresh 2026-05-07 (later) — post LookAtAssociations × CmdNeo4j coverage merge

Triggering events (the full Associations × CmdNeo4j unblock chain, all on `main`):

- PR #112 (`probe/associations-cmdneo4j`, commit `1145219`) — runtime probe; observed JET 3061 + 0 files; verdict `probe_found_new_runtime_bug_candidate`.
- PR #114 (`investigate/associations-cmdneo4j-c-index-addr-type-code`, commit `68cfa9b`) — static investigation; pinned target-side schema mismatch on `ZZ_SCRATCH_PEOPLE.c_index_addr_type_code`.
- PR #115 (`issue/file-associations-cmdneo4j-target-column-mismatch`, commit `97ff1d8`) — canonical Issue #23 filed (P1_visible_crash); 4 anchored static markers in `tests/test_known_bugs.py`.
- PR #116 (`driver/associations-cmdneo4j-c-addr-type-rewrite`, commit `9040daa`) — narrow driver patch `_rewrite_associations_cmdneo4j_target_column` rewrites `c_index_addr_type_code, c_female` → `c_addr_type, c_female`; verification probe `patch_resolved_issue23_but_exposed_msgbox_blocker` (8 files, 5 watchdog dialogs).
- PR #117 (`driver/associations-cmdneo4j-debug-msgbox-suppress`, commit `fed8734`) — narrow driver patch `_suppress_associations_cmdneo4j_debug_msgbox` suppresses the 5 concat-form debug MsgBox prefixes; verification probe `patch_verified_chain_clean` (8 files, 0 watchdog dialogs, 0 :ERR).
- PR #118 (`test/associations-cmdneo4j-coverage`, commit `7741059`) — coverage test on the matrix `assoc_<top_code>_unfiltered` fixture; per-shape pin `_assert_lookatassociations_neo4j_shape` rules `len(files) == 8` exactly with no missing / no extra header first-cols; classifier gained `OccasionCode`, `TopicCode`, and a 4-col `(AssociationCode, AssociationTypeID)` 2-col disambiguator.

Read-only refresh on top of `main = 7741059`. No COM, no tests / driver / canonical reports / issue severity changes.

### Current truth deltas (vs `refresh_2026_05_07` earlier)

| Cell | Before this refresh | After this refresh |
|---|---|---|
| `LookAtAssociations × CmdNeo4j` | rank-1 probe-first candidate (still skipped / investigation line) | **covered via repo-local driver workarounds** (PR #116 + #117 driver patches + PR #118 coverage test); the canonical Issue #23 **stays P1** — the upstream `.mdb` defect is NOT fixed, only the test driver's in-memory VBA is rewritten so the chain runs |
| `LookAtAssociationPairs × CmdNeo4j` | covered (PR #110) | covered (no change) |
| `LookAtPlace × CmdNeo4j` | skipped in `tests/test_vba_cmdneo4j_cross_form.py::_spec_skip_marks` with reason "Item not found in this collection. mid-body — looks like a real CBDB bug" | still skipped / investigation line (no change this refresh; rank-1 candidate this refresh — see below) |
| `LookAtStatus × CmdNeo4j` | skipped with reason "chain post-cleanup invalidates the subform recordset rebind; downstream CmdNeo4j reads RecordCount=0; same root family as Pajek/Gephi Status skip" | still skipped / investigation line (no change this refresh; deferred — see "Place vs Status" below) |
| `Issue #23` (canonical P1) | open, P1_visible_crash | **open, P1_visible_crash (UNCHANGED)** — coverage via local workaround does NOT silently downgrade or close the canonical issue |

### Coverage-via-local-workaround vs upstream-fixed (a hedge that must not slip)

**Coverage landed ≠ upstream fixed.** This is the most important clarification of this refresh:

- The Associations × CmdNeo4j cell can now be exercised by the test suite end-to-end.  This is achieved via **two driver-side workarounds in `tests/cbdb_driver/vba_session.py`** that rewrite the in-memory VBA at test time:
  - `_rewrite_associations_cmdneo4j_target_column` (PR #116) — fixes the JET 3061 column-not-found INSERT
  - `_suppress_associations_cmdneo4j_debug_msgbox` (PR #117) — suppresses the 5 concat-form debug MsgBox dialogs
- A real CBDB end user opening the un-patched `.mdb` would **still hit the JET 3061** when clicking CmdNeo4j on Associations.  The user-visible defect is unchanged.
- The canonical Issue #23 in `reports/generate_report.py::ISSUES` therefore stays at `P1_visible_crash`.  Its 4 static markers in `tests/test_known_bugs.py::test_bug23_associations_cmdneo4j_target_column_mismatch` still reproduce against the unpatched VBA dump on disk; the marker test will fire only when the upstream `.mdb` is actually fixed.
- `triage should not silently drop Issue #23 because the test now passes under driver patching` — this refresh explicitly preserves the canonical issue and the static-marker test.  Both states are coherent: the cell IS test-covered AND the upstream defect IS open; they are independent assertions on the same underlying source.

### Bucket distribution after this refresh

| Bucket | Count | Cells |
|---|---:|---|
| A small_candidate | 0 | — |
| B blocked_by_known_driver_issue | 3 *(unchanged)* | AssocPairs CmdGIS; Networks CmdGIS / CmdPajek |
| C blocked_by_form_query_timeout | 1 *(unchanged)* | GroupData × CmdNeo4j *(investigation line, Issue #21 canonical)* |
| D new_export_family_needs_design | 4 *(unchanged)* | Associations / Place / Kinship / AssociationPairs × CmdUCINet |
| D + B stacked | 1 *(unchanged)* | Networks × CmdUCINet |

**Remaining gaps: 8** *(was 9 in `refresh_2026_05_07` earlier)*. The 1-cell drop is `LookAtAssociations × CmdNeo4j` (this refresh, just merged via PR #118).

Note: the bucket categories above only cover cells that are *gaps in the cross-form Neo4j coverage*. The Place / Status × CmdNeo4j skips were already in the `refresh_2026_05_07` ledger as "investigation line" cells, parallel to GroupData × CmdNeo4j (Issue #21).

### Next work, ranked (max 3)

#### Rank 1 — probe-first investigation: `LookAtPlace × CmdNeo4j`

- **Category:** probe-first investigation (NOT coverage PR, NOT canonical issue yet).
- **Why:** the cheapest genuinely-unfinished local PR available today. `tests/test_vba_cmdneo4j_cross_form.py::_spec_skip_marks` describes the cell as "fires `Item not found in this collection.` mid-body — looks like a real CBDB bug (SQL or recordset field reference against a renamed/missing column)". That is a **JET 3061 column-not-found family**, structurally identical to Issue #23's symptom. The end-to-end methodology that worked for Associations CmdNeo4j (#112 probe → #114 static investigation → #115 canonical issue → #116 driver rewrite → #117 MsgBox suppress → #118 coverage) directly transfers as a playbook.
- **Shape:** read-only probe MD + JSON only; same shape as PR #112 (`probe/associations-cmdneo4j` artifacts). NOT a coverage PR; NOT a driver PR.
- **Risk:** low; outcome buckets identical to PR #112's scheme — `clean_probe_promote_to_coverage_candidate` / `probe_found_new_runtime_bug_candidate` / `probe_hit_existing_known_failure_family` / `still_not_cheap_needs_deeper_investigation`.

#### Rank 2 — maintainer-line: Issue #22 upstream-fix coordination

- **Category:** maintainer-line (out-of-band; NOT a PR action in this repo).
- **Why:** Issue #22 (P1) is `LookAtAssociations.CmdUCINet`'s FSO `CreateTextFile` missing-Unicode-arg crash on CJK Han characters; recommended fix is `CreateTextFile(..., True, True)`.  Reaching the CBDB maintainer to land the fix on the upstream `.mdb` is the next forward step that meaningfully changes the cell-state of `LookAtAssociations × CmdUCINet` and the Kinship sibling fragility.  Not a local PR.
- **Shape:** out-of-band coordination; no repo file touched.

#### Rank 3 — maintainer-line: Issue #21 + Issue #23 upstream-fix coordination

- **Category:** maintainer-line (out-of-band; NOT a PR action in this repo).
- **Why bundled:** both are `Form_LookAt*.CmdNeo4j_Click` source-level defects with concrete fix recommendations in their canonical entries; both have driver-side workarounds already in place (Issue #21 has no workaround yet but is canonically tracked; Issue #23 was just unblocked by PR #116 + #117).  Reaching the maintainer with these alongside Issue #22 keeps the upstream-fix line coherent.
- **Shape:** out-of-band coordination; no repo file touched.
- **Note:** Issue #23 specifically — the driver workarounds in PR #116/#117 are repo-local; they do NOT remove the need for an upstream fix.  When the upstream fix lands, the static markers in `tests/test_known_bugs.py::test_bug23_associations_cmdneo4j_target_column_mismatch` will fire (signal that the canonical issue can be marked fixed AND the driver workarounds can be removed).

### Direct answers to the brief

**Q1: After Associations × CmdNeo4j covered, what is the new rank-1?**

**`LookAtPlace × CmdNeo4j` probe-first investigation.** The previous refresh's rank-1 (`LookAtAssociations × CmdNeo4j` probe-first) is now done end-to-end (chain through PR #118).  Place inherits the rank-1 slot because it has the strongest fit with the just-proven methodology.

**Q2: Why is Place better than continuing to tighten assertions on already-covered Neo4j forms?**

The 6 already-covered Neo4j forms (Entry / Texts / Office / Kinship / AssociationPairs / Associations) all have per-shape pinning today (`_assert_lookatentry_neo4j_shape`, etc., or 1-shape-per-host equivalent).  Tighter pinning would catch more regressions but does NOT expand the coverage surface.  Promoting Place from skip → covered moves the family from 6/8 → 7/8 covered — a much higher leverage move.  Per the brief's `不顺手处理别的 Neo4j hosts` boundary, the choice is between Place and Status, not "tighter pinning on covered forms".

**Q3: Place vs Status for CmdNeo4j next — which one, or neither?**

**Place is more valuable to do first.  Status should NOT be the next probe.**

| | LookAtPlace × CmdNeo4j | LookAtStatus × CmdNeo4j |
|---|---|---|
| skip reason | "Item not found in this collection. mid-body — looks like a real CBDB bug (SQL or recordset field reference against a renamed/missing column)" | "chain post-cleanup invalidates the subform recordset rebind; downstream CmdNeo4j reads RecordCount=0; same root family as Pajek/Gephi Status skip" |
| failure family | likely **JET 3061 column-not-found** (same as Issue #23) | **CmdQuery cleanup-rebind architectural issue** affecting multiple chains — different family |
| methodology fit | high — directly mirrors the Associations chain (#112 → #118) | low — Status's CmdQuery cleanup section is the root cause and it would need driver/meta-level work, not per-form column rewrite |
| likely outcome of probe | one of the four PR #112-style buckets, with high probability of `probe_found_new_runtime_bug_candidate` (reproducing Issue #23-shape canonicalization) | likely a confirmed pre-existing architectural issue family with no per-form workaround playbook |

Status is also explicitly skipped on the Pajek / Gephi cross-form tests for the same root cause; tackling Status × CmdNeo4j without first tackling the broader CmdQuery-cleanup issue would be premature.

### Explicitly NOT to do (this refresh + onward)

- ❌ Do NOT reopen the Associations × CmdNeo4j or AssocPairs × CmdNeo4j Neo4j probes — both are covered, scope locked.
- ❌ Do NOT downgrade or close Issue #23 just because Associations × CmdNeo4j is now test-covered.  Coverage via local workaround ≠ upstream fixed; the canonical entry stays `P1_visible_crash` and the static markers stay live.
- ❌ Do NOT write the workaround-covered cell as upstream-fixed in any artifact.  The README cross-form table row for LookAtAssociations × CmdNeo4j (updated in PR #118) explicitly cites "Issue #23 stays P1 canonical, the workarounds make the cell *testable* on the existing source, not *fixed* upstream" — that hedge must not slip.
- ❌ Do NOT touch README / canonical reports / issue severity in this triage refresh.
- ❌ Do NOT promote `LookAtStatus × CmdNeo4j` ahead of Place — different failure family, requires driver/meta work, premature.
- ❌ Do NOT promote `LookAtAssociationPairs × CmdUCINet` to coverage — still D-only (CmdUCINet family infra still does not exist).
- ❌ Do NOT continue dissecting `LookAtPlace × CmdUCINet` COM bridge instability — it needs its own driver/meta brief, separate from the CmdNeo4j line.

### Self-review (per `docs/skills/programmer-self-review-template.md`)

**A. Branch shape.** Branch `refresh/queue-after-associations-cmdneo4j-covered` cut from clean `main = 7741059`. Only the two triage files touched (this MD + paired JSON). No tests, driver, README, canonical reports, issue severity, or other artifacts changed — matches the brief's `read-only analysis only` boundary exactly.

**B. Source-of-truth sync.** This MD section ↔ the new JSON `refresh_2026_05_07_later` block carry the same truth deltas, bucket distribution, ranked list, brief Q-A, do-not-touch list, and coverage-vs-upstream-fixed hedge. Truth-delta items checked against canonical truth on `main`: all six unblock-chain PRs (#112 / #114 / #115 / #116 / #117 / #118) merged with verified commit SHAs; `Issue #23` confirmed `tier = P1_visible_crash` via direct read of `reports/generate_report.py::ISSUES`; static marker test still in `tests/test_known_bugs.py`. No source-of-truth file is being changed by this refresh.

**C. Evidence vs claim.** All truth-delta items cite specific commit SHAs (`1145219` / `68cfa9b` / `97ff1d8` / `9040daa` / `fed8734` / `7741059`) and PR numbers (#112 / #114 / #115 / #116 / #117 / #118). The "Issue #23 stays P1" claim is grounded in a direct read of `ISSUES`, not paraphrased. The "Place is better than Status" judgement is grounded in a 4-row comparison table; "Place's skip reason is JET 3061-family" is hedged with "likely" because no probe has run yet. No claim that Place's probe will succeed; only that the methodology fit is strong.

**D. Residual risk.** Triage-document refresh, not an implementation; residual risk is purely advisory. Specifically: (1) Place's actual probe outcome is not yet known — if the probe surfaces a structural issue more like Status's cleanup-rebind class than Issue #23's column typo, the rank-1 framing may need correcting via a follow-up refresh.  (2) Bundling Issue #23 into the maintainer-line at Rank 3 is a judgement call — if maintainer prioritizes the upstream-fix coordination differently, ranks 2 and 3 can be reordered.  No code path or test altered, so no runtime regression risk.

---

## Refresh 2026-05-08 — post LookAtPlace × CmdNeo4j coverage merge

Triggering events (the full Place × CmdNeo4j unblock chain, all on `main`):

- PR #120 (`probe/place-cmdneo4j`, commit `dd2ed37`) — runtime probe; observed `LookAtPlace:ERR Item not found in this collection.` mid-body (DAO 3265 family) + 0 files; verdict `probe_found_new_runtime_bug_candidate`.
- PR #121 (`investigate/place-cmdneo4j-trstpeople-projection`, commit `97e1162`) — static investigation; pinned `tRstPeople` SELECT projection mismatch — sub reads `c_dynasty / c_dynasty_chn / c_female` via `!c_dynasty` / `!c_female` field accessors, but the SELECT lists only 4 columns from `ZZ_SCRATCH_P_TEXT`.
- PR #122 (`issue/file-place-cmdneo4j-trstpeople-projection-mismatch`, commit `aaffa4b`) — canonical Issue #24 filed (P1_visible_crash); 5 anchored static markers in `tests/test_known_bugs.py::test_bug24_*`.
- PR #123 (`driver/place-cmdneo4j-trstpeople-projection-rewrite`, commit `dbd0236`) — narrow driver patch `_rewrite_place_cmdneo4j_trstpeople_projection` extends the SELECT projection inside `Form_LookAtPlace.CmdNeo4j_Click` to add `DYNASTIES.c_dynasty / c_dynasty_chn / BIOG_MAIN.c_female`; verification probe `patch_verified_chain_clean` (6 files, 0 watchdog dialogs, 0 `:ERR`).
- PR #124 (`coverage/place-cmdneo4j-real-vba`, commit `a38d353`) — coverage test on the matrix `place_addr_<top_addr_id>` fixture; per-shape pin `_assert_lookatplace_neo4j_shape` rules `len(files) == 6` exactly with no missing / no extra header first-cols; classifier gained `PersonID`, `PersonPlaceRelCode`, `IndexAddrTypeCode` (3 new single-col entries; no 2-col disambiguator needed).

Read-only refresh on top of `main = a38d353`. **No COM, no tests / driver / canonical reports / issue severity / README changes.**

**Methodology note:** Place's chain ran the same 5-PR shape as Associations (#112 → #118 in 2026-05-07 later), but **needed only 1 driver-side workaround** instead of 2 — Place's `Form_LookAtPlace.CmdNeo4j_Click` has no concat-form debug-MsgBox layer, so the analog of PR #117 was not needed.

### Current truth deltas (vs `refresh_2026_05_07_later`)

| Cell | Before this refresh | After this refresh |
|---|---|---|
| `LookAtPlace × CmdNeo4j` | rank-1 probe-first candidate (still skipped / investigation line) | **covered via repo-local driver workaround** (PR #123 driver patch + PR #124 coverage test); the canonical Issue #24 **stays P1** — the upstream `.mdb` defect is NOT fixed, only the test driver's in-memory VBA is rewritten so the chain runs |
| `LookAtAssociations × CmdNeo4j` | covered via repo-local driver workarounds (PR #116 + #117 + #118) | covered (no change) |
| `LookAtAssociationPairs × CmdNeo4j` | covered (PR #110) | covered (no change) |
| `LookAtStatus × CmdNeo4j` | skipped with reason "chain post-cleanup invalidates the subform recordset rebind; downstream CmdNeo4j reads RecordCount=0; same root family as Pajek/Gephi Status skip" | **only remaining skipped / investigation line in the CmdNeo4j family** (no change to skip; rank explicitly DEFERRED by this refresh — see `place_vs_status_decision` carry-forward and "rank-1 changed" Q below) |
| `Issue #24` (canonical P1) | filed 2026-05-08 via PR #122; open, P1_visible_crash | **open, P1_visible_crash (UNCHANGED)** — coverage via local workaround does NOT silently downgrade or close the canonical issue |
| `Issue #23` (canonical P1) | open, P1_visible_crash (carry-forward from `refresh_2026_05_07_later`) | **open, P1_visible_crash (UNCHANGED)** — same hedge applies |

### Coverage-via-local-workaround vs upstream-fixed (the same hedge, restated for Issue #24)

**Coverage landed ≠ upstream fixed.**  This rule applied to Issue #23 in `refresh_2026_05_07_later`; it now applies identically to Issue #24:

- The Place × CmdNeo4j cell can now be exercised by the test suite end-to-end via **one driver-side workaround in `tests/cbdb_driver/vba_session.py`** that rewrites the in-memory VBA at test time:
  - `_rewrite_place_cmdneo4j_trstpeople_projection` (PR #123) — extends the `tRstPeople` SELECT projection to include the 3 columns the downstream field reads need
- A real CBDB end user opening the un-patched `.mdb` would **still hit the JET 3265** (`Item not found in this collection.`) when clicking CmdNeo4j on Place.  The user-visible defect is unchanged.
- The canonical Issue #24 in `reports/generate_report.py::ISSUES` therefore stays at `P1_visible_crash`.  Its 5 static markers in `tests/test_known_bugs.py::test_bug24_*` still reproduce against the unpatched VBA dump on disk; the marker tests will fire only when the upstream `.mdb` is actually fixed.
- **`triage should not silently drop Issue #24 because the test now passes under driver patching`** — this refresh explicitly preserves the canonical issue and the static-marker tests.  Both states are coherent: the cell IS test-covered AND the upstream defect IS open; they are independent assertions on the same underlying source.
- **The same rule still applies to Issue #23** (carry-forward; not re-asserted in any new way) — it stays P1 canonical with its 4 static markers still live.  Closing or downgrading either Issue #23 or Issue #24 because their respective cells are now test-covered is explicitly disallowed without an upstream-fix-landed signal.

### Bucket distribution after this refresh

| Bucket | Count | Cells |
|---|---:|---|
| A small_candidate | 0 *(unchanged)* | — |
| B blocked_by_known_driver_issue | 3 *(unchanged)* | AssocPairs CmdGIS; Networks CmdGIS / CmdPajek |
| C blocked_by_form_query_timeout | 1 *(unchanged)* | GroupData × CmdNeo4j *(investigation line, Issue #21 canonical)* |
| D new_export_family_needs_design | 4 *(unchanged)* | Associations / Place / Kinship / AssociationPairs × CmdUCINet |
| D + B stacked | 1 *(unchanged)* | Networks × CmdUCINet |

**Remaining gaps: 8** *(unchanged from `refresh_2026_05_07_later`)*.  Place × CmdNeo4j was never in the cells-array bucket distribution above — it was always in the parallel `investigation-line` category alongside Status × CmdNeo4j and GroupData × CmdNeo4j (Issue #21).  The covered-via-workaround flip moves Place out of the investigation-line list but does NOT affect the cells-array gap count.

**CmdNeo4j family snapshot after this refresh:** **7 covered / 1 skipped**.

| Cell | State | Notes |
|---|---|---|
| LookAtEntry × CmdNeo4j | covered | 7-file set; pins Issue #9 LATENT-gate |
| LookAtTexts × CmdNeo4j | covered | 4-file floor |
| LookAtOffice × CmdNeo4j | covered | 4-file floor + OfficeCodes shape |
| LookAtKinship × CmdNeo4j | covered | 4-file floor |
| LookAtAssociationPairs × CmdNeo4j | covered | 6-file pin (PR #110); requires PR #109 driver patch |
| LookAtAssociations × CmdNeo4j | covered (workaround) | 8-file pin (PR #118); requires PR #116 + #117 driver patches; Issue #23 P1 canonical |
| LookAtPlace × CmdNeo4j | **covered (workaround) — NEW this refresh** | 6-file pin (PR #124); requires PR #123 driver patch; Issue #24 P1 canonical |
| LookAtStatus × CmdNeo4j | skipped | architectural CmdQuery cleanup-rebind family; shared root with Status × CmdPajek + CmdGephi |

### Next work, ranked (max 3) — with explicit category

#### Rank 1 — maintainer-line: bundled CmdNeo4j-family upstream-fix coordination (Issues #21 + #23 + #24)

- **Category:** maintainer-line / upstream-fix coordination (out-of-band; NOT a PR action in this repo).
- **Why rank-1:** three canonical P1 entries on the CmdNeo4j family now sit on the upstream side, all with concrete fix recommendations in their canonical entries:
  - **Issue #21** (`LookAtGroupData × CmdNeo4j`) — unguarded `.MoveFirst` on empty recordset (DAO 3021); fix: empty-recordset guards on blocks #9 and #10
  - **Issue #23** (`LookAtAssociations × CmdNeo4j`) — INSERT target column typo `c_index_addr_type_code` → `c_addr_type`; fix: rename in source VBA
  - **Issue #24** (`LookAtPlace × CmdNeo4j`) — `tRstPeople` SELECT projection missing 3 columns; fix: extend SELECT to add `DYNASTIES.c_dynasty / c_dynasty_chn / BIOG_MAIN.c_female`
  Each issue has 4-5 anchored static markers in `tests/test_known_bugs.py` that will fire automatically when the upstream `.mdb` ships the fix — that's the agreed signal for marking each canonical issue resolved AND removing its corresponding driver workaround (where one exists).  Reaching the CBDB maintainer with all three together is the single highest-leverage forward step right now.
- **Why bundled, not individual ranks:** the three issues share a pattern (all `Form_LookAt*.CmdNeo4j_Click` source-level defects; all P1; all with verified driver workarounds OR canonical issues already filed); a single maintainer interaction is more efficient than three sequential ones.
- **Shape:** out-of-band coordination; no repo file touched.
- **Is local PR:** no.
- **Note on Issue #22:** Issue #22 (`LookAtAssociations × CmdUCINet` FSO Unicode crash) also remains an open P1 maintainer-line item.  Treated as a separate channel because it's a different export family (CmdUCINet, not CmdNeo4j) and a different defect class (FSO encoding, not source-level VBA bug).

#### Rank 2 — driver/meta investigation: Status CmdQuery cleanup-rebind family root cause

- **Category:** driver / meta investigation (NOT coverage PR, NOT canonical issue, NOT per-form workaround).
- **Why rank-2:** the CmdNeo4j family now has **only one remaining skipped cell** — `LookAtStatus × CmdNeo4j` — and its skip reason explicitly identifies the root cause as **shared with Status × CmdPajek + CmdGephi** (per `tests/test_vba_pajek_gephi_cross_form.py` Status skip reason and the "same root family" cross-reference in this triage's CmdNeo4j skip reason).  A single driver/meta investigation that lifts the CmdQuery cleanup-rebind family blocker would unblock **3 cells at once**:
  - `LookAtStatus × CmdNeo4j` (this triage)
  - `LookAtStatus × CmdPajek` (per Pajek/Gephi cross-form test skip)
  - `LookAtStatus × CmdGephi` (per Pajek/Gephi cross-form test skip)
- **Why NOT a single-cell coverage PR for Status × CmdNeo4j:** the per-form column-typo playbook proven in #112 → #118 (Associations) and #120 → #124 (Place) does **not** transfer.  Status's failure is architectural (cleanup-rebind invalidation pattern), not a per-form VBA literal typo.  A probe-first investigation on Status × CmdNeo4j alone would not add new information beyond the already-documented skip reason — both `_spec_skip_marks` (CmdNeo4j) and the Pajek/Gephi Status skip reason already name the architectural root cause.  Skipping straight to the per-form workaround pattern would also be premature because a per-form Status workaround would not generalize and would solve only 1 of 3 cells.
- **Shape:** read-only static investigation of Status's CmdQuery cleanup section + the existing `_SUBFORMS_TO_REQUERY` driver dict pre-existing Status warning comment + a small probe characterizing the rebind invalidation pattern across all three Status export buttons (CmdNeo4j, CmdPajek, CmdGephi).  Output: a probe MD + JSON like the AssocPairs CmdGIS not-cheap analysis (`analysis/assocpairs_cmdgis_note.md`), NOT a driver patch.  If the probe surfaces a viable driver-side fix, that's a separate PR.
- **Is local PR:** yes (probe / investigation, no driver or test changes).
- **Risk:** medium — investigation may conclude (similar to AssocPairs CmdGIS) that no narrow driver fix exists and the cell stays skipped pending broader meta work.  That's still a valid outcome — it would close the question of whether Status is "stuck for the same reason as AssocPairs CmdGIS" or "stuck for a different reason".

#### Rank 3 — probe-first investigation: `LookAtPlace × CmdUCINet` resume

- **Category:** probe-first investigation (NOT coverage PR, NOT canonical issue).
- **Why rank-3:** this was `refresh_2026_05_06`'s rank-1 candidate, then paused in `refresh_2026_05_07` on a "timeout-path-correlated COM bridge instability" concern.  After two CmdNeo4j unblock chains shipped (#112 → #118 for Associations, #120 → #124 for Place) and AssocPairs CmdNeo4j's per-shape pin proved solid, the COM bridge stack is empirically more stable than at the time of pause.  Resuming this probe would add the last unprobed cell of the CmdUCINet family and create a clean fork: either coverage candidate, new canonical issue, or deferred gap with a recorded reason.
- **Important hedge:** resume is **subject to a maintainer brief** confirming the COM bridge instability concern is now scoped down enough — this refresh does NOT itself authorize the probe, only ranks it as the next probe-shape candidate IF a brief lands.
- **Shape:** read-only probe MD + JSON only; same shape as the original rank-1 framing in `refresh_2026_05_06`.
- **Is local PR:** yes.
- **Risk:** low; Place's CmdUCINet writer is structurally distinct from Issue #22's failing FSO writer (uses ADO Stream, UTF-8 by default — see `refresh_2026_05_06` Place static signal section).

### Direct answers to the brief

**Q1: Is `LookAtStatus × CmdNeo4j` (the only remaining skipped cell in the CmdNeo4j family) rank-1?**

**No.**  Demoted from rank-1 to rank-2 (and re-categorized from `coverage_pr` / `probe_first_investigation` to `driver_meta_investigation`).  The shared-root-cause observation (Q3 below) is the deciding factor.

**Q2: If Status × CmdNeo4j is not rank-1, what is, and why?**

The new rank-1 is **maintainer-line: bundled CmdNeo4j-family upstream-fix coordination (Issues #21 + #23 + #24)**.

Why it ranks above Status driver/meta investigation:

1. **Three canonical P1 issues are now ripe.**  Issues #21 / #23 / #24 are all on the CmdNeo4j family, all with concrete fix recommendations, all already canonicalized.  This refresh moves Issue #24 to that ripe state alongside #21 and #23 — making a bundled maintainer-line move uniquely well-supported right now (it wasn't in `refresh_2026_05_07_later`, when only Issue #21 and #23 were ripe).
2. **Each issue has automatic upstream-fix detection.**  All three sets of static markers in `tests/test_known_bugs.py` will fire when the upstream fix lands.  No additional repo-side work needed to verify; the tests do it.
3. **The Status driver/meta investigation is a slower, more uncertain path.**  AssocPairs CmdGIS taught us that some architectural blockers don't have narrow fixes (`refresh_2026_05_05_later` final state).  Status's CmdQuery cleanup-rebind family may have the same shape — investigation may conclude no narrow fix exists.  That's a valid outcome but not a guaranteed unblock.
4. **The CmdNeo4j family is the highest-leverage upstream-fix bundle right now** — 3 canonical P1's, all fixable upstream with low-risk targeted edits.  Other families (CmdUCINet at Issue #22) are split across single issues with higher-risk fixes.

**Q3: Does Status × CmdNeo4j sharing its root cause with Status × CmdPajek + Status × CmdGephi change the ranking judgement?**

**Yes — directly.**  Three points:

1. **Three cells share one blocker.**  This is the highest unlock-per-effort ratio among remaining open cells.  A successful Status driver/meta investigation would close 3 cells; a per-form Status × CmdNeo4j workaround (analogous to PR #123 / PR #116) would close only 1.
2. **The per-form workaround pattern does NOT fit.**  PR #117 / #123 / #116 all rewrite single VBA literals (column names / SELECT projections / MsgBox prefixes) inside one `CmdNeo4j_Click` body.  Status's CmdQuery cleanup-rebind issue lives in `CmdQuery_Click`'s post-INSERT recordset rebind section — same code path used by all three downstream export buttons (CmdNeo4j, CmdPajek, CmdGephi).  Per-form workaround would either (a) re-implement the cleanup section in 3 places or (b) miss the actual root cause.  Both are bad.
3. **Rank-2 (NOT rank-1) is the right slot for this work.**  It's leverage-positive (3 cells unlocked vs 1) but uncertainty-heavy (may find no narrow fix, like AssocPairs CmdGIS did).  Pairing it with a higher-confidence rank-1 (the maintainer-line) hedges the queue.

The shared-root-cause observation thus elevates Status work from "ignore until a separate brief" to "rank-2 driver/meta investigation, tied to the broader CmdQuery-cleanup family" — but does NOT elevate it to rank-1.

### Explicitly NOT to do (this refresh + onward)

- ❌ Do NOT reopen the Place / Associations / AssociationPairs Neo4j probes — all three are covered, scope locked.  Re-running probes wastes COM cycles and produces no new triage signal.
- ❌ Do NOT downgrade or close Issue #23 because Associations × CmdNeo4j is now test-covered.  Coverage via local workaround ≠ upstream fixed.
- ❌ Do NOT downgrade or close Issue #24 because Place × CmdNeo4j is now test-covered.  Same rule, same reason.  Static markers stay live.
- ❌ Do NOT write any workaround-covered cell (Associations / Place × CmdNeo4j) as upstream-fixed in any artifact — the README rows for both, updated in PR #118 and PR #124 respectively, explicitly cite the "Issue #N stays P1 canonical, the workaround makes the cell *testable* on the existing source, not *fixed* upstream" hedge.  That hedge must not slip.
- ❌ Do NOT touch README / canonical reports / issue severity in this triage refresh — explicit brief boundary.
- ❌ Do NOT process `analysis/report_screenshot_audit.md` drift — explicit brief boundary; that drift belongs to a separate PR (or to the auto-sync gate's next run).
- ❌ Do NOT promote `LookAtStatus × CmdNeo4j` as a single-cell coverage PR — see Q3 above; the per-form workaround playbook does NOT fit and would solve only 1 of 3 affected cells.
- ❌ Do NOT promote `LookAtAssociationPairs × CmdUCINet` to coverage — still D-only (CmdUCINet family infra still does not exist).
- ❌ Do NOT continue dissecting `LookAtPlace × CmdUCINet` COM bridge instability inside this triage line — it remains its own driver/meta brief, separate from the CmdUCINet probe ranking.

### Self-review (per `docs/skills/programmer-self-review-template.md`)

**A. Branch shape.** Branch `triage/cmdneo4j-place-covered-refresh` cut from clean `main = a38d353`. Only the two triage files touched (this MD + paired JSON). No tests, driver, README, canonical reports, issue severity, or other artifacts changed — matches the brief's `read-only analysis only` boundary exactly.  Pre-existing `analysis/report_screenshot_audit.md` drift left alone per brief boundary.

**B. Source-of-truth sync.** This MD section ↔ the new JSON `refresh_2026_05_08` block carry the same truth deltas, bucket distribution, ranked list, brief Q-A, and do-not-touch list.  Truth-delta items checked against canonical truth on `main = a38d353`: all five unblock-chain PRs (#120 / #121 / #122 / #123 / #124) merged with verified commit SHAs (`dd2ed37` / `97e1162` / `aaffa4b` / `dbd0236` / `a38d353`); `Issue #24` confirmed `tier = P1_visible_crash` via direct read of `reports/generate_report.py::ISSUES`; 5 static markers still in `tests/test_known_bugs.py::test_bug24_*`; CmdNeo4j family covered count `6 → 7` and skipped count `2 → 1` aligned with both `tests/test_vba_cmdneo4j_cross_form.py::_spec_skip_marks` (only Status remains) and the README roadmap-8 line updated by PR #124. No source-of-truth file is being changed by this refresh.

**C. Evidence vs claim.** All five unblock-chain PRs cited with explicit commit SHAs (`dd2ed37` / `97e1162` / `aaffa4b` / `dbd0236` / `a38d353`) and PR numbers (#120 / #121 / #122 / #123 / #124). The "Issue #24 stays P1" claim is grounded in a direct read of `ISSUES`, not paraphrased.  The "Status shares root cause with CmdPajek + CmdGephi" claim is grounded in a direct cross-reference between `_spec_skip_marks` (CmdNeo4j) and `tests/test_vba_pajek_gephi_cross_form.py` skip reasons — not invented.  Rank-2 (Status driver/meta investigation) is hedged with the AssocPairs CmdGIS analog ("may find no narrow fix"); no claim that the investigation will succeed, only that the leverage justifies attempting it.  Rank-3 (Place × CmdUCINet resume) is explicitly hedged as `subject to a maintainer brief` — not authorizing resumption, only ranking it.

**D. Residual risk.** Triage-document refresh, not an implementation; residual risk is purely advisory. Specifically: (1) Rank-1 being a maintainer-line item means the queue's top item is intentionally out-of-band; if no maintainer interaction happens, ranks 2 and 3 are the next local PR candidates.  (2) Rank-2 (Status driver/meta investigation) may conclude no narrow fix exists, in which case Status × CmdNeo4j stays skipped and the next refresh would re-rank.  (3) Bundling Issues #21 + #23 + #24 into a single maintainer-line item is a judgement call — if the maintainer prefers issue-by-issue handoff, the bundle decomposes naturally; nothing in this refresh prevents that.  (4) Rank-3's unpause is conditional; if the COM bridge instability is still considered active, rank-3 stays parked.  No code path or test altered, so no runtime regression risk.

---

## Refresh 2026-05-08 (later) — post Status × CmdPajek/Gephi local-workaround line exhaustion

Triggering events (the full Status × CmdPajek/Gephi local-workaround investigation chain, all on `main`):

- PR #127 (`investigate/status-cleanup-rebind-family`, commit `9461472`) — driver/meta probe; pinned `Object required` on Pajek/Gephi to subform Recordset Nothing post-CmdQuery cleanup; CmdNeo4j unaffected.
- PR #128 (`coverage/status-cmdneo4j-real-vba`, commit `22667eb`) — CmdNeo4j unskip (false-positive skip removed; no driver patch needed).
- PR #129 (`investigate/status-cleanup-rebind-set-to-requery`, commit `a783b72`) — candidate (a) Set→Requery tested; removed Object required but exposed RecordCount=0; driver REVERTED.
- PR #130 (`investigate/status-subform-recordsource-binding`, commit `5f100b4`) — static investigation of subform RecordSource binding; surfaced ZZ_SCRATCH_P_STATUS form's c_dynasty Filter as candidate; needed runtime confirmation.
- PR #131 (`investigate/status-p-status-runtime-microcheck`, commit not pinned) — runtime micro-check (3 reads); decisive `H_chain_timing_supported`: Q1 c_dynasty='unknown' = 7/17022 (filter immaterial); Q2 FilterOn=False (filter dormant); Q3 RecordCount = 17023/17022 with explicit Requery + 1.5 s sleep.
- PR #132 (`driver/chain-dispatcher-settle`, commit not pinned) — VBA-side 250 ms DoEvents settle in chain dispatcher; Object required reappeared; driver REVERTED.
- PR #133 (`investigate/status-explicit-requery-variant`, commit not pinned) — explicit `Form.Requery` + DoEvents only (no sleep); Object required reappeared; driver REVERTED.
- PR #134 (`investigate/status-settle-bisect`, commit not pinned) — VBA-side settle bisect 500/750/1000 ms; ALL failed → `no_value_up_to_1000ms_unblocks`; identified VBA-side DoEvents ≠ COM-side `time.sleep` mechanism boundary; driver REVERTED.
- PR #135 (`investigate/status-com-split-dispatch`, commit not pinned) — COM-side split dispatch (two separate `click_via_timer` calls with Python `time.sleep(1.5)` between); Phase B's sub didn't run (Phase B `ZZ_TEST_DEBUG` empty) → infrastructure limitation surfaced.
- PR #136 (`investigate/status-direct-invoke`, commit `948978d`) — direct invoke via `Application.Run` + Public-wrapper trick; rejected by Access (form/class-module subs not addressable from external `Application.Run`); boundary finding.
- PR #137 (`investigate/click-via-timer-second-dispatch`, commit `7b76c8e`) — raw-COM `Form_Timer` instrumentation + `acCmdCompileAndSaveAllModules` force-compile fix attempt; force-compile succeeded but timer body STILL did not dispatch; verdict pinned the layer at Access form-class-instance event-binding cache.

Read-only refresh on top of `main = 7b76c8e`. **No COM, no tests / driver / canonical reports / issue severity / README changes.**

### Current truth deltas (vs `refresh_2026_05_08`)

| Cell | Before this refresh | After this refresh |
|---|---|---|
| `LookAtStatus × CmdNeo4j` | rank-2 driver/meta investigation candidate (still skipped) | **covered** since PR #128 (no driver patch needed; the prior skip was a false-positive copy-paste from the Pajek/Gephi cross-form Status skip — PR #127's Q3 directly refuted "same root family"; CmdNeo4j opens fresh `dbOpenDynaset` on the underlying scratch tables and bypasses the subform recordset entirely) |
| `LookAtStatus × CmdPajek` | skipped (cross-form Pajek/Gephi test) | **still skipped** — local test-driver workaround line EXHAUSTED across PR #127 + #129/#132/#133/#134/#135/#136/#137 (7 mechanism / sleep / dispatcher / direct-invoke / infra layers all tested; all failed); next forward step is maintainer-line / upstream-fix coordination |
| `LookAtStatus × CmdGephi` | skipped (cross-form Pajek/Gephi test) | **still skipped** — same as CmdPajek (sibling cell, identical failure shape across all 7 attempts) |
| **CmdNeo4j family snapshot** | 7 covered / 1 skipped | **8 covered / 0 skipped** ✓ |
| Issue #21 / #23 / #24 (canonical P1 each) | open, P1_visible_crash | open, P1_visible_crash (UNCHANGED) — no test-driver coverage state change is allowed to silently downgrade these |

### Why CmdNeo4j was a false-positive skip and is now covered

Pre-PR #127 the test `tests/test_vba_cmdneo4j_cross_form.py::_spec_skip_marks` skipped `LookAtStatus × CmdNeo4j` with reason *"chain post-cleanup invalidates the subform recordset rebind; downstream CmdNeo4j reads RecordCount=0; same root family as Pajek/Gephi Status skip"*.  PR #127's 3-phase driver/meta probe RAN CmdNeo4j on the same fixture as Pajek + Gephi.  Result: CmdNeo4j produced 6 files in 11.04 s with `[ENTER, MSGBOX, DONE]` markers AND no `:ERR`; Pajek + Gephi failed with `Object required`.

Static reading of `Form_LookAtStatus.vb` confirmed CmdNeo4j is structurally distinct: at lines 527+528 it opens fresh `dbOpenDynaset` directly on the underlying scratch tables (`ZZ_SCRATCH_STATUS` / `ZZ_SCRATCH_P_STATUS`), bypassing the subform recordset entirely.  The cleanup-rebind issue at `Form_LookAtStatus.vb:1457+1460` (which leaves the subform recordset in an Object-Required state and breaks Pajek/Gephi) does not affect CmdNeo4j AT ALL.

The "same root family" claim in the original skip-reason text was a copy-paste from the Pajek/Gephi skip without verification.  PR #128 directly unskipped CmdNeo4j with per-shape pinning of the 6-file output and chain-completion gates; no driver patch was needed.

### Why Status × CmdPajek/Gephi remain blocked

Both cells read `<subform>.Form.Recordset.RecordCount` upfront in their `_Click` handlers:
- CmdPajek lines 2156 + 2161
- CmdGephi lines 45 + 50

After CmdQuery's cleanup section runs `Set <subform>.Form.Recordset = <Dim'd-local-recordset>` (lines 1457+1460), the subform's Recordset reads as Nothing in any subsequent test-driver-driven access → VBA error 424 'Object required' → bail.  This is the same observable across all 7 test-driver-side attempts.

The sibling cell distinction is real and structural — CmdNeo4j bypasses the subform recordset; Pajek/Gephi don't.  Coverage state cannot be aligned across the family because the underlying CBDB code paths are different.

### Why the recent 7-probe chain changes the ranking/disposition

The local-workaround line tested across PR #127 + #129 → #137 covered each plausible test-driver-side intervention layer:

| Layer | PR(s) tested | Result |
|---|---|---|
| Per-form VBA literal rewrite (Set→Requery) | #129 | Removes Object required but exposes RecordCount=0 (refuted) |
| VBA-side DoEvents settle window | #132 (250ms generic), #133 (0ms explicit), #134 (500/750/1000ms bisect) | All failed; mechanism boundary identified (VBA-side ≠ COM-side) |
| COM-side sleep + sequential dispatch | #135 | Phase B sub didn't even run (infra limitation surfaced) |
| Direct invoke via `Application.Run` | #136 | Rejected by Access (class-module subs not addressable) |
| Raw-COM `Form_Timer` re-injection + force-compile | #137 | Timer-binding cache pinned; force-compile insufficient |

7 consecutive negative results across mechanism / sleep / dispatcher / direct-invoke / infra layers.  Test-driver-side attempts are exhausted in the realistic sense — the remaining theoretical candidates (close+reopen form between dispatches, standard-module Form_Timer redirect, pywinauto button click) each carry significant implementation cost AND uncertain outcome.

The realistic next forward step is **maintainer-line / upstream-fix coordination** — fix the underlying CBDB pattern (Dim'd-local `Set <subform>.Form.Recordset` rebind in CmdQuery cleanup) at the source level, not work around it at the test-driver level.

### Bucket distribution after this refresh

Unchanged from `refresh_2026_05_08`: 8 remaining gaps in the cells-array.  Status × CmdPajek/Gephi are NOT in the cells-array (they're not gaps in cross-form CmdNeo4j coverage — they live on the cross-form Pajek/Gephi test); they remain in the parallel investigation-line category.

| Bucket | Count | Cells |
|---|---:|---|
| A small_candidate | 0 | — |
| B blocked_by_known_driver_issue | 3 | AssocPairs CmdGIS; Networks CmdGIS / CmdPajek |
| C blocked_by_form_query_timeout | 1 | GroupData × CmdNeo4j (investigation line, Issue #21 canonical) |
| D new_export_family_needs_design | 4 | Associations / Place / Kinship / AssociationPairs × CmdUCINet |
| D + B stacked | 1 | Networks × CmdUCINet |

**Investigation-line cells (separate ledger):**

| Cell | State | Disposition |
|---|---|---|
| LookAtGroupData × CmdNeo4j | skipped | Issue #21 P1 canonical; maintainer-line |
| LookAtAssociations × CmdNeo4j | covered (driver workaround PR #116/#117) | Issue #23 P1 canonical (workaround, not upstream fix) |
| LookAtPlace × CmdNeo4j | covered (driver workaround PR #123) | Issue #24 P1 canonical (workaround, not upstream fix) |
| LookAtStatus × CmdNeo4j | covered | NO canonical issue (false-positive skip; cell ran clean all along) |
| **LookAtStatus × CmdPajek** | **skipped (cross-form Pajek/Gephi test)** | **local test-driver line EXHAUSTED (PR #127 + #129/.../#137); next action is maintainer-line** |
| **LookAtStatus × CmdGephi** | **skipped (cross-form Pajek/Gephi test)** | **same as CmdPajek (sibling cell)** |

### Next work, ranked (max 3) — with explicit category

#### Rank 1 — maintainer-line: bundled CmdNeo4j-family + Status-Pajek/Gephi-cleanup-rebind upstream-fix coordination

- **Category:** maintainer-line / upstream-fix coordination (out-of-band; NOT a PR action in this repo).
- **Why rank-1:** three canonical P1 entries on the CmdNeo4j family (#21 + #23 + #24 unchanged) PLUS a now-empirically-pinned Status × CmdPajek/Gephi root cause that's structurally a CBDB-side defect (Dim'd-local `Set` rebind pattern in cleanup section).  Reaching the CBDB maintainer with all four concerns at once is high-leverage:
  - **Issue #21** (`LookAtGroupData × CmdNeo4j`) — empty-recordset guards on blocks #9/#10
  - **Issue #23** (`LookAtAssociations × CmdNeo4j`) — INSERT target column rename
  - **Issue #24** (`LookAtPlace × CmdNeo4j`) — `tRstPeople` SELECT projection extension
  - **NEW: Status × CmdPajek + CmdGephi cleanup-rebind** — `Form_LookAtStatus.vb:1457+1460` `Set <subform>.Form.Recordset = <Dim'd-local-recordset>` pattern; suggested fix: switch to globals (per `gRstPeople` precedent in line 1184) OR refactor away the rebind entirely.  This last item is NOT yet a canonical Issue — that's a separate decision (whether to file it now or only after maintainer feedback).
- **Why bundled:** all four share a "cross-form export-button blocker rooted in CBDB source pattern" theme; a single maintainer interaction is more efficient than four sequential ones.  Also: the fix surface for the Status cleanup-rebind is small (10 lines or fewer of VBA).
- **Shape:** out-of-band coordination; no repo file touched.
- **Is local PR:** no.

#### Rank 2 — driver/meta investigation: Status × CmdPajek/Gephi close+reopen-form between dispatches (CONDITIONAL)

- **Category:** driver / meta investigation.
- **Why rank-2 (and conditional):** PR #137's verdict-note ranks "close + reopen form between dispatches" as the most likely remaining test-driver-side fix candidate.  This drops the form's class instance entirely, forcing Access to re-resolve event handlers from the freshly-compiled module on Form_Open — directly addresses the per-instance event-binding cache that PR #137 pinned.
- **Why conditional:** explicit maintainer guidance is needed before this is opened.  The brief just told us "next local work should stop"; this PR records the candidate but does NOT authorize.  If maintainer says "try one more local thing before maintainer-line", this is the candidate.  Otherwise it stays parked.
- **Shape:** read-only probe with the form-close + reopen pattern between Phase A and Phase B; no public driver edit at first.  If probe succeeds, separate driver-PR + verification-probe pair lands the public driver change.
- **Is local PR:** yes (probe + driver), but only if explicitly authorized.
- **Risk:** medium — Access form close+reopen has its own complications (fixture state may need re-seeding); also requires heavier driver shape change than any prior attempt.

#### Rank 3 — probe-first investigation: `LookAtPlace × CmdUCINet` resume

- **Category:** probe-first investigation (NOT coverage PR, NOT canonical issue).
- **Why rank-3:** carries forward from `refresh_2026_05_06`/`refresh_2026_05_07_later`; was paused on COM bridge instability concerns; nothing in the past month has changed that calculus.  Listed here only because rank-1 is out-of-band and rank-2 is conditional — if maintainer wants a local PR available, this is the cleanest cheap one.
- **Shape:** read-only probe MD + JSON only; same shape as the original `refresh_2026_05_06` framing.
- **Is local PR:** yes (conditional on COM bridge unpause brief).
- **Risk:** low.

### Direct answers to the brief

**Q1: After CmdNeo4j family closure, what is the new rank-1?**

**Maintainer-line: bundled upstream-fix coordination** for Issues #21 + #23 + #24 + the newly-pinned Status × CmdPajek/Gephi cleanup-rebind root cause.  This is the same rank-1 framing as `refresh_2026_05_08`, now expanded with the Status-Pajek/Gephi addition.

**Q2: Is any local repo PR line still worth pursuing before maintainer-line?**

**Not at rank-1 priority.**  The 7-probe chain produced sufficient negative evidence that further test-driver-side carving has steeply diminishing returns.  The two remaining theoretical candidates (close+reopen, standard-module redirect) carry implementation cost AND uncertain outcomes.

If the maintainer's preference is "try one more local thing before maintainer-line", **rank-2 (close+reopen probe)** is the candidate.  Otherwise, **the realistic next step is rank-1 (maintainer-line) directly**.

**Q3: Should `Status × CmdPajek/Gephi` now be considered parked?**

**Yes, parked at the local-workaround layer.**  Both cells stay skipped; the cross-form Pajek/Gephi test's `_case_skip_marks` retains its existing skip reason without revision in this PR.  The disposition is: **parked pending maintainer-line outcome**.  If maintainer fixes the Set-rebind upstream, the cells become unskippable AND a coverage PR can land.  If maintainer declines and explicitly authorizes more local work, rank-2 (close+reopen) is the candidate.  If nothing happens, both cells stay skipped indefinitely — but with a clean disposition record explaining why.

### Explicitly NOT to do (this refresh + onward)

- ❌ Do NOT continue the local test-driver workaround line for Status × CmdPajek/Gephi.  PR #127 + #129 → #137 inclusive is the exhaustion record; further narrow tweaks in that direction would not produce new evidence.
- ❌ Do NOT re-open the CmdNeo4j Status / Place / Associations Neo4j probes — all covered, scope locked.
- ❌ Do NOT downgrade or close Issue #21 / #23 / #24 — coverage via local workaround does NOT silently change canonical issue truth.  Same rule.
- ❌ Do NOT promote Status × CmdPajek/Gephi to coverage at the cross-form Pajek/Gephi test surface.  They remain skipped pending maintainer-line.
- ❌ Do NOT touch README / canonical reports / issue severity in this triage refresh.
- ❌ Do NOT process `analysis/report_screenshot_audit.md` drift.

### Self-review (per `docs/skills/programmer-self-review-template.md`)

**A. Branch shape.** Branch `triage/status-pajek-gephi-disposition` cut from current `main = 7b76c8e` (post PR #137 merge — same baseline as the refresh prose above).  Only the two triage files touched (this MD + paired JSON).  No tests, driver, README, canonical reports, issue severity, or other artifacts changed.  Pre-existing `analysis/report_screenshot_audit.md` drift left alone per standing instruction.

**B. Source-of-truth sync.** This MD section ↔ the new JSON `refresh_2026_05_08_later` block carry the same truth deltas, ranked list, brief Q-A, and do-not-touch list.  The 7-probe chain is cited PR-by-PR (#127 / #128 / #129 / #130 / #131 / #132 / #133 / #134 / #135 / #136 / #137); the CmdNeo4j family closure (8/0) is cited against the actual current state of `tests/test_vba_cmdneo4j_cross_form.py::_spec_skip_marks` (post-PR #128: empty body).  No source-of-truth file is being changed by this refresh.

**C. Evidence vs claim.** Each PR's outcome is summarized factually (the Status family's 7-PR exhaustion table); no claim that maintainer-line will succeed (rank-1 is presented as "the realistic next step", not a guaranteed unblock).  Rank-2 (close+reopen) is explicitly hedged as "conditional" — not a recommendation to open it; the rank-3 carry-forward has the same hedge it had in prior refreshes.  CmdNeo4j family false-positive narrative is grounded in PR #127's runtime evidence + static read of `Form_LookAtStatus.vb:527+528`, not paraphrased.

**D. Residual risk.** Triage refresh, not implementation; residual risk is purely advisory: (1) rank-1 maintainer-line item is intentionally out-of-band — if no maintainer interaction happens, the queue's top item is "wait"; that's a deliberate design, not an oversight.  (2) rank-2 conditional-unpause is a judgement call; maintainer may prefer skipping local work entirely.  (3) the disposition "parked" for Status × CmdPajek/Gephi could be re-considered if a new fundamentally-different mechanism emerges (e.g., Access COM API extension we haven't discovered yet); the refresh leaves this open.  (4) bundling the Status cleanup-rebind into the same maintainer-line as Issues #21/#23/#24 is a packaging decision; the maintainer may want to handle it as a 5th issue with its own canonical entry — that's fine, the refresh doesn't prevent it.  No code path or test altered, so no runtime regression risk.
