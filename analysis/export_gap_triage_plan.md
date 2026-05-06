# Export coverage gap triage plan

**Date:** 2026-05-04 · **Refreshed:** 2026-05-05 (post GroupData × CmdGIS coverage merge)
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
| D + B stacked | 2 *(state as of 2026-05-05; see Refresh 2026-05-06 second — AssocPairs × CmdUCINet is now D-only)* | ~~AssociationPairs × CmdUCINet~~ *(now D-only)*; Networks × CmdUCINet |

*(State as of 2026-05-05. Superseded by Refresh 2026-05-06 (second):
AssocPairs × CmdUCINet moved from D+B to D-only (SetFocus patch landed,
commits `3bb69ef` + `0c0eaf1`); CmdPajek + CmdGephi covered (commit
`4b8a927`); CmdGIS has a second independent stale-subform-RecordCount
blocker; CmdNeo4j is now probe-first candidate. See Refresh 2026-05-06
(second) for current truth.)*

*(Historical context for B-bucket count: B-bucket was 5 — 4 cells gated on
the AssociationPairs SetFocus driver patch (now resolved), 2 cells gated on
Networks Form_Open landmine #3.5.  CmdNeo4j AssociationPairs sat in B; the
4 vs 2 summed to 6 because Networks CmdUCINet was counted under "D + B
stacked", not in the pure-B row.)*

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
| LookAtAssociationPairs × CmdGIS / CmdNeo4j / CmdPajek / CmdGephi | AssociationPairs (per refresh brief); also gated on CmdQuery SetFocus driver patch *(historical; patch landed commits `3bb69ef`+`0c0eaf1`; CmdPajek+CmdGephi now covered commit `4b8a927`; CmdGIS has second independent blocker; CmdNeo4j is now probe-first candidate — see Refresh 2026-05-06 second)* |
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

### LookAtAssociationPairs × CmdGIS — bucket B (stale-subform-RecordCount blocker)

> **Historical (2026-05-04):** re-classified A → B; blocker was CmdQuery
> SetFocus driver bug (same as CmdPajek/Gephi).
> **Updated per Refresh 2026-05-06 (second):** SetFocus patch landed
> (commits `3bb69ef` + `0c0eaf1`); but CmdGIS does **NOT** drop back to
> bucket A. A second independent blocker was confirmed when CmdGIS was
> attempted post-patch (branch `cover/assocpairs-cmdgis-1x3`, no commits,
> abandoned): `CmdQuery` cleanup at line 2000 opens a fresh `dbOpenDynaset`
> whose `RecordCount` returns 0 until visited; `.Form.Requery` after that
> rebind likely invalidates the fresh recordset. Documented in
> `analysis/assocpairs_cmdgis_note.md`. The wording below in italics
> is superseded historical context.

*Historical context (2026-05-04):*
*Original triage: bucket A. Re-classified because CmdGIS is downstream
of the same CmdQuery whose SetFocus bug blocked INSERTs. Investigation
didn't directly probe CmdGIS, but the blocker was upstream of the export.*

- **handler exists?** yes
- **button exists?** yes
- **current blocker (2026-05-06):** stale-subform-RecordCount — NOT SetFocus
  (resolved). `_PER_FORM_CMDGIS_PATCHES` SetFocus suppression is in place;
  CmdQuery now runs its INSERTs. But `ZZ_SOCIAL_NETWORK.RecordCount`
  returns 0 before `.MoveLast`, and the subform Requery after CmdQuery's
  fresh `dbOpenDynaset` rebind likely invalidates it. Distinct from the
  SetFocus fix. Full context: `analysis/assocpairs_cmdgis_note.md`.
- **nearest existing test pattern:** `tests/test_vba_cmdgis_other_forms.py`
  (5 forms × CmdGIS). Does not transfer directly — subform-RecordCount
  issue is AssocPairs-specific.
- **recommended next action (2026-05-06):** needs own investigation brief
  for the stale-subform-RecordCount blocker; waiting for the SetFocus patch
  is no longer the constraint.
- **risk:** **medium** — second independent driver-level blocker confirmed.

### LookAtAssociationPairs × CmdNeo4j — **probe-first candidate** (SetFocus lifted)

> **Historical (2026-05-04):** bucket C — matrix-CmdQuery timeout + 0-file
> Associations sibling failure stacked.
> **Re-classified (2026-05-05 later):** bucket B — actual blocker was the
> SetFocus driver issue, not timeout (same as CmdPajek/Gephi).
> **Updated per Refresh 2026-05-06 (second):** SetFocus patch landed
> (commits `3bb69ef` + `0c0eaf1`). SetFocus blocker resolved. This cell
> is now the **rank-1 probe-first candidate** — the prior triage's own
> condition "probe after SetFocus patch lands" is now met. Wording below
> in italics is superseded historical context.

*Historical context (2026-05-04): matrix-CmdQuery timeout on 4×5 fixture;
two unknowns stacked (timeout + 0-file Associations sibling mode). The
timeout explanation was later corrected — actual blocker was the same
SetFocus driver bug as CmdPajek/Gephi (see Refresh 2026-05-05 later).*

- **handler exists?** yes
- **button exists?** yes
- **current blocker (2026-05-06):** multi-file Neo4j chain timing on small
  known-edged fixture — unknown. Sibling LookAtAssociations CmdNeo4j
  "0-file mode" is an open companion question.
- **nearest existing test pattern:** `tests/test_vba_cmdneo4j_cross_form.py`
  cross-form host; 1×3 known-edged fixture shape from CmdPajek/CmdGephi
  coverage (`4b8a927`).
- **recommended next action (2026-05-06):** probe-first — drive CmdQuery +
  CmdNeo4j on a small known-edged fixture; time the chain; count files;
  snapshot ZZ_TEST_DEBUG. Promote to coverage only if chain completes
  within watcher timeout AND ≥1 file AND no mid-chain `:ERR`. **Ranked
  Rank 1** in Refresh 2026-05-06 (second).
- **risk:** **medium** (multi-file chain timing unknown; Associations
  CmdNeo4j 0-file mode is an open companion question).

### LookAtAssociationPairs × CmdPajek — **COVERED (commit `4b8a927`)**

> **Superseded per Refresh 2026-05-06 (second):** this cell is **covered**.
> SetFocus driver patch landed (commits `3bb69ef` + `0c0eaf1`;
> `_PER_FORM_CMDGIS_PATCHES["Form_LookAtAssociationPairs"]`), followed by
> coverage PR commit `4b8a927` (`tests/test_vba_pajek_gephi_cross_form.py`,
> 1×3 known-edged fixture). The blocker narrative below is historical
> context only — the recommended next action is **no longer actionable**
> for this cell.

**Historical blocker narrative (2026-05-04, now resolved):**

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
- **historical blocker (superseded):** `Me.CmdQuery.SetFocus` inside
  CmdQuery_Click failed under Form_Timer dispatch. Resolved by
  `_PER_FORM_CMDGIS_PATCHES["Form_LookAtAssociationPairs"]`
  (commits `3bb69ef` + `0c0eaf1`).
- **current status (2026-05-06):** **covered** —
  `tests/test_vba_pajek_gephi_cross_form.py` `_CASES` includes
  AssociationPairs CmdPajek via 1×3 known-edged fixture (`4b8a927`).
- **recommended next action (2026-05-06):** none — cell is closed.
- **risk:** —

### LookAtAssociationPairs × CmdGephi — **COVERED (commit `4b8a927`)**

> **Superseded per Refresh 2026-05-06 (second):** this cell is **covered**.
> Same resolution path as CmdPajek above: SetFocus patch landed
> (commits `3bb69ef` + `0c0eaf1`) + coverage PR commit `4b8a927`.
> The wording below is historical context only.

- **handler exists?** yes
- **button exists?** yes
- **historical blocker (superseded):** identical to CmdPajek — same
  CmdQuery SetFocus bug; `Form_LookAtAssociationPairs.vb:113-119`
  shows CmdGephi has the same `RecordCount = 0` early-bail. Resolved
  by `_PER_FORM_CMDGIS_PATCHES["Form_LookAtAssociationPairs"]`.
- **current status (2026-05-06):** **covered** —
  `tests/test_vba_pajek_gephi_cross_form.py` `_CASES` includes
  AssociationPairs CmdGephi via 1×3 known-edged fixture (`4b8a927`).
- **recommended next action (2026-05-06):** none — cell is closed.
- **risk:** —

### LookAtAssociationPairs × CmdUCINet — bucket D (CmdUCINet family only)

> **Updated per Refresh 2026-05-06 (second):** SetFocus patch landed
> (commits `3bb69ef` + `0c0eaf1`); the inner B-stack (SetFocus / matrix-
> CmdQuery) blocker is resolved. Bucket was D+B stacked (re-classified
> from D+C in the 2026-05-04 summary JSON); now **D-only** — CmdUCINet
> family design is the sole remaining blocker.

- **handler exists?** yes
- **button exists?** yes
- **current blocker (2026-05-06):** CmdUCINet family blocker only (bucket D).
  SetFocus stack removed. CmdUCINet family design + probe pass still needed
  before this cell can be addressed.
- **recommended next action (2026-05-06):** NOT in next 1-3 PRs. Resolve
  CmdUCINet family design first; AssociationPairs is the highest-risk member
  of that family once the design is in place.
- **risk:** **high** — CmdUCINet family design still needed; AssocPairs is
  structurally complex.

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
| LookAtAssociationPairs × CmdGephi | ~~A small_candidate~~ → ~~B blocked_by_known_driver_issue~~ → **COVERED** (commit `4b8a927`, Refresh 2026-05-06 second) | — |
| LookAtGroupData × CmdGIS | ~~A small_candidate~~ → **COVERED 2026-05-05** | — |
| LookAtAssociationPairs × CmdNeo4j | ~~C blocked_by_form_query_timeout~~ → ~~B~~ → **probe-first candidate** (SetFocus lifted, Refresh 2026-05-06 second) | medium |
| LookAtGroupData × CmdNeo4j | C blocked_by_form_query_timeout | medium |
| LookAtNetworks × CmdGIS | B blocked_by_known_driver_issue | medium |
| LookAtNetworks × CmdPajek | B blocked_by_known_driver_issue | medium |
| LookAtAssociations × CmdUCINet | D new_export_family_needs_design | medium |
| LookAtPlace × CmdUCINet | D new_export_family_needs_design | medium |
| LookAtKinship × CmdUCINet | D new_export_family_needs_design | medium |
| LookAtAssociationPairs × CmdUCINet | ~~D + B stacked~~ → **D-only** (SetFocus lifted, Refresh 2026-05-06 second) | high |
| LookAtNetworks × CmdUCINet | D + B stacked | high |

*(The two tables below show counts as of 2026-05-05; superseded by
Refresh 2026-05-06 (second). For current truth see the Refresh
2026-05-06 (second) section. Key changes: CmdPajek + CmdGephi covered;
AssocPairs × CmdUCINet is D-only; CmdNeo4j is probe-first candidate.)*

| Bucket (12 remaining gaps after 2026-05-05) — **historical; see note above** | Count |
|---|---:|
| A small_candidate | **0** (was 1 on 2026-05-04; GroupData × CmdGIS now covered) |
| B blocked_by_known_driver_issue | 5 *(historical; CmdPajek+CmdGephi now covered; counts pre-date Refresh 2026-05-06 second)* |
| C blocked_by_form_query_timeout | 2 |
| D new_export_family_needs_design | 3 |
| D + B stacked | 2 *(historical; AssocPairs × CmdUCINet is now D-only)* |

*(Counts above use the post-2026-05-04 reclassifications.  See the
JSON for canonical bucket-per-cell assignments — the per-cell rows
in this MD's table above predate the JSON's CmdNeo4j AssocPairs
re-class and are kept as historical record.)*

*(Historical note: LookAtAssociationPairs × CmdUCINet was previously
stacked D + C; the underlying matrix-CmdQuery blocker was later
corrected to the same SetFocus driver issue as the Pajek/Gephi cells,
so its inner blocker moved from C → B.  Stack was D + B — but the
SetFocus patch has since landed (commits `3bb69ef` + `0c0eaf1`), making
it D-only.  Current truth: D-only.  See Refresh 2026-05-06 (second).)*

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

### ~~Driver-side meta-PR~~ — **COMPLETED (commits `3bb69ef` + `0c0eaf1`)**

> **Superseded per Refresh 2026-05-06 (second):** this direction has been
> implemented. `_ASSOCPAIRS_SETFOCUS_TARGETS_BY_SUB` dict +
> `_PER_FORM_CMDGIS_PATCHES["Form_LookAtAssociationPairs"]` in
> `cbdb_driver/vba_session.py` suppresses 12 SetFocus call sites across
> `CmdQuery_Click`, `Link1stOrder`, and `Link2ndOrder`. Coverage for
> CmdPajek + CmdGephi followed in commit `4b8a927`. See per-cell entries
> above for current state of remaining cells (CmdGIS has second independent
> blocker; CmdNeo4j is now probe-first; CmdUCINet is D-only).

*(Historical wording for audit trail:)*

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
| ~~LookAtAssociationPairs × CmdGIS / CmdPajek / CmdGephi~~ | *(Historical; SetFocus patch landed. CmdPajek+CmdGephi are **covered** (`4b8a927`). CmdGIS has independent stale-subform-RecordCount blocker — see per-cell entry and `analysis/assocpairs_cmdgis_note.md`.)* |
| LookAtAssociationPairs × CmdNeo4j | SetFocus lifted; now **probe-first candidate** (Rank 1, Refresh 2026-05-06 second). NOT autopilot — chain-timing probe required before any coverage PR. |
| LookAtGroupData × CmdNeo4j | Multi-file Neo4j chain on heavy form; watcher-timeout risk.  Sequence after GroupData CmdGIS. |
| LookAtNetworks × CmdGIS / CmdPajek | Form_Open landmine #3.5 — needs minimal-injection scaffolding either via extending `tests/test_vba_networks_small_fixture.py` or refactoring cross-form test driver.  Either path is design work, not mechanical. |
| LookAtAssociations × CmdUCINet | New export family — needs design + probe pass to identify file format and assertion strictness BEFORE writing the test. |
| LookAtPlace × CmdUCINet | Same family blocker. |
| LookAtKinship × CmdUCINet | Same family blocker. |
| LookAtAssociationPairs × CmdUCINet | CmdUCINet family design only (bucket D-only; SetFocus stack removed per Refresh 2026-05-06 second). |
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

## Refresh 2026-05-06 (second) — post Place-probe-pause + infra landing

**Correction from initial draft:** the initial draft incorrectly
listed the AssocPairs CmdQuery SetFocus driver patch as Rank-1
next work.  That patch is already merged on main (commits
`3bb69ef` + `0c0eaf1`; `_PER_FORM_CMDGIS_PATCHES["Form_LookAt
AssociationPairs"]` in `cbdb_driver/vba_session.py`).
AssocPairs × CmdPajek + CmdGephi coverage is also already
merged (commit `4b8a927`; `test_vba_pajek_gephi_cross_form.py`
1×3 known-edged fixture).  This corrected version removes all
forward-looking references to that work and re-ranks from
current main truth.

Triggering event: main is now at `bd6337f` (PR AV #106 —
`test_inputs.json` auto-refresh gate). PRs since prior refresh
(`14099a7`):

- **PR AT (#104)** Kinship CmdUCINet runtime `:ERR` pin —
  Rank-3 deferred item from prior refresh is now **done**.
- **PR AS (#103)** Place × CmdUCINet probe; 4 iterations; all
  concluded with COM RPC unavailable on second `set_form_tag`;
  classified `still_needs_better_fixture`.
- **PR AU (#105)** 5-trial Place COM bridge instability matrix;
  `long_click_via_timer_polling_loop_correlated`; unresolved
  confound — T3 vs T5 differ along two dimensions (polling
  return mode AND post-CmdQuery COM-touch type).
- **PR AV (#106)** `pytest_configure` auto-refresh gate;
  stale fixture-input drift is no longer a top queue concern.

**AssocPairs context (landed before 14099a7; canonical on main
— NOT next steps):**

- SetFocus driver patch: merged (`3bb69ef` + `0c0eaf1`;
  `_ASSOCPAIRS_SETFOCUS_TARGETS_BY_SUB` dict +
  `_PER_FORM_CMDGIS_PATCHES["Form_LookAtAssociationPairs"]`
  in `cbdb_driver/vba_session.py`).
- AssocPairs × CmdPajek + CmdGephi: covered (`4b8a927`;
  `test_vba_pajek_gephi_cross_form.py`, 1×3 known-edged
  fixture).  2 cells closed.
- AssocPairs × CmdGIS: SetFocus blocker is lifted, but a
  **second independent blocker** remains — stale-subform-
  RecordCount (unguarded `.MoveFirst` on freshly-opened
  `dbOpenDynaset`; documented in
  `analysis/assocpairs_cmdgis_note.md`).  NOT a coverage
  candidate under current state.
- AssocPairs × CmdNeo4j: SetFocus blocker is now lifted.
  Prior triage condition was "probe after patch lands" — that
  condition is now met.  **New probe-first candidate.**
- AssocPairs × CmdUCINet: was D+B stacked (SetFocus +
  CmdUCINet family); SetFocus now lifted → **D-only**
  (CmdUCINet family design still needed).

**Read-only analysis. No Access COM. No tests / driver /
README / canonical-report changes.**

### CmdUCINet family — current state on main (bd6337f)

| Cell | Status | Key update since 14099a7 |
|---|---|---|
| `LookAtAssociations × CmdUCINet` | gap | canonical Issue #22 (P1_visible_crash); unchanged |
| `LookAtKinship × CmdUCINet` | covered (fixture-fragile) | Rank-3 item done: runtime `:ERR` pin now in `test_vba_bug_behaviors.py` (PR AT) |
| `LookAtPlace × CmdUCINet` | gap — **paused** | Probe tried and stopped; COM bridge instability; unresolved confound; more matrix work prohibited |

### AssocPairs line — updated state on main (bd6337f)

| Cell | Current status | Notes |
|---|---|---|
| CmdPajek | **covered** | `4b8a927`; 1×3 known-edged fixture |
| CmdGephi | **covered** | `4b8a927`; 1×3 known-edged fixture |
| CmdGIS | bucket B (stale-subform-RecordCount blocker) | SetFocus lifted; independent second blocker remains; `analysis/assocpairs_cmdgis_note.md` |
| CmdNeo4j | **probe-first candidate** | SetFocus lifted; remaining: chain timing on small known-edged fixture |
| CmdUCINet | bucket D-only | SetFocus lifted; CmdUCINet family design still needed |

### New next-work ranking (max 3 items)

#### Rank 1 — probe-first investigation: AssocPairs × CmdNeo4j

- **Category:** `probe-first investigation`
- **Why now:** the SetFocus blocker that kept CmdNeo4j in
  bucket B has been resolved.  The prior triage's own
  condition — "probe after SetFocus patch lands" — is now
  met.  This is the cheapest remaining AssocPairs cell
  without an additional independent blocker.
- **Shape:** drive CmdQuery + CmdNeo4j on a small known-edged
  fixture (analog of the 1×3 pair used for CmdPajek/CmdGephi);
  time the multi-file chain; count files; snapshot
  ZZ_TEST_DEBUG.  Probe MD + JSON only — no `test_*`
  additions.  Note: also check whether LookAtAssociations ×
  CmdNeo4j "0-file mode" (open question in the triage doc)
  is resolved, since CmdQuery source-data behavior is now
  confirmed to work.
- **Promote to coverage condition:** chain completes within
  watcher timeout AND ≥1 file produced AND no mid-chain
  `:ERR`.
- **Reject to investigation condition:** chain trips watcher
  OR mid-chain `:ERR`.
- **Risk:** medium (multi-file Neo4j chain on AssocPairs is
  a new shape; sibling Associations CmdNeo4j 0-file mode is
  an open companion question).

#### Rank 2 — maintainer-line / upstream-fix coordination: Issue #22 upstream fix

- **Category:** `maintainer-line / upstream-fix coordination`
- **Scope:** out-of-band; reach CBDB maintainer to land
  `CreateTextFile(..., True, True)` Unicode arg fix in
  `Form_LookAtAssociations.CmdUCINet_Click` and
  `Form_LookAtKinship.CmdUCINet_Click`.  Not a PR in this
  repo.
- **Why second:** Issue #22 is P1 visible crash for end
  users.  Fix recommendation is canonical.  Only upstream
  action remains.

#### Rank 3 — infra: Networks Form_Open minimal-injection scaffold

- **Category:** `infra` (design work)
- **Scope:** extend `tests/test_vba_networks_small_fixture.py`
  to cover CmdGIS / CmdPajek slices using the proven Cao Zhi
  minimal-injection pattern
  (`skip_inject_autodetect_forms=SKIP_SIBLINGS`), OR refactor
  the cross-form CmdGIS / CmdPajek test infrastructure to
  thread `skip_inject_autodetect_forms` through.
- **Unblocks:** Networks × CmdGIS + CmdPajek → coverage
  candidates (2 cells from bucket B).
- **Not autopilot:** scaffold approach choice (extend vs
  refactor) is a design decision; needs scope-defining brief.

### Explicitly NOT to touch (this refresh + onward)

- ❌ **Continue Place × CmdUCINet COM bridge matrix** —
  unresolved confound; more isolation work hits diminishing
  returns; prohibited.
- ❌ **Downgrade Kinship × CmdUCINet from covered** —
  fixture-fragile caveat is the truth; runtime pin in place.
- ❌ **Reopen Issue #22 / #23 wording** — canonical; agreed.
- ❌ **AssocPairs × CmdGIS coverage PR** — stale-subform-
  RecordCount is the remaining independent blocker (SetFocus
  is no longer the constraint); needs its own investigation
  brief before a coverage PR is meaningful.
- ❌ **AssocPairs × CmdNeo4j coverage PR without probe** —
  chain timing on small fixture is still unknown; probe
  outcome determines next step.
- ❌ **Matrix/probe batches without a clear target** — infra
  is now stable; more probes need a specific hypothesis.

### Self-review (per `docs/skills/programmer-self-review-template.md`)

**A. Branch shape.** Branch `queue/recheck-2026-05-06` cut
from clean main `bd6337f`.  Only two files touched: this MD
and the paired JSON.  No tests / driver / README /
canonical-report changes — matches brief boundary exactly.

**B. Source-of-truth sync.** MD section ↔ JSON
`refresh_2026_05_06_second` block carry the same CmdUCINet
family state, AssocPairs line state, ranked list, and
do-not-touch list.  AssocPairs state verified against:
`cbdb_driver/vba_session.py` `_PER_FORM_CMDGIS_PATCHES
["Form_LookAtAssociationPairs"]`; `test_vba_pajek_gephi_cross
_form.py` `_CASES` including AssociationPairs rows; prior
triage `Refresh 2026-05-05 (later)` section documenting the
stale-subform-RecordCount blocker for CmdGIS.  No source-of-
truth file is changed.

**C. Evidence vs claim.** CmdNeo4j "probe-first candidate"
claim is grounded in the prior triage's own explicit
condition ("probe after SetFocus patch lands"), now met —
not an inference.  AssocPairs × CmdGIS "independent second
blocker" claim cites `analysis/assocpairs_cmdgis_note.md`
already on main.  AssocPairs × CmdUCINet "D-only" reflects
the logical consequence of SetFocus being lifted from a D+B
stack — the CmdUCINet family design requirement is unchanged.
All state claims are grounded in existing canonical docs on
main `bd6337f`.

**D. Residual risk.** This is a triage-document refresh, not
an implementation.  Residual risk is advisory error: ranking
AssocPairs × CmdNeo4j probe above Networks scaffold or
Issue #22 maintainer-line could be wrong if the maintainer's
priority differs.  Mitigated by: all three ranked items with
explicit do-not-touch list give the reviewer full context to
redirect.  No code path or test is altered; no runtime
regression risk.
