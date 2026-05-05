# Export coverage gap triage plan

**Date:** 2026-05-04
**Branch:** `plan/export-gap-triage` (off main `434168a`)
**Source data (read-only):**
- `reports/export_coverage_inventory.json` — 13 `gap` cells
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

### LookAtGroupData × CmdGIS — bucket A (small_candidate)

- **handler exists?** yes
- **button exists?** yes
- **current blocker:** matrix CmdRun timeout (GroupData aggregates across many tables on heavy fixtures).  But person_1 small fixture works (`tests/test_vba_matrix_hard_forms.py::_HARD_FIXTURES`).
- **nearest existing test pattern:** `tests/test_vba_cmdgis_other_forms.py` covers 6 other forms × CmdGIS; LookAtGroupData was explicitly excluded (`"omitted because their CmdQuery / CmdRun ... matrix"` per docstring).
- **recommended next action:** wire GroupData into a small-fixture variant of the cross-form CmdGIS test (similar shape to the AssociationPairs.CmdGIS recommendation above — both share the matrix-timeout-but-small-fixture-works pattern).
- **risk:** **low-medium** — small fixture proven for CmdRun; CmdGIS export chain on the same small input is uncharted but conceptually the standard CmdGIS `.tab` write loop.

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
| LookAtGroupData × CmdGIS | A small_candidate | low-medium |
| LookAtAssociationPairs × CmdNeo4j | C blocked_by_form_query_timeout | medium-high |
| LookAtGroupData × CmdNeo4j | C blocked_by_form_query_timeout | medium |
| LookAtNetworks × CmdGIS | B blocked_by_known_driver_issue | medium |
| LookAtNetworks × CmdPajek | B blocked_by_known_driver_issue | medium |
| LookAtAssociations × CmdUCINet | D new_export_family_needs_design | medium |
| LookAtPlace × CmdUCINet | D new_export_family_needs_design | medium |
| LookAtKinship × CmdUCINet | D new_export_family_needs_design | medium |
| LookAtAssociationPairs × CmdUCINet | D + B stacked | high |
| LookAtNetworks × CmdUCINet | D + B stacked | high |

| Bucket | Count (updated) |
|---|---:|
| A small_candidate | **1** (was 4) |
| B blocked_by_known_driver_issue | **5** (was 2) |
| C blocked_by_form_query_timeout | 2 |
| D new_export_family_needs_design | 3 |
| D + B stacked | **2** (was 1 D+C, 1 D+B) |

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

### PR 3 (now PR 1) — GroupData × CmdGIS via person_1 small fixture
- **Bucket:** A
- **Risk:** low-medium
- **Scope sketch:** add `LookAtGroupData` to `tests/test_vba_cmdgis_other_forms.py::_FORMS_WITH_CMDGIS_TESTABLE_HERE`; wire person_1 small fixture similarly.  Single-file `.tab` export.
- **Caveat:** GroupData's CmdRun behaviour (UPDATE-style backfill) differs from Associations/Place/Kinship CmdQuery — and given that AssociationPairs's CmdQuery turned out to have a hidden SetFocus blocker, **probe GroupData's CmdRun + CmdGIS chain end-to-end FIRST** before promising this is "low risk".  The triage plan repeatedly under-estimated risk for matrix-blocked forms; assume the same may be true here until probed.
- **Status:** **the only remaining bucket-A cell**, but its risk profile is no longer "guaranteed low" — re-baseline via probe before the next implementer commits to it.

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
