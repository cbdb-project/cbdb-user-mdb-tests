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

### LookAtAssociationPairs × CmdGIS — bucket A (small_candidate)

- **handler exists?** yes
- **button exists?** yes
- **current blocker:** matrix CmdQuery for AssociationPairs times out (Link1stOrder JET self-join — see `_xfail_marks` docstring).  But a small fixture (4×5, in `tests/test_vba_matrix_hard_forms.py`) DOES complete CmdQuery to `ZZ_SOCIAL_NETWORK`.
- **nearest existing test pattern:** `tests/test_vba_cmdgis_other_forms.py` (5 forms × CmdGIS, single-file output, structural assertion).  AssociationPairs's CmdGIS would write a single `.tab` GIS export, similar shape.
- **recommended next action:** add AssociationPairs to `_FORMS_WITH_CMDGIS_TESTABLE_HERE` and wire `_make_assoc_pairs_fixtures` to `_all_fixtures()` with the existing 4×5 small-fixture controls (or thread a small-fixture variant via a follow-up of the matrix-hard-forms pattern).  Single-file CmdGIS shape is well-understood and matches the existing test's assertion machinery.
- **risk:** **low** — small fixture proven; single-file export; structural assertion already generic.

### LookAtAssociationPairs × CmdNeo4j — bucket C (blocked_by_form_query_timeout)

- **handler exists?** yes
- **button exists?** yes
- **current blocker:** declared in `tests/test_vba_cmdneo4j_cross_form.py::_spec_skip_marks`-style intent but actually missing from `_SPECS`; the underlying matrix-CmdQuery timeout would block it just like it does CmdGIS, but Neo4j's multi-file chain (6-10 SaveAs blocks) on a 4×5 fixture might still complete OR might blow past 180s — this is the unknown.
- **nearest existing test pattern:** `tests/test_vba_cmdneo4j_cross_form.py` is the cross-form host; LookAtAssociations is the closest sibling and is currently *skipped* with `"produces 0 files in directory mode — needs investigation alongside Place"` (NOT a timeout — a different kind of failure).
- **recommended next action:** **NOT in the next 1-3 PRs.**  Two unknowns stack here (matrix timeout + 0-file Associations sibling failure).  Probe first to confirm small-fixture chain runtime AND whether the AssociationPairs-specific Neo4j has the same 0-file mode as Associations.  If both probes are clean, promote to bucket A.  If not, open a separate investigation PR.
- **risk:** **medium-high** — cumulative unknowns.

### LookAtAssociationPairs × CmdPajek — bucket A (small_candidate)

- **handler exists?** yes
- **button exists?** yes
- **current blocker:** not in `_CASES` of `tests/test_vba_pajek_gephi_cross_form.py` (which only has Kinship/Place/Status/Associations for CmdPajek).  AssociationPairs needs the small-fixture (4×5) wiring.
- **nearest existing test pattern:** `tests/test_vba_pajek_gephi_cross_form.py::Case("LookAtAssociations", "CmdPajek", ".net", "*vertices")` — same file format (.net), same `*vertices` header anchor.
- **recommended next action:** add `Case("LookAtAssociationPairs", "CmdPajek", ".net", "*vertices")` to `_CASES` AND add a small-fixture path so `_fixture_for("LookAtAssociationPairs")` returns the 4×5 fixture instead of the matrix one (or use the existing matrix_hard_forms fixture directly).  Single-file `.net` output, shape known.
- **risk:** **low** — Pajek `.net` shape verified; small fixture proven; shape-class shared with already-passing Associations.CmdPajek.

### LookAtAssociationPairs × CmdGephi — bucket A (small_candidate)

- **handler exists?** yes
- **button exists?** yes
- **current blocker:** same as CmdPajek above — not in `_CASES`.
- **nearest existing test pattern:** `Case("LookAtAssociations", "CmdGephi", ".gdf", "nodedef")` already passes.
- **recommended next action:** add `Case("LookAtAssociationPairs", "CmdGephi", ".gdf", "nodedef")` alongside the CmdPajek slice (same PR is reasonable since both share the small-fixture seeding).
- **risk:** **low** — Gephi `.gdf` shape verified.

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

| Cell | Bucket | Risk |
|---|---|---|
| LookAtAssociationPairs × CmdGIS | A small_candidate | low |
| LookAtAssociationPairs × CmdPajek | A small_candidate | low |
| LookAtAssociationPairs × CmdGephi | A small_candidate | low |
| LookAtGroupData × CmdGIS | A small_candidate | low-medium |
| LookAtAssociationPairs × CmdNeo4j | C blocked_by_form_query_timeout | medium-high |
| LookAtGroupData × CmdNeo4j | C blocked_by_form_query_timeout | medium |
| LookAtNetworks × CmdGIS | B blocked_by_known_driver_issue | medium |
| LookAtNetworks × CmdPajek | B blocked_by_known_driver_issue | medium |
| LookAtAssociations × CmdUCINet | D new_export_family_needs_design | medium |
| LookAtPlace × CmdUCINet | D new_export_family_needs_design | medium |
| LookAtKinship × CmdUCINet | D new_export_family_needs_design | medium |
| LookAtAssociationPairs × CmdUCINet | D + C stacked | high |
| LookAtNetworks × CmdUCINet | D + B stacked | high |

| Bucket | Count |
|---|---:|
| A small_candidate | **4** |
| B blocked_by_known_driver_issue | 2 |
| C blocked_by_form_query_timeout | 2 |
| D new_export_family_needs_design | 3 |
| D + C stacked | 1 |
| D + B stacked | 1 |

---

## Recommended next 1-3 PRs (ranked by risk + ROI)

### PR 1 — AssociationPairs × CmdPajek + CmdGephi via small 4×5 fixture
- **Bucket:** A (×2 cells closed)
- **Risk:** low
- **Scope sketch:** add `Case("LookAtAssociationPairs", "CmdPajek", ".net", "*vertices")` and `Case("LookAtAssociationPairs", "CmdGephi", ".gdf", "nodedef")` to `tests/test_vba_pajek_gephi_cross_form.py::_CASES`; arrange `_fixture_for("LookAtAssociationPairs")` to return the 4×5 small fixture from `tests/test_vba_matrix_hard_forms.py` (or wire a small-fixture variant in `_all_fixtures()` gated for AssociationPairs only).
- **Closes 2 of 4 bucket-A cells in one PR; both are single-file exports with classifier shapes already proven on the sibling Associations form.**

### PR 2 — AssociationPairs × CmdGIS via the same 4×5 fixture
- **Bucket:** A
- **Risk:** low
- **Scope sketch:** add `LookAtAssociationPairs` to `tests/test_vba_cmdgis_other_forms.py::_FORMS_WITH_CMDGIS_TESTABLE_HERE`; same small-fixture wiring pattern as PR 1.  Single-file `.tab` export.
- **Could be combined with PR 1 if reviewer prefers; kept separate here for narrower diff and to surface any AssociationPairs-CmdGIS-specific quirks before bundling.**

### PR 3 — GroupData × CmdGIS via person_1 small fixture
- **Bucket:** A
- **Risk:** low-medium
- **Scope sketch:** add `LookAtGroupData` to `tests/test_vba_cmdgis_other_forms.py::_FORMS_WITH_CMDGIS_TESTABLE_HERE`; wire person_1 small fixture similarly.  Single-file `.tab` export.
- **Slight extra risk vs PRs 1/2 because GroupData's CmdRun behaviour (UPDATE-style backfill) differs from Associations/Place/Kinship CmdQuery — verify the chain works end-to-end before adding the assertion machinery.**

---

## Explicitly NOT recommended for autopilot implementation

The following 9 cells should **NOT** be picked up by an implementer without first opening a scope-defining brief from the maintainer:

| Cell | Why NOT autopilot |
|---|---|
| LookAtAssociationPairs × CmdNeo4j | Two stacked unknowns (matrix-timeout + Associations-sibling 0-file mode).  Probe first. |
| LookAtGroupData × CmdNeo4j | Multi-file Neo4j chain on heavy form; watcher-timeout risk.  Sequence after PR 3. |
| LookAtNetworks × CmdGIS / CmdPajek | Form_Open landmine #3.5 — needs minimal-injection scaffolding either via extending `tests/test_vba_networks_small_fixture.py` or refactoring cross-form test driver.  Either path is design work, not mechanical. |
| LookAtAssociations × CmdUCINet | New export family — needs design + probe pass to identify file format and assertion strictness BEFORE writing the test. |
| LookAtPlace × CmdUCINet | Same family blocker. |
| LookAtKinship × CmdUCINet | Same family blocker. |
| LookAtAssociationPairs × CmdUCINet | Stacked blockers (CmdUCINet + matrix-timeout). |
| LookAtNetworks × CmdUCINet | Stacked blockers (CmdUCINet + Networks Form_Open). |

The 4 bucket-A small_candidates above are the only cells worth promoting in narrow PRs without first investing in design/probe work.  The rest will close eventually but each requires a PR brief from the maintainer that names the specific blocker being lifted.
