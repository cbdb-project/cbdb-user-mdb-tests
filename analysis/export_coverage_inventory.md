# Export coverage inventory (LookAt form × export button)

**Generated:** 2026-05-05T08:24:33+00:00
**Generator:** `analysis/inventory_export_coverage.py`
**Companion JSON:** `reports/export_coverage_inventory.json`

Read-only inventory.  No MDB.  No Access COM.  Reads the VBA dump (`analysis/dump/vba/`), the control inventory (`analysis/dump/control_inventory.json`), and a curated test-coverage manifest declared in this script.  The manifest is cross-validated against the VBA dump on every run; drift surfaces under § Manifest drift.

## Legend

| Glyph | Status | Meaning |
|---|---|---|
| `✓` | `real_vba_covered` | Real-VBA test exists and passes for this (form, button). Both handler and UI button present. |
| `✓*` | `real_vba_covered_via_handler_dispatch` | Test passes via `Form_Timer` handler dispatch — handler exists, UI button is missing (P3 family).  Test exercises handler logic; missing button stays documented as a P3 issue. |
| `skip` | `real_vba_skipped` | Real-VBA test exists but is `pytest.mark.skip`'d (see per-cell skip_reason in JSON). |
| `skip*` | `real_vba_skipped_via_handler_dispatch` | Same as `skip`, but for a handler-dispatched cell whose UI button is missing. |
| `FAIL` | `real_vba_failing` | Real-VBA test exists, runs (does NOT skip), and fails on the current dump.  See per-cell `skip_reason` in JSON for the failure mode (typically a depth-check classifier gap for an unfamiliar file-shape family).  **NOT counted as covered.**  Distinct from `skip` (no `pytest.mark.skip`) and from `GAP` (test does exist).  See § Status semantics below. |
| `GAP` | `gap` | Both handler + button present, no real-VBA test, no static-only test either — true uncovered slice. |
| `static` | `unit_or_static_only` | Only static / source-level tests cover this cell (`tests/test_known_bugs.py` family); no real CmdQuery → CmdX chain. |
| `no-btn` | `missing_ui_button` | Handler exists in source but no control on the form (P3 missing-UI family — Issues #15-19). |
| `orphan` | `orphan_button_no_handler` | Button exists on the form but no `Sub <Cmd>_Click` in the dumped VBA — clicking would be a no-op. |
| `—` | `not_applicable` | Neither handler nor button — this form just doesn't host this export. |

**CmdKML caveat:** CmdKML is included in the matrix per the investigation brief, but the codebase has neither a `CmdKML` button nor a `CmdKML_Click` handler on any LookAt form.  KML output is implemented as a `ChkKML` checkbox option that other exports honour (e.g. LookAtOffice has `ChkPeopleKML` / `ChkOfficeKML`).  Every CmdKML cell therefore renders `—` (not_applicable); future KML coverage would be checkbox-driven, not a separate button slice.

**CmdUCInet → CmdUCINet:** the brief used `CmdUCInet` (lower-case `i`); the actual control + handler is `CmdUCINet` (capital `N`).  The matrix uses the real casing.

## Status semantics

Four statuses describe "there is or is not real-VBA test coverage for this cell".  They are NOT interchangeable; in particular `real_vba_failing` is its own bucket, never rolled into `real_vba_covered`.

| Status | Test exists? | Test skipped? | Test passes? | Counted as covered? | Eligible as Tier-1 (low-hanging) candidate? |
|---|:---:|:---:|:---:|:---:|:---:|
| `real_vba_covered` (✓ / ✓*) | yes | no | **yes** | **yes** | n/a (already covered) |
| `real_vba_skipped` (skip / skip*) | yes | **yes** (`pytest.mark.skip`) | n/a — not run | no | **yes**, but only when `skip_reason` matches a mechanical-fix pattern (currently: "no matrix fixture") |
| `real_vba_failing` (FAIL) | yes | no | **no** | **no** | **no** — failing tests are not skips with a mechanical fix; they're tests that run and fail.  Each failing cell must carry a `skip_reason` (re-used as failure-mode description) or substantive `notes` (>= 20 chars) so the failure mode is machinery-readable. |
| `gap` (GAP) | **no** | n/a | n/a | no | yes (always — ranked by family priority in Tier 2) |

Three deterministic invariants are checked at script-exit time and printed to stderr (with non-zero exit code) if violated:

  - **I1** — `real_vba_failing` cells are never counted as covered (`by_status` rolls them up as their own bucket).
  - **I2** — `real_vba_failing` cells never appear in `low_hanging_skips` (which is gated on status == `real_vba_skipped` by construction; this invariant pins that gate).
  - **I3** — every `real_vba_failing` cell's manifest entries carry either a `skip_reason` (re-used as failure-mode description) or substantive `notes` (>= 20 chars).

If a `FAIL` cell ever wants to graduate to `real_vba_covered`, the path is to (a) fix the underlying failure mode (typically: extend the depth-check classifier in `tests/test_vba_cmdneo4j_cross_form.py` for a new file-shape family) and (b) flip the manifest entry's `status` from `real_vba_failing` to `covered`.  The flip must come AFTER the test actually passes, not before.

## Summary

- **Cells total:** 80 (10 forms × 8 buttons)
- `not_applicable` (—): **40**
- `real_vba_covered` (✓): **17**
- `gap` (GAP): **12**
- `real_vba_skipped` (skip): **5**
- `real_vba_covered_via_handler_dispatch` (✓*): **3**
- `real_vba_skipped_via_handler_dispatch` (skip*): **2**
- `missing_ui_button` (no-btn): **1**

## Coverage matrix

| Form | CmdGIS | CmdNeo4j | CmdPajek | CmdGephi | CmdGUESS | CmdKML | CmdUCINet | CmdGISPeople |
|---|---|---|---|---|---|---|---|---|
| **LookAtEntry** | ✓ | ✓ | — | — | — | — | — | — |
| **LookAtTexts** | ✓ | ✓ | — | — | — | — | — | — |
| **LookAtAssociations** | ✓ | skip | ✓ | ✓ | — | — | GAP | — |
| **LookAtOffice** | ✓ | ✓ | — | — | ✓* | — | — | ✓ |
| **LookAtPlace** | ✓* | skip | ✓ | ✓ | — | — | GAP | — |
| **LookAtKinship** | ✓ | ✓ | ✓* | — | ✓ | — | GAP | — |
| **LookAtStatus** | ✓ | skip | skip* | skip* | — | — | no-btn | — |
| **LookAtAssociationPairs** | GAP | GAP | GAP | GAP | — | — | GAP | — |
| **LookAtNetworks** | GAP | skip | GAP | — | skip | — | GAP | — |
| **LookAtGroupData** | ✓ | GAP | — | — | — | — | — | — |

## Per-cell detail (non-trivial cells)

Cells with status `not_applicable` are omitted to keep the noise down.

### LookAtEntry × CmdGIS — `real_vba_covered`

- Handler in source: yes
- Button on form: yes
- Why: Real-VBA test exists and passes for this cell.
- Test: `tests/test_vba_export.py::test_lookatentry_cmd_gis` — **covered**; notes: byte-level golden compare; the original CmdGIS test

### LookAtEntry × CmdNeo4j — `real_vba_covered`

- Handler in source: yes
- Button on form: yes
- Why: Real-VBA test exists and passes for this cell.
- Test: `tests/test_vba_cmdneo4j_cross_form.py::test_cmd_neo4j_produces_files[LookAtEntry]` — **covered**; notes: min_files=7 + per-shape depth + LookAtEntry-specific structural assertion that the file set is exactly {People, PeopleEntry, Places, PeoplePlaces, PersonPlaceCodes, EntryCodes, AssocCodes} AND no InstitutionCodes file (Issue #9 LATENT-gate pin: `ENTRY_DATA.c_inst_code > 0 = 0` on this dump).  Promoted from skip 2026-05-04 after the Issue #9 reverification probe verified the chain end-to-end with c_entry_code=101.
- Static-only note: tests/test_known_bugs.py::test_bug9_lookat_entry_cmdneo4j_with_wrong_var — pins source typo + LATENT-gate (Issue #9)
- Static-only note: tests/test_vba_bug_behaviors.py::test_bug9_lookat_entry_cmdneo4j_with_institutions_fixture — runtime: chain finishes cleanly without ERR for non-inst fixture

### LookAtTexts × CmdGIS — `real_vba_covered`

- Handler in source: yes
- Button on form: yes
- Why: Real-VBA test exists and passes for this cell.
- Test: `tests/test_vba_cmdgis_other_forms.py::test_cmd_gis_produces_file` — **covered**; notes: structural

### LookAtTexts × CmdNeo4j — `real_vba_covered`

- Handler in source: yes
- Button on form: yes
- Why: Real-VBA test exists and passes for this cell.
- Test: `tests/test_vba_cmdneo4j_cross_form.py::test_cmd_neo4j_produces_files[LookAtTexts]` — **covered**; notes: min_files=4 + per-shape depth

### LookAtAssociations × CmdGIS — `real_vba_covered`

- Handler in source: yes
- Button on form: yes
- Why: Real-VBA test exists and passes for this cell.
- Test: `tests/test_vba_cmdgis_other_forms.py::test_cmd_gis_produces_file` — **covered**; notes: structural

### LookAtAssociations × CmdNeo4j — `real_vba_skipped`

- Handler in source: yes
- Button on form: yes
- Why: Real-VBA test exists but is currently skipped.  See skip_reason on each manifest entry.
- Test: `tests/test_vba_cmdneo4j_cross_form.py::test_cmd_neo4j_produces_files[LookAtAssociations]` — **skipped**; skip_reason: produces 0 files in directory mode — needs investigation alongside Place

### LookAtAssociations × CmdPajek — `real_vba_covered`

- Handler in source: yes
- Button on form: yes
- Why: Real-VBA test exists and passes for this cell.
- Test: `tests/test_vba_pajek_gephi_cross_form.py::test_export_produces_file[LookAtAssociations-CmdPajek]` — **covered**; notes: shape: .net / *vertices header

### LookAtAssociations × CmdGephi — `real_vba_covered`

- Handler in source: yes
- Button on form: yes
- Why: Real-VBA test exists and passes for this cell.
- Test: `tests/test_vba_pajek_gephi_cross_form.py::test_export_produces_file[LookAtAssociations-CmdGephi]` — **covered**; notes: shape: .gdf / nodedef header

### LookAtAssociations × CmdUCINet — `gap`

- Handler in source: yes
- Button on form: yes
- Why: Button + handler both present, no entry in the real-VBA test manifest, no static-only test note either.

### LookAtOffice × CmdGIS — `real_vba_covered`

- Handler in source: yes
- Button on form: yes
- Why: Real-VBA test exists and passes for this cell.
- Test: `tests/test_vba_cmdgis_other_forms.py::test_cmd_gis_produces_file` — **covered**; notes: structural

### LookAtOffice × CmdNeo4j — `real_vba_covered`

- Handler in source: yes
- Button on form: yes
- Why: Real-VBA test exists and passes for this cell.
- Test: `tests/test_vba_cmdneo4j_cross_form.py::test_cmd_neo4j_produces_files[LookAtOffice]` — **covered**; notes: min_files=4 + per-shape depth.  Promoted from `real_vba_failing` 2026-05-04 in PR cover/lookatoffice-cmdneo4j-peopleoffice: added `(NameID, OfficeCode) -> PeopleOffice` 2-col entry to `_NEO4J_SHAPES_BY_TWO_COLS` with required cols `[NameID, OfficeCode, OfficeAddrID, SocialInstID, PostingFirstYear, PostingLastYear, PostingDynasty]`.  Header literal verified at `Form_LookAtOffice.vb:947`.  Live-verified end-to-end with --include-vba: 6 files produced (People, PeopleOffice, Places, PeoplePlaces, PersonPlaceCodes, OfficeCode-codes-via-loose-check), 5/6 classified strictly via the depth check; the 6th (`OfficeCode_*.csv`, header `OfficeCode,OfficeTrans,OfficePinyin[,OfficeHZ]` per `Form_LookAtOffice.vb:1324-1326`) passes via the loose-check fallback because the classifier doesn't have a single-column entry for `OfficeCode` yet — non-failing today; tightening it is a future hygiene follow-up if needed.  InstitutionCodes block (line 1399) is gated like LookAtEntry's and is absent on this dump (no `c_inst_code > 0` rows in the office-relevant scratch table).  Inventory previously (PR 89d9a63) had marked this `covered` based on an assumption from the test file's _SPECS list; PR fix/cmdneo4j-classifier-lookattexts surfaced the actual classifier-side failure and correctly downgraded to `real_vba_failing`; this PR fixes the classifier and honestly re-promotes after live verification.

### LookAtOffice × CmdGUESS — `real_vba_covered_via_handler_dispatch`

- Handler in source: yes
- Button on form: no
- Why: Handler exists in source but the form has no control with this name (P3 missing-UI family).  The cross-form test exercises the handler via Form_Timer dispatch (no UI click needed), so handler logic IS covered; the missing UI button is documented separately as a P3 issue.
- Test: `tests/test_vba_cmdguess_cross_form.py::test_cmd_guess_produces_file[LookAtOffice]` — **covered**; notes: .gdf shape
- Static-only note: (plus Issue #19: P3 missing UI button)

### LookAtOffice × CmdGISPeople — `real_vba_covered`

- Handler in source: yes
- Button on form: yes
- Why: Real-VBA test exists and passes for this cell.
- Test: `tests/test_vba_cmdgispeople_office.py::test_cmd_gis_people` — **covered**; notes: only host; people-side GIS export distinct from office-side CmdGIS

### LookAtPlace × CmdGIS — `real_vba_covered_via_handler_dispatch`

- Handler in source: yes
- Button on form: no
- Why: Handler exists in source but the form has no control with this name (P3 missing-UI family).  The cross-form test exercises the handler via Form_Timer dispatch (no UI click needed), so handler logic IS covered; the missing UI button is documented separately as a P3 issue.
- Test: `tests/test_vba_cmdgis_other_forms.py::test_cmd_gis_produces_file` — **covered**; notes: structural; passes only thanks to driver-side _PER_FORM_CMDGIS_PATCHES rewrite of GISFrame -> CodeFrame (Issue #4 latent typo workaround)
- Static-only note: (plus Issue #15: P3 missing UI button — covered by tests/test_known_bugs.py::test_bugs_15_to_19_orphan_export_handlers)

### LookAtPlace × CmdNeo4j — `real_vba_skipped`

- Handler in source: yes
- Button on form: yes
- Why: Real-VBA test exists but is currently skipped.  See skip_reason on each manifest entry.
- Test: `tests/test_vba_cmdneo4j_cross_form.py::test_cmd_neo4j_produces_files[LookAtPlace]` — **skipped**; skip_reason: fires 'Item not found in this collection' mid-body — Issue #7 (real CBDB bug)
- Static-only note: tests/test_known_bugs.py::test_bug7_lookat_place_cmdneo4j_select_missing_dynasty_female — pins source SELECT (Issue #7)

### LookAtPlace × CmdPajek — `real_vba_covered`

- Handler in source: yes
- Button on form: yes
- Why: Real-VBA test exists and passes for this cell.
- Test: `tests/test_vba_pajek_gephi_cross_form.py::test_export_produces_file[LookAtPlace-CmdPajek]` — **covered**; notes: shape: .net / *vertices header

### LookAtPlace × CmdGephi — `real_vba_covered`

- Handler in source: yes
- Button on form: yes
- Why: Real-VBA test exists and passes for this cell.
- Test: `tests/test_vba_pajek_gephi_cross_form.py::test_export_produces_file[LookAtPlace-CmdGephi]` — **covered**; notes: shape: .gdf / nodedef header

### LookAtPlace × CmdUCINet — `gap`

- Handler in source: yes
- Button on form: yes
- Why: Button + handler both present, no entry in the real-VBA test manifest, no static-only test note either.

### LookAtKinship × CmdGIS — `real_vba_covered`

- Handler in source: yes
- Button on form: yes
- Why: Real-VBA test exists and passes for this cell.
- Test: `tests/test_vba_cmdgis_other_forms.py::test_cmd_gis_produces_file` — **covered**; notes: structural; subform requery via _SUBFORMS_TO_REQUERY

### LookAtKinship × CmdNeo4j — `real_vba_covered`

- Handler in source: yes
- Button on form: yes
- Why: Real-VBA test exists and passes for this cell.
- Test: `tests/test_vba_cmdneo4j_cross_form.py::test_cmd_neo4j_produces_files[LookAtKinship]` — **covered**; notes: min_files=4 + per-shape depth

### LookAtKinship × CmdPajek — `real_vba_covered_via_handler_dispatch`

- Handler in source: yes
- Button on form: no
- Why: Handler exists in source but the form has no control with this name (P3 missing-UI family).  The cross-form test exercises the handler via Form_Timer dispatch (no UI click needed), so handler logic IS covered; the missing UI button is documented separately as a P3 issue.
- Test: `tests/test_vba_pajek_gephi_cross_form.py::test_export_produces_file[LookAtKinship-CmdPajek]` — **covered**; notes: shape: .net / *vertices header

### LookAtKinship × CmdGUESS — `real_vba_covered`

- Handler in source: yes
- Button on form: yes
- Why: Real-VBA test exists and passes for this cell.
- Test: `tests/test_vba_cmdguess_cross_form.py::test_cmd_guess_produces_file[LookAtKinship]` — **covered**; notes: .gdf shape

### LookAtKinship × CmdUCINet — `gap`

- Handler in source: yes
- Button on form: yes
- Why: Button + handler both present, no entry in the real-VBA test manifest, no static-only test note either.

### LookAtStatus × CmdGIS — `real_vba_covered`

- Handler in source: yes
- Button on form: yes
- Why: Real-VBA test exists and passes for this cell.
- Test: `tests/test_vba_cmdgis_other_forms.py::test_cmd_gis_produces_file` — **covered**; notes: structural assertion (header + non-empty cols)

### LookAtStatus × CmdNeo4j — `real_vba_skipped`

- Handler in source: yes
- Button on form: yes
- Why: Real-VBA test exists but is currently skipped.  See skip_reason on each manifest entry.
- Test: `tests/test_vba_cmdneo4j_cross_form.py::test_cmd_neo4j_produces_files[LookAtStatus]` — **skipped**; skip_reason: chain post-cleanup invalidates subform recordset rebind; downstream CmdNeo4j reads RecordCount=0 (same family as Pajek/Gephi Status skip)

### LookAtStatus × CmdPajek — `real_vba_skipped_via_handler_dispatch`

- Handler in source: yes
- Button on form: no
- Why: Handler exists in source but the form has no control with this name (P3 missing-UI family).  The cross-form test exercises the handler via Form_Timer dispatch (no UI click needed), so handler logic IS covered; the missing UI button is documented separately as a P3 issue.
- Test: `tests/test_vba_pajek_gephi_cross_form.py::test_export_produces_file[LookAtStatus-CmdPajek]` — **skipped**; skip_reason: chain to CmdPajek/CmdGephi reads RecordCount=0 after subform rebind cleanup; same as Status Neo4j skip
- Static-only note: (plus Issue #16: P3 missing UI button)

### LookAtStatus × CmdGephi — `real_vba_skipped_via_handler_dispatch`

- Handler in source: yes
- Button on form: no
- Why: Handler exists in source but the form has no control with this name (P3 missing-UI family).  The cross-form test exercises the handler via Form_Timer dispatch (no UI click needed), so handler logic IS covered; the missing UI button is documented separately as a P3 issue.
- Test: `tests/test_vba_pajek_gephi_cross_form.py::test_export_produces_file[LookAtStatus-CmdGephi]` — **skipped**; skip_reason: same chain-cleanup family as Status CmdPajek
- Static-only note: (plus Issue #17: P3 missing UI button)

### LookAtStatus × CmdUCINet — `missing_ui_button`

- Handler in source: yes
- Button on form: no
- Why: Handler exists in source but the form has no control with this name (P3 missing-UI family — Issues #15-19 documented).  No real-VBA test exercises this handler.
- Static-only note: (plus Issue #18: P3 missing UI button)

### LookAtAssociationPairs × CmdGIS — `gap`

- Handler in source: yes
- Button on form: yes
- Why: Button + handler both present, no entry in the real-VBA test manifest, no static-only test note either.

### LookAtAssociationPairs × CmdNeo4j — `gap`

- Handler in source: yes
- Button on form: yes
- Why: Button + handler both present, no entry in the real-VBA test manifest, no static-only test note either.

### LookAtAssociationPairs × CmdPajek — `gap`

- Handler in source: yes
- Button on form: yes
- Why: Button + handler both present, no entry in the real-VBA test manifest, no static-only test note either.

### LookAtAssociationPairs × CmdGephi — `gap`

- Handler in source: yes
- Button on form: yes
- Why: Button + handler both present, no entry in the real-VBA test manifest, no static-only test note either.

### LookAtAssociationPairs × CmdUCINet — `gap`

- Handler in source: yes
- Button on form: yes
- Why: Button + handler both present, no entry in the real-VBA test manifest, no static-only test note either.

### LookAtNetworks × CmdGIS — `gap`

- Handler in source: yes
- Button on form: yes
- Why: Button + handler both present, no entry in the real-VBA test manifest, no static-only test note either.

### LookAtNetworks × CmdNeo4j — `real_vba_skipped`

- Handler in source: yes
- Button on form: yes
- Why: Real-VBA test exists but is currently skipped.  See skip_reason on each manifest entry.
- Test: `tests/test_vba_cmdneo4j_cross_form.py::(implicit — LookAtNetworks is in _FORMS_WITH_CMDGUESS-style skip family; no explicit Spec entry in _SPECS)` — **skipped**; skip_reason: matrix CmdRun skipped (high-degree anchor expansion) + default full injection Form_Open deadlock (AGENTS landmine #3.5; minimal injection works for Form_Open per tests/test_vba_networks_small_fixture.py)
- Static-only note: tests/test_known_bugs.py::test_bug8_lookat_networks_cmdneo4j_select_missing_xy — pins source SELECT (Issue #8)

### LookAtNetworks × CmdPajek — `gap`

- Handler in source: yes
- Button on form: yes
- Why: Button + handler both present, no entry in the real-VBA test manifest, no static-only test note either.

### LookAtNetworks × CmdGUESS — `real_vba_skipped`

- Handler in source: yes
- Button on form: yes
- Why: Real-VBA test exists but is currently skipped.  See skip_reason on each manifest entry.
- Test: `tests/test_vba_cmdguess_cross_form.py::test_cmd_guess_produces_file[LookAtNetworks]` — **skipped**; skip_reason: CmdRun times out on high-degree anchors (AGENTS landmine #3.5)

### LookAtNetworks × CmdUCINet — `gap`

- Handler in source: yes
- Button on form: yes
- Why: Button + handler both present, no entry in the real-VBA test manifest, no static-only test note either.

### LookAtGroupData × CmdGIS — `real_vba_covered`

- Handler in source: yes
- Button on form: yes
- Why: Real-VBA test exists and passes for this cell.
- Test: `tests/test_vba_cmdgis_other_forms.py::test_cmd_gis_groupdata_clean_branches` — **covered**; notes: Clean-branches coverage (Status / Office / Addr) using matrix_hard_forms's groupdata_person_1_small fixture.  Explicitly excludes Entry (Issue #6 P1 fires JET 3061; pinned by tests/test_known_bugs.py::test_bug6_groupdata_query_entry_wrong_field [static] AND tests/test_vba_bug_behaviors.py::test_bug6_lookat_groupdata_query_entry_fires_no_such_field [runtime]).  Excludes Text (person_1 has 0 BIOG_TEXT_DATA rows -> WriteGIS_Text bails on RecCount=0; benign 0-files state, not a coverage gap).  Excludes ChkGisOfficePeople (the Office_OfficeOffice variant alone exercises queryOffice -> WriteGIS_OfficeOffice cleanly; OfficePeople adds a second writer without coverage value).  Probe evidence: analysis/probe_groupdata_cmdgis_subcalls.py + analysis/groupdata_cmdgis_subcall_trace.md.

### LookAtGroupData × CmdNeo4j — `gap`

- Handler in source: yes
- Button on form: yes
- Why: Button + handler both present, no entry in the real-VBA test manifest, no static-only test note either.

## Manifest drift

**Clean — manifest entries all match the VBA dump and control inventory.**

## Recommended next real-VBA slice

*Suggestion only — this PR ships inventory and does NOT implement.  Review the cell's manifest entry + static notes before committing to a slice.*

### Tier 1 — low-hanging skipped tests (mechanical fix)

*(No skipped tests with a mechanical fix class.  All current skips are blocked on harder issues — chain cleanup family, matrix CmdRun timeout family, Form_Open deadlock family.)*

### Tier 2 — pure `gap` cells (button + handler exist, no test of any kind)

Cells ranked by family priority (lower = lower-risk, smaller blast radius).  Note: many gap cells in this list (Networks family, AssociationPairs family, GroupData family, the entirely-untested CmdUCINet family) sit behind known blockers and are NOT good candidates for a small first slice — Tier 1 is preferred.

| # | Form | Button | Family priority | Known family blocker |
|---:|---|---|---:|---|
| 1 | LookAtAssociationPairs | CmdGIS | 1 | matrix CmdQuery times out — no CrossFixture promoted to a passing assertion |
| 2 | LookAtNetworks | CmdGIS | 1 | Form_Open hang / CmdRun timeout (AGENTS landmine #3.5) |
| 3 | LookAtAssociationPairs | CmdGephi | 3 | matrix CmdQuery times out — no CrossFixture promoted to a passing assertion |
| 4 | LookAtAssociationPairs | CmdPajek | 3 | matrix CmdQuery times out — no CrossFixture promoted to a passing assertion |
| 5 | LookAtNetworks | CmdPajek | 3 | Form_Open hang / CmdRun timeout (AGENTS landmine #3.5) |
| 6 | LookAtAssociationPairs | CmdNeo4j | 5 | matrix CmdQuery times out — no CrossFixture promoted to a passing assertion |
| 7 | LookAtGroupData | CmdNeo4j | 5 | matrix CmdQuery has issues; depends on a CrossFixture that doesn't exist for GroupData |
| 8 | LookAtAssociationPairs | CmdUCINet | 6 | matrix CmdQuery times out — no CrossFixture promoted to a passing assertion |
| 9 | LookAtAssociations | CmdUCINet | 6 | entirely untested handler family — no existing test infrastructure to extend |
| 10 | LookAtKinship | CmdUCINet | 6 | entirely untested handler family — no existing test infrastructure to extend |
| 11 | LookAtNetworks | CmdUCINet | 6 | Form_Open hang / CmdRun timeout (AGENTS landmine #3.5) |
| 12 | LookAtPlace | CmdUCINet | 6 | entirely untested handler family — no existing test infrastructure to extend |
