# Design: Standardized Fixed-Skeleton Report

**Status**: Design approved, implementation pending.

## Problem

The current report structure is entirely driven by the `ISSUES` dict in
`reports/generate_report.py`.  Since `ISSUES` is rebuilt from scratch on every
build, a build where all tests pass produces a near-empty report — no coverage
matrix, no appendix, no way to tell whether a form was tested at all vs.
simply not mentioned.

Additionally, a previous implementation error introduced a `cbdb_replay/`
module that re-implements VBA SQL logic in Python.  This approach is strictly
prohibited (see AGENTS.md "ABSOLUTE PROHIBITION" section) because a Python
rewrite of buggy VBA has the same bug, producing false-green test results.

---

## Design Goals

1. **Fixed structure every time.** Every build report has the same sections
   regardless of whether bugs are found.  A reader can immediately see which
   forms/buttons were tested and what the result was.

2. **Real VBA only.** All coverage data comes from running the original VBA
   via Access COM (`--include-vba`).  No Python replays, no SQL translations.
   Index-year and address calculations are verified by running the original
   source (VBA/PHP), not Python derivations.

3. **Automatically populated.** The coverage matrix is filled from the pytest
   JSON report (`--json-report`) — no manual data entry.

4. **Appendix always present.** The index-year/addr drift appendix appears in
   every report, with a clear "data not available — run step X" message if the
   analysis scripts have not been run.

---

## Report Structure (target)

```
# CBDB User MDB — Issues Report (build YYYYMMDD)

## Coverage Matrix                          ← NEW: always present
[table: 10 forms × N buttons, PASS/FAIL/N/A]

## Issues Found in This Build               ← existing ISSUES dict
[P0 / P1 / P2 / … sections, empty if none]

## Severity Legend                          ← existing

## Appendix A: index_year / index_addr drift  ← existing, always shown
[populated by collect_index_year_diffs.py]

## Appendix B: Schema changes vs previous build  ← NEW (optional)
[columns added/removed, populated by compare_schema.py]

## Closing Note                             ← existing
```

---

## Coverage Matrix Specification

### Forms and buttons to always test

| Form | CmdQuery | CmdGIS | CmdNeo4j | CmdPajek | CmdGephi | CmdUCINet | CmdKML | CmdGUESS | CmdRun |
|------|----------|--------|----------|----------|----------|-----------|--------|----------|--------|
| LookAtEntry        | ✓ | ✓ | ✓ | — | — | — | ✓ | — | — |
| LookAtStatus       | ✓ | — | — | ✓ | ✓ | ✓ | — | — | — |
| LookAtTexts        | ✓ | — | — | — | — | — | — | — | — |
| LookAtPlace        | ✓ | ✓ | ✓ | — | — | — | — | — | — |
| LookAtAssociations | ✓ | ✓ | ✓ | — | — | ✓ | — | — | — |
| LookAtOffice       | ✓ | ✓ | — | — | — | — | — | ✓ | — |
| LookAtKinship      | — | ✓ | — | ✓ | — | ✓ | — | ✓ | ✓ |
| LookAtNetworks     | — | — | ✓ | — | — | — | — | — | ✓ |
| LookAtGroupData    | — | — | ✓ | — | — | — | — | — | ✓ |
| LookAtAssocPairs   | ✓ | — | ✓ | ✓ | — | — | — | — | — |

Cell values:
- `PASS (N rows)` — test ran and passed, row count shown
- `FAIL — <short reason>` — test ran and failed, links to Issue #N
- `ERROR` — test raised an exception (infra failure, not a bug)
- `SKIP` — test marked skip (known infra limitation)
- `N/A` — button does not exist on this form
- `NOT RUN` — test not yet written for this cell

### Test → matrix cell mapping

The coverage matrix is populated by parsing `reports/pytest_report_YYYYMMDD.json`
(produced by `pytest --json-report`).  Each test name maps to a cell:

```
test_vba_matrix[*LookAtEntry*]              → LookAtEntry / CmdQuery
test_vba_export::test_lookatentry_cmd_gis   → LookAtEntry / CmdGIS
test_vba_export::test_lookatentry_cmd_neo4j → LookAtEntry / CmdNeo4j
test_vba_pajek_gephi_cross_form[*Status*Pajek*]  → LookAtStatus / CmdPajek
test_vba_cmducinet_kinship[*]               → LookAtKinship / CmdUCINet
...
```

A mapping config file (`docs/test-matrix-mapping.json`) will enumerate all
mappings explicitly so the matrix builder is not fragile to test name changes.

---

## Implementation Plan

### Step 1: Mapping config (no code changes to generate_report.py)

Create `docs/test-matrix-mapping.json`:

```json
{
  "LookAtEntry": {
    "CmdQuery": ["test_vba_matrix[top_entry_code_"],
    "CmdGIS":   ["test_lookatentry_cmd_gis"],
    "CmdNeo4j": ["test_vba_cmdneo4j_cross_form[LookAtEntry"]
  },
  ...
}
```

### Step 2: Matrix builder script

Create `analysis/build_coverage_matrix.py`:
- Reads `reports/pytest_report_YYYYMMDD.json`
- Applies `docs/test-matrix-mapping.json`
- Emits `reports/coverage_matrix.json`

### Step 3: generate_report.py additions

Add two new sections rendered before ISSUES:

```python
COVERAGE_MATRIX_JSON = REPO / "reports" / "coverage_matrix.json"

def _add_coverage_matrix(doc, is_en, Z):
    """Render form × button coverage table, always present."""
    data = json.loads(COVERAGE_MATRIX_JSON.read_text()) \
           if COVERAGE_MATRIX_JSON.exists() \
           else {}
    # ... render table with PASS/FAIL/N/A cells ...
    # If data is empty: show "Coverage matrix not generated — run step X"
```

### Step 4: Update canonical workflow (AGENTS.md step 5b)

Add after step 5 (Run tests):

```
5b. **Build coverage matrix**:
    python analysis/build_coverage_matrix.py \
        --report reports/pytest_report_YYYYMMDD.json
    (reads the JSON report from step 5, emits reports/coverage_matrix.json)
```

### Step 5: Always-present appendix

`_add_index_drift_appendix()` already exists.  Make it show a placeholder when
the source JSON files are missing, instead of silently omitting the section.

Add step to workflow:
```
5c. **Regenerate index-year drift appendix**:
    python analysis/collect_index_year_diffs.py
    (queries DATA mdb via pyodbc, emits reports/index_drift_*.json)
```

---

## No-Python-Rewrite Constraint

All tests that populate the coverage matrix MUST drive real VBA via Access COM.
Specifically:

- **Forbidden**: Reading `analysis/dump/vba/Form_LookAt*.vb`, translating the
  SQL to Python, running it via pyodbc against the DATA mdb.  This is what
  `cbdb_replay/` does and it is prohibited.

- **Required**: Opening the form in Access via `VbaSession`, calling the real
  `Cmd*_Click` handler via `click_via_timer`, reading results from
  `ZZ_SCRATCH_*` tables.

- **For index-year/addr cross-checks**: Run the original calculation
  (VBA formula or server-side PHP), never re-derive in Python.  The
  `collect_index_year_diffs.py` script compares the mdb's stored
  `c_index_year` / `c_index_addr_id` against the cbdb-online SQLite
  snapshot — it does NOT recalculate either side in Python.

### `cbdb_replay/` migration path

| File | Status | Action |
|------|--------|--------|
| `cbdb_replay/lookatentry.py` | Legacy, DO NOT extend | Replace with real-VBA test when capacity allows |
| `cbdb_replay/TEMPLATE_lookat.py` | Legacy template for Python replays | **Delete** — new tests use VBA COM, not this template |
| `cbdb_replay/exports.py` | Used by `test_exports.py` for format-only unit tests (no SQL) | Retain: this tests the export *format* not the VBA SQL |
| `cbdb_replay/common.py` | Shared helpers | Retain if still needed by tests not slated for replacement |

---

## Appendix B: Schema Changes (optional, future)

Compare `analysis/dump/tables.json` (current build) against the previous
build's archived copy.  Surface:
- New tables/columns (data model expansion)
- Removed tables/columns (potentially breaking changes)

Script: `analysis/compare_schema.py` (not yet written).

---

## Open Questions

1. **`cbdb_replay/` test_lookatentry.py etc.** — these non-VBA tests currently
   run in the fast suite (no `--include-vba`).  Until they are replaced with
   real-VBA equivalents, the fast suite will continue to run them.  The
   coverage matrix should mark their cells as `REPLAY (not VBA)` to distinguish
   from real-VBA results.

2. **LookAtNetworks** — known to deadlock under default injection (AGENTS.md
   landmine #4).  The matrix cell will show `SKIP (deadlock risk)` until the
   underlying issue is resolved.

3. **Test-matrix-mapping.json maintenance** — every new VBA test must update
   this file.  Consider making the mapping auto-detected from test markers
   (`@pytest.mark.form("LookAtEntry")`, `@pytest.mark.button("CmdGIS")`)
   instead of a separate config file.
