# CBDB User MDB — Issues Report

_A respectful summary of issues uncovered during regression testing._

Dear maintainer,

Below is a summary of the issues we uncovered while building an automated regression-test suite for the CBDB User MDB. We hope this report is useful as you continue your wonderful stewardship of this dataset, and we sincerely thank you for the immense work that has gone into building it.

The issues are ordered by severity (P0 highest). Each entry includes a concise description, step-by-step user reproduction, screenshots where the issue is visible in the Access UI, and a suggested fix. None of these are urgent; they are documented so they can be addressed at the maintainer's convenience.

## Coverage Matrix — Form × Button Test Results

| Form | CmdQuery | CmdGIS | CmdNeo4j | CmdPajek | CmdGephi | CmdUCINet | CmdKML | CmdGUESS | CmdRun |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| LookAtEntry | ✗ FAIL | ✓ | ⚠ ERR | — | — | — | —? | — | — |
| LookAtStatus | ⚠ ERR | — | — | ✗ FAIL | ⚠ ERR | ⚠ ERR | — | — | — |
| LookAtTexts | ⚠ ERR | — | — | — | — | — | — | — | — |
| LookAtPlace | ⚠ ERR | ✗ FAIL | ⚠ ERR | — | — | — | — | — | — |
| LookAtAssociations | ⚠ ERR | ⚠ ERR | ⚠ ERR | — | — | ⚠ ERR | — | — | — |
| LookAtOffice | ⚠ ERR | ⚠ ERR | — | — | — | — | — | ⚠ ERR | — |
| LookAtKinship | — | ⚠ ERR | — | ✓ | — | ⚠ ERR | — | ⚠ ERR | ⚠ ERR |
| LookAtNetworks | — | — | ~ SKIP | — | — | — | — | — | ~ SKIP |
| LookAtGroupData | — | — | ⚠ ERR | — | — | — | — | — | ✗ FAIL |
| LookAtAssocPairs | ⚠ ERR | — | ⚠ ERR | ⚠ ERR | — | — | — | — | — |

_PASS: 2 · FAIL: 4 · ERROR: 22 · SKIP: 2 · NOT RUN: 1 · N/A: 59_

## Table of Contents

- [P2 — Silent display](#p2--silent-display)
  - [Issue #1 — LookAtEntry: c_entry_desc backfill is NULL for all rows when entry_code = 36 (jinshi general)](#issue-1--lookatentry-c_entry_desc-backfill-is-null-for-all-rows-when-entry_code--36-jinshi-general)
  - [Issue #2 — LookAtGroupData: CmdRun does not backfill c_name from BIOG_MAIN](#issue-2--lookatgroupdata-cmdrun-does-not-backfill-c_name-from-biog_main)
- [Severity legend](#severity-legend)
- [Appendix — c_index_year / c_index_addr_id drift vs the cbdb-online-main-server snapshot (differences need per-row classification before being filed as bugs)](#appendix--c_index_year--c_index_addr_id-drift-vs-the-cbdb-online-main-server-snapshot-differences-need-per-row-classification-before-being-filed-as-bugs)
- [Closing note](#closing-note)

## Severity legend

- P0 — Silent data corruption: data is wrong or missing without an error popup.
- P1 — Visible runtime crash: a popup appears, the operation aborts.
- P2 — Silent display: form fields render blank when they should show data.
- P3 — Missing UI: a feature exists in code but no button invokes it.
- P4 — Setup: one-time hurdle on each new install.
- P5 — Dormant / latent / not currently reproducible: kept as historical record; we re-checked on the current dump and could not trigger the symptom.

## P2 — Silent display

### Issue #1 — LookAtEntry: c_entry_desc backfill is NULL for all rows when entry_code = 36 (jinshi general)

**Affected sub:** `Form_LookAtEntry.CmdQuery_Click`

**Severity:** P2 — Silent display issue: 92,545 rows affected.  The user can see the blank c_entry_desc column in the result grid, but Access shows no error — making it easy to overlook.  Exports (GIS, Neo4j, KML) that reference this column will also carry the blank.

#### Description

When the user runs a LookAtEntry query filtered to entry code 36 (examination: jinshi general), the result table ZZ_SCRATCH_ENTRY is populated with 92,545 rows but the c_entry_desc column is NULL for every row.  The expected value is 'examination: jinshi (general)'.

The CmdQuery_Click handler successfully inserts rows from ENTRY_DATA joined to ENTRY_CODES, but the c_entry_desc backfill step does not write the description for this specific entry code.  All other columns appear to be filled normally.  The missing description means the on-screen result grid shows a blank entry-type column for every record, which is misleading — the user sees results but cannot identify what type of examination each record represents.

Detected by: test_vba_full_matrix[top_entry_code_36_unfiltered] — assertion 'c_entry_desc backfill wrong' with 92,545 affected rows.

#### Steps to reproduce

1. Open CBDB_BJ_User.mdb in Microsoft Access.
2. From the Navigation Pane, open the form **LookAtEntry**.
3. In the Entry Code picker, select entry code **36** (label: 'examination: jinshi (general)').
4. Leave dynasty, address, and year filters blank.
5. Click **Run Query** (CmdQuery button).
6. When the query completes, inspect the result grid: the entry-type description column (c_entry_desc) is blank for every row.
7. SQL verification: `SELECT TOP 5 c_entry_code, c_entry_desc FROM ZZ_SCRATCH_ENTRY` returns (36, NULL) for all rows.

#### Suggested fix

Locate the backfill step in Form_LookAtEntry.CmdQuery_Click that sets c_entry_desc for ZZ_SCRATCH_ENTRY rows.  Verify that the JOIN to ENTRY_CODES on c_entry_code = 36 is not inadvertently filtered out or that the UPDATE / backfill SQL matches the column name exactly.  After the fix, `SELECT c_entry_desc FROM ZZ_SCRATCH_ENTRY WHERE c_entry_code = 36 LIMIT 1` should return 'examination: jinshi (general)'.

### Issue #2 — LookAtGroupData: CmdRun does not backfill c_name from BIOG_MAIN

**Affected sub:** `Form_LookAtGroupData.CmdRun_Click`

**Severity:** P2 — Silent display issue: CmdRun completes without any error message, but the c_name column in the result is blank.  The user has no indication that the backfill failed.

#### Description

When the user seeds a person ID into LookAtGroupData and clicks CmdRun, the handler is expected to run an UPDATE query that joins ZZ_SCRATCH_IMPORT_PEOPLE to BIOG_MAIN and fills in c_name (and c_dynasty) for each seeded row.  In this build the UPDATE does not execute successfully: after CmdRun completes, c_name remains NULL in ZZ_SCRATCH_IMPORT_PEOPLE.

The result is that the group-data import display shows empty name cells.  The user has no indication that the backfill failed — CmdRun does not surface an error.

Detected by: test_hard_form_query_small_fixture[groupdata_person_1_small] — assertion 'CmdRun didn't backfill c_name for c_person_id=1', c_name is None after CmdRun completes.

#### Steps to reproduce

1. Open CBDB_BJ_User.mdb in Microsoft Access.
2. From the Navigation Pane, open the form **LookAtGroupData**.
3. In the import person list, enter a valid person ID (e.g. **1**).
4. Click **Run** (CmdRun button).
5. When CmdRun completes, inspect the result: the Name column is blank.
6. SQL verification: `SELECT c_person_id, c_name FROM ZZ_SCRATCH_IMPORT_PEOPLE` returns (1, NULL) — c_name was not backfilled from BIOG_MAIN.

#### Suggested fix

Locate the UPDATE statement in Form_LookAtGroupData.CmdRun_Click that joins ZZ_SCRATCH_IMPORT_PEOPLE to BIOG_MAIN and sets c_name. Check that the JOIN condition matches the correct key column and that the UPDATE target column name is spelled correctly.  After the fix, running CmdRun with any valid person ID should populate c_name in ZZ_SCRATCH_IMPORT_PEOPLE.

## Appendix A — c_index_year / c_index_addr_id drift vs the cbdb-online-main-server snapshot (differences need per-row classification before being filed as bugs)

> Appendix data not yet generated.  Run step 5c: `python reports/collect_index_year_diffs.py`

## Closing note

Thank you for taking the time to read this report. None of the items above is urgent; we hope having them all in one place makes it easy to address them at your own pace.

If any of the descriptions or suggested fixes are unclear, we would be glad to discuss further. The corresponding regression tests in this repository will automatically flip from PASS to FAIL the moment any regression marker stops reproducing in the source dump — that is a signal to investigate, not an automatic confirmation that the bug is fixed (the marker could fail because of an upstream fix, a fixture / driver change on our side, or a misclassification we made earlier).
