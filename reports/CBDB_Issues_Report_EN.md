# CBDB User MDB — Issues Report

_A respectful summary of issues uncovered during regression testing._

Dear maintainer,

Below is a summary of the issues we uncovered while building an automated regression-test suite for the CBDB User MDB. We hope this report is useful as you continue your wonderful stewardship of this dataset, and we sincerely thank you for the immense work that has gone into building it.

The issues are ordered by severity (P0 highest). Each entry includes a concise description, step-by-step user reproduction, screenshots where the issue is visible in the Access UI, and a suggested fix. None of these are urgent; they are documented so they can be addressed at the maintainer's convenience.

## Coverage Matrix — Form × Button Test Results

| Form | CmdQuery | CmdGIS | CmdNeo4j | CmdPajek | CmdGephi | CmdUCINet | CmdKML | CmdGUESS | CmdRun | CmdUTF8Pajek |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| LookAtEntry | ✗ FAIL | ✓ | ✓ | — | — | — | — | — | — | — |
| LookAtStatus | ✓ | ✓ | ✓ | ✗ FAIL | ~ SKIP | ✓ | — | — | — | — |
| LookAtTexts | ✓ | ✓ | ✓ | — | — | — | — | — | — | — |
| LookAtPlace | ✓ | ✗ FAIL | ✗ FAIL | ✓ | ✓ | — | — | — | — | — |
| LookAtAssociations | ✓ | ✓ | ✓ | ✗ FAIL | ✓ | ✗ FAIL | — | — | — | — |
| LookAtOffice | ✓ | ✗ FAIL | ✓ | — | — | — | — | ✓ | — | — |
| LookAtKinship | — | ✓ | ✓ | ✓ | — | ✗ FAIL | — | ✗ FAIL | ✓ | ✗ FAIL |
| LookAtNetworks | — | — | ~ SKIP | — | — | — | — | — | ~ SKIP | — |
| LookAtGroupData | — | ✓ | ✗ FAIL | — | — | — | — | — | ✗ FAIL | — |
| LookAtAssocPairs | ~ SKIP | ✓ | ✓ | ✓ | — | ✓ | — | — | — | — |

_PASS: 28 · FAIL: 12 · ERROR: 0 · SKIP: 4 · NOT RUN: 0 · N/A: 56_

## Table of Contents

- [P0 — Silent data corruption](#p0--silent-data-corruption)
  - [Issue #21 — LookAtOffice: CmdGIS output IndexYear column is nearly empty (0.2% fill rate) — likely silent column-bind regression](#issue-21--lookatoffice-cmdgis-output-indexyear-column-is-nearly-empty-02-fill-rate--likely-silent-column-bind-regression)
  - [Issue #26 — c_index_addr_id disagreement between User MDB and cbdb-online snapshot exceeds 0.5% threshold](#issue-26--c_index_addr_id-disagreement-between-user-mdb-and-cbdb-online-snapshot-exceeds-05-threshold)
  - [Issue #23 — LookAtAssociations: CmdPajek vertex section has off-by-N count — header declares 501 vertices but exports 8,093 rows](#issue-23--lookatassociations-cmdpajek-vertex-section-has-off-by-n-count--header-declares-501-vertices-but-exports-8093-rows)
  - [Issue #24 — LookAtKinship: CmdGUESS Gephi output has wrong field count per node row (nodedef declares 15 columns)](#issue-24--lookatkinship-cmdguess-gephi-output-has-wrong-field-count-per-node-row-nodedef-declares-15-columns)
- [P1 — Visible runtime crash](#p1--visible-runtime-crash)
  - [Issue #22 — LookAtAssociations / LookAtKinship: CmdUCINet crashes with 'Invalid procedure call or argument' when c_name contains CJK Han characters](#issue-22--lookatassociations--lookatkinship-cmducinet-crashes-with-invalid-procedure-call-or-argument-when-c_name-contains-cjk-han-characters)
  - [Issue #6 — LookAtGroupData.queryEntry crashes with 'No such field' — ENTRY_DATA.c_parental_status missing _code suffix](#issue-6--lookatgroupdataqueryentry-crashes-with-no-such-field--entry_datac_parental_status-missing-_code-suffix)
  - [Issue #7 — LookAtPlace.CmdNeo4j: People-CSV loop reads c_dynasty / c_dynasty_chn / c_female not in SELECT — crashes JET 3265 'Item not found'](#issue-7--lookatplacecmdneo4j-people-csv-loop-reads-c_dynasty--c_dynasty_chn--c_female-not-in-select--crashes-jet-3265-item-not-found)
  - [Issue #25 — LookAtKinship / LookAtGroupData / LookAtAssociationPairs: CmdImport round-trip fails — ZZ_SCRATCH_IMPORT_PEOPLE stays empty](#issue-25--lookatkinship--lookatgroupdata--lookatassociationpairs-cmdimport-round-trip-fails--zz_scratch_import_people-stays-empty)
- [P2 — Silent display](#p2--silent-display)
  - [Issue #2 — LookAtGroupData: CmdRun does not backfill c_name from BIOG_MAIN](#issue-2--lookatgroupdata-cmdrun-does-not-backfill-c_name-from-biog_main)
  - [Issue #3 — LookAtEntry: c_entry_desc backfill is NULL for all rows when entry_code = 36 (jinshi general)](#issue-3--lookatentry-c_entry_desc-backfill-is-null-for-all-rows-when-entry_code--36-jinshi-general)
  - [Issue #10 — EVENT_ADDR_2 Subform: TxtAddrCHN / TxtAddrPY bound to unaliased column names not in View_EventAddrData — render blank](#issue-10--event_addr_2-subform-txtaddrchn--txtaddrpy-bound-to-unaliased-column-names-not-in-view_eventaddrdata--render-blank)
- [P3 — Missing UI](#p3--missing-ui)
  - [Issue #15 — LookAtPlace is missing its CmdGIS button — handler exists but no UI control](#issue-15--lookatplace-is-missing-its-cmdgis-button--handler-exists-but-no-ui-control)
- [P5 — Dormant / latent / not currently reproducible](#p5--dormant--latent--not-currently-reproducible)
  - [Issue #20 — BOM-prefixed address names would produce embedded TAB delimiters in GIS exports — DORMANT (BOM data cleaned upstream, 0 affected rows in this dump)](#issue-20--bom-prefixed-address-names-would-produce-embedded-tab-delimiters-in-gis-exports--dormant-bom-data-cleaned-upstream-0-affected-rows-in-this-dump)
  - [Issue #1 — View_StatusData would display last-year range in the first-year column — DORMANT (no source rows trigger it on this dump)](#issue-1--view_statusdata-would-display-last-year-range-in-the-first-year-column--dormant-no-source-rows-trigger-it-on-this-dump)
  - [Issue #4 — LookAtPlace.CmdGIS_Click references non-existent control GISFrame — latent, masked by missing button (Issue #15)](#issue-4--lookatplacecmdgis_click-references-non-existent-control-gisframe--latent-masked-by-missing-button-issue-15)
  - [Issue #5 — LookAtStatus.CmdPajek_Click references missing control ChkIDs and three non-existent columns — latent, masked by missing button (Issue #16)](#issue-5--lookatstatuscmdpajek_click-references-missing-control-chkids-and-three-non-existent-columns--latent-masked-by-missing-button-issue-16)
  - [Issue #9 — LookAtEntry.CmdNeo4j Institutions block uses wrong recordset variable tRstAssocCodes — latent (no ENTRY_DATA row has c_inst_code > 0)](#issue-9--lookatentrycmdneo4j-institutions-block-uses-wrong-recordset-variable-trstassoccodes--latent-no-entry_data-row-has-c_inst_code--0)
  - [Issue #11 — EVENTS_DATA_2 Subform: c_event_record_id control bound to non-existent column — hidden, so latent](#issue-11--events_data_2-subform-c_event_record_id-control-bound-to-non-existent-column--hidden-so-latent)
  - [Issue #12 — POSTED_TO_OFFICE_DATA_2 Subform: c_appt_type_code control bound to non-projected column — hidden, so latent](#issue-12--posted_to_office_data_2-subform-c_appt_type_code-control-bound-to-non-projected-column--hidden-so-latent)
- [Severity legend](#severity-legend)
- [Appendix A — c_index_year / c_index_addr_id drift vs the cbdb-online-main-server snapshot (differences need per-row classification before being filed as bugs)](#appendix-a--c_index_year--c_index_addr_id-drift-vs-the-cbdb-online-main-server-snapshot-differences-need-per-row-classification-before-being-filed-as-bugs)
- [Appendix B — TablesFields: documentation vs. actual structure](#appendix-b--tablesfields-documentation-vs-actual-structure)
- [Appendix C — ForeignKeys: documentation vs. actual structure](#appendix-c--foreignkeys-documentation-vs-actual-structure)
- [Closing note](#closing-note)

## Severity legend

- P0 — Silent data corruption: data is wrong or missing without an error popup.
- P1 — Visible runtime crash: a popup appears, the operation aborts.
- P2 — Silent display: form fields render blank when they should show data.
- P3 — Missing UI: a feature exists in code but no button invokes it.
- P4 — Setup: one-time hurdle on each new install.
- P5 — Dormant / latent / not currently reproducible: kept as historical record; we re-checked on the current dump and could not trigger the symptom.

## P0 — Silent data corruption

### Issue #21 — LookAtOffice: CmdGIS output IndexYear column is nearly empty (0.2% fill rate) — likely silent column-bind regression

**Affected sub:** `Form_LookAtOffice.CmdGIS_Click`

**Severity:** P0 — Silent data corruption: the GIS export appears to succeed but IndexYear data is missing for 99.8% of rows.  Downstream GIS workflows that depend on year-based filtering will silently receive null years.

#### Description

When CmdGIS runs for LookAtOffice with person 80944 (unfiltered), the GIS output file is produced but the IndexYear column contains non-empty values in only 64 of 36,602 rows (0.2%), well below the 80% threshold expected for a correctly-populated GIS export.  This pattern is consistent with the silent column-bind regressions documented in Bugs #10, #11, and #12 — a column name in the CmdGIS SELECT is mismatched against the actual ZZ_SCRATCH table schema.

Detected by: test_cmd_gis_produces_file[office_80944_unfiltered] — assertion [LookAtOffice] CmdGIS column 'IndexYear' is non-empty in only 64/36602 rows (0.2%), below 80% threshold.

#### Steps to reproduce

1. Open CBDB_BJ_User.mdb in Microsoft Access.
2. Open the LookAtOffice form.
3. Set person ID to 80944 and leave all other filters blank.
4. Click CmdGIS.  The file is produced without an error popup.
5. Open the GIS output file and inspect the IndexYear column: the vast majority of rows will be empty.

#### Suggested fix

Inspect Form_LookAtOffice.CmdGIS_Click: locate the SELECT that populates the IndexYear column in the GIS output and verify the source column name matches the actual schema (check ZZ_SCRATCH_OFFICE or the equivalent table).

### Issue #26 — c_index_addr_id disagreement between User MDB and cbdb-online snapshot exceeds 0.5% threshold

**Affected sub:** `BIOG_MAIN (c_index_addr_id)`

**Severity:** P0 — Silent data drift: ~25 persons have a different primary address ID than the online system.  Geographic analyses and GIS exports that use c_index_addr_id will silently place these persons at the wrong location.

#### Description

A cross-check of BIOG_MAIN.c_index_addr_id between the User MDB and the cbdb-online-main-server SQLite snapshot found a disagreement rate of 0.500%, exactly at the maximum acceptable threshold.  At the default 5,000-row sample this means approximately 25 persons have a different c_index_addr_id in the two systems, indicating that either the User MDB has not fully applied recent upstream address assignments or the snapshot is ahead of the current data export.

Detected by: test_index_year_addr_xcheck_sample — assertion c_index_addr_id disagreement 0.500% exceeds 0.5% threshold.

#### Steps to reproduce

1. Run: python reports/collect_index_year_diffs.py
2. Inspect reports/index_drift_examples.json for rows where the bucket is 'addr_only' — these are persons where c_index_addr_id differs between the User MDB and the online snapshot.
3. For each differing person, query BIOG_MAIN.c_index_addr_id and compare against the online server to determine which value is authoritative.

#### Suggested fix

Apply the latest c_index_addr_id assignments from the cbdb-online server to BIOG_MAIN in the User MDB.  The differing rows are enumerated in reports/index_drift_examples.json (bucket: 'addr_only').

### Issue #23 — LookAtAssociations: CmdPajek vertex section has off-by-N count — header declares 501 vertices but exports 8,093 rows

**Affected sub:** `Form_LookAtAssociations.CmdPajek_Click`

**Severity:** P0 — Silent data corruption: the exported Pajek file is structurally invalid. Network analyses that ingest the file will operate on a truncated vertex set, producing incorrect centrality / community detection results without any warning.

#### Description

The Pajek .net file produced by Form_LookAtAssociations.CmdPajek_Click declares '*Vertices 501' in the header but the actual vertex section contains 8,093 rows before the next `*` section marker.  Pajek and other network analysis tools that rely on the vertex count header will either truncate the vertex list after 501 rows or raise a parse error, silently discarding the remaining ~7,592 vertices from the network.

Detected by: test_export_button_produces_file[LookAtAssociations_CmdPajek] — assertion header declared 501 vertices but found 8093 vertex rows before the next `*` section.

#### Steps to reproduce

1. Open CBDB_BJ_User.mdb in Microsoft Access.
2. Open LookAtAssociations, select a query that returns a large association network.
3. Click CmdPajek.  The .net file is written without an error popup.
4. Open the .net file and count the lines in the *Vertices section: the count exceeds the number declared in the '*Vertices N' header.

#### Suggested fix

In Form_LookAtAssociations.CmdPajek_Click, locate where the '*Vertices N' header is written and ensure N is derived from the actual count of vertex rows written, not a pre-computed estimate or a separate query result.

### Issue #24 — LookAtKinship: CmdGUESS Gephi output has wrong field count per node row (nodedef declares 15 columns)

**Affected sub:** `Form_LookAtKinship.CmdGUESS_Click`

**Severity:** P0 — Silent data corruption: the Gephi file is structurally invalid. Node attributes are silently misaligned, making all imported node metadata unreliable.

#### Description

The Gephi .gdf file produced by Form_LookAtKinship.CmdGUESS_Click declares 15 columns in the nodedef header but the actual node data rows contain a different number of fields (column/value misalignment).  Gephi and downstream tools will either fail to load the file or silently map node attributes to the wrong columns.

Detected by: test_cmd_guess_produces_file[kinship_person_3211] — assertion [LookAtKinship] Gephi: node rows with bad field count (nodedef has 15 cols).

#### Steps to reproduce

1. Open CBDB_BJ_User.mdb in Microsoft Access.
2. Open LookAtKinship.  Set a person ID that returns a kinship network.
3. Click CmdGUESS.  The .gdf file is written without error.
4. Open the .gdf file: count the columns declared in the 'nodedef>' header and compare with the number of comma-separated values in the first node data row.

#### Suggested fix

In Form_LookAtKinship.CmdGUESS_Click, ensure the nodedef header column list and the per-row value list are generated from the same ordered column definition. A mismatch typically occurs when a column is added to one list but not the other.

## P1 — Visible runtime crash

### Issue #22 — LookAtAssociations / LookAtKinship: CmdUCINet crashes with 'Invalid procedure call or argument' when c_name contains CJK Han characters

**Affected sub:** `Form_LookAtAssociations.CmdUCINet_Click / Form_LookAtKinship.CmdUCINet_Click`

**Severity:** P1 — Visible runtime crash: a popup aborts the export.  Any UCINet workflow on an association network that includes persons with Han-character names will fail.  Most CBDB persons have CJK names, making this effectively a blanket failure for real-world LookAtAssociations / LookAtKinship → UCINet usage.

#### Description

Form_LookAtAssociations.CmdUCINet_Click and Form_LookAtKinship.CmdUCINet_Click both call CreateTextFile with the 2-argument signature (filename, overwrite).  VBA's CreateTextFile raises 'Invalid procedure call or argument' (runtime error 5) when the output path contains a CJK Han character in a c_name value — the 2-arg form does not accept a Unicode flag, so Access silently uses the system ANSI code page, which cannot encode Han characters.  The error fires as a popup and aborts the export.  Fixtures using association code c_assoc_code = 437 ('Presented literary composition as gift to' / '贈詩、文') reliably trigger this because the associated persons include Han-character names.

Detected by: test_bug22_associations_cmducinet_fires_invalid_procedure_call and test_bug22_kinship_cmducinet_sibling_form_fires_invalid_procedure_call.

#### Steps to reproduce

1. Open CBDB_BJ_User.mdb in Microsoft Access.
2. Open the LookAtAssociations form.
3. Pick association code c_assoc_code = 437 ('Presented literary composition as gift to').
4. Click CmdUCINet.  A popup appears: 'Invalid procedure call or argument'.  The UCINet export file is not created.
5. The same error occurs in LookAtKinship.CmdUCINet when the kinship network contains a person whose c_name includes CJK Han characters (e.g. '取' / 贈詩、文).

#### Suggested fix

Change CreateTextFile calls in Form_LookAtAssociations.CmdUCINet_Click and Form_LookAtKinship.CmdUCINet_Click to the 3-argument form: CreateTextFile(filename, True, True) — the third argument enables Unicode output.  Test with a fixture that includes a person whose c_name contains CJK Han characters.

### Issue #6 — LookAtGroupData.queryEntry crashes with 'No such field' — ENTRY_DATA.c_parental_status missing _code suffix

**Affected sub:** `Form_LookAtGroupData.queryEntry`

**Severity:** P1 — Visible crash on a common path: any user who ticks the Entry checkbox in LookAtGroupData will hit this error.  ZZ_SCRATCH_ENTRY stays at 0 rows, so no Entry data is available for downstream export steps (GIS, Neo4j, etc.).

#### Description

Form_LookAtGroupData.vb line ~2621 has an INSERT INTO whose target column list ends with c_parental_status_code but whose SELECT projection ends with ENTRY_DATA.c_parental_status (no _code suffix).  The actual column on ENTRY_DATA is c_parental_status_code, so the SQL crashes with JET error 3061 'No value given for one or more required parameters' the moment the user ticks the Entry checkbox and clicks Run.  ZZ_SCRATCH_ENTRY remains at 0 rows.  The identical query in Form_LookAtEntry.vb uses the correct name; this is a single-character drift between the two forms.

Detected by: test_bug6_lookat_groupdata_query_entry_fires_no_such_field — asserts LookAtGroupData:ERR with JET 3061 signature and ZZ_SCRATCH_ENTRY stays at 0.  Also: test_bug6_groupdata_query_entry_wrong_field — static source-string assertion.

#### Steps to reproduce

1. Open CBDB_BJ_User.mdb in Microsoft Access.
2. Open the form **LookAtGroupData**.
3. In the import person list, enter person ID **1** (An Dun 安惇; has 2 ENTRY_DATA rows).
4. Tick only the **Entry** checkbox; leave Status / Office / Text / Addr unchecked.
5. Click **Run**.
6. A popup appears reporting JET error 3061 ('No value given for one or more required parameters' or 'No such field').  ZZ_SCRATCH_ENTRY stays empty.

#### Suggested fix

Change ENTRY_DATA.c_parental_status to ENTRY_DATA.c_parental_status_code on line ~2621 of Form_LookAtGroupData.vb.  One-character fix; the corrected form already appears in Form_LookAtEntry.vb.

### Issue #7 — LookAtPlace.CmdNeo4j: People-CSV loop reads c_dynasty / c_dynasty_chn / c_female not in SELECT — crashes JET 3265 'Item not found'

**Affected sub:** `Form_LookAtPlace.CmdNeo4j_Click`

**Severity:** P1 — Visible crash on a normal user click.  Any LookAtPlace → Neo4j export with a non-empty place-people result hits this deterministically.  0 CSV files are produced despite the SaveAs dialog having already fired.

#### Description

The People-CSV section of Form_LookAtPlace.CmdNeo4j_Click opens a recordset via a SELECT that projects only four ZZ_SCRATCH_P_TEXT columns, but the row-write loop reads !c_dynasty, !c_dynasty_chn, and !c_female from that recordset.  DAO's Recordset.Fields collection contains only SELECT-projected columns; the JOIN brings DYNASTIES and BIOG_MAIN into scope for filtering but does not expose their fields.  JET raises 3265 'Item not found in this collection' on the first !c_dynasty read; the error trap exits the sub before any file is written, so the user sees a popup AND the export produces 0 CSV files.

Detected by: test_bug7_lookat_place_cmdneo4j_fires_item_not_found — asserts 'Item not found' in ZZ_TEST_DEBUG :ERR markers.  Also: test_bug7_lookat_place_cmdneo4j_select_missing_dynasty_female — static SELECT projection assertion.

#### Steps to reproduce

1. Open CBDB_BJ_User.mdb in Microsoft Access.
2. Open the form **LookAtPlace**.
3. Use the address picker to select a substantive address (e.g. c_addr_id = 100658, Kaifeng 開封).  Click **Run Query**.
4. Click the **Neo4j** export button and choose a save location.
5. A Run-time error 3265 — 'Item not found in this collection' popup appears.  The chosen folder contains no Neo4j export files.

#### Suggested fix

Extend the SELECT in Form_LookAtPlace.CmdNeo4j_Click (lines ~643-647) to also project DYNASTIES.c_dynasty, DYNASTIES.c_dynasty_chn, and BIOG_MAIN.c_female.  The FROM / JOIN structure already brings those source tables into scope; three column references added to the SELECT projection is the complete fix.

### Issue #25 — LookAtKinship / LookAtGroupData / LookAtAssociationPairs: CmdImport round-trip fails — ZZ_SCRATCH_IMPORT_PEOPLE stays empty

**Affected sub:** `Form_LookAtKinship.CmdImport_Click / Form_LookAtGroupData.CmdImport_Click / Form_LookAtAssociationPairs.CmdImportList_Click`

**Severity:** P1 — Silent import failure: the import appears to complete successfully but the target table is empty.  Any subsequent query or export that depends on the imported person list will operate on an empty dataset without warning.

#### Description

After seeding person IDs [1, 2, 3] and clicking CmdImport (or CmdImportList for LookAtAssociationPairs), the handler is expected to populate ZZ_SCRATCH_IMPORT_PEOPLE with the seeded IDs.  In all three forms the target table remains empty (c_person_id = []) after the import completes.  No error popup is shown — the import appears to succeed silently but writes nothing.

Detected by: test_cmd_import_round_trip[LookAtKinship.CmdImport], test_cmd_import_round_trip[LookAtGroupData.CmdImport], and test_cmd_import_round_trip[LookAtAssociationPairs.CmdImportList] — all assert ZZ_SCRATCH_IMPORT_PEOPLE.c_person_id = []; expected [1, 2, 3].

#### Steps to reproduce

1. Open CBDB_BJ_User.mdb in Microsoft Access.
2. Open LookAtKinship (or LookAtGroupData / LookAtAssociationPairs).
3. Enter person IDs 1, 2, 3 in the import field.
4. Click CmdImport.  No error popup appears.
5. Query ZZ_SCRATCH_IMPORT_PEOPLE: SELECT c_person_id FROM ZZ_SCRATCH_IMPORT_PEOPLE — the table is empty.

#### Suggested fix

Inspect the CmdImport_Click / CmdImportList_Click handlers in each affected form.  Verify the INSERT INTO ZZ_SCRATCH_IMPORT_PEOPLE statement: check that the source control (text box or list box) is correctly read and that the INSERT executes within an active transaction that is committed.

## P2 — Silent display

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

### Issue #3 — LookAtEntry: c_entry_desc backfill is NULL for all rows when entry_code = 36 (jinshi general)

**Affected sub:** `Form_LookAtEntry.CmdQuery_Click`

**Severity:** P2 — Silent display issue: 92,545 rows affected.  The user can see the blank c_entry_desc column in the result grid, but Access shows no error — making it easy to overlook.  Exports (GIS, Neo4j, KML) that reference this column will also carry the blank.

#### Description

When the user runs a LookAtEntry query filtered to entry code 36 (examination: jinshi general), the result table ZZ_SCRATCH_ENTRY is populated with 92,545 rows but the c_entry_desc column is NULL for every row.  The expected value is 'examination: jinshi (general)'.

The CmdQuery_Click handler successfully inserts rows from ENTRY_DATA joined to ENTRY_CODES, but the c_entry_desc backfill step does not write the description for this specific entry code.  All other columns appear to be filled normally.  The missing description means the on-screen result grid shows a blank entry-type column for every record, which is misleading — the user sees results but cannot identify what type of examination each record represents.

Detected by: test_vba_full_matrix[top_entry_code_36_unfiltered] — assertion 'c_entry_desc backfill wrong' with 92,545 affected rows.  Also: test_vba_full_matrix[entry_39_dy_20], [entry_36_dy_20], [entry_36_dy_15].

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

### Issue #10 — EVENT_ADDR_2 Subform: TxtAddrCHN / TxtAddrPY bound to unaliased column names not in View_EventAddrData — render blank

**Affected sub:** `EVENT_ADDR_2 Subform`

**Severity:** P2 — Silent display: both address controls on the EVENT_ADDR_2 sub-form render blank for every row, with no error popup.  Users see no indication that the address name is available; the parent row's address display is unaffected.

#### Description

The EVENT_ADDR_2 sub-form's TxtAddrCHN control has ControlSource c_name_chn and TxtAddrPY has ControlSource c_name, but the form's RecordSource is View_EventAddrData, which aliases ADDR_CODES.c_name_chn as c_event_addr_chn and ADDR_CODES.c_name as c_event_addr_name.  Neither c_name nor c_name_chn is in the projection, so both controls silently render blank for every row on the Events-with-Addresses sub-datasheet.  A SQL probe confirms: SELECT c_name_chn FROM View_EventAddrData raises 'Too few parameters. Expected 2.' — JET treats the unknown identifier as a parameter.

Detected by: test_subform_control_source_unresolved[bug10_TxtAddrCHN] — opens EVENT_ADDR_2 Subform via COM and asserts TxtAddrCHN.ControlSource is still 'c_name_chn'.

#### Steps to reproduce

1. Open CBDB_BJ_User.mdb in Microsoft Access.
2. Open CBDB_Browser_2 and navigate to person **c_personid = 44872** (Sun Cai 孫才) — this person has an EVENT_ADDR row pointing to c_addr_id = 12603 (Anfeng 安豐).
3. Switch to the **Events** sub-tab.
4. Observe the EVENT_ADDR_2 sub-form nested inside the event row: TxtAddrCHN and TxtAddrPY render blank even though the parent row's address controls (bound to c_addr_chn / c_addr_name in View_EventsData) show '安豐'.

#### Suggested fix

In the form designer for EVENT_ADDR_2 Subform, change TxtAddrCHN.ControlSource from c_name_chn to c_event_addr_chn, and TxtAddrPY.ControlSource from c_name to c_event_addr_name — the actual alias names in View_EventAddrData.

## P3 — Missing UI

### Issue #15 — LookAtPlace is missing its CmdGIS button — handler exists but no UI control

**Affected sub:** `LookAtPlace`

**Severity:** P3 — Missing UI: the GIS export feature is completely unavailable to LookAtPlace users even though the underlying handler is functional (with the Issue #4 fix applied).

#### Description

Form_LookAtPlace.vb defines a fully functional CmdGIS_Click handler — it builds and writes a GIS .tab export identical in shape to the GIS button on Status / Texts / Associations / Office / Kinship.  But LookAtPlace's form design has no CmdGIS button.  Users on Place can use Pajek / Gephi / Neo4j export but cannot use GIS export; the handler is there, just unreachable from the UI.  Note: if the button is added, Issue #4 (GISFrame vs CodeFrame typo in the same handler) must be fixed at the same time.

Detected by: test_orphan_export_button_truly_missing[bug15_LookAtPlace_CmdGIS] — opens LookAtPlace via COM, calls Controls('CmdGIS'), and asserts the lookup raises 'Item not found'.

#### Steps to reproduce

1. Open CBDB_BJ_User.mdb in Microsoft Access.
2. Open the form **LookAtPlace**.
3. Look at the export-buttons row at the bottom right.  There is no GIS button.
4. Compare with LookAtStatus / LookAtAssociations / LookAtOffice etc., all of which have a GIS button.

#### Suggested fix

In LookAtPlace's form design, add a CmdGIS button next to the existing CmdPajek / CmdGephi buttons with OnClick = [Event Procedure].  Also fix Issue #4 (GISFrame → CodeFrame typo) in the same patch.

## P5 — Dormant / latent / not currently reproducible

_Items in this tier are kept as historical / latent record.  They fall into three categories: (a) DORMANT — verified that current source data doesn't trigger the symptom; (b) NOT CURRENTLY REPRODUCIBLE — the symptom no longer surfaces even though the suspect code is still present (we have NOT confirmed an upstream source-level fix; could be a JET / Office behaviour change, a fixture / driver change on our side, or the original diagnosis was a false positive); (c) LATENT — the source-code defect is real, but the user can't reach it because another issue (e.g. a missing UI button) blocks the path.  None of these are user-facing today; **none have been verified as fixed upstream** — please consult before treating any of them as either urgent or closed._

### Issue #20 — BOM-prefixed address names would produce embedded TAB delimiters in GIS exports — DORMANT (BOM data cleaned upstream, 0 affected rows in this dump)

**Affected sub:** `ADDR_CODES + Form_LookAt*.CmdGIS_Click`

**Severity:** P5 — Dormant on this dump (BOM data cleaned upstream, 0 affected rows in data-20260602).  Would re-activate as P0 silent data corruption if any future ADDR_CODES row re-introduces a U+FEFF prefix in c_name or c_name_chn.  The unescaped CmdGIS write code is the second half of the fix still outstanding.

#### Description

The ADDR_CODES table previously contained rows where c_name and c_name_chn carried a leading U+FEFF (BOM) prefix (almost certainly from a UTF-8-with-BOM paste at data-import time), which when passed through JET produced embedded TAB characters in the CmdGIS output.  The key reachable example was c_addr_id = 702559 (Wei Shi 尉氏), reachable from persons with c_status_code = 40 (civil office / [為官者：文]).

In this build (data-20260602) the BOM data has been cleaned upstream: test_addr_codes_has_known_bom_dirty_rows now finds 0 BOM-prefixed rows, and test_known_reachable_dirty_addr_present also returns 0 rows.  The GIS unescaped-write pattern remains in the code (CmdGIS of LookAtTexts / LookAtPlace / LookAtAssociations / LookAtOffice / LookAtKinship still does no TAB escaping), so the structural risk re-activates if future data introduces another BOM row.

GOLDEN_STALE: BOM golden tests now expect 0 rows; update goldens.  This issue is retained as P5_dormant_or_latent because the unescaped-write code defect is still present.

#### Steps to reproduce

1. In this build the bug cannot be triggered — SELECT COUNT(*) FROM ADDR_CODES WHERE Left(c_name, 1) = ChrW(65279) returns 0.
2. The structural risk: open Form_LookAtOffice (or any LookAt form), run CmdGIS with status code **40** (civil office / [為官者：文]).  If a BOM row were present for c_addr_id = 702559 (Wei Shi 尉氏) the output file would have an extra TAB column around row 11476.
3. (Dormant verification) Confirm: SELECT COUNT(*) FROM ADDR_CODES WHERE Left(c_name, 1) = ChrW(65279) returns 0 on this dump.

#### Suggested fix

The data-side fix has been applied upstream (0 BOM rows remain).  The code-side fix is still needed: before each tStr = tStr + value + tC append in the CmdGIS bodies of all LookAt forms, replace any embedded Chr(9), Chr(10), Chr(13), or U+FEFF in value with a space.  This prevents re-occurrence if future imports bring BOM-prefixed rows.

### Issue #1 — View_StatusData would display last-year range in the first-year column — DORMANT (no source rows trigger it on this dump)

**Affected sub:** `View_StatusData`

**Severity:** P5 — Dormant on this dump (would be P2 silent display if any STATUS_DATA row has both c_fy_range and c_ly_range set differently).  The SQL alias swap is a confirmed source-level defect; the symptom simply has no trigger row today.

#### Description

The saved query View_StatusData joins YEAR_RANGE_CODES twice, aliasing the second copy as YEAR_RANGE_CODES_1 and joining it on STATUS_DATA.c_ly_range (the last-year range).  However the SELECT clause pulls c_fy_range_desc and c_fy_range_chn from YEAR_RANGE_CODES_1 — the wrong alias — so every status row would display the last-year range text in the first-year range column.  On the current dump no STATUS_DATA row has both c_fy_range and c_ly_range populated with different values, so the symptom is invisible in the UI today but will surface the moment future data introduces such a row.

Detected by: test_bug_view_statusdata_fy_alias_swap — assertion 'YEAR_RANGE_CODES_1.c_range AS c_fy_range_desc' still present in View_StatusData SQL.  Also: test_bug_view_statusdata_fy_value_equals_ly_value — asserts c_fy_range_desc == c_ly_range_desc for all rows with non-NULL range descriptions.

#### Steps to reproduce

1. Open CBDB_BJ_User.mdb in Microsoft Access.
2. Press F11 to show the Navigation Pane, then double-click query **View_StatusData**.
3. Inspect the SELECT clause: both c_fy_range_desc and c_fy_range_chn reference YEAR_RANGE_CODES_1, but the FROM clause joins YEAR_RANGE_CODES_1 on STATUS_DATA.c_ly_range — not c_fy_range.
4. (Dormant verification) Run: SELECT c_personid, c_fy_range_desc, c_ly_range_desc FROM View_StatusData WHERE c_fy_range > 0 OR c_ly_range > 0.  On the current dump the result is empty, confirming the bug is latent.

#### Suggested fix

In View_StatusData, change YEAR_RANGE_CODES_1.c_range AS c_fy_range_desc and YEAR_RANGE_CODES_1.c_range_chn AS c_fy_range_chn to reference the un-aliased YEAR_RANGE_CODES copy (which the FROM clause already joins on STATUS_DATA.c_fy_range).  One-line fix per column.

### Issue #4 — LookAtPlace.CmdGIS_Click references non-existent control GISFrame — latent, masked by missing button (Issue #15)

**Affected sub:** `Form_LookAtPlace.CmdGIS_Click`

**Severity:** P5 — Latent (would be P1 visible crash if Issue #15 were fixed without also fixing this line).  The test driver works around this via a per-form GISFrame→CodeFrame substitution patch so integration tests pass; the underlying CBDB bug remains.

#### Description

Form_LookAtPlace.CmdGIS_Click reads GISFrame.Value on line ~1539, but LookAtPlace has no control named GISFrame — the actual encoding selector is named CodeFrame.  If the button were ever added (fixing Issue #15) without first correcting this line, every click would raise Run-time error 424 'Object required' and the GIS export would never run.  Today the bug is masked because no CmdGIS button exists on the form (Issue #15), so users cannot click it at all.

Detected by: test_bug4_lookat_place_cmdgis_fires_object_required — disables the driver's GISFrame→CodeFrame patch and confirms the un-patched code raises 'Object required'.  Also: test_bug4_lookat_place_cmdgis_references_nonexistent_gisframe — asserts GISFrame.Value is still present in Form_LookAtPlace.vb.

#### Steps to reproduce

1. (Hypothetical, requires Issue #15 fixed first.) Open LookAtPlace.
2. Run any query so the scratch table has data.
3. Click the GIS button.
4. A Run-time error 424 — Object required popup appears; the export produces no file.

#### Suggested fix

Change GISFrame.Value to CodeFrame.Value on line ~1539 of Form_LookAtPlace.vb.  The same form's CmdNeo4j_Click, CmdGephi_Click, and CmdPajek_Click already use CodeFrame correctly — this is a single-identifier drift.  Fix in the same patch as Issue #15 (adding the CmdGIS button).

### Issue #5 — LookAtStatus.CmdPajek_Click references missing control ChkIDs and three non-existent columns — latent, masked by missing button (Issue #16)

**Affected sub:** `Form_LookAtStatus.CmdPajek_Click`

**Severity:** P5 — Latent (would be P1 visible crash if Issue #16 were fixed without also fixing these two defects).  The whole sub looks copy-pasted from LookAtAssociations.CmdPajek_Click without adapting column names to Status's schema.

#### Description

Form_LookAtStatus.CmdPajek_Click contains two related defects copied from LookAtAssociations without adapting names: (a) line ~2308 reads ChkIDs.Value, but LookAtStatus has no ChkIDs control; (b) the SELECT inside CmdPajek_Click references ZZ_SCRATCH_STATUS.c_person_id, c_status_id, and c_status_count — none of which exist on ZZ_SCRATCH_STATUS (the real columns are c_personid, c_status_code; there is no count column).  Both defects are moot today because LookAtStatus has no CmdPajek button (Issue #16), but adding the button without fixing these would expose both failures to users.

Detected by: test_bug5_lookat_status_cmdpajek_sql_fires_field_error — pre-seeds ZZ_SCRATCH_STATUS and fires CmdPajek directly, asserting an ERR marker with an object-required or missing-field signature.  Also: test_bug5_lookat_status_cmdpajek_references_nonexistent_chkids — static source-string assertion.

#### Steps to reproduce

1. (Hypothetical, requires Issue #16 fixed first.) Open LookAtStatus.
2. Run any query so ZZ_SCRATCH_STATUS has data.
3. Click the Pajek button.
4. First: a Run-time error 424 'Object required' popup appears (ChkIDs.Value).
5. If worked around: a 'No such field' error fires from the SELECT referencing c_person_id / c_status_id / c_status_count.

#### Suggested fix

Two fixes required: (a) replace ChkIDs.Value with False (or add a real ChkIDs control to LookAtStatus) and (b) rewrite the SELECT to use ZZ_SCRATCH_STATUS.c_personid and ZZ_SCRATCH_STATUS.c_status_code, dropping or computing c_status_count differently.  In practice the entire sub likely needs a thoughtful rewrite rather than spot fixes.

### Issue #9 — LookAtEntry.CmdNeo4j Institutions block uses wrong recordset variable tRstAssocCodes — latent (no ENTRY_DATA row has c_inst_code > 0)

**Affected sub:** `Form_LookAtEntry.CmdNeo4j_Click`

**Severity:** P5 — Latent source-level typo (would re-promote to P1 if any future ENTRY_DATA row has c_inst_code > 0).  Test test_bug9_lookat_entry_cmdneo4j PASSED this build (gate shut).  The missing InstitutionCodes CSV is not a user-visible bug today.

#### Description

Form_LookAtEntry.vb line ~1415 opens an institutions recordset as Set tRstInstitutions = CurrentDb.OpenRecordset(tQueryStr), but line ~1425 says With tRstAssocCodes — referencing a recordset that was already Close'd in the AssocCodes block upstream.  If executed, .MoveFirst would raise DAO 3021 'No current record'.  The entire block sits inside If tRecDeleted > 0 Then at line ~1389, where tRecDeleted is the count of ENTRY_DATA rows with c_inst_code > 0.  On the current dump 0 of 263,454 ENTRY_DATA rows have c_inst_code > 0, so the gate evaluates false and the buggy block is unreachable from any LookAtEntry fixture today.  Test test_bug9_lookat_entry_cmdneo4j PASSED because the gate was shut, confirming the latent state.

Representative fixtures: c_entry_code = 36 (examination: jinshi general / 進士) and c_entry_code = 101 (recommendation / 薦舉).  Both yield ENTRY_DATA rows with c_inst_code = 0, confirming the gate remains shut on this dump.

Included per MANIFEST requirement (test_report_code_labels_audit_clean).

#### Steps to reproduce

1. On the current dump this bug cannot be triggered through the UI — the If tRecDeleted > 0 Then gate at Form_LookAtEntry.vb:~1389 is false for every possible LookAtEntry fixture (0 of 263,454 ENTRY_DATA rows have c_inst_code > 0).
2. Verify the source-level typo statically: open analysis/dump/vba/Form_LookAtEntry.vb and inspect lines ~1415-1425.  Line ~1415: Set tRstInstitutions = CurrentDb.OpenRecordset(tQueryStr).  Line ~1425: With tRstAssocCodes (intended: With tRstInstitutions).
3. (Optional) Confirm the gate condition: SELECT COUNT(*) FROM ENTRY_DATA WHERE c_inst_code > 0 returns 0.
4. Representative fixtures: c_entry_code = 36 (jinshi / 進士) and c_entry_code = 101 (recommendation / 薦舉).  Both confirm c_inst_code = 0.

#### Suggested fix

Change With tRstAssocCodes on line ~1425 to With tRstInstitutions.  Single-identifier fix; the correct variable was opened just a few lines above.

### Issue #11 — EVENTS_DATA_2 Subform: c_event_record_id control bound to non-existent column — hidden, so latent

**Affected sub:** `EVENTS_DATA_2 Subform`

**Severity:** P5 — Latent (would be P2 silent display if the control were ever made Visible=True or widened).  Code-hygiene only; no user-visible impact today.

#### Description

The EVENTS_DATA_2 sub-form has a control named c_event_record_id whose ControlSource is also c_event_record_id.  Neither EVENTS_DATA nor View_EventsData projects a column of that name.  If the control were visible, it would render blank.  A live COM probe confirms the control is Visible=False with width 240 twips (~4 mm) — a hidden internal control, almost certainly a leftover join-key field never meant to be shown.  Real users see no blank column because the control is not displayed.

Detected by: test_subform_control_source_unresolved[bug11_c_event_record_id] — opens EVENTS_DATA_2 Subform via COM and asserts the ControlSource is still c_event_record_id.

#### Steps to reproduce

1. Verification is static + COM probe only — there is no UI symptom.
2. Static evidence: SELECT c_event_record_id FROM View_EventsData raises 'Too few parameters. Expected 1.' — confirming the column is absent from the projection.
3. Visibility evidence: the COM probe confirms Visible=False and width=240 twips for the c_event_record_id control on EVENTS_DATA_2 Subform.

#### Suggested fix

Either delete the hidden c_event_record_id control, or change its ControlSource to a real column (e.g. c_event_code) so it does not carry a stale binding.  Either change is invisible to users; this is code-hygiene only.

### Issue #12 — POSTED_TO_OFFICE_DATA_2 Subform: c_appt_type_code control bound to non-projected column — hidden, so latent

**Affected sub:** `POSTED_TO_OFFICE_DATA_2 Subform`

**Severity:** P5 — Latent (would be P2 silent display if the control were made visible).  The user-facing appointment-type controls on the form work correctly.  Code-hygiene only.

#### Description

The POSTED_TO_OFFICE_DATA_2 sub-form has a control c_appt_type_code with ControlSource c_appt_type_code, but View_PostingOfficeData projects c_appt_code (no _type infix) — not c_appt_type_code.  A live COM probe confirms the control is Visible=False, so the blank rendering is not user-visible today.  The user-facing appointment-type controls on the same form work correctly.  This is a code-hygiene issue only.

Detected by: test_subform_control_source_unresolved[bug12_c_appt_type_code] — opens POSTED_TO_OFFICE_DATA_2 Subform via COM and asserts the ControlSource is still c_appt_type_code.

#### Steps to reproduce

1. Verification is static + COM probe only — there is no UI symptom.
2. Static evidence: in control_inventory.json, POSTED_TO_OFFICE_DATA_2 Subform has a control with control_source = 'c_appt_type_code', but View_PostingOfficeData projects c_appt_code.
3. Visibility evidence: the COM probe confirms Visible=False for c_appt_type_code.

#### Suggested fix

Either delete the hidden c_appt_type_code control, or change its ControlSource to c_appt_code (the actual column projected by View_PostingOfficeData).

## Appendix A — c_index_year / c_index_addr_id drift vs the cbdb-online-main-server snapshot (differences need per-row classification before being filed as bugs)

When we compare BIOG_MAIN's `c_index_year` and `c_index_addr_id` between this User MDB and the weekly cbdb-online-main-server SQLite snapshot, a small fraction of persons disagree.

**The two sides are independent implementations.**  The SQLite snapshot's `c_index_year` is produced by cbdb-online-main-server's PHP `IndexYearRebuildService.php` and its `c_index_addr_id` by `IndexAddressRebuildService.php` (both at <https://github.com/cbdb-project/cbdb-online-main-server>); the User MDB-side: `c_index_addr_id` rebuilt by VBA in `Form_frmIndexAddr` (front-end mdb); `c_index_year` rebuilt by **37 saved QueryDefs named `BM IY Rule …`** in the linked-tables backend `data/CBDB_<YYYYMMDD>_DATA.mdb`, driven by `frmBaseMaintenance`.  Both algorithms now extracted to `analysis/dump_data/querydefs_index/*.sql`; form / module driver VBA still needs an interactive Access SaveAsText pass.  PHP is intended to mirror the VBA but they are separate code paths.  Per-row differences can come from at least four sources, and a diff alone doesn't tell us which: (1) source-data snapshot drift; (2) algorithm / porting divergence between PHP and VBA; (3) priority / tie-break differences; (4) null / default handling differences.

**We have not classified the steady ~575 / 657 246 diffs we currently observe.**  The examples below are a small sample (currently 13 rows across 3 buckets, from `reports/index_drift_examples.json`) — illustrative of the shapes of disagreement, not statistically representative.  They are a starting point for per-row triage, not a verdict.

### Examples where only c_index_year disagrees

**`c_personid = 3501` — 李孝稱 (Li Xiaocheng)**

| Field | User MDB | cbdb-online-main-server snapshot |
|---|---|---|
| `c_index_year` | 1018 | 1028 |
| `c_index_addr_id` | 100658 | 100658 |
| `c_birthyear` | 0 | 0 |
| `c_deathyear` | 0 | 0 |
| `c_index_year_type_code` | 19 | 1912 |
| `c_index_year_source_id` | 19149 | 3479 |

**`c_personid = 15971` — 郭世隆 (Guo Shilong)**

| Field | User MDB | cbdb-online-main-server snapshot |
|---|---|---|
| `c_index_year` | 960 |  |
| `c_index_addr_id` | 100658 | 100658 |
| `c_birthyear` | 0 | 0 |
| `c_deathyear` | 0 | 0 |
| `c_index_year_type_code` | 14 |  |
| `c_index_year_source_id` | 24426 |  |

**`c_personid = 16266` — 錢孟回 (Qian Menghui)**

| Field | User MDB | cbdb-online-main-server snapshot |
|---|---|---|
| `c_index_year` | 1004 | 992 |
| `c_index_addr_id` | 12723 | 12723 |
| `c_birthyear` | 0 | 0 |
| `c_deathyear` | 0 | 0 |
| `c_index_year_type_code` | 13 | 0512 |
| `c_index_year_source_id` | 700103 | 3035 |

**`c_personid = 16267` — 錢知雄 (Qian Zhixiong)**

| Field | User MDB | cbdb-online-main-server snapshot |
|---|---|---|
| `c_index_year` | 1034 | 1071 |
| `c_index_addr_id` | 12723 | 12723 |
| `c_birthyear` | 0 | 0 |
| `c_deathyear` | 0 | 0 |
| `c_index_year_type_code` | 1312 | 14 |
| `c_index_year_source_id` | 16266 | 16269 |

**`c_personid = 19771` — 李彭 (Li Peng)**

| Field | User MDB | cbdb-online-main-server snapshot |
|---|---|---|
| `c_index_year` | 1253 |  |
| `c_index_addr_id` | 100185 | 100185 |
| `c_birthyear` | 0 | 0 |
| `c_deathyear` | 0 | 0 |
| `c_index_year_type_code` | 11 |  |
| `c_index_year_source_id` | 40822 |  |

### Examples where only c_index_addr_id disagrees

**`c_personid = 1` — 安惇 (An Dun)**

| Field | User MDB | cbdb-online-main-server snapshot |
|---|---|---|
| `c_index_year` | 1042 | 1042 |
| `c_index_addr_id` | 101117 |  |
| `c_birthyear` | 1042 | 1042 |
| `c_deathyear` | 1104 | 1104 |
| `c_index_year_type_code` | 01 | 01 |
| `c_index_year_source_id` |  |  |

**`c_personid = 470` — 金君卿 (Jin Junqing)**

| Field | User MDB | cbdb-online-main-server snapshot |
|---|---|---|
| `c_index_year` | 1012 | 1012 |
| `c_index_addr_id` | 12879 | 12785 |
| `c_birthyear` | 0 | 0 |
| `c_deathyear` | 0 | 0 |
| `c_index_year_type_code` | 05 | 05 |
| `c_index_year_source_id` |  |  |

**`c_personid = 481` — 周秩 (Zhou Zhi)**

| Field | User MDB | cbdb-online-main-server snapshot |
|---|---|---|
| `c_index_year` | 1043 | 1043 |
| `c_index_addr_id` | 100416 | 12785 |
| `c_birthyear` | 0 | 0 |
| `c_deathyear` | 0 | 0 |
| `c_index_year_type_code` | 05 | 05 |
| `c_index_year_source_id` |  |  |

**`c_personid = 485` — 周穜 (Zhou Tong)**

| Field | User MDB | cbdb-online-main-server snapshot |
|---|---|---|
| `c_index_year` | 1046 | 1046 |
| `c_index_addr_id` | 100416 | 12785 |
| `c_birthyear` | 0 | 0 |
| `c_deathyear` | 0 | 0 |
| `c_index_year_type_code` | 05 | 05 |
| `c_index_year_source_id` |  |  |

**`c_personid = 562` — 范沖 (Fan Chong)**

| Field | User MDB | cbdb-online-main-server snapshot |
|---|---|---|
| `c_index_year` | 1067 | 1067 |
| `c_index_addr_id` | 100658 | 13292 |
| `c_birthyear` | 1067 | 1067 |
| `c_deathyear` | 1141 | 1141 |
| `c_index_year_type_code` | 01 | 01 |
| `c_index_year_source_id` |  |  |

### Examples where the SOURCE data itself differs (birthyear / deathyear)

**`c_personid = 263` — 張穆之 (Zhang Muzhi)**

| Field | User MDB | cbdb-online-main-server snapshot |
|---|---|---|
| `c_index_year` | 1016 |  |
| `c_index_addr_id` |  |  |
| `c_birthyear` | 1016 | 0 |
| `c_deathyear` | 1079 | 0 |
| `c_index_year_type_code` | 01 |  |
| `c_index_year_source_id` |  |  |

**`c_personid = 1455` — 沈邈 (Shen Miao)**

| Field | User MDB | cbdb-online-main-server snapshot |
|---|---|---|
| `c_index_year` | 1001 | 1008 |
| `c_index_addr_id` | 12887 | 12887 |
| `c_birthyear` | 1001 | 0 |
| `c_deathyear` | 1047 | 0 |
| `c_index_year_type_code` | 01 | 05 |
| `c_index_year_source_id` |  |  |

**`c_personid = 19149` — 李孝基 (Li Xiaoji)**

| Field | User MDB | cbdb-online-main-server snapshot |
|---|---|---|
| `c_index_year` | 1016 | 1026 |
| `c_index_addr_id` | 100658 | 100658 |
| `c_birthyear` | 1016 | 0 |
| `c_deathyear` | 1076 | 0 |
| `c_index_year_type_code` | 01 | 11 |
| `c_index_year_source_id` |  | 41030 |

## Appendix B — TablesFields: documentation vs. actual structure

This section compares the contents of the `TablesFields` table in `CBDB_20260602_DATA.mdb` against the database schema reconstructed from Access DAO (TableDefs) by `reports/collect_schema_diffs.py`. Discrepancies indicate the documentation table may be out of date.

Total rows in TablesFields: 875. Reconstructed from DB: 997.

Reconstructed schema: [tables_fields_regen.csv](tables_fields_regen.csv)

### Rows in TablesFields not found in actual DB (stale)

| AccessTblNm | AccessFldNm |
|---|---|
| ADMIN_CAT_CODE_TYPE_REL | c_admin_type_code |
| ADMIN_CAT_TYPES | c_admin_type_code |
| ADMIN_CAT_TYPES | c_admin_type_hz |
| ADMIN_CAT_TYPES | c_admin_type_trans |
| ENTRY_DATA | c_addr_id |
| ENTRY_DATA | c_posting_id |
| MERGED_PERSON_DATA | c_merged_to_personid |
| PersonIDSource | LineNum |
| PersonIDSource | SourceTable |
| TMP_ADDR_C | Max_c_belongs_first_year |

### Columns in actual DB not documented in TablesFields

| AccessTblNm | AccessFldNm | DataFormat | NULL_allowed |
|---|---|---|---|
| ADDRESSES | belongs1_ID | Long | True |
| ADDRESSES | belongs1_Name | Text | True |
| ADDRESSES | belongs2_ID | Long | True |
| ADDRESSES | belongs2_Name | Text | True |
| ADDRESSES | belongs3_ID | Long | True |
| ADDRESSES | belongs3_Name | Text | True |
| ADDRESSES | belongs4_ID | Long | True |
| ADDRESSES | belongs4_Name | Text | True |
| ADDRESSES | belongs5_ID | Long | True |
| ADDRESSES | belongs5_Name | Text | True |
| ADDRESSES | c_addr_cbd | Text | True |
| ADDRESSES | c_addr_id | Long | True |
| ADDRESSES | c_admin_type | Text | True |
| ADDRESSES | c_firstyear | Integer | True |
| ADDRESSES | c_lastyear | Integer | True |
| ADDRESSES | c_name | Text | True |
| ADDRESSES | c_name_chn | Text | True |
| ADDRESSES | x_coord | Double | True |
| ADDRESSES | y_coord | Double | True |
| ADMIN_CAT_CODE_TYPE_REL | c_admin_cat_type_code | Text | False |
| ADMIN_CAT_TYPES | c_admin_cat_type_code | Text | False |
| ADMIN_CAT_TYPES | c_admin_cat_type_hz | Text | False |
| ADMIN_CAT_TYPES | c_admin_cat_type_trans | Text | False |
| ASSOC_DATA | c_tertiary_type_notes | Text | True |
| BIOG_ADDR_DATA | c_delete | Integer | True |
| BIOG_MAIN | c_name_fixed | Text | True |
| CopyTables | NotProcessed | Yes/No | True |
| CopyTables | TableName | Text | False |
| CopyTablesDefault | ID | Long | True |
| CopyTablesDefault | TableName | Text | True |
| ENTRY_DATA | c_entry_addr_id | Long | True |
| ETHNICITY_TRIBE_CODES | c_sortorder | Integer | True |
| ForeignKeys | AccessFldNm | Text | True |
| ForeignKeys | AccessTblNm | Text | True |
| ForeignKeys | DataFormat | Text | True |
| ForeignKeys | FKName | Text | True |
| ForeignKeys | FKString | Text | True |
| ForeignKeys | ForeignKey | Text | True |
| ForeignKeys | ForeignKeyBaseField | Text | True |
| ForeignKeys | IndexOnField | Text | True |
| ForeignKeys | NULL_allowed | Yes/No | True |
| ForeignKeys | skip | Integer | True |
| FormLabels | c_english | Text | True |
| FormLabels | c_fanti | Text | True |
| FormLabels | c_form | Text | True |
| FormLabels | c_jianti | Text | True |
| FormLabels | c_label_id | Integer | True |
| MERGED_PERSON_DATA | c_merged_from_personid | Long | False |
| OFFICE_CODES_CONVERSION | c_office_chn | Text | True |
| OFFICE_CODES_CONVERSION | c_office_chn_backup | Text | True |
| OFFICE_CODES_CONVERSION | c_office_id | Long | True |
| OFFICE_CODES_CONVERSION | c_office_id_backup | Long | True |
| OFFICE_TYPE_TREE_backup | c_office_type_desc | Text | True |
| OFFICE_TYPE_TREE_backup | c_office_type_desc_chn | Text | True |
| OFFICE_TYPE_TREE_backup | c_office_type_node_id | Text | True |
| OFFICE_TYPE_TREE_backup | c_parent_id | Text | True |
| OFFICE_TYPE_TREE_backup | c_tts_node_id | Text | True |
| Paste Errors | c_bibl_cat_code | Long | True |
| Paste Errors | c_created_by | Text | True |
| Paste Errors | c_created_date | Date/Time | True |
| Paste Errors | c_extant | Long | True |
| Paste Errors | c_modified_by | Text | True |
| Paste Errors | c_modified_date | Date/Time | True |
| Paste Errors | c_notes | Memo | True |
| Paste Errors | c_pages | Text | True |
| Paste Errors | c_source | Long | True |
| Paste Errors | c_textid | Long | True |
| Paste Errors | c_text_country | Long | True |
| Paste Errors | c_text_dy | Long | True |
| Paste Errors | c_text_nh_code | Long | True |
| Paste Errors | c_text_nh_year | Long | True |
| Paste Errors | c_text_range_code | Long | True |
| Paste Errors | c_text_type_id | Text | True |
| Paste Errors | c_text_year | Long | True |
| Paste Errors | c_title | Text | True |
| Paste Errors | c_title_alt_chn | Text | True |
| Paste Errors | c_title_chn | Text | True |
| Paste Errors | c_title_trans | Text | True |
| Paste Errors | c_url_api | Text | True |
| Paste Errors | c_url_api_coda | Text | True |
| Paste Errors | c_url_homepage | Text | True |
| POSTED_TO_OFFICE_DATA | c_posting_id_old | Long | True |
| SOCIAL_INSTITUTION_ALTNAME_CODES | c_inst_altname_chn | Text | True |
| SOCIAL_INSTITUTION_ALTNAME_CODES | c_inst_altname_desc | Text | True |
| SOCIAL_INSTITUTION_ALTNAME_CODES | c_inst_altname_type | Integer | True |
| SOCIAL_INSTITUTION_ALTNAME_CODES | c_notes | Text | True |
| SOCIAL_INSTITUTION_ALTNAME_DATA | c_inst_altname_hz | Text | False |
| SOCIAL_INSTITUTION_ALTNAME_DATA | c_inst_altname_py | Text | True |
| SOCIAL_INSTITUTION_ALTNAME_DATA | c_inst_altname_type | Integer | False |
| SOCIAL_INSTITUTION_ALTNAME_DATA | c_inst_code | Integer | False |
| SOCIAL_INSTITUTION_ALTNAME_DATA | c_inst_name_code | Integer | False |
| SOCIAL_INSTITUTION_ALTNAME_DATA | c_notes | Memo | True |
| SOCIAL_INSTITUTION_ALTNAME_DATA | c_pages | Text | True |
| SOCIAL_INSTITUTION_ALTNAME_DATA | c_source | Long | True |
| SOCIAL_INSTITUTION_CODES | c_inst_end_dy | Integer | True |
| SOCIAL_INSTITUTION_CODES | c_inst_end_year | Integer | True |
| SOCIAL_INSTITUTION_CODES_CONVERSION | c_inst_code | Long | True |
| SOCIAL_INSTITUTION_CODES_CONVERSION | c_inst_code_new | Integer | True |
| SOCIAL_INSTITUTION_CODES_CONVERSION | c_inst_name_code | Integer | True |
| SOCIAL_INSTITUTION_CODES_CONVERSION | c_new_new_code | Long | True |
| STATUS_TYPES | c_status_type_parent_code | Text | True |
| TablesFields | AccessFldNm | Text | False |
| TablesFields | AccessTblNm | Text | False |
| TablesFields | DataFormat | Text | True |
| TablesFields | DumpFldNm | Text | True |
| TablesFields | DumpTblNm | Text | True |
| TablesFields | ForeignKey | Text | True |
| TablesFields | ForeignKeyBaseField | Text | True |
| TablesFields | IndexOnField | Text | True |
| TablesFields | NULL_allowed | Yes/No | True |
| TablesFields | RowNum | Long | True |
| TablesFieldsChanges | Change | Text | True |
| TablesFieldsChanges | ChangeDate | Text | True |
| TablesFieldsChanges | ChangeNotes | Text | True |
| TablesFieldsChanges | FieldName | Text | True |
| TablesFieldsChanges | TableName | Text | True |
| TEXT_BIBLCAT_CODES | c_text_cat_level | Text | True |
| TEXT_BIBLCAT_CODES | c_text_cat_parent_id | Text | True |
| TEXT_CODES | c_text_type_id | Text | True |
| TMP_ADDR_C | Min_c_belongs_first_year | Integer | True |
| TMP_ADDR_D | c_addr_cbd | Text | True |
| TMP_ADDR_E | c_addr_cbd | Text | True |
| TMP_DISTANCE_DATA | assoc_xcoord | Double | True |
| TMP_DISTANCE_DATA | assoc_ycoord | Double | True |
| TMP_DISTANCE_DATA | c_assoc_id | Long | False |
| TMP_DISTANCE_DATA | c_distance | Double | True |
| TMP_DISTANCE_DATA | c_personid | Long | False |
| TMP_DISTANCE_DATA | c_t_dist | Double | True |
| TMP_DISTANCE_DATA | x_coord | Double | True |
| TMP_DISTANCE_DATA | y_coord | Double | True |
| ZZZ_DY_DATA | c_dy | Integer | False |
| ZZZ_DY_DATA | c_personid | Long | False |

### Attribute mismatches

Full list: `reports/schema_diff_tables_fields_mismatches.csv` (143 rows)

## Appendix C — ForeignKeys: documentation vs. actual structure

This section covers the `ForeignKeys` table and the FK relationships it documents.

Total rows in ForeignKeys: 188. Reconstructed from DB (via Access.Application DAO): 223.

Reconstructed FK list: [foreign_keys_regen.csv](foreign_keys_regen.csv)

### Rows in ForeignKeys not found in actual DB (stale)

| AccessTblNm | AccessFldNm | ForeignKey | ForeignKeyBaseField |
|---|---|---|---|
| ADDR_BELONGS_DATA | c_source | TEXT_CODES | c_textid |
| assoc_data | c_assoc_day_gz | GANZHI_CODES | c_ganzhi_code |
| assoc_data | c_assoc_nh_code | nian_hao | c_nianhao_id |
| assoc_data | c_assoc_range | year_range_codes | c_range_code |
| assoc_data | c_inst_name_code | SOCIAL_INSTITUTION_NAME_CODES | c_inst_name_code |
| assoc_data | c_inst_name_code,c_inst_code | SOCIAL_INSTITUTION_CODES | c_inst_code |
| biog_addr_data | c_addr_id | ADDR_CODES | c_addr_id |
| biog_addr_data | c_fy_day_gz | GANZHI_CODES | c_ganzhi_code |
| biog_addr_data | c_fy_nh_code | nian_hao | c_nianhao_id |
| biog_addr_data | c_fy_range | year_range_codes | c_range_code |
| biog_addr_data | c_ly_day_gz | GANZHI_CODES | c_ganzhi_code |
| biog_addr_data | c_ly_nh_code | nian_hao | c_nianhao_id |
| biog_addr_data | c_ly_range | year_range_codes | c_range_code |
| biog_addr_data | c_personid | BIOG_MAIN | c_personid |
| biog_addr_data | c_source | TEXT_CODES | c_textid |
| BIOG_INST_DATA | c_inst_name_code | SOCIAL_INSTITUTION_NAME_CODES | c_inst_name_code |
| BIOG_INST_DATA | c_inst_name_code,c_inst_code | SOCIAL_INSTITUTION_CODES | c_inst_code |
| biog_main | c_death_age_range | year_range_codes | c_range_code |
| biog_main | c_index_year_source_id | BIOG_MAIN | c_personid |
| biog_main | c_index_year_type_code | INDEXYEAR_TYPE_CODES | c_index_year_type_code |
| ENTRY_DATA | c_entry_dy | DYNASTIES | c_dy |
| ENTRY_DATA | c_inst_name_code | SOCIAL_INSTITUTION_NAME_CODES | c_inst_name_code |
| ENTRY_DATA | c_inst_name_code,c_inst_code | SOCIAL_INSTITUTION_CODES | c_inst_code |
| EVENTS_ADDR | c_event_code | EVENT_CODES | c_event_code |
| EVENTS_ADDR | c_personid | BIOG_MAIN | c_personid |
| EVENTS_ADDR | c_personid,c_sequence,c_event_code | EVENTS_DATA | c_event_code |
| POSTED_TO_OFFICE_DATA | c_inst_name_code,c_inst_code | SOCIAL_INSTITUTION_CODES | c_inst_code |
| SOCIAL_INSTITUTION_ADDR | c_inst_name_code,c_inst_code | SOCIAL_INSTITUTION_CODES | c_inst_code |

### Columns in actual DB not documented in ForeignKeys

| AccessTblNm | AccessFldNm | ForeignKey | ForeignKeyBaseField |
|---|---|---|---|
| ADDRESSES | c_addr_id | ADDR_CODES | c_addr_id |
| ASSOC_CODES | c_assoc_pair | ASSOC_CODES | c_assoc_code |
| Assoc_data | c_assoc_fy_day_gz | GANZHI_CODES | c_ganzhi_code |
| Assoc_data | c_assoc_fy_nh_code | NIAN_HAO | c_nianhao_id |
| Assoc_data | c_assoc_fy_range | YEAR_RANGE_CODES | c_range_code |
| ASSOC_TYPES | c_assoc_type_parent_id | ASSOC_TYPES | c_assoc_type_code |
| ENTRY_DATA | c_entry_addr_id | ADDR_CODES | c_addr_id |
| EVENTS_DATA | c_event_code | EVENTS_ADDR | c_event_code |
| EVENTS_DATA | c_personid | EVENTS_ADDR | c_personid |
| EVENTS_DATA | c_sequence | EVENTS_ADDR | c_sequence |
| POSTED_TO_OFFICE_DATA | c_inst_code | SOCIAL_INSTITUTION_CODES | c_inst_code |
| POSTED_TO_OFFICE_DATA | c_inst_name_code | SOCIAL_INSTITUTION_CODES | c_inst_name_code |
| SOCIAL_INSTITUTION_ADDR | c_inst_code | SOCIAL_INSTITUTION_CODES | c_inst_code |
| SOCIAL_INSTITUTION_ADDR | c_inst_name_code | SOCIAL_INSTITUTION_CODES | c_inst_name_code |
| SOCIAL_INSTITUTION_CODES | c_inst_end_dy | DYNASTIES | c_dy |

## Closing note

Thank you for taking the time to read this report. None of the items above is urgent; we hope having them all in one place makes it easy to address them at your own pace.

If any of the descriptions or suggested fixes are unclear, we would be glad to discuss further. The corresponding regression tests in this repository will automatically flip from PASS to FAIL the moment any regression marker stops reproducing in the source dump — that is a signal to investigate, not an automatic confirmation that the bug is fixed (the marker could fail because of an upstream fix, a fixture / driver change on our side, or a misclassification we made earlier).
