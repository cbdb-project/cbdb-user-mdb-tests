# `CBDB_BJ_User.mdb` Logic Audit Report

Audit date: 2026-04-30
Scope: `CBDB_BJ_User.mdb` (41 MB) — 79 forms, 21 saved queries, 66 form
VBA modules, 2 class modules, ~50,000 lines of VBA total.

## Confirmed bugs

### Bug #1 — `View_StatusData` alias swap (display only, does not affect row selection)

`View_StatusData` is the `RecordSource` of `STATUS_DATA_2 Subform`, which
is shown every time the user opens a person's "Status" sub-datasheet.

**Symptom**: returned rows show `c_fy_range_desc` / `c_fy_range_chn`
(first-year range descriptions) populated with the *last-year* range
value rather than the first-year value.

**Root cause** (`analysis/dump/queries.json`, `View_StatusData`):

```sql
... LEFT JOIN YEAR_RANGE_CODES ON STATUS_DATA.c_fy_range = YEAR_RANGE_CODES.c_range_code
... LEFT JOIN YEAR_RANGE_CODES AS YEAR_RANGE_CODES_1 ON STATUS_DATA.c_ly_range = YEAR_RANGE_CODES_1.c_range_code
```

But all four range-related output aliases in the SELECT list pull from
`YEAR_RANGE_CODES_1`:

```
YEAR_RANGE_CODES_1.c_range     AS c_fy_range_desc   ← should use YEAR_RANGE_CODES
YEAR_RANGE_CODES_1.c_range_chn AS c_fy_range_chn    ← should use YEAR_RANGE_CODES
YEAR_RANGE_CODES_1.c_range     AS c_ly_range_desc   ← correct
YEAR_RANGE_CODES_1.c_range_chn AS c_ly_range_chn    ← correct
```

**Fix**: change to

```
YEAR_RANGE_CODES.c_range     AS c_fy_range_desc
YEAR_RANGE_CODES.c_range_chn AS c_fy_range_chn
```

**Regression tests**: `tests/test_known_bugs.py::test_bug_view_statusdata_fy_alias_swap`
and `test_bug_view_statusdata_fy_value_equals_ly_value`. Once the bug is
fixed, both tests will start failing — that is the signal that the fix
worked, and the assertions should be flipped to expect the corrected
behaviour as instructed in their docstrings.

---

### Bug #9 — `LookAtEntry.CmdNeo4j_Click` reads `!c_inst_*` from the wrong recordset (`tRstAssocCodes` instead of `tRstInstitutions`)

🔴 silent: identical "Item not found in this collection" symptom as
Bugs #3 / #4 — the InstitutionCodes file isn't written and the user
gets a vague popup.

`Form_LookAtEntry.vb:1415` opens the institutions recordset under
the right name:

```vba
Set tRstInstitutions = CurrentDb.OpenRecordset(tQueryStr)
```

But the very next loop at line 1425 reaches for the wrong handle:

```vba
With tRstAssocCodes        ← typo, should be tRstInstitutions
    .MoveFirst
    Do While Not .EOF
        If Not IsNull(!c_inst_code) Then
            tStr = ... + Trim(Str(!c_inst_name_code)) + ...
        ...
```

`tRstAssocCodes` was last bound much earlier (line 1335) to the
AssocCodes SELECT — it doesn't have `c_inst_*` columns.

**Fix**: change `With tRstAssocCodes` to `With tRstInstitutions`
on line 1425.

**Static detection**: `analysis/audit_recordset_sql_projection.py`
catches this — same scanner as Bug #8.

---

### Bug #8 — `LookAtNetworks.CmdNeo4j_Click` SQL projections are missing fields the loops read

🔴 silent: same family as Bug #3 — Networks's CmdNeo4j export dies
mid-stream with "Item not found in this collection."

Two distinct mismatches:

1. `Form_LookAtNetworks.vb:2458` — `tRstPlace` SQL projects three
   columns (`c_index_addr_id`, `c_index_addr_name`,
   `c_index_addr_chn`).  The loop at lines 2495 / 2498 / 2502 / 2505
   also reads `!x_coord` and `!y_coord`.

2. The companion `tRstPeoplePlace` SQL similarly omits
   `c_person_id` and `c_index_addr_id` that the loop at 2570 / 2572
   / 2574 reads.

**Fix**: extend each SELECT to project the columns the loop reads
(use the appropriate JOIN to ADDR_CODES / BIOG_MAIN to expose
`x_coord` / `y_coord` / `c_person_id` / `c_index_addr_id`).

**Static detection**: `analysis/audit_recordset_sql_projection.py`
flags both occurrences.

---

### Bug #7 — `LookAtPlace.CmdNeo4j_Click` SQL projection is missing fields the loop reads (silent runtime "Item not found")

🔴 silent: Neo4j export from LookAtPlace dies mid-stream and the
user gets a vague "Item not found in this collection." popup.

`Form_LookAtPlace.vb:322` builds a People-CSV for the Neo4j export:

```sql
SELECT DISTINCT
    ZZ_SCRATCH_P_TEXT.c_person_id,
    ZZ_SCRATCH_P_TEXT.c_name,
    ZZ_SCRATCH_P_TEXT.c_name_chn,
    ZZ_SCRATCH_P_TEXT.c_index_year
FROM ZZ_SCRATCH_P_TEXT
INNER JOIN (DYNASTIES RIGHT JOIN BIOG_MAIN ON DYNASTIES.c_dy = BIOG_MAIN.c_dy)
    ON ZZ_SCRATCH_P_TEXT.c_person_id = BIOG_MAIN.c_personid
```

But the per-row loop at lines 379-393 reads `!c_dynasty`,
`!c_dynasty_chn`, and `!c_female` from the recordset — none of
which are projected.  As soon as the loop hits a row, JET raises
"Item not found in this collection" → the `Err_CmdNeo4j_Click`
handler silences itself with `MsgBox Err.Description` → the
SaveToFile never runs and downstream files (PeopleIndexAddr,
Place, PeoplePlaceRelations, ...) never write either.

**Fix**: extend the SELECT projection to include the three fields:

```sql
SELECT DISTINCT
    ZZ_SCRATCH_P_TEXT.c_person_id,
    ZZ_SCRATCH_P_TEXT.c_name,
    ZZ_SCRATCH_P_TEXT.c_name_chn,
    ZZ_SCRATCH_P_TEXT.c_index_year,
    DYNASTIES.c_dynasty,
    DYNASTIES.c_dynasty_chn,
    BIOG_MAIN.c_female
FROM ...
```

(`c_dynasty` / `c_dynasty_chn` live on `DYNASTIES`, not `BIOG_MAIN`
— the existing join already exposes them; `c_female` is on
`BIOG_MAIN`.)

**Regression test**: `tests/test_vba_cmdneo4j_cross_form.py::
test_cmd_neo4j_produces_files[LookAtPlace]` — currently skipped
with this bug as the documented reason; once the fix lands, remove
the `LookAtPlace` arm of `_spec_skip_marks` and the test will pass.

**Why none of the static auditors caught this**: the SQL is built
at runtime via VBA string concatenation (`tQueryStr = "SELECT ..."
+ _ "FROM ..."`).  The per-Sub recordset-field auditor
(`analysis/audit_recordset_fields.py`) deliberately only tracks
recordsets opened on a *literal* table name; it can't reason about
runtime SQL projections.  Catching this needs a separate scanner
that pairs `OpenRecordset(<sql_string>)` with subsequent
`<rstvar>!<field>` reads against the projected column list.

---

### Bug #2 — VBA reference to `dao360.dll` is broken (blocks form automation and breaks new machines)

The VBA project in `CBDB_BJ_User.mdb` carries a reference to
`C:\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll`, the DAO
3.6 location used by Access 2003. Modern Office (2016+) ships
`ACEDAO.DLL` instead.

**Symptoms**:
- On any machine without legacy DAO installed, every form's `Form_Open`
  raises "Can't find project or library".
- Automated scripts (including the COM-driven UI tests this project
  originally intended to build) cannot open any form.

**Fix**:
- Open the .mdb in Access, press `Alt+F11` to enter the VBE.
- `Tools → References`, uncheck the entry shown as `MISSING: dao360.dll`.
- Check `Microsoft Office 16.0 Access Database Engine Object Library`
  (i.e. ACEDAO.DLL).
- Save.

**Automated**: `analysis/check_vba_refs.py` detects and patches the
broken reference (against a working copy).

---

## Heuristic scans (no further issues)

| Scanner | Result |
|---|---|
| `analysis/audit_view_aliases.py` — YEAR_RANGE_CODES alias mismatches | One false positive in `View_PeopleData` (the `c_da_*` prefix is shorthand for "death age"); the only real positive is Bug #1 |
| `analysis/audit_duplicate_aliases.py` — duplicate aliases in SELECT lists | None |

---

## Architectural observations (not bugs, but required context for the test suite)

1. **Linked tables**: `BIOG_MAIN`, `*_DATA`, `*_CODES`, `ADDR_CODES`, and
   ~120 other tables are linked from `CBDB_*_DATA.mdb` (DAO marks them
   as `record_count == -1` in `tables.json`). `CBDB_BJ_User.mdb` itself
   only owns ~63 local `ZZ_*` / `Z_*` working tables.

2. **Query entry point**: each `LookAtXxx` form has a single
   `Private Sub CmdQuery_Click` that:
   - Clears the `ZZ_SCRATCH_<XXX>` output table
   - Reads inputs from form controls + picker-populated `ZZ_SCRATCH_<CODE>`
     tables + form-module public globals like `gUseADDRID`
   - Builds and executes an
     `INSERT INTO ZZ_SCRATCH_<XXX> SELECT ... FROM <linked tables>`
   - Runs follow-up `UPDATE` statements to backfill descriptive columns
     from the `*_CODES` lookup tables
   - Re-binds the form's `RecordSource` to `ZZ_SCRATCH_<XXX>`

3. **Picker-form convention**: every `frmPickXxx_multi` writes the
   user's selection to its corresponding `ZZ_SCRATCH_<XXX>` table
   (e.g. `ZZ_SCRATCH_ENTRY_CODE`, `ZZ_SCRATCH_ADDR`). Tests can bypass
   picker UI by INSERT-ing directly into those tables.

4. **Private events cannot be reached via `Application.Run`**:
   `CmdQuery_Click` is `Private` by default, so calling
   `Application.Run "Form_X.CmdQuery_Click"` from outside fails. Options
   are (a) change them to `Public`, (b) use SendKeys, or (c) replay the
   SQL in Python — this test suite takes option (c).

5. **HelpFile reference numbers drift as CBDB data is updated**:
   `HelpFile_LookAtEntry.pdf` documents the example "Kaifeng yin general
   900-1100 = 104 people"; on the current data the replay returns 103
   (drift -1). The Use-Entry-Years variant is 12 today vs the
   HelpFile's 11 (drift +1). This confirms two things: (a) our SQL
   replay is faithful to the VBA, (b) HelpFile numbers should be soft
   references with 5-20% tolerance rather than hard assertions.

---

## Not yet covered / handed off to the user

- Full `CmdQuery_Click` translations for the other 9 LookAt forms
  (only LookAtEntry is complete). `tests/cbdb_replay/TEMPLATE_lookat.py`
  contains the recipe.
- Edge cases for complex pickers (multi-select, sub-unit expansion,
  XY-radius expansion). LookAtEntry implements the sub-unit and XY paths
  but no fixtures exercise them yet.
- Recursive kin networks (`LookAtKinship` / `LookAtNetworks` rely on the
  deep recursion in `clsTreeView` to derive kinship, which is the
  highest-risk audit surface).
- Bilingual switching (`changeDisplayLanguage` rewrites many label
  `Caption`s and is currently uncovered).
