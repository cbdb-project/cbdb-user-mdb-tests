# Investigation: `c_index_addr_type_code` source vs target side

**Date:** 2026-05-07  ·  **Branch:** `investigate/associations-cmdneo4j-c-index-addr-type-code` (off main `6f80d0f`)

Static follow-up to PR #112's verdict `probe_found_new_runtime_bug_candidate` on `LookAtAssociations × CmdNeo4j`.  PR #112 left open the exact question of whether the JET 3061 `unknown field name: 'c_index_addr_type_code'` is a source-side (`BIOG_MAIN`) drift or a target-side (`ZZ_SCRATCH_PEOPLE`) mismatch.  This investigation answers that statically — no Access COM, no probe rerun.

Source data: `analysis/dump/tables.json` (canonical metadata dump produced by `analysis/dump_metadata.py`) + `analysis/dump/vba/Form_LookAtAssociations.vb` (VBA dump).

## Raw observed facts

(Direct from the metadata dump; not interpreted.)

### Suspect column on each side

| Table | Column count | Has `c_index_addr_type_code`? |
|---|---:|:---:|
| `BIOG_MAIN` (source) | 55 | **YES** |
| `ZZ_SCRATCH_PEOPLE` (target) | 22 | **NO** |

Cross-check: `tests/test_schema.py::REQUIRED_COLUMNS` requires `BIOG_MAIN.c_index_addr_type_code` (line 47-48) and `pytest tests/test_schema.py` passes on current `main`, independently confirming `BIOG_MAIN` has the column.

### Target table full column list (22 cols)

  - `c_addr_chn`
  - `c_addr_desc`
  - `c_addr_desc_chn`
  - `c_addr_id`
  - `c_addr_name`
  - `c_addr_type`
  - `c_delete`
  - `c_dy`
  - `c_dynasty`
  - `c_dynasty_chn`
  - `c_female`
  - `c_index_year`
  - `c_index_year_type_code`
  - `c_index_year_type_desc`
  - `c_index_year_type_hz`
  - `c_name`
  - `c_name_chn`
  - `c_node_dist`
  - `c_person_id`
  - `x_coord`
  - `xy_count`
  - `y_coord`

### VBA INSERT under investigation

`Form_LookAtAssociations.vb`, `CmdNeo4j_Click` body (near lines 1287-1299; the INSERT that builds `ZZ_SCRATCH_PEOPLE` and references `BIOG_MAIN.c_index_addr_type_code`):

```vb
tQueryStr = "INSERT INTO ZZ_SCRATCH_PEOPLE ( c_person_id, c_name, c_name_chn, c_index_year, c_index_year_type_code, c_dy, c_addr_id, c_index_addr_type_code, c_female ) " + _
            "SELECT DISTINCT ZZ_SCRATCH_P_TEXT.c_person_id, BIOG_MAIN.c_name, BIOG_MAIN.c_name_chn, BIOG_MAIN.c_index_year, BIOG_MAIN.c_index_year_type_code, " + _
            "BIOG_MAIN.c_dy, BIOG_MAIN.c_index_addr_id, BIOG_MAIN.c_index_addr_type_code, BIOG_MAIN.c_female " + _
            "FROM ZZ_SCRATCH_P_TEXT INNER JOIN BIOG_MAIN ON ZZ_SCRATCH_P_TEXT.c_person_id = BIOG_MAIN.c_personid"
```

Per-INSERT-column existence on target `ZZ_SCRATCH_PEOPLE`:

  - `c_person_id`: OK
  - `c_name`: OK
  - `c_name_chn`: OK
  - `c_index_year`: OK
  - `c_index_year_type_code`: OK
  - `c_dy`: OK
  - `c_addr_id`: OK
  - `c_index_addr_type_code`: **MISSING**
  - `c_female`: OK

### Follow-up UPDATE (next stmt after the INSERT)

After the INSERT, `CmdNeo4j_Click` runs an UPDATE that LEFT JOINs four code tables and SETs descriptive columns:

```vb
tQueryStr = "UPDATE ( ( ( ZZ_SCRATCH_PEOPLE LEFT JOIN INDEXYEAR_TYPE_CODES ON ZZ_SCRATCH_PEOPLE.c_index_year_type_code = INDEXYEAR_TYPE_CODES.c_index_year_type_code )" + _
            " LEFT JOIN DYNASTIES ON ZZ_SCRATCH_PEOPLE.c_dy = DYNASTIES.c_dy ) " + _
            "LEFT JOIN ADDR_CODES ON ZZ_SCRATCH_PEOPLE.c_addr_id = ADDR_CODES.c_addr_id ) " + _
            "LEFT JOIN BIOG_ADDR_CODES ON ZZ_SCRATCH_PEOPLE.c_addr_type = BIOG_ADDR_CODES.c_addr_type " + _
            "SET ZZ_SCRATCH_PEOPLE.c_addr_desc = [BIOG_ADDR_CODES].[c_addr_desc], ..."
```

Per-UPDATE-target-column existence on `ZZ_SCRATCH_PEOPLE`:

  - `c_index_year_type_code`: OK
  - `c_dy`: OK
  - `c_addr_id`: OK
  - `c_addr_type`: OK
  - `c_index_year_type_desc`: OK
  - `c_index_year_type_hz`: OK
  - `c_dynasty`: OK
  - `c_dynasty_chn`: OK
  - `c_addr_name`: OK
  - `c_addr_chn`: OK
  - `c_addr_desc`: OK
  - `c_addr_desc_chn`: OK

INSERT-vs-UPDATE join-key population check (does the INSERT populate every column the UPDATE later joins on?):

  - `c_index_year_type_code`: POPULATED by INSERT
  - `c_dy`: POPULATED by INSERT
  - `c_addr_id`: POPULATED by INSERT
  - `c_addr_type`: **NEVER POPULATED**

## Classification

Strict gate evaluation (the four buckets are mutually exclusive; the first matching one wins):

| Bucket | Required | Match |
|---|---|:---:|
| `still_needs_deeper_investigation` | source or target missing from dump | — |
| `source_schema_drift_candidate` | suspect missing on source AND present on target | — |
| `target_table_schema_mismatch_candidate` | suspect present on source AND missing on target | ✅ |
| `sql_projection_mismatch_needs_runtime_confirmation` | suspect present on both | — |

**Outcome bucket:** `target_table_schema_mismatch_candidate`

## Brief Q1-Q4 answers

**Q1 — Does `BIOG_MAIN` have `c_index_addr_type_code`?**  **YES**

**Q2 — `ZZ_SCRATCH_PEOPLE` schema (key suspects):**

- `c_index_addr_type_code`: **NO** (missing)
- `c_addr_type`: **YES**
- `c_addr_id`: **YES**
- All UPDATE-used target columns present? `{'c_index_year_type_code': True, 'c_dy': True, 'c_addr_id': True, 'c_addr_type': True, 'c_index_year_type_desc': True, 'c_index_year_type_hz': True, 'c_dynasty': True, 'c_dynasty_chn': True, 'c_addr_name': True, 'c_addr_chn': True, 'c_addr_desc': True, 'c_addr_desc_chn': True}`

**Q3 — INSERT / SELECT / UPDATE self-consistent?**

- Self-consistent overall? **False**
- INSERT target columns vs target table: `{'c_person_id': True, 'c_name': True, 'c_name_chn': True, 'c_index_year': True, 'c_index_year_type_code': True, 'c_dy': True, 'c_addr_id': True, 'c_index_addr_type_code': False, 'c_female': True}`
- UPDATE join keys populated by INSERT: `{'c_index_year_type_code': True, 'c_dy': True, 'c_addr_id': True, 'c_addr_type': False}`

If insert_target_columns_present_on_target has any False, the INSERT references a target column that does not exist (root cause of the JET 3061 in this run).  If update_join_keys_populated_by_insert has any False (especially c_addr_type), even if the INSERT succeeded, the follow-up UPDATE would silently produce wrong/partial data.

**Q4 — Outcome bucket:** `target_table_schema_mismatch_candidate`

## Verdict: `target_table_schema_mismatch_candidate`

The unknown field is on the **target** table.  `ZZ_SCRATCH_PEOPLE` does not have a `c_index_addr_type_code` column on the current dump (verified against analysis/dump/tables.json — target has 22 columns and `c_index_addr_type_code` is not among them).  `BIOG_MAIN` does have it (per the same dump and per tests/test_schema.py REQUIRED_COLUMNS, which would fail loudly if BIOG_MAIN lost the column).

Strong static inference about *intent*: the follow-up UPDATE joins on `ZZ_SCRATCH_PEOPLE.c_addr_type = BIOG_ADDR_CODES.c_addr_type` (the BIOG_ADDR_CODES LEFT JOIN — see VBA fragment in MD).  Target table DOES have `c_addr_type`.  But the INSERT never populates `c_addr_type` — and `c_addr_type` is the natural rename target for the source column `BIOG_MAIN.c_index_addr_type_code`.  This pattern is consistent with a CBDB-side typo: the VBA author copied the source column name verbatim into the INSERT target list when they meant to rename it to `c_addr_type`.  Same shape as canonical Bugs #4 (`GISFrame` typo on LookAtPlace), #5 (`ChkIDs` typo on LookAtStatus), and #6 (queryEntry column typo on LookAtGroupData).

Sufficient evidence to support a canonical issue filing PR: yes.  Bucket: `target_table_schema_mismatch_candidate`.  Recommended fix path (for the maintainer brief, not this PR): either (a) upstream CBDB fix renames the INSERT target `c_index_addr_type_code` to `c_addr_type` so the INSERT populates the column the UPDATE later joins on, OR (b) driver-side per-form patch entry in `_PER_FORM_CMDGIS_PATCHES` for `Form_LookAtAssociations` mapping the literal `c_index_addr_type_code` -> `c_addr_type` inside CmdNeo4j_Click only.  The latter mirrors the existing GISFrame -> CodeFrame and ChkIDs -> False workaround patterns.

## Direct answers to the brief

**1. Source or target?**  **Target** (`ZZ_SCRATCH_PEOPLE`).  Static evidence is unambiguous: the metadata dump shows `ZZ_SCRATCH_PEOPLE` has 22 columns and `c_index_addr_type_code` is not among them; `BIOG_MAIN` has 55 columns including the suspect.  Cross-checked against `tests/test_schema.py` REQUIRED_COLUMNS which would fail loudly if the source side lost the column.

**2. Why static layer is enough to narrow this?**  Three independent static signals converge:

  - The metadata dump (`tables.json`) directly enumerates every column on every table; SUSPECT is missing on target, present on source.
  - The `tests/test_schema.py` REQUIRED_COLUMNS for `BIOG_MAIN` includes the suspect (line 47-48); schema test passes ⇒ source side cannot be the missing one.
  - The follow-up UPDATE in CmdNeo4j_Click LEFT JOINs `ZZ_SCRATCH_PEOPLE.c_addr_type = BIOG_ADDR_CODES.c_addr_type`, requiring `ZZ_SCRATCH_PEOPLE.c_addr_type` populated.  But the INSERT does not populate `c_addr_type` — and `c_addr_type` is the natural rename target for the source's `c_index_addr_type_code` (the existing INSERT for `c_addr_id` already does the analogous rename: BIOG_MAIN.c_index_addr_id ↦ target.c_addr_id).  This third signal goes beyond 'which side is missing the column' to suggest *what the VBA author meant*.

**3. Next step: issue filing or smaller confirmation?**  **Issue filing PR is sufficient as the next step.**  No smaller confirmation probe needed because:

  - The schema mismatch is a binary fact already pinned by `analysis/dump/tables.json` + a passing `test_schema.py`.
  - The VBA fragment is already fully transcribed in this artifact (and traceable via the existing PR #112 probe + `analysis/dump/vba/Form_LookAtAssociations.vb`).
  - The bug class (per-form column-name mismatch in a Cmd*_Click sub) is already canonical in the issue list as Bugs #4 / #5 / #6.  An issue-filing PR would mirror those entries' `fix_en` / `fix_zh` / static-marker test shape; a runtime pin in `tests/test_vba_bug_behaviors.py` is also feasible (the JET 3061 reproduces deterministically on the matrix Associations fixture per PR #112).
  - A driver-side `_PER_FORM_CMDGIS_PATCHES` workaround is feasible too (mirror the GISFrame -> CodeFrame and ChkIDs -> False patterns), but that is implementation work that belongs in a separate brief AFTER the issue is filed canonically.

## Constraints honoured per brief

- ✅ Investigation artifacts only — no `tests/` changes
- ✅ No driver / README / canonical reports / issue severity / triage docs touched
- ✅ No new issue filed (deferred to maintainer brief)
- ✅ No coverage PR
- ✅ No Access COM / no probe rerun — schema dump + VBA dump + test_schema.py REQUIRED_COLUMNS are sufficient
- ✅ Raw facts and conclusion separated: `## Raw observed facts` is dump-only, `## Verdict` and `## Direct answers to the brief` are interpretation