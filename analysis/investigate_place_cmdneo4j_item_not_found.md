# Investigation: which `Recordset!c_<col>` triggers JET 3265 in `Form_LookAtPlace.CmdNeo4j_Click`

**Date:** 2026-05-08  ·  **Branch:** `investigate/place-cmdneo4j-item-not-found` (off main `8f94276`)

Static-only follow-up to PR #120's verdict `probe_found_new_runtime_bug_candidate` on `LookAtPlace × CmdNeo4j`.  PR #120 left open the exact question of which `!c_<col>` reference inside CmdNeo4j_Click body actually fires the JET 3265 "Item not found in this collection.".  This investigation answers that statically — no Access COM, no probe rerun.

Source data:
- `analysis/dump/vba/Form_LookAtPlace.vb` (VBA source)
- `analysis/dump/tables.json` (canonical metadata dump)
- `tests/test_schema.py` REQUIRED_COLUMNS (cross-check)

## Methodology

1. Find every `Set tRst<X> = CurrentDb.OpenRecordset(...)` inside `CmdNeo4j_Click` body (lines 435-1778).
2. For each binding, extract the bound SQL — either the literal table name or the upstream `tQueryStr` build (multi-line `+ _` continuations supported).
3. Parse the SQL's SELECT projection (last identifier per comma-separated item; AS aliases respected).
4. Walk the body forward from each binding, attributing every `Recordset!c_<col>` field reference to the most recent rs binding (Set or With).
5. Cross-check used fields vs projected fields per binding.  Fields used but NOT projected are JET 3265 candidates.
6. For each candidate, check whether the underlying source table (named in the SQL's FROM/JOIN clause) actually has the column on the current dump.  This distinguishes `recordset_projection_mismatch` from `source_column_rename_or_removal`.

## Raw observed facts

- **OpenRecordset bindings inside CmdNeo4j_Click (lines 435-1778):** 7
- **Total `!c_<col>` field references inside the sub:** 54
- **Used-but-NOT-projected candidates (across all bindings):** 5
- **Candidates on the chain-order-first failing binding (line 651):** 5 (static pre-analysis listed 54 raw `!c_<col>` sites; this narrows to 5 chain-order-first candidates)

### Per-binding inventory

| line | rs_var | binding_kind | proj cols | from tables |
|---:|---|---|---|---|
| 531 | `tRstPeopleStatus` | `table_literal` | `(none parsed)` | `ZZ_SCRATCH_STATUS` |
| 651 | `tRstPeople` | `query_string` | `c_person_id, c_name, c_name_chn, c_index_year` | `ZZ_SCRATCH_P_TEXT, DYNASTIES, BIOG_MAIN` |
| 881 | `tRstPeoplePlace` | `query_string` | `c_person_id, c_index_addr_id, c_index_addr_type_code` | `ZZ_SCRATCH_P_TEXT, BIOG_MAIN` |
| 1073 | `tRstPlace` | `query_string` | `(none parsed)` | `?` |
| 1259 | `tRstPeoplePlace` | `query_string` | `(none parsed)` | `?` |
| 1479 | `tRstPeoplePlace` | `query_string` | `(none parsed)` | `?` |
| 1671 | `tRstPeoplePlace` | `query_string` | `c_index_addr_type_code, c_index_addr_type_desc, c_index_addr_type_chn` | `ZZ_SCRATCH_P_TEXT, BIOG_MAIN, BIOG_ADDR_CODES` |

### Candidates (chain-order-first binding only)

| use_line | rs_var | field | candidate_class | context |
|---:|---|---|---|---|
| 757 | `tRstPeople` | `c_dynasty` | `recordset_projection_mismatch` | `If IsNull(!c_dynasty) Then` |
| 765 | `tRstPeople` | `c_dynasty` | `recordset_projection_mismatch` | `tStr = tStr + !c_dynasty + tC` |
| 769 | `tRstPeople` | `c_dynasty_chn` | `recordset_projection_mismatch` | `tStr = tStr + !c_dynasty_chn + tC` |
| 777 | `tRstPeople` | `c_female` | `recordset_projection_mismatch` | `If IsNull(!c_female) Then` |
| 783 | `tRstPeople` | `c_female` | `recordset_projection_mismatch` | `tStr = tStr + IIf(!c_female, "F", "M")` |

### Source-side schema check (per candidate)

- `c_dynasty` (used at line 757):
  - sources_with_field (column DOES exist): `['ZZ_SCRATCH_P_TEXT', 'DYNASTIES']`
  - sources_lacking_field (column does NOT exist): `['BIOG_MAIN']`
- `c_dynasty` (used at line 765):
  - sources_with_field (column DOES exist): `['ZZ_SCRATCH_P_TEXT', 'DYNASTIES']`
  - sources_lacking_field (column does NOT exist): `['BIOG_MAIN']`
- `c_dynasty_chn` (used at line 769):
  - sources_with_field (column DOES exist): `['ZZ_SCRATCH_P_TEXT', 'DYNASTIES']`
  - sources_lacking_field (column does NOT exist): `['BIOG_MAIN']`
- `c_female` (used at line 777):
  - sources_with_field (column DOES exist): `['BIOG_MAIN']`
  - sources_lacking_field (column does NOT exist): `['ZZ_SCRATCH_P_TEXT', 'DYNASTIES']`
- `c_female` (used at line 783):
  - sources_with_field (column DOES exist): `['BIOG_MAIN']`
  - sources_lacking_field (column does NOT exist): `['ZZ_SCRATCH_P_TEXT', 'DYNASTIES']`

## Classification

Strict gate evaluation against the 4 brief-allowed buckets; first match wins.  See `_classify_outcome` docstring for the rules.

**Outcome bucket:** `recordset_projection_mismatch_candidate`

## Q1-Q5 answers

**Q1 — Most likely failing `!c_<col>` candidates (chain-order-first):**

- chain-order-first failure: **`tRstPeople!c_dynasty` at line 757**
- siblings on same binding (would fire next if head were fixed):
    - `tRstPeople!c_dynasty` at line 765
    - `tRstPeople!c_dynasty_chn` at line 769
    - `tRstPeople!c_female` at line 777
    - `tRstPeople!c_female` at line 783

- narrowed from static pre-analysis: `54 -> 5 (narrowed by per-binding cross-check; chain-order-first binding only)`

**Q2 — Which recordset binding each candidate is on:**

all chain-order-first candidates are on the same binding: line 651 (tRstPeople)

**Q3 — SQL projection check:**

Per-binding SELECT projections are listed in `bindings_full_inventory`.  The chain-order-first binding's SELECT does NOT project the candidate fields, BUT the source tables in its FROM clause DO have those columns on the current dump.

**Q4 — source-side rename vs recordset projection mismatch:**

recordset_projection_mismatch — the SELECT statement does not project the candidate fields, but the columns DO exist on the source tables in the FROM clause.  The VBA loop assumed the JOINed tables' columns would be automatically in the recordset, but DAO requires explicit SELECT projection.

**Q5 — Outcome bucket:** `recordset_projection_mismatch_candidate`

## Verdict: `recordset_projection_mismatch_candidate`

**Static evidence pinpoints the failing reference to a small candidate set inside one binding.**

Chain-order-first failure (the JET 3265 fires here, before any other binding's loop is reached): `tRstPeople!c_dynasty` at line 757.

Sibling unprojected uses on the same binding (would fire next if the head were fixed): [(765, 'c_dynasty'), (769, 'c_dynasty_chn'), (777, 'c_female'), (783, 'c_female')]

Binding: `Set tRstPeople = CurrentDb.OpenRecordset(tQueryStr, dbOpenDynaset)` at line 651; the `tQueryStr` build starts at line 643.

Failing SELECT projection (4 cols): `['c_person_id', 'c_name', 'c_name_chn', 'c_index_year']`.

Source-side schema check: every candidate field DOES exist on at least one of the binding's FROM tables — `['ZZ_SCRATCH_P_TEXT', 'DYNASTIES', 'BIOG_MAIN']`.  Specifically, `sources_with_field` for each candidate:
  - `c_dynasty`: present on ['ZZ_SCRATCH_P_TEXT', 'DYNASTIES']
  - `c_dynasty`: present on ['ZZ_SCRATCH_P_TEXT', 'DYNASTIES']
  - `c_dynasty_chn`: present on ['ZZ_SCRATCH_P_TEXT', 'DYNASTIES']
  - `c_female`: present on ['BIOG_MAIN']
  - `c_female`: present on ['BIOG_MAIN']

This rules out source-side rename / removal — the columns ARE on the source tables; the SELECT just doesn't project them.  The VBA loop assumes the JOINed tables' columns are automatically in the recordset's Fields collection, but DAO requires explicit SELECT projection.

**Sufficient for canonical issue filing:** YES.  Same shape as Issue #23 (LookAtAssociations × CmdNeo4j target-column mismatch, P1):
  - static evidence is binary (cols present on source, absent from SELECT projection)
  - failing line cited unambiguously
  - fix is a clear single-statement SELECT extension: add the missing cols to the SELECT projection (see fix_en sketch below)
  - per-form workaround would mirror PR #116's `.replace()` shape but rewrite the SELECT identifier list, not the INSERT target list

Recommended fix sketch (for a separate issue-filing PR, NOT this one):

```sql
-- BEFORE (Form_LookAtPlace.vb lines 643-647):
SELECT DISTINCT
    ZZ_SCRATCH_P_TEXT.c_person_id,
    ZZ_SCRATCH_P_TEXT.c_name,
    ZZ_SCRATCH_P_TEXT.c_name_chn,
    ZZ_SCRATCH_P_TEXT.c_index_year
FROM ZZ_SCRATCH_P_TEXT INNER JOIN
     ( DYNASTIES RIGHT JOIN BIOG_MAIN ON ... ) ON ...

-- AFTER (extend SELECT to project the cols the loop reads):
SELECT DISTINCT
    ZZ_SCRATCH_P_TEXT.c_person_id,
    ZZ_SCRATCH_P_TEXT.c_name,
    ZZ_SCRATCH_P_TEXT.c_name_chn,
    ZZ_SCRATCH_P_TEXT.c_index_year,
    DYNASTIES.c_dynasty,
    DYNASTIES.c_dynasty_chn,
    BIOG_MAIN.c_female
FROM ZZ_SCRATCH_P_TEXT INNER JOIN
     ( DYNASTIES RIGHT JOIN BIOG_MAIN ON ... ) ON ...
```

## Direct answers to the brief

**1. Most likely failing `!c_<col>` (narrowed from 54):**

- `tRstPeople!c_dynasty` at line 757 (recordset_projection_mismatch)
- `tRstPeople!c_dynasty` at line 765 (recordset_projection_mismatch)
- `tRstPeople!c_dynasty_chn` at line 769 (recordset_projection_mismatch)
- `tRstPeople!c_female` at line 777 (recordset_projection_mismatch)
- `tRstPeople!c_female` at line 783 (recordset_projection_mismatch)

In chain order, **the first one to fire is `tRstPeople!c_dynasty` at line 757** — the JET 3265 raises here and the error trap exits before any other candidate is reached.

**2. Which recordset they're on:** all chain-order-first candidates hang on the binding at line 757's upstream `Set tRstX = ...` (see `Q1_chain_order_first.binding_open_recordset_line` in JSON).

**3. Sufficient for canonical issue filing?**

**Yes.**  Same shape as Issue #23 (LookAtAssociations × CmdNeo4j target-column mismatch, P1):
- Static evidence is binary — the candidate fields are confirmed present on the binding's source tables (per `tables.json`), and confirmed absent from the SELECT projection (per the parsed VBA).  No runtime confirmation needed to file.
- Failing line cited unambiguously (chain-order-first).
- Fix is a clear single-statement SELECT extension (see verdict_note for the SQL sketch).
- Per-form workaround would mirror PR #116's `.replace()` shape (rewrite the SELECT identifier list, not the INSERT target list).

**4. Smallest next step:** open a canonical issue filing PR analogous to PR #115 (which filed Issue #23).  No smaller confirmation step needed.

## Constraints honoured per brief

- ✅ Investigation artifacts only — no `tests/` changes
- ✅ No driver / README / canonical reports / issue severity / triage docs touched
- ✅ No new issue filed (deferred to maintainer brief)
- ✅ No coverage PR
- ✅ No Access COM / no probe rerun — VBA dump + tables.json + parsed SELECT projections are sufficient
- ✅ Raw facts and inference separated (`## Raw observed facts` is dump/parse output only; `## Verdict` and `## Direct answers` are interpretation)
- ✅ Candidate set narrowed from 54 raw sites to a minimal chain-order-first set