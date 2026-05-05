# Static audit catalogue

**Status:** authoritative catalogue of every static audit script
under `analysis/audit_*.py`.  Extracted from `AGENTS.md` 2026-05-05
to keep that file's default-context surface lean.

## When to read this

- You are about to add or extend a static audit script
- You are triaging a CBDB release and want to know which audits
  fire on which kinds of regression
- You are tracing a confirmed bug back to which audit caught it
- You are reviewing a PR that touches anything under
  `analysis/audit_*.py` or `analysis/run_all_audits.py`

If you are not in any of those situations, you do not need to
read this file — the AGENTS.md summary + bug-attribution
sentence is enough default context.

## Runner

`analysis/run_all_audits.py` runs every static audit and prints a
FLAGGED / CLEAN summary.  Run on every CBDB release; cheap (a few
seconds total).  Latest run on the shipped dump: **6 of 19 audits
flagged, 6.5 s total**.

```
python analysis/run_all_audits.py        # human-readable summary
python analysis/run_all_audits.py --ci   # exit non-zero if any
                                          # audit is above baseline
```

All audits share `analysis/audit_lib.read_vba_lines` for proper
`\r\r\n` handling, so reported line numbers match the line numbers
seen in `grep` and the VBE.

## Per-audit catalogue

Each entry: what the audit does + current state + which bugs (if
any) it has surfaced + what it guards against long-term.

### `analysis/audit_missing_controls.py`

Control-name typos in form VBA bodies.

- **Found:** Bug #5 (`LookAtStatus.CmdPajek_Click` references
  non-existent `ChkIDs` control)

### `analysis/audit_sql_columns.py`

SQL `<Table>.<Column>` typos in literal SQL strings inside VBA
bodies.

- **Found:** Bug #6 (`LookAtGroupData::queryEntry` references
  `ENTRY_DATA.c_parental_status` which doesn't exist).  Also
  expanded the scope of Bug #5.

### `analysis/audit_insert_select_columns.py`

INSERT/SELECT cardinality check companion to `audit_sql_columns
.py`.

- **Status:** currently clean (0 findings on 295 pure-literal
  INSERT statements)
- **Long-term guard:** future regressions where an INSERT's
  column list and the SELECT's projection drift apart silently

### `analysis/audit_sql_table_names.py`

Table-name typos in literal SQL strings.

- **Status:** currently surfaces only `frmIndexAddr` (orphan
  maintenance form, end users can't reach), filed as 🟢 LOW
- **Long-term guard:** SQL strings that reference renamed /
  removed tables

### `analysis/audit_saved_queries.py`

Same `<Table>.<Column>` check as `audit_sql_columns.py` but
applied to the 21 saved queries in `analysis/dump/queries.json`.

- **Status:** currently clean
- **Long-term guard:** view definitions drifting away from the
  underlying table schema

### `analysis/audit_recordset_fields.py`

Tracks `Set <var> = CurrentDb.OpenRecordset("<TABLE>", ...)` per
Sub and flags `<var>!<field>` references where `<field>` isn't
on `<TABLE>`.

- **Scope:** per-sub; invalidates on any `Set <var> = ...`
  reassignment (including reassignment to a SQL string the
  scanner can't statically evaluate)
- **Status:** currently clean

### `analysis/audit_recordset_sql_projection.py`

Sister scanner to `audit_recordset_fields.py` for the SQL-string
case.  Tracks `tQueryStr = "SELECT ..."` literal-only concats
and `Set <var> = CurrentDb.OpenRecordset(tQueryStr)`, parses the
SELECT projection, and flags `<var>!field` AND bare `!field`
(inside `With <var>`) reads where `field` isn't projected.

- **Found:** **Bugs #7 / #8 / #9** (CmdNeo4j family across
  LookAtPlace / LookAtNetworks / LookAtEntry — wrong recordset
  variable / SELECT missing required columns).
- **Run on every release.**

### `analysis/audit_subform_control_sources.py`

For every sub-form whose RecordSource is a saved query (`View_*`),
check each bound control's ControlSource exists in the saved
query's SELECT projection.

- **Found:** **Bugs #10 / #11 / #12** (silent display bugs in
  EVENT_ADDR_2 / EVENTS_DATA_2 / POSTED_TO_OFFICE_DATA_2 sub-forms
  — controls bound to columns the projection doesn't carry).

### `analysis/audit_error_label_targets.py`

Every `On Error GoTo <label>` / `Resume <label>` / `GoTo <label>`
must point at a label defined in the same Sub.

- **Status:** currently clean
- **Long-term guard:** typo'd error-handler renames

### `analysis/audit_event_handlers_exist.py`

Every form-control event handler named in `analysis/dump/control
_inventory.json` must have a matching `Sub <name>()` defined in
the form's VBA module.

- **Status:** currently clean
- **Long-term guard:** "renamed Sub but forgot to update OnClick
  property" silent-no-op bugs

### `analysis/audit_dcount_where_columns.py`

Every D-aggregate call (`DCount` / `DLookup` / `DSum` / etc.) with
a literal table+criteria must reference columns that exist on the
named table.

- **Status:** currently clean
- **Long-term guard:** stale-criteria silent-False bugs

### `analysis/audit_cross_form_references.py`

Every `Forms!<form>!<ctl>` reference must resolve to an existing
form AND existing control on that form (case-insensitive).  Skips
`Form__TMPCLP*.vb` auto-backup snapshots.

- **Found:** **Bugs #13 / #14** (BIOG_MAIN_2_Subform /
  KIN_DATA_Subform reference picker forms that don't exist in
  the .mdb).

### `analysis/audit_doc_md_open_form.py`

Every literal `DoCmd.OpenForm "<form>"` must resolve.

- **Status:** currently clean.  Bug #13's reference uses a string
  variable so it's caught by `audit_cross_form_references.py`
  instead.
- **Long-term guard:** direct-literal regressions

### `analysis/audit_dlookup_fields.py`

Every `DLookup("<field>", "<table>", ...)` literal call must
reference a valid field on the table.

- **Status:** currently clean
- **Long-term guard:** companion to `audit_dcount_where_columns
  .py`

### `analysis/audit_orphan_event_handlers.py`

Find Subs named like `<Control>_<Event>` where `<Control>` doesn't
exist on the form.  Code-smell signal (exit 0, informational).

- **Found:** **Bugs #15–#19** (LookAtPlace / LookAtStatus /
  LookAtOffice each have export-button event handlers with no
  matching button on the form design — silent missing UI).

### `analysis/audit_blocking_msgbox.py`

List every `If MsgBox(...) = vb<Yes|No|...>` confirmation prompt.
Not a bug check; informational guard so `tests/cbdb_driver/
vba_session._inject_autodetect` knows which prompts to pre-arrange
for tests.

### `analysis/audit_control_row_sources.py`

For every ListBox / ComboBox with a non-empty RowSource SQL,
verify each `<Table>.<Column>` reference is in the schema.

- **Status:** currently clean
- **Long-term guard:** third leg of the SQL-column-resolution
  stool (alongside `audit_sql_columns.py` and `audit_saved_queries
  .py`)

## Bug attribution rollup

For quick reverse-lookup ("which audit caught Bug #N?"):

| Bug | Caught by |
|---|---|
| #5  | `audit_missing_controls.py` (also expanded by `audit_sql_columns.py`) |
| #6  | `audit_sql_columns.py` |
| #7 / #8 / #9    | `audit_recordset_sql_projection.py` |
| #10 / #11 / #12 | `audit_subform_control_sources.py` |
| #13 / #14       | `audit_cross_form_references.py` |
| #15 / #16 / #17 / #18 / #19 | `audit_orphan_event_handlers.py` |

The remaining audits are currently clean; they're long-term
guards against future regressions in their respective scopes.

## Re-running

```
python analysis/run_all_audits.py
```

Add this to the per-release workflow:

```
1. analysis/dump_metadata.py + analysis/dump_vba.py (refresh dumps)
2. analysis/run_all_audits.py (this catalogue)
3. python analysis/discover_test_inputs.py
4. fast suite + slow VBA suite
```

See `AGENTS.md § Standard workflow after a .mdb update` for the
full release recipe.
