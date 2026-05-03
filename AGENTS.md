# AGENTS.md — context for future agent sessions on this project

This file is the **operational handoff** for any agent (or human) coming
into this repo cold.  Read it BEFORE making changes.

## What this project is

Automated regression testing for **CBDB_BJ_User.mdb** — Harvard's CBDB
(China Biographical Database) user-facing Access database.  The .mdb
overlays 10 query forms (LookAtEntry / LookAtKinship / ...) over a
separate `CBDB_*_DATA.mdb` (~1 GB linked tables).

User goal: after every CBDB data update or VBA change, automatically
verify that all query/export logic still produces correct results.

User's stated pain points (in priority order):
1. Export functions silently break — wrong columns, missing data
2. Export columns get misaligned (off-by-one)
3. Some queries trigger error message-boxes (which kill any unattended run)
4. Some buttons require preconditions before working (button enable-state)
5. Specific parameter combinations (entry method × dynasty × address)
   are the real test surface — random fixtures often hit sparse data

## Single source of truth (two of them, scoped)

This project keeps **two** authoritative documents and nothing
else.  Do not duplicate their content elsewhere — copies drift.

1. **Roadmap, coverage table, fixture status** →
   **`README.md` § Plan & status**.
2. **Issues / bugs (all 19 of them, content + tier + severity +
   reproduction steps)** → the `ISSUES` dict in
   **`reports/generate_report.py`**, which auto-generates the four
   `reports/CBDB_Issues_Report_*.md` outputs (en + zh; the same
   script also emits `.docx` siblings on demand, but those are
   gitignored — see PR O policy).

**Hard rules for any agent (or human contributor):**

- If your change moves a roadmap item from open → in-progress →
  done, or completes / discovers / skips a fixture, **update the
  corresponding row in `README.md`'s Plan & status table in the
  same PR**.  CI/reviewers should fail PRs that change `tests/`
  substantively without touching `README.md`.
- If your change discovers, fixes, reclassifies, or invalidates
  an issue, **edit the entry in `reports/generate_report.py`'s
  `ISSUES` dict and re-run `python reports/generate_report.py`
  in the same PR**.  Don't paste bug content into `README.md`,
  `AGENTS.md`, or any other markdown — link to the report instead.
  `analysis/reverify_all_issues.py` cross-checks the classifier
  against the report and should be run when you touch any issue.
- This file (`AGENTS.md`) holds *operational* knowledge — the COM
  landmines, JET quirks, driver patterns.  Plan/status and issue
  content do *not* belong here.

## Repo layout

```
cbdb-user-mdb-tests/
├── data/                         # gitignored
│   ├── CBDB_BJ_User.mdb          # the .mdb under test (~48 MB)
│   ├── CBDB_<YYYYMMDD>_DATA.mdb  # linked data (~1 GB) — DO NOT MODIFY
│   └── HelpFiles/                # PDF refs (HelpFile_LookAt*.pdf)
├── analysis/                     # one-shot scripts + dumped metadata
│   ├── dump/                     #   tables.json / queries.json /
│   │                             #   forms.json / vba_modules.json /
│   │                             #   vba/Form_*.vb / test_inputs.json
│   ├── dump_metadata.py          # re-extract schema + queries
│   ├── dump_vba.py               # re-extract VBA source
│   ├── discover_test_inputs.py   # ⭐ scan DATA for high-density inputs
│   ├── probe_pywinauto.py        # ⭐ standalone-working VBA driver demo
│   ├── audit_view_aliases.py     # heuristic SQL audit
│   ├── ...                       # see reports/CBDB_Issues_Report_EN.md for what each found
├── tests/
│   ├── conftest.py               # session/fn-scoped fixtures
│   ├── cbdb_driver/              # COM + pywinauto drivers
│   │   ├── access_app.py         # AccessApp class (COM + ODBC)
│   │   ├── vba_session.py        # ⭐ THE working VBA driver
│   │   ├── vba_inject.py         # MsgBox suppression + ZZ_TEST tables
│   │   └── form_driver.py        # legacy abstraction (kept for ref)
│   ├── cbdb_replay/              # Python re-impl of CmdQuery_Click SQL
│   │   ├── lookat<X>.py          # one per form
│   │   ├── exports.py            # GIS / Neo4j / KML byte-level writers
│   │   └── common.py             # shared helpers
│   ├── golden/                   # frozen CSV / .tab / .kml snapshots
│   ├── test_schema.py            # table/col existence
│   ├── test_saved_views.py       # 18 View_* queries smoke + FK
│   ├── test_lookatentry.py       # SQL replay golden tests
│   ├── test_other_lookat_forms.py
│   ├── test_exports.py           # export byte format unit tests
│   ├── test_known_bugs.py        # regression for confirmed bugs
│   ├── test_vba_inline.py        # ⭐ first VBA-driven test (probe inline)
│   ├── test_vba_differential.py  # ⭐ VBA vs Python replay
│   ├── test_vba_integrity.py     # 12 data-integrity dimensions
│   ├── test_vba_matrix.py        # ⭐ data-driven param matrix
│   └── test_other_forms_skeletons.py
├── reports/                      # auto-generated bilingual issue report
│   ├── generate_report.py        #   ⭐ source of truth for all 19 issues
│   └── CBDB_Issues_Report_*.md   (committed; .docx is generated on demand and gitignored)
├── tests/MANUAL_SMOKE.md         # 5-min checklist (UI things tests miss)
├── tests/VALIDATION_DIMENSIONS.md  # menu of test dimensions
└── FINAL_STATE.md                # last regenerated state
```

## ⭐ Mission-critical landmines (DO NOT relearn these the hard way)

### 1. (legacy) Probe vs pytest: `pywinauto.click_input()` only fires if Access
window is in the FOREGROUND, and silently drops on disabled controls
or locked-screen sessions.  We've migrated to Form_Timer trigger
universally — see #4 below.  Original gotcha kept here as reference:

```python
# REQUIRED before any click:
main = pwa.window(title="Welcome to CBDB!")
main.wait("ready", timeout=10).set_focus()    # ← without this, click is dropped silently
time.sleep(0.5)
```

Symptom of missing this: click reports success, but `ZZ_SCRATCH_<X>` stays empty.

### 2. `gUseADDRID` and other Public globals don't auto-set when you
INSERT into picker scratch tables.

The form's `CmdQuery_Click` gates the address branch on the Public
global `gUseADDRID`, which is normally set by `CmdSelectPlace_Click`
or `CmdImportPlaces_Click`.  Tests that bypass pickers (by direct
INSERT) leave the global False and the address filter is ignored.

**Fix used here**: `tests/cbdb_driver/vba_session.py:_inject_autodetect()`
prepends a 3-line auto-detect to `CmdQuery_Click` that sets
`gUseADDRID = (DCount("ZZ_SCRATCH_ADDR") > 0)`.

If you change the autodetect, **do NOT use variable names starting
with underscore** — VBA forbids them and will crash with a compile
dialog mid-test (e.g. don't use `_td`, use `tdAddrCount`).

### 3. NAVIGATION_PANE.Form_Open hangs forever if `LinkListInit.c_path`
doesn't equal the working-copy file path.

The form's startup tries to relink all data tables to
`<path>_<version>_DATA.mdb`, which doesn't exist at our working-copy
location.  **Pre-patch via pyodbc BEFORE Access opens**:

```python
conn.cursor().execute(
    f"UPDATE LinkListInit SET c_path = '{work_path}'"
)
```

### 4. `Application.Run "Form_X.SubName"` does NOT work for form-module
subs on this Office install.

We tried 7 variants.  Use `vba_session.click_via_timer(form, ctl)`
instead — it injects a `Form_Timer` sub that calls the private
`<ctl>_Click` then disables the timer.  Trigger by setting
`Forms(form).TimerInterval = 100`.  Access fires Form_Timer itself,
bypassing both the Application.Run unreachability and any
disabled-button click drop.

(Earlier we used `pywinauto.click_input` for forms with enabled
buttons, but its UIA tree corrupts after a few open/close cycles in
one pytest session — assoc tests fail with 0 rows after biblcat,
plus desktop-locked sessions silently drop clicks.  Form_Timer is
the universal trigger now.)

### 5. `VBComponents.Add(1)` (new standard module) fails with
COM error 0x800AC471 even with `AccessVBOM=1`.

Workaround: append helpers to an EXISTING form module
(`Form_LookAtEntry`).  See `vba_inject.py:HOST_FORM_MODULE`.

### 6. Office's `AutomationSecurity` defaults to `2` (ByUI), which
blocks macros for COM-opened mdbs.

```python
app.AutomationSecurity = 1  # msoAutomationSecurityLow
# MUST be set BEFORE OpenCurrentDatabase()
```

### 7. The shipped .mdb has a broken DAO 3.6 reference.

`dao360.dll` doesn't exist on Office 2016+.  `vba_session.py:open()`
auto-detects + replaces with `ACEDAO.DLL`.  This is also a real bug
in the .mdb that the user's report found.

### 8b. LookAtOffice's CmdQuery starts `Enabled=False` and pywinauto
clicks are silently dropped on disabled controls.

Setting `Forms("LookAtOffice").Controls("CmdQuery").Enabled = True` via
COM works (verified — `Properties("Enabled") = True` afterwards), but
pywinauto's UIA tree caches the *disabled* state.  `click_input()`
returns success but Windows blocks the actual click (any click on a
disabled control is dropped).  After `force_enable`, `Repaint()`,
`time.sleep` — pywinauto STILL sees disabled.

Also: `Application.Run "Form_LookAtOffice.PublicWrapperSub"` returns
success but doesn't actually invoke (same as #4).

**FIX (use Form_Timer trigger):** `vba_session.click_via_timer(form, ctl)`
injects a `Form_Timer` sub into the form module that calls `<ctl>_Click`
on first tick, then sets `TimerInterval=0` to stop.  Python triggers
by setting `Forms("LookAtOffice").TimerInterval = 100`.  Access fires
the timer event itself — no click, no Application.Run, no UIA cache
issue.  Required for any form whose CmdQuery starts disabled.

The autodetect injection ALSO appends an `INSERT INTO ZZ_TEST_DEBUG
VALUES ('<form>:DONE')` right before the `Exit_<sub>:` label so
callers can poll for *true* completion: row_count alone is insufficient
because backfill UPDATEs don't change row count.  Office's CmdQuery_Click
runs ~3 backfill UPDATE statements (5-7 table joins each on 37k rows)
that take 30-60s each — the DONE marker is the only reliable signal
the chain finished.

After the chain finishes, **don't use pandas.read_sql for large
results** — it deadlocks with Access's internal recordset binding on
big local tables.  Use raw `vba.conn.cursor().execute(...).fetchall()`
instead.  See `tests/test_vba_matrix_all_forms.py` step 5+.

### 8e. `win32com.client.Dispatch("Access.Application")` can hit
Windows fatal exception 0x800706ba (RPC server unavailable) when the
ROT has stale entries from killed Access processes.  Use
`DispatchEx` instead — it forces a fresh out-of-proc instance.
`vba_session.open()` already does this.

### 8g. Real export tests — Form_Timer fires only once per OpenForm,
so CmdQuery + CmdGIS must run in the SAME fire.

Chain pattern: pass `"<chain>|<path>"` to `Forms(form).Tag` (in-process
property, no JET cache delay).  Form_Timer calls CmdQuery_Click.  An
autodetect-injected post-body block in CmdQuery_Click reads Me.Tag,
parses chain after the first comma, calls `CmdGIS_Click`.  CmdGIS_Click
uses `GetTestExportPath()` (also reads Me.Tag) instead of popping
SaveAs dialog (patch by `vba.patch_filedialog`).

Two non-obvious VBA gotchas the chain hit:
- VBA `Or` is NOT short-circuited.  `If GetTestExportPath() <> "" Or
  dlgSaveAs.Show = -1 Then` STILL pops the dialog because `.Show` is
  evaluated even when the test path is non-empty.  Use `If ... Then ...
  ElseIf .Show = -1 Then ... End If` instead.
- After `Call CmdQuery_Click` from inside Form_Timer, control DOES
  return cleanly — but only if the called sub completes.  Chaining
  TWO calls (Call A; Call B) inside Form_Timer hangs (CmdGIS_Click's
  internal MsgBox or recordset binding blocks).  Hence chain inside
  CmdQuery_Click body, not in Form_Timer.

Working impl: `tests/test_vba_export.py` + `vba_session.set_form_tag`
+ `vba_session.patch_filedialog` + autodetect-injected chain block in
`vba_session._inject_autodetect`.

### 8f. Don't try to gracefully close Access via `DoCmd.Close /
CloseCurrentDatabase / Quit` after a heavy CmdQuery_Click.  Those
COM calls hang for minutes while Access finishes background subform
renders / UPDATE chains.  `vba_session.close()` skips them and goes
straight to `taskkill /F` which is reliable.

### 8c. `Form_LookAtOffice.Form_Open` wipes `ZZ_OFFICE_CODE` on every
open.  No other LookAt form's Form_Open touches its own picker table.

Implication: must populate ZZ_OFFICE_CODE *after* `open_form()`, not
before.  See `tests/test_vba_matrix_all_forms.py::test_cross_form_matrix`.

### 8d. JET page-cache coherence: pyodbc INSERTs are invisible to
Access's internal SQL/DCount for several seconds.

After `set_picker_codes` (which writes via pyodbc), call
`app.DBEngine.Idle(8)` (dbRefreshCache) + `app.RefreshDatabaseWindow()`
to flush.  Without it, autodetect's `DCount("*", "ZZ_OFFICE_CODE")`
returns 0 even though pyodbc just inserted 1 row.  Already wired into
`vba_session.set_picker_codes`.

The current Python SQL-replay tests CANNOT find VBA bugs.

The `cbdb_replay/lookat<X>.py` files were written by reading the VBA
SQL and translating to Python.  If the VBA has a bug, the Python
replay almost certainly has the SAME bug.  Self-consistent =/ correct.

To find VBA bugs use `test_vba_differential.py` /
`test_vba_matrix.py` patterns: drive REAL VBA via pywinauto, compare
with INDEPENDENT source SQL (not the Python replay).

## Confirmed bugs in the live .mdb

→ **All 19 issues + reproduction steps live in
[`reports/CBDB_Issues_Report_EN.md`](reports/CBDB_Issues_Report_EN.md)**
([中文](reports/CBDB_Issues_Report_ZH-Hant.md)). The `ISSUES` dict
in `reports/generate_report.py` is the source of truth — do **not**
reproduce bug content in this file or in `README.md`.

Triage convention used by the report:

- **P0** silent data corruption — user can't detect from the UI
- **P1** visible runtime crash — error popup on a normal click
- **P2** silent display — sub-form column shows blank where data exists
- **P3** missing UI — handler exists but no button to fire it
- **P4** setup — one-time install fix
- **P5** dormant / latent / not currently reproducible — defect real
  but doesn't fire on the current dump or has no UI trigger today;
  **none have been verified as upstream-fixed** (see marker-failure
  policy below)

When triaging future findings, weight P0/P1 heavier than P2-P5.
Static scans tend to surface a lot of low-priority noise — flag it,
but don't let it crowd out the P0/P1 stuff.

Re-run ALL FOUR static auditors on every CBDB release — they're
cheap (seconds) and they keep finding bugs:
- `analysis/audit_missing_controls.py` — control-name typos
  (found Bug #5)
- `analysis/audit_sql_columns.py` — SQL `<Table>.<Column>` typos
  (found Bug #6, also expanded scope of Bug #5)
- `analysis/audit_sql_columns.py`-style INSERT/SELECT cardinality
  check is in `analysis/audit_insert_select_columns.py` — currently
  clean (0 findings on 295 pure-literal INSERT statements);
  guards against future regressions
- `analysis/audit_sql_table_names.py` — table-name typos in SQL
  strings; currently surfaces only `frmIndexAddr` (orphan
  maintenance form, end users can't reach), filed as 🟢 LOW
- `analysis/audit_saved_queries.py` — same `<Table>.<Column>` check
  as `audit_sql_columns.py` but applied to the 21 saved queries in
  `queries.json`.  Currently clean; guards against view definitions
  drifting away from the underlying table schema.
- `analysis/audit_recordset_fields.py` — tracks `Set <var> =
  CurrentDb.OpenRecordset("<TABLE>", ...)` per Sub and flags
  `<var>!<field>` references where `<field>` isn't on `<TABLE>`.
  Per-sub scope, invalidates on any `Set <var> = ...` reassignment
  (including reassignment to a SQL string we can't statically
  evaluate).  Currently clean.
- `analysis/audit_recordset_sql_projection.py` — sister scanner to
  the above for the SQL-string case.  Tracks `tQueryStr = "SELECT
  ..."` literal-only concats and `Set <var> = CurrentDb.OpenRecordset
  (tQueryStr)`, parses the SELECT projection, and flags `<var>!field`
  AND bare `!field` (inside `With <var>`) reads where `field` isn't
  projected.  **Found Bugs #7 / #8 / #9** (CmdNeo4j family across
  LookAtPlace / LookAtNetworks / LookAtEntry).  Run on every release.
- `analysis/audit_subform_control_sources.py` — for every sub-form
  whose RecordSource is a saved query (View_*), check each bound
  control's ControlSource exists in the saved query's SELECT
  projection.  **Found Bugs #10 / #11 / #12** (silent display bugs
  in EVENT_ADDR_2 / EVENTS_DATA_2 / POSTED_TO_OFFICE_DATA_2 sub-
  forms).
- `analysis/audit_error_label_targets.py` — every `On Error GoTo
  <label>` / `Resume <label>` / `GoTo <label>` must point at a
  label defined in the same Sub.  Currently clean — long-term
  guard for typo'd error-handler renames.
- `analysis/audit_event_handlers_exist.py` — every form-control
  event handler named in `control_inventory.json` must have a
  matching `Sub <name>()` defined in the form's VBA module.
  Currently clean — long-term guard for "renamed Sub but forgot to
  update OnClick property" silent-no-op bugs.
- `analysis/audit_dcount_where_columns.py` — every D-aggregate
  call (DCount / DLookup / DSum / etc.) with a literal table+criteria
  must reference columns that exist on the named table.  Currently
  clean — long-term guard for stale-criteria silent-False bugs.
- `analysis/audit_cross_form_references.py` — every
  `Forms!<form>!<ctl>` reference must resolve to an existing form
  AND existing control on that form (case-insensitive).  Skips
  Form__TMPCLP*.vb auto-backup snapshots.  **Found Bugs #13 / #14**
  (BIOG_MAIN_2_Subform / KIN_DATA_Subform reference picker forms
  that don't exist in the .mdb).
- `analysis/audit_doc_md_open_form.py` — every literal
  `DoCmd.OpenForm "<form>"` must resolve.  Currently clean — Bug
  #13's reference uses a string variable so it's caught by the
  cross-form-references audit instead.  Long-term guard for direct-
  literal regressions.
- `analysis/audit_dlookup_fields.py` — every `DLookup("<field>",
  "<table>", ...)` literal call must reference a valid field on
  the table.  Currently clean — companion to dcount-where-columns.
- `analysis/audit_orphan_event_handlers.py` — find Subs named
  like `<Control>_<Event>` where `<Control>` doesn't exist on the
  form.  Code-smell signal (exit 0, informational).  **Found
  Bugs #15-#19** (LookAtPlace / LookAtStatus / LookAtOffice each
  have export-button event handlers with no matching button on
  the form design — silent missing UI).
- `analysis/audit_blocking_msgbox.py` — list every
  `If MsgBox(...) = vb<Yes|No|...>` confirmation prompt.  Not bugs;
  informational guard so `tests/cbdb_driver/vba_session._inject_autodetect`
  knows which prompts to pre-arrange.
- `analysis/audit_control_row_sources.py` — for every ListBox /
  ComboBox with a non-empty RowSource SQL, verify each
  `<Table>.<Column>` reference is in the schema.  Currently clean —
  third leg of the SQL-column-resolution stool.

**Runner**: `analysis/run_all_audits.py` runs every static audit and
prints a FLAGGED / CLEAN summary.  Per-release workflow.  Latest run
on the shipped dump: 6 of 19 audits flagged, 6.5 s total.

All audits share `analysis/audit_lib.read_vba_lines` for proper
`\\r\\r\\n` handling so reported line numbers match grep / VBE.

These are guarded by `tests/test_known_bugs.py`; if those tests
start failing, the marker no longer reproduces — that's a signal
to investigate, **not** an automatic confirmation that upstream
fixed the bug.  The candidates are: (a) upstream actually patched
the source .mdb / VBA, (b) the input fixture or Access driver
behaviour changed out from under the test, (c) the original bug was
misclassified.  Only (a) justifies marking the issue as fixed in
`reports/generate_report.py`'s `ISSUES` dict, and only after
inspecting the new VBA / queries dump or hearing from the
maintainer.  Until then, prefer re-classifying (Dormant / Latent /
Not currently reproducible) over removing the issue.

### Index-year cross-check: User MDB ≠ cbdb-online-main-server SQLite (classification still open)

`tests/test_index_year_xcheck.py` compares the User MDB's
`BIOG_MAIN` against the weekly cbdb-online-main-server SQLite
snapshot.  As of 2026-05-02 it reports ~12 source-data drifts and
~575 derived-field drifts on a 657 246-person sweep.

**Two independent implementations, not one shared algorithm.**

- The cbdb-online-main-server SQLite snapshot's `c_index_year` is
  produced by the PHP service `IndexYearRebuildService.php`, and
  its `c_index_addr_id` by `IndexAddressRebuildService.php`.  Both
  PHP services live in
  <https://github.com/cbdb-project/cbdb-online-main-server>.
- The User MDB-side rebuild lives in TWO places:
    - `c_index_addr_id` → `analysis/dump/vba/Form_frmIndexAddr.vb`
      in the front-end (`CBDB_BJ_User.mdb`).
    - `c_index_year`    → 37 saved QueryDefs named `BM IY Rule …`
      in the linked-tables BACKEND
      (`data/CBDB_<YYYYMMDD>_DATA.mdb`), driven from
      `frmBaseMaintenance` in the same file.  `analysis/
      dump_data_mdb_algorithms.py` extracts them to
      `analysis/dump_data/querydefs_index/*.sql`; the form/module
      VBA source still needs an interactive Access SaveAsText
      pass.  (PR G originally claimed the year rebuild was
      missing — that grep was looking only at the front-end VBA
      dump.)
- The PHP services are intended to mirror the MDB/VBA maintenance
  logic, but they are **independent implementations** (likely
  ported, not the same code).  Any of the following can produce a
  per-row difference and they all look the same in a diff:

  1. **Source-data snapshot drift** — the online system updates its
     source rows continuously; the User MDB ships a point-in-time
     snapshot.
  2. **Algorithm / porting divergence** — the PHP service may pick
     up tweaks the VBA hasn't, or vice versa.
  3. **Priority / tie-break differences** — when multiple
     candidate years apply, the two implementations may pick
     differently.
  4. **Null / default handling differences** — e.g. how a missing
     birthyear collapses to 0 vs NULL vs "use deathyear minus 60".

**Do not assume any individual diff is "just data drift".**  We
have not done a full classification.  The thresholds in the test
(0.5 % on derived fields, 0.1 % on source fields) are coarse
guards — they would catch a sudden jump (e.g. a wholesale
algorithm regression) but say nothing about whether the steady
~575 we already see are drift, divergence, or genuine bugs.

If you're triaging a specific divergence, you need to walk the
input rows on both sides and reproduce by hand which implementation
yields which value.  The drift appendix in
`reports/CBDB_Issues_Report_EN.md` shows a handful of representative
examples to start from, but the sampled set (currently 13 rows
across 3 buckets) is illustrative, not statistically
representative.  Don't open or close CBDB issues based on it
without classification.

**Classification framework (`analysis/classify_index_drift.py`).**
Runs the full per-personid taxonomy and writes
`reports/index_drift_classification.json`.  Latest run on the
shipped dump (657 245 common personids):

| Bucket | Count | Meaning |
|---|---:|---|
| `exact_match` | 656 682 | all four compared fields agree |
| `source_drift_index_agrees` | 2 | source drift but indices agreed |
| `source_drift_index_diffs_too` | 14 | source drift AND ≥1 index differs (consistent with simple data-drift hypothesis) |
| `index_year_only_diff` | 59 | source matched, only c_index_year differs — needs follow-up |
| `index_addr_only_diff` | 478 | source matched, only c_index_addr_id differs — needs follow-up |
| `index_both_diff` | 10 | source matched, both indices differ — strongest single-row signal |

Net diffs 563 / 657 245.  16 attributable to source drift; **547
unclassified**.  The unclassified buckets are not automatically
bugs — they could be PHP↔VBA divergence OR drift in evidence
tables (BIOG_ADDR_DATA, ENTRY_DATA, NIAN_HAO, fl_*_year, etc.)
that the classifier does not compare.  Per-row triage required.

Algorithm pointers (each side, where the source code lives) are in
`analysis/index_drift_algorithm_notes.md`.  PR N (replacing PR I) compares the **runtime** Access source
(`GetBirthIndexYearSQL` inside `frmBaseMaintenance`, dumped by
PR M) against PHP `IndexYearRebuildService.php` (pinned commit
`a642f7a`) → `analysis/index_year_rule_comparison.{md,json}`.
Pairing is by emitted `c_index_year_type_code` (both sides emit
them).  Verdicts: **22 matched, 8 matched_minor_diff, 0
logic_diff**, 3 access-only (concubine wife variants 31/32/33),
0 php-only.  PR I's `+N`/`-N` sign-flip flag was an artefact of
comparing PHP against the wrong Access source (the 37 saved BM
IY Rule QueryDefs, which `CmdIndexYear_Click` does not call —
they're vestigial).  At the rule level the runtime Access path
matches PHP almost everywhere; the closest thing to a real
divergence is off-by-1 (Rule 29) / off-by-3 (Rule 30) on
deathyear-default offsets.  None confirmed as bugs.

PR K1 took those rule-level findings and ran them against the 69
year-only diffs (59 `index_year_only_diff` + 10 `index_both_diff`)
from PR G — see `analysis/classify_index_year_drift_by_rule.py`
and `reports/index_year_drift_rule_classification.json`.  Headline:

  - 19 `php_did_not_compute`         (Access has a value, PHP wrote 0/null)
  -  7 `access_did_not_compute`      (PHP has a value, Access wrote 0/null)
  -  5 `iteration_order_diff`        (Phase-C propagation differs;
                                       e.g. PHP tcode '11' vs Access '1112')
  - 14 `consistent_within_rule`      (multiple rows share the same
                                       (php_tcode, access_tcode, diff)
                                       triple — single rule-level
                                       cause to investigate)
  -  1 `php_returned_sentinel`       (PHP value ≥ 9999, looks like
                                       overflow / garbage)
  -  5 `candidate_algorithm_divergence` (row's signature matches
                                          one of K1's dormant
                                          historical probes but
                                          can't reconstruct from
                                          a single evidence row)
  - 18 `unclassified`

Total 69, ~74 % bucketed into named patterns (still none labelled
as confirmed bugs; the buckets are evidence categories for the
maintainer).

PR K2 (`analysis/triage_index_year_drift_groups.py` →
`reports/index_year_drift_rule_groups.json`) does a second triage
pass on what K1 left under-named:

  - **consistent_within_rule** (14 rows): 5 (php_tcode,
    access_tcode, diff) signature groups, all named
    `candidate_same_rule_tie_break_or_aggregation_diff` (both
    sides chose the same rule but produced different values; tie-
    break or aggregation diff suspected).  A recurring **diff =
    -20 across Rules 11 / 13 / 15 / 19** stands out.
  - **unclassified** (18 rows): all 18 named after triage; 17
    flagged `blocked_by_runtime_priority_triage_pending`
    (different rules picked on each side; PR M dumped
    `frmBaseMaintenance` so the source is in repo, but
    deciding which side's choice is intentional still needs a
    per-row walk of the runtime priority / iteration order in
    `GetBirthIndexYearSQL`).  PR X renamed this label from
    the older `blocked_by_missing_frmBaseMaintenance_vba`,
    which became stale once PR M shipped.
  - **php_did_not_compute** (19 rows): 6 groups by Access tcode.
    Biggest is `access_tcode='05'` × 7 — `candidate_php_entry_
    code_mapping_gap`.  PR N has runtime Rule 05 matched
    against `sqlEntryRule('040101', 30, '05')`; the row
    evidence supports PHP missing the entry-code → '040101'
    map for those 7 personids.

After K2, the 17 leftover rows are blocked on **runtime
priority / iteration-order triage** (not on missing source —
that was PR M's job).  Re-run the comparator + classifier +
triage after every fresh SQLite snapshot or any change to the
index-recompute path on either side.

PR L (`analysis/classify_index_addr_drift.py` →
`reports/index_addr_drift_classification.json`) classifies the
488 c_index_addr diffs (the 478 `index_addr_only_diff` plus 10
`index_both_diff` from PR G).  PHP source pinned at commit
`e31fba7` of `app/Services/IndexAddressRebuildService.php` →
`analysis/php_source/IndexAddressRebuildService.php`.  Headline:

| Bucket | Count | Meaning |
|---|---:|---|
| `mdb_stale_index_addr` | **412** | SQLite stored value matches what we'd recompute from its BIOG_ADDR_DATA; User MDB stored value does NOT.  Most likely cause: the User MDB shipped with stale c_index_addr_id — BIOG_ADDR_DATA was updated after the last frmBaseMaintenance rebuild but nobody re-ran it.  PHP re-runs weekly so SQLite stays fresh. |
| `mdb_value_php_null` | 47 | User MDB has a value, PHP wrote 0/null.  Most are personids whose BIOG_ADDR_DATA is missing from the SQLite snapshot. |
| `same_candidates_diff_winner` | 10 | Identical BIOG_ADDR_DATA on both sides; both pick addr_type=1 but a different addr_id within it.  Tie-break / null-handling diff. |
| `both_stale_recompute_mismatch` | 10 | Neither side's stored value matches the rank+MAX(c_sequence) recompute.  Could be older snapshots or an algorithm feature we don't model. |
| `both_sides_match_recomputed` | 6 | Each side's stored value matches its own recompute; diff is BIOG_ADDR_DATA snapshot drift. |
| `sqlite_stale_index_addr` | 2 | Reverse direction; rare. |
| `mdb_null_php_value` | 1 | |
| `unclassified` | 0 | |

The 412 `mdb_stale_index_addr` is **not** a bug in either
algorithm — it's a maintenance-cadence diff (the User MDB needs
its `frmBaseMaintenance` rebuild re-run before the next release).
The 10 `same_candidates_diff_winner` rows are the only candidate
algorithm-divergence rows.

PR S deep-dived those 10 (`analysis/deep_dive_addr_same_candidates.py`
→ `reports/index_addr_same_candidates_deep_dive.json`).
**10 / 10 have MAX(c_sequence) ties** — multiple BIOG_ADDR_DATA
rows of the winning addr_type all sharing the same c_sequence.
PHP, Access maintenance, and the PR L recompute each pick a
different row depending on the underlying engine's storage order.
Both implementations follow the same documented rule (rank-
priority + MAX c_sequence); neither is "wrong"; the result is
just non-deterministic when ties occur.  Candidate mitigation:
add an explicit secondary tie-break (e.g. MIN(c_addr_id)) to both
sides.  Treated as `candidate_release_process_or_algorithm_
improvement`, not a confirmed bug.

Verified separately that BIOG_ADDR_CODES (the rank table) is
**identical** between the two sides for all 22 addr_types, so the
divergence isn't coming from rank-priority configuration.

PR M dumped `frmBaseMaintenance` and the 4 DATA-mdb modules via
`Access.Application.SaveAsText` →
`analysis/dump_data/vba/Form_frmBaseMaintenance.vb` etc.  Two
findings rewrite parts of PR I and PR L (see
`analysis/index_drift_algorithm_notes.md` § "Maintenance trigger
path" for the long version):

  1. The **`CmdIndexYear` button calls `GetBirthIndexYearSQL`**,
     which is inline VBA closely mirroring PHP
     `IndexYearRebuildService.php` (subtractive offsets like
     `c_year - 30`, `ENTRY_CODE_TYPE_REL` joins, etc.).  It does
     NOT call the 37 saved `BM IY Rule` QueryDefs PR I compared
     against.  PR I's "logic_diff" sign-flip flag is
     **invalidated**: the runtime path matches PHP on those
     rules.  PR I's JSON / markdown stay in repo as historical
     evidence + a methodological warning.  PR X re-derived K1 /
     K2's docstrings + in-script rationales + JSON output
     strings against PR N's runtime comparison and renamed
     the stale `blocked_by_missing_frmBaseMaintenance_vba`
     label to `blocked_by_runtime_priority_triage_pending`
     (the source is in repo since PR M; the blocker is now
     the priority/iteration-order walk).
  2. The **`CmdIndexAddress` button** uses a per-rank UPDATE that
     does **NOT** explicitly aggregate by `MAX(c_sequence)` (the
     way PHP and the front-end `Form_frmIndexAddr.vb` do).  When
     a person has multiple BIOG_ADDR_DATA rows of the same
     addr_type, this UPDATE picks whichever row JET surfaces
     first — a candidate algorithmic divergence between the
     maintenance code and PHP, on top of the maintenance-cadence
     issue PR L flagged.

Operational implication for the User MDB release process: the
maintenance buttons are manual, with no scheduler / AutoExec
trigger.  Adding a release-checklist step that runs
`CmdIndexYear` then `CmdIndexAddress` on the DATA mdb before
shipping a new User MDB is the most direct mitigation for the 412
`mdb_stale_index_addr` rows.  (Candidate, not prescribed — the
maintainer's call.)

## Test inventory snapshot (run this to get current state)

```bash
python -m pytest tests/ -W ignore --ignore=tests/test_infra_smoke.py \
    --ignore=tests/test_vba_inline.py \
    --ignore=tests/test_vba_differential.py \
    --ignore=tests/test_vba_integrity.py \
    --ignore=tests/test_vba_matrix.py
# fast suite (~30s): SQL replay + saved views + schema + exports
```

```bash
python -m pytest tests/test_vba_*.py -W ignore --tb=short -p no:cacheprovider
# slow VBA-driven suite (~10-30 min): real VBA + pywinauto + integrity
```

## ⭐ Data-driven testing (the right pattern)

The user explicitly asked for this — DO follow it.

**Wrong**: pick fixture inputs (entry codes, dynasties) by hand.
Half the time they hit sparse data and tests pass with 0-row results,
which proves nothing.

**Right**: query the live `DATA.mdb` for high-density combinations,
serialize to `analysis/dump/test_inputs.json`, drive parametrized
tests off it.

```bash
# After every DATA.mdb update:
python analysis/discover_test_inputs.py   # ~5 sec
# Then run the matrix tests; new high-density combos drive new fixtures.
```

`analysis/discover_test_inputs.py` finds:
- top entry codes by row count
- top addresses by indexed-person count
- top dynasties
- (entry × dynasty), (entry × addr), (status × dynasty) high-density combos
- top persons by kin / association count (good network seeds)
- top assoc / kin / office codes
- well-connected (person, person) pairs for AssociationPairs

`tests/test_vba_matrix.py` consumes the JSON to build VBA-driven
parametrized fixtures automatically.

### CRITICAL: re-run discovery on EVERY data update

CBDB grows over time.  The "top entry code" today may not be top in
six months.  Plus new addresses / dynasties / status codes can
appear.  Stale `test_inputs.json` defeats the whole point.

`tests/conftest.py` should auto-run discovery if `test_inputs.json`
is older than `data/CBDB_BJ_User.mdb` (TODO — see Issues below).

## Standard workflow after a `.mdb` update

```bash
# 1. Re-extract metadata + VBA source (so analysis/dump/* is fresh)
python analysis/dump_metadata.py
python analysis/dump_vba.py

# 2. Re-discover high-density test inputs
python analysis/discover_test_inputs.py

# 3. Fast tests (~30s) — schema, saved views, exports
python -m pytest tests/ -W ignore \
    --ignore=tests/test_vba_inline.py \
    --ignore=tests/test_vba_differential.py \
    --ignore=tests/test_vba_integrity.py \
    --ignore=tests/test_vba_matrix.py \
    --ignore=tests/test_infra_smoke.py

# 4. Review golden CSV diffs
git diff tests/golden/

# 5. If diffs are intentional data updates (not bugs), bless them
python -m pytest tests/ --regenerate-goldens

# 6. Slow VBA-driven tests (~10-30 min) — actually run the form code
python -m pytest tests/test_vba_*.py -W ignore -p no:cacheprovider

# 7. Manual 5-min smoke (the few things tests can't cover):
#    open Access, click each LookAt form, verify exports, see MANUAL_SMOKE.md
```

## Open issues / TODOs

1. ✅ DONE — `test_vba_matrix_all_forms.py` covers 7 LookAt forms
   (Status / Texts / Associations / Office / Place / Kinship +
   LookAtEntry separately in `test_vba_matrix.py`)
2. ✅ DONE — `test_vba_export.py` real CmdGIS export, byte-level diff
3. Auto-run `discover_test_inputs.py` from conftest if json is stale
4. Three forms still skipped in matrix (need smaller fixtures + lower
   distance constraints to avoid recursive expansion timeout):
   - LookAtAssociationPairs (CmdQuery times out — TxtID1+TxtID2 set
     but more preconditions needed)
   - LookAtNetworks (CmdRun times out on Zhu Xi 2471 assocs — try
     someone with <100 assocs and constrain distance)
   - LookAtGroupData (CmdRun aggregates many tables)
5. Picker dialog tests (frmPickEntry_multi etc. — currently bypassed
   by direct ZZ_SCRATCH_<CODE> writes)
6. Real export tests for the OTHER buttons (CmdNeo4j / CmdKML / CmdPajek
   / CmdGUESS / CmdGephi); current test_vba_export.py only covers
   LookAtEntry CmdGIS

## Operating principles (these are the ones that bit me hard — internalize them)

- **Probe scripts in `analysis/probe_*.py` always work first**; pytest
  comes second.  If a probe works but the pytest equivalent doesn't,
  the difference is in fixture lifecycle / windowing.
- **Differential testing is the only way to find VBA bugs.**  Same-source
  Python replay finds nothing.  Independent SQL or HelpFile-documented
  values are the source of truth.
- **NEVER assert on randomly-chosen test inputs.**  Use
  `discover_test_inputs.py` to pick high-density real-world params.
  An assertion of "≥ N rows" with N=0 is a useless test.
- **Function-scoped Access fixtures > session-scoped**.  Each VBA test
  pays ~12s for a fresh Access instance, but isolation is worth it.
  Trying to share Access across tests led to days of debugging
  (the `test_infra_smoke.py @skip` is a relic).
- **VBA injection should be MINIMAL** — every line we add is a
  potential parser failure that crashes Access mid-test (and pops a
  modal dialog that requires user click to dismiss).  Every change
  must be valid VBA syntax (e.g. variable names can't start with `_`).
- **Use existing form modules, not new modules**, for any helper VBA
  injection.  `VBComponents.Add(1)` doesn't work on this install.

## Key contacts / context

- The `.mdb` is the canonical Harvard CBDB User-Interface mdb.
- Schema docs: `data/HelpFiles/CBDB Users Guide.pdf` (English), `CH.pdf` (Chinese).
- Per-form behaviour docs: `data/HelpFiles/HelpFile_LookAt*.pdf`.

## Memory

This file complements `~/.claude/projects/.../memory/` notes (see
`MEMORY.md` index).  When in doubt, this file (AGENTS.md) is
authoritative for THIS project; memory is general user/preference info.
