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

## Single source of truth: README "Plan & status"

The roadmap, current coverage table, confirmed bugs, and per-item
status all live in **`README.md` § Plan & status**. That is the
*single* place to look — and the *single* place to update.

**Hard rule for any agent (or human contributor):** if your change
moves a roadmap item from open → in-progress → done, or completes /
discovers / skips a fixture, **you must update the corresponding row
in the README's Plan & status table in the same PR**. Do not leave
the README out of sync. Do not duplicate plan content into other
files. CI/reviewers should fail PRs that change `tests/` substantively
without touching `README.md`.

This file (`AGENTS.md`) holds *operational* knowledge — the COM
landmines, JET quirks, driver patterns. Plan/status content does *not*
belong here.

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
│   ├── ...                       # see findings.md for what each found
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
├── findings.md / findings_en.md  # confirmed bugs report
├── MANUAL_SMOKE.md               # 5-min checklist (UI things tests miss)
├── PHASE1_BREAKTHROUGH.md        # why pywinauto works in probe
├── VALIDATION_DIMENSIONS.md      # menu of test dimensions
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

## Confirmed bugs in the live .mdb (findings.md)

1. **`View_StatusData` alias swap** — `c_fy_range_desc` /
   `c_fy_range_chn` pull from `YEAR_RANGE_CODES_1` (the LY join).
   First-year range value displayed is actually the LY value.
2. **DAO 3.6 reference broken** — see #7 above.

These are guarded by `tests/test_known_bugs.py`; if those tests start
failing, it means upstream fixed the bug (good — flip the asserts).

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
