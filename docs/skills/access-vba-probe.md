# Skill: access-vba-probe

**Status:** repo-local draft (2026-05-05).  Not installed globally;
read this file before driving Access via COM or writing a new
real-VBA test.

## When to use

Trigger this skill any time the work involves:

- Investigating whether a documented CBDB bug actually fires on the
  current dump
- Writing a new real-VBA test (anything in `tests/test_vba_*.py`)
- Probing a form's CmdQuery / CmdRun / Cmd<Export> chain to see
  what files / errors / popups it produces
- Capturing fresh screenshots for a report issue
- Reproducing a maintainer-reported manual symptom

If the task can be answered from the VBA dump (`analysis/dump/vba/`),
saved-query JSON (`analysis/dump/queries.json`), control inventory
(`analysis/dump/control_inventory.json`), or pure pyodbc against the
MDB, do that FIRST and skip this skill.  Access COM is the heaviest
hammer; the per-session cost is ~12 s and the failure modes are
nasty (RPC death, focus loss, Form_Open deadlock).

## Authoritative files

| File | What it owns |
|---|---|
| `tests/cbdb_driver/vba_session.py` | The VbaSession class — production driver.  Owns Access lifecycle (DispatchEx, AutomationSecurity, LinkListInit pre-patch, DAO ref repair), VBA injection (`_inject_autodetect`, `_inject_timer_trigger`), Form_Timer dispatch (`click_via_timer`), dialog patching (`patch_filedialog`), config blackboard (`set_form_tag` / `set_timer_target`).  Read this before adding any new test. |
| `tests/cbdb_driver/access_app.py` | Lower-level helpers: `kill_orphan_access` (gated on `CBDB_KILL_ALL_ACCESS=1`), `make_working_copy`, broken-DAO repair. |
| `tests/cbdb_driver/vba_inject.py` | MsgBox suppression patterns + `ZZ_TEST_DEBUG` / `ZZ_TEST_CONFIG` table setup. |
| `tests/test_vba_networks_small_fixture.py` | Reference test for the **minimal-injection** (Networks) pattern.  Copy its scaffolding when adding any Networks-touching test. |
| `tests/test_vba_matrix_hard_forms.py` | Reference test for **small-fixture** matrix coverage (AssociationPairs 4×5, GroupData person_1).  CAVEAT: its assertions are intentionally loose; do NOT read its silent-pass as evidence that an export chain works. |
| `tests/test_vba_cmdneo4j_cross_form.py` | Reference for the multi-file CmdNeo4j chain pattern + `_NEO4J_SHAPES` / `_NEO4J_SHAPES_BY_TWO_COLS` depth classifier + per-form structural assertions (e.g. LookAtEntry's InstitutionCodes-absent assertion). |
| `analysis/dump/vba/Form_*.vb` | The dumped VBA source.  AUTHORITATIVE for "does this Sub exist", "what columns does this SELECT project", "what's the early-bail condition".  Read this first; do not rely on memory. |
| `analysis/dump/control_inventory.json` | Per-form control catalogue.  AUTHORITATIVE for "does this button exist on the form".  Read this before claiming a button is missing. |
| `AGENTS.md` § Mission-critical landmines | The full landmine list.  Read it before your first probe; refer back when something hangs / errors. |

## Pure SQL first, COM only when necessary

**Pure SQL via pyodbc** answers most questions without ever opening
Access:

```python
import pyodbc
conn = pyodbc.connect(
    "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    "DBQ=data/CBDB_BJ_User.mdb;",
    autocommit=True, readonly=True,
)
```

Use pyodbc when the question is:

- Does any row in `<TABLE>` satisfy `<predicate>`?  (e.g. PR
  `investigate/issue9-neo4j-institutioncodes` proved
  `ENTRY_DATA.c_inst_code > 0 == 0` across all 263,454 rows)
- What's the column shape of a saved query / view?
- Are two persons in the same dynasty?  Does this picker code
  appear in the codes table?

Escalate to Access COM only when:

- The bug only fires inside `Cmd<X>_Click` runtime logic
- The maintainer reports a popup / hung dialog the static analysis
  can't predict
- You need to verify a multi-file export chain end-to-end (Neo4j /
  Pajek / Gephi / GIS shape + content)
- A bug's gate condition can be confirmed statically but the actual
  user-visible behaviour (popup / no popup, file present / absent)
  needs runtime evidence to ship in the report

When you DO open COM, drive via `VbaSession`, never via raw
`win32com.client.Dispatch("Access.Application")` — the production
driver handles a dozen landmines for you (DispatchEx, security
mode, LinkListInit pre-patch, DAO ref repair, Form_Timer
injection, dialog redirect, scoped per-PID kill on close).

## Standard probe scaffolding

Place probes under `analysis/_*` (gitignored).  Pattern:

```python
import os, sys, time
sys.path.insert(0, r'<repo>\tests')
os.environ['CBDB_KILL_ALL_ACCESS'] = '1'

from pathlib import Path
from cbdb_driver.vba_session import VbaSession
from cbdb_driver.access_app import kill_orphan_access

ROOT = Path(r'<repo>')
SRC = ROOT / 'data' / 'CBDB_BJ_User.mdb'
WORK = ROOT / 'analysis' / '_my_probe.mdb'
OUT_DIR = ROOT / 'analysis' / '_my_probe_out'

kill_orphan_access(); time.sleep(45)   # see "Common traps" §
OUT_DIR.mkdir(exist_ok=True, parents=True)
for f in OUT_DIR.iterdir():
    try: f.unlink()
    except OSError: pass

s = VbaSession(SRC, WORK).open()
try:
    s.open_form('LookAt<Form>')
    s.set_picker_codes('ZZ_SCRATCH_<X>', [<id>], '<col>')
    s.set_control('LookAt<Form>', '<ctl>', <val>)
    s.patch_filedialog('LookAt<Form>')
    s.set_form_tag('LookAt<Form>', 'CmdQuery,Cmd<Export>',
                   str(OUT_DIR) + '\\')
    n = s.click_via_timer('LookAt<Form>', ctl='CmdQuery',
                          result_table='ZZ_SCRATCH_<X>', timeout=180)
    cur = s.conn.cursor()
    cur.execute('SELECT id, msg FROM ZZ_TEST_DEBUG ORDER BY id')
    for r in cur.fetchall():
        print(r[0], '|', (r[1] or '')[:120])
    for f in sorted(OUT_DIR.glob('f*')):
        head = f.read_bytes()[:200].decode('utf-8', errors='replace').lstrip('﻿')
        print(f.name, head.split('\n', 1)[0])
finally:
    try: s.close()
    except Exception: pass
```

Key VbaSession primitives:

| Primitive | Purpose |
|---|---|
| `open_form(name)` | Opens the form (handles per-form quirks like LookAtOffice's `Form_Open` wiping `ZZ_OFFICE_CODE` — populate pickers AFTER open). |
| `set_picker_codes(table, ids, col)` | Writes ids into a `ZZ_SCRATCH_<X>` picker table via pyodbc + flushes JET page cache (`DBEngine.Idle(8)` + `RefreshDatabaseWindow`). |
| `set_control(form, ctl, val)` | Sets a control value via COM. |
| `patch_filedialog(form)` | Rewrites every `dlgSaveAs.Show` in the form module to short-circuit through `GetTestExportPath()`.  Required before any export-chain test; supports directory mode (trailing `\` → `f<n>.out` per call). |
| `set_form_tag(form, chain, path)` | Sets the form's `.Tag` to `<chain>|<path>`.  The autodetect chain block reads this and dispatches to `Cmd<X>_Click` calls after CmdQuery. |
| `click_via_timer(form, ctl, result_table, timeout)` | Sets `TimerInterval=100`; the injected Form_Timer body calls `<ctl>_Click`; polls `result_table` for change, then waits for the `<form>:DONE` marker in `ZZ_TEST_DEBUG`. |
| `set_timer_target(target)` | Tells Form_Timer which `<ctl>_Click` to dispatch to (used when the chain has multiple buttons). |

## Mission-critical landmines (must internalize)

These are the ones that have eaten the most agent-hours.  Read
`AGENTS.md` § Mission-critical landmines for the full list; the
ones below are the absolute floor.

### LinkListInit pre-patch

Without it, `NAVIGATION_PANE.Form_Open` hangs forever trying to
relink data tables to a non-existent `CBDB_<ver>_DATA.mdb` at the
working-copy path.  `VbaSession.open()` does this automatically
via pyodbc BEFORE `OpenCurrentDatabase`.  If you write a probe
that bypasses `VbaSession.open()`, you MUST do:

```python
conn.cursor().execute(
    f"UPDATE LinkListInit SET c_path = '{work_path}'"
)
```
…before calling `OpenCurrentDatabase`, or Access locks up.

### AutomationSecurity = 1 BEFORE OpenCurrentDatabase

```python
app = win32com.client.DispatchEx("Access.Application")
app.AutomationSecurity = 1     # MUST be before OpenCurrentDatabase
app.Visible = True
app.OpenCurrentDatabase(...)
```

`VbaSession.open()` already does this.  Default `2` (ByUI) blocks
macros for COM-opened mdbs.

### DispatchEx, never Dispatch

`Dispatch("Access.Application")` can hit fatal exception 0x800706ba
(RPC server unavailable) when the ROT has stale entries from
previously-killed Access processes.  `DispatchEx` forces a fresh
out-of-proc instance.  `VbaSession.open()` already uses it.

### Form_Timer is the universal trigger

`Application.Run "Form_X.SubName"` does NOT work for form-module
subs on this Office install (we tried 7 variants).  pywinauto
`click_input` is fragile: it drops on disabled controls AND
silently misses on locked-screen sessions AND its UIA cache
corrupts after a few open/close cycles in one pytest session.

Use `click_via_timer(form, ctl)` exclusively for new tests.  It
injects a `Form_Timer` body that calls `<ctl>_Click` directly,
bypasses focus / disabled-state issues, and is the only reliable
trigger for forms whose CmdQuery starts disabled (LookAtOffice).

### Networks Form_Open landmine (#3.5)

LookAtNetworks's `Form_Open` deadlocks under default
`_inject_autodetect` because ANY sibling `Form_LookAt*` module
modification dirties the VBA project, and Networks's Form_Open
self-references on its own subform recordsets during open, which
hits a project-wide auto-compile interaction.  PR AR-AY bisected
this fully; the only viable workaround is **minimal injection**:

```python
SKIP_SIBLINGS = {
    "Form_LookAtEntry", "Form_LookAtOffice", "Form_LookAtStatus",
    "Form_LookAtTexts", "Form_LookAtAssociations",
    "Form_LookAtPlace", "Form_LookAtKinship",
    "Form_LookAtAssociationPairs", "Form_LookAtGroupData",
}
sess = VbaSession(SRC, WORK,
                  skip_inject_autodetect_forms=SKIP_SIBLINGS)
```

Live reference: `tests/test_vba_networks_small_fixture.py`.  The
general matrix Networks case stays skipped.  Fixing it requires
either re-architecting the matrix harness for per-form minimal
injection, or a deeper Access-side workaround — both are scope-
defining design work, NOT autopilot.

### `_inject_autodetect` variable-name rule

Variable names in injected VBA must NOT start with underscore
(e.g. don't use `_td`; use `tdAddrCount`).  VBA forbids it;
crash dialog mid-test requires manual click to dismiss.

### LookAtOffice quirks

- `Form_Open` wipes `ZZ_OFFICE_CODE` on every open.  Populate
  AFTER `open_form`.
- CmdQuery starts disabled — must use Form_Timer trigger; don't
  bother trying pywinauto-click.
- CmdQuery_Click runs ~3 backfill UPDATE statements (5-7 table
  joins each on 37k rows) that take 30-60 s each.  The `:DONE`
  marker in `ZZ_TEST_DEBUG` is the only reliable completion
  signal — row count alone misses backfills.

### AssociationPairs CmdQuery SetFocus blocker (PR cover/assocpairs-pajek-gephi)

`Form_LookAtAssociationPairs.CmdQuery_Click:1635` calls
`Me.CmdQuery.SetFocus`.  Under Form_Timer dispatch the active form
is welcome / NAVIGATION_PANE (NOT LookAtAssociationPairs), so
SetFocus to a control on a non-active form raises VBA error 2110
("Welcome to CBDB! can't move the focus to the control CmdQuery").
CmdQuery exits via the error handler before its INSERT statements
run; `ZZ_SOCIAL_NETWORK = 0`; downstream CmdGIS / CmdPajek /
CmdGephi / CmdNeo4j all bail on `RecordCount = 0`.

`tests/test_vba_matrix_hard_forms.py` silently swallows this
because `_check_assoc_pairs` doesn't assert row count.  **Don't
read matrix_hard_forms's "passing" state as evidence that an
AssociationPairs export chain works.**  Until a driver-side patch
strips the SetFocus calls (out of scope for autopilot per AGENTS),
all 4 AssociationPairs export cells are bucket B.

### Don't `DoCmd.Close / CloseCurrentDatabase / Quit`

After a heavy CmdQuery_Click those COM calls hang for minutes
while Access finishes background subform renders / UPDATE chains.
`VbaSession.close()` skips them and goes straight to scoped
`taskkill /F`.

### Cross-COM-session orphans need cooldown

After a hard COM crash (RPC unavailable, fatal exception) MSACCESS
.exe processes may linger.  Use `kill_orphan_access()` (gated on
`CBDB_KILL_ALL_ACCESS=1`) + a `time.sleep(30-60)` cooldown before
the next probe.  Without the sleep, COM injection (`AddFromString`)
intermittently fails with RPC errors.  Heavier forms (Office's
37k-row CmdQuery, AssociationPairs' Link1stOrder) need longer
cooldown — 60-120 s is not unusual.

## Probe-result interpretation

When a probe completes, distinguish:

- **Genuine pass** — `<form>:DONE` marker present in
  `ZZ_TEST_DEBUG`; expected output files present; no `<form>:ERR`
  marker.
- **Genuine fail** — `<form>:ERR` marker present with VBA error
  description; or expected output file absent.
- **Driver flake** — `click_via_timer` returns `n=0` even though
  `ZZ_SOCIAL_NETWORK / ZZ_SCRATCH_*` is non-empty when queried via
  `vba.conn.cursor()`.  Common race: the watcher polls the table
  during the DELETE-then-INSERT window of CmdQuery and sees 0.
  Re-check the table count via pyodbc before declaring failure.
- **Silent swallow** — chain "passes" but nothing meaningful was
  asserted (matrix_hard_forms's loose assertions are the textbook
  example).  Trust direct table reads + file inventory, not the
  test's exit code.

For multi-file exports always classify file shapes (see
`tests/test_vba_cmdneo4j_cross_form.py::_NEO4J_SHAPES`).  An
unrecognized first-column header indicates the classifier is
missing a shape entry — extend it (PRs `fix/cmdneo4j-classifier-
lookattexts` / `cover/lookatoffice-cmdneo4j-peopleoffice` /
`hygiene/cmdneo4j-classifier-officecode-codes` are the reference
patches).

## Screenshot probes

- Real Access screenshots live in `reports/screenshots/`.  Do NOT
  add a screenshot whose caption asserts present-tense user
  triggering for a P5 latent issue (the
  `analysis/audit_report_screenshot_consistency.py` regression
  test will block the PR).
- Faux MsgBox popups (composited in PIL) are acceptable for P5
  latent issues IFF the caption explicitly says "Hypothetical" /
  "Users currently CAN'T trigger this".  See Issue #4's
  `bug4_step3_faux_popup.png` for the gold-standard caption shape.
- For runtime captures, drive via the existing
  `reports/capture_screenshots.py` helpers (LinkListInit
  pre-patch + ACEDAO repair + autodetect — all the same landmines).
- If the symptom is "missing file" rather than "popup", the right
  screenshot is the output folder showing the missing file (NOT a
  faux popup).  PR `reclassify/issue9-latent-source-typo` removed
  three misleading bug9 screenshots for exactly this reason.

## When NOT to change the issue from an investigation PR

Investigation PRs (probe-driven, read-only-on-the-canonical-report)
should ship:

- `analysis/<investigation_topic>.md` — human-readable findings
- `reports/<investigation_topic>.json` — machine-readable evidence
- `analysis/<investigation_script>.py` — reproducible probe (with
  optional `--com` mode if Access is needed)
- (rarely) `reports/screenshots/<new_screenshot>.png` — only if the
  finding fundamentally changes what the user sees today

The investigation PR should NOT touch `reports/generate_report.py`
or `tests/test_known_bugs.py` directly.  The reclassification +
screenshot replacement is a SEPARATE PR after maintainer review,
following the issue-report-maintainer skill's 6-step sync.

This separation is load-bearing: it lets the maintainer review
the evidence independently from the consequence (downgrade /
removal).  See PR pair `investigate/issue9-neo4j-institutioncodes`
+ `reclassify/issue9-latent-source-typo` for the canonical
example.

## Required validation commands

For any probe / test addition that touches Access COM:

```bash
# Fast sanity (no Access).
python -m pytest tests/ -W ignore

# The specific probe / new test.  Reap orphans first; cooldown.
python -c "import os, time; os.environ['CBDB_KILL_ALL_ACCESS']='1'; \
  import sys; sys.path.insert(0,'tests'); \
  from cbdb_driver.access_app import kill_orphan_access; \
  kill_orphan_access(); time.sleep(45); print('reaped')"
python -m pytest tests/<new_or_changed_test>.py -W ignore --include-vba -v -s

# Inventory still consistent after any test-coverage change.
python analysis/inventory_export_coverage.py
```

For artifacts-only investigation PRs:

```bash
python <your_probe_script>.py
python <your_probe_script>.py --com    # if it has opt-in COM mode
python -m pytest tests/ -W ignore       # confirm fast suite unchanged
```

## Common traps (do NOT relearn the hard way)

- **Don't trust `n` returned by `click_via_timer`.**  Re-check via
  `vba.conn.cursor().execute("SELECT COUNT(*) FROM <table>")`.
  The watcher races CmdQuery's DELETE-then-INSERT window.
- **Don't run multiple --include-vba pytest sessions back-to-back
  without `kill_orphan_access` + cooldown.**  ROT staleness +
  Access process leakage will manifest as RPC errors in the next
  session's `_inject_autodetect`.
- **Don't write tests that drive Networks via the standard cross-
  form harness.**  Use the minimal-injection scaffold in
  `tests/test_vba_networks_small_fixture.py` instead.
- **Don't probe AssociationPairs' export chain expecting success
  via Form_Timer dispatch.**  CmdQuery's SetFocus calls fail; you
  need a driver-side patch first (or evidence that Access has
  changed behaviour).
- **Don't put probe artifacts in version control under non-`_*`
  paths.**  `analysis/_*` is gitignored; the heavy `_<probe>.mdb`
  / `_<probe>_out/` directories blow up repo size if committed.
- **Don't change the canonical report from a probe PR.**  Always
  ship the investigation as artifacts; the reclassification is a
  separate PR after review.
- **Don't add new VBA injection lines without testing them in
  isolation first.**  Every line is a potential VBA parser
  failure that crashes Access mid-test (modal dialog requires
  human dismissal).  Variable names must NOT start with
  underscore.
- **Don't assume `pywinauto.click_input` works.**  It silently
  drops on disabled controls + locked-screen sessions; we
  migrated to Form_Timer dispatch universally.  New tests should
  use `click_via_timer` exclusively.
- **Don't share Access across tests.**  Function-scoped fixtures
  cost ~12 s each but isolation is worth it; trying to share
  Access across tests led to days of debugging (the
  `test_infra_smoke.py @skip` is a relic).
