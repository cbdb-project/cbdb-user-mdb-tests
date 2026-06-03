# Design: Standardized Fixed-Skeleton Report

**Status**: Implemented (Steps 1-5 complete).

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

---

## Phase 6: Fix VBA fixture orphan-process cascade

**Status**: Implemented (Fixes A, B, C complete as of 2026-06-03).

### Problem

Every pytest run with `--include-vba` produces a cascade of ~120 `ERROR`
results in the VBA test files.  The cascade starts at the second parametrized
case of the first alphabetical VBA test file
(`test_vba_associationpairs_probe.py`) and propagates through every subsequent
VBA fixture setup:

```
tests\test_vba_associationpairs_probe.py .E
tests\test_vba_bilingual_toggle.py       .EEEEEEEEEEEEE
tests\test_vba_bilingual_ui.py           EEEEEEEEE
...
```

The errors are all `PermissionError: cannot remove stale working copy
..._copy.mdb: [WinError 32] The process cannot access the file because it is
being used by another process.`

### Root cause analysis

There are two independent root causes that compound each other.

**Root cause A — `kill_access_pid()` returns before the process exits.**

`VbaSession.close()` calls `kill_access_pid(self._pid)` which runs
`taskkill /F /PID <n>`.  `taskkill /F` delivers the termination signal
*synchronously* but the Access process may remain alive for 100–500 ms
afterwards as Windows drains its I/O completion ports and releases file
handles.  The next test's `VbaSession.open()` calls `self.work.unlink()`
immediately, which races against the dying process.  When it loses the race
it raises `PermissionError`, errors the test in setup, and leaves the
working-copy file on disk — permanently locking out every subsequent test
that uses the same path.

Location: `tests/cbdb_driver/access_app.py::kill_access_pid` (lines 47-57).

**Root cause B — `test_vba_inline.py` uses `Dispatch` instead of
`DispatchEx`.**

`test_vba_inline.py` fixture opens Access with
`win32com.client.Dispatch("Access.Application")`.  `Dispatch` can reuse
ROT (Running Object Table) entries from a previously killed Access process,
which causes a Windows fatal exception 0x800706be when the stale ROT entry
points to a dead server.  This crashes the test with a non-fatal thread
exception that pytest survives, but the working copy
`analysis/_inline_test_copy.mdb` is left on disk with no PID to kill.

`VbaSession` already uses `DispatchEx` for exactly this reason (landmine #9
in AGENTS.md).  `AccessApp.open()` in `access_app.py` also uses `Dispatch`
and has the same latent bug.

Location: `tests/test_vba_inline.py` line 59;
`tests/cbdb_driver/access_app.py` line 157.

### Fix plan (three changes)

#### Fix A — wait for process exit in `kill_access_pid()`

After `taskkill /F /PID`, poll until the PID is gone or a 3-second deadline
passes.  Use `psutil.pid_exists(pid)` if psutil is available; fall back to
attempting `OpenProcess` via `win32api` and checking if it returns
`ERROR_INVALID_PARAMETER` (process gone).

```python
def kill_access_pid(pid: int, wait_s: float = 3.0) -> bool:
    if not pid:
        return False
    rc = subprocess.run(["taskkill", "/F", "/PID", str(int(pid))],
                        capture_output=True, check=False)
    deadline = time.monotonic() + wait_s
    while time.monotonic() < deadline:
        try:
            import psutil
            if not psutil.pid_exists(pid):
                return True
        except ImportError:
            # fallback: try to open the process; if it's gone, OSError
            try:
                import win32api, win32con
                h = win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
                win32api.WaitForSingleObject(h, 100)
                win32api.CloseHandle(h)
            except Exception:
                return True   # process gone
        time.sleep(0.05)
    return rc.returncode == 0
```

#### Fix B — kill-and-retry in `VbaSession.open()` and `make_working_copy()`

When `self.work.unlink()` raises `PermissionError`, instead of immediately
surfacing the error, find the PID holding the file (via `psutil`) and kill it
with `kill_access_pid()`, which now waits.  Retry `unlink()` once.  Only if
the retry also fails should the exception be re-raised.

```python
if self.work.exists():
    try:
        self.work.unlink()
    except PermissionError:
        _kill_file_holder(self.work)   # find PID via psutil, kill+wait
        self.work.unlink()             # retry — should succeed now
```

`_kill_file_holder(path)` implementation:

```python
def _kill_file_holder(path: Path) -> None:
    """Kill the MSACCESS.EXE process holding `path` open, if identifiable."""
    try:
        import psutil
        for proc in psutil.process_iter(["pid", "name", "open_files"]):
            if proc.info["name"] and "MSACCESS" in proc.info["name"].upper():
                try:
                    files = proc.open_files()
                except (psutil.AccessDenied, psutil.NoSuchProcess):
                    continue
                if any(Path(f.path).resolve() == path.resolve()
                       for f in files):
                    kill_access_pid(proc.pid)   # includes wait
                    return
    except ImportError:
        pass
    # psutil unavailable or process not found — fall back to kill-all if env var set
    kill_orphan_access()
```

#### Fix C — replace `Dispatch` with `DispatchEx`

In `test_vba_inline.py` line 59:
```python
# Before:
app = win32com.client.Dispatch("Access.Application")
# After:
app = win32com.client.DispatchEx("Access.Application")
```

In `access_app.py` line 157 (same fix for `AccessApp.open()`):
```python
# Before:
self._app = win32com.client.Dispatch("Access.Application")
# After:
self._app = win32com.client.DispatchEx("Access.Application")
```

### Expected outcome

After all three fixes:
- Each VBA test's teardown waits for Access to fully release file handles
  before returning → next test's `unlink()` always succeeds
- `test_vba_inline.py` no longer crashes with 0x800706be even when ROT has
  stale entries → no orphan working copy left behind
- The ERROR cascade of ~120 tests collapses to zero infrastructure errors;
  only genuine PASS / FAIL / SKIP results remain

### Prerequisites

`psutil` must be installed in the test environment:
```
pip install psutil
```

Add to `requirements.txt` / `environment.yml` alongside existing deps.

### Regression risk

- Fix A adds up to 3 s of polling overhead per test teardown in the worst
  case.  In practice Access exits in < 300 ms after `taskkill /F`; the
  50 ms polling interval means typical overhead is one extra poll (50 ms).
- Fix B only activates on `PermissionError` (the error path) — zero cost on
  clean runs.
- Fix C is strictly safer: `DispatchEx` always creates a fresh process;
  `Dispatch` only reuses ROT if a stale entry exists.

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
