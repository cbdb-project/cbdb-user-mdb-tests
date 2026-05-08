# cbdb-user-mdb-tests

Automated regression tests for the **CBDB User MDB** (the Microsoft Access
front-end shipped by the [China Biographical
Database](https://projects.iq.harvard.edu/cbdb), file name
`CBDB_BJ_User.mdb` in the current edition). The test suite runs the
real VBA inside Access via COM automation, captures the resulting
scratch tables and exports, and compares them against frozen goldens —
catching regressions in queries, column backfills, and exports after
each `.mdb` data refresh.

> 中文版說明請見 [§ 中文簡介](#中文簡介-traditional-chinese)。

---

## What this is

CBDB ships a Windows-only Access database to non-technical historians.
The `.mdb` is updated periodically with new biographical data; each
update can silently break a query, drop a column from an export, or
shift a code lookup. There are 10 `LookAt*` query forms, hundreds of
saved queries, and ~50 000 lines of VBA — far too much to revalidate by
hand on every release.

This project instruments Access from Python, fires the real
`CmdQuery_Click` / `CmdRun_Click` / `CmdGIS_Click` event handlers (so
test results match what an actual user sees), and asserts:

- Result row counts vs. independent SQL replay
- Column-set equality (a dropped backfill is the #1 user-reported bug)
- Byte-level equality of GIS / Neo4j / KML exports
- Foreign-key integrity across joined lookup tables

Confirmed bugs found by the suite live in
[`reports/CBDB_Issues_Report_EN.md`](./reports/CBDB_Issues_Report_EN.md)
([中文](./reports/CBDB_Issues_Report_ZH-Hant.md) — both Markdown
files are auto-generated from
[`reports/generate_report.py`](./reports/generate_report.py),
which can also emit a `.docx` sibling per language; the .docx
files are intentionally NOT committed (their zip-timestamp
metadata is non-deterministic and would dirty every regen).
Run `python reports/generate_report.py` locally to produce
fresh .docx if you need a Word file.

---

## Project structure

```
cbdb-user-mdb-tests/
├── data/                       # gitignored — drop CBDB MDB + HelpFiles here
│   ├── CBDB_BJ_User.mdb
│   ├── CBDB_<YYYYMMDD>_DATA.mdb
│   └── HelpFiles/
├── tests/
│   ├── cbdb_driver/            # Python ↔ Access COM driver
│   ├── cbdb_replay/            # independent SQL replay of each LookAt form
│   ├── golden/                 # frozen reference outputs
│   ├── test_vba_matrix.py
│   ├── test_vba_matrix_all_forms.py
│   ├── test_vba_export.py      # real CmdGIS export, byte-level diff
│   ├── test_vba_integrity.py   # 12-dimension data-integrity battery
│   ├── test_vba_differential.py
│   ├── test_known_bugs.py
│   ├── test_schema.py
│   ├── test_saved_views.py
│   ├── test_lookatentry.py
│   ├── test_other_lookat_forms.py
│   ├── test_exports.py
│   ├── DESIGN.md
│   ├── VALIDATION_DIMENSIONS.md
│   └── MANUAL_SMOKE.md
├── analysis/                   # one-off discovery + dump scripts
│   ├── dump_metadata.py        # writes analysis/dump/{tables,queries,…}.json
│   ├── dump_vba.py             # writes analysis/dump/vba/Form_*.vb
│   ├── discover_test_inputs.py # picks fresh fixtures from current data
│   └── dump/                   # frozen metadata snapshots (committed)
├── AGENTS.md                   # contributor / agent guide; landmines
├── reports/                    # auto-generated bilingual issue report
│   ├── generate_report.py      #   ⭐ SOURCE OF TRUTH for all 21 issues
│   └── CBDB_Issues_Report_*.{md,docx}   # EN / ZH-Hant × md / docx
└── FINAL_STATE.md              # snapshot of test-suite coverage
```

---

## Setup

### Requirements

| Tool | Version | Notes |
|------|---------|-------|
| Windows | 10 / 11 | Access COM is Windows-only (the `--include-vba` suite); the fast suite collects on Linux too once `pip install -r requirements.txt` succeeds |
| Microsoft Access | 2016+ (ACE engine) | Office 365 click-to-run works; only needed for `--include-vba` |
| Python | 3.11+ | tested on 3.12 |
| Python deps | pinned in [`requirements.txt`](./requirements.txt) | `pip install -r requirements.txt` (covers driver: `pyodbc`, `pywin32`, `pywinauto`, `pandas`; tests: `pytest`, `pytest-json-report`; report generator: `python-docx`, `Pillow`, `opencc`) |
| Microsoft Access Driver (`*.mdb, *.accdb`) | bundled with Office | required for ODBC access |

### Drop the MDB + HelpFiles into `data/`

The data files are not redistributed in this repo (proprietary CBDB
release). Obtain them from
[projects.iq.harvard.edu/cbdb/download-cbdb](https://projects.iq.harvard.edu/cbdb/download-cbdb)
and place them as:

```
data/CBDB_BJ_User.mdb            # the User-facing front-end
data/CBDB_<YYYYMMDD>_DATA.mdb    # the linked data backend
data/HelpFiles/                  # CBDB Users Guide + per-form HelpFile_*.pdf
```

### (Optional) Pull the cbdb-online-main-server SQLite snapshot

For the cross-check test (`tests/test_index_year_xcheck.py`,
roadmap item 12) you also need the upstream weekly SQLite dump:

```bash
python analysis/download_hf_sqlite.py
# downloads ~130 MB into data/cbdb_online_sqlite/, extracts ~550 MB
```

Re-run weekly when CBDB pushes a new dump.

### One-time setup of the User MDB

`CBDB_BJ_User.mdb` ships with a broken VBA reference to legacy DAO
(`dao360.dll`). On a new machine you must remove it once before tests
will run — see [`AGENTS.md` § DAO 3.6 reference](./AGENTS.md) for the
3-click fix in the VBE, or run:

```bash
python analysis/check_vba_refs.py
```

---

## Running the tests

```bash
# Fast SQL-replay + static-data tests (no Access process; ~30s).
# This is the CI-safe entry point — every test_vba_*.py file (which
# needs Windows + Access COM) is automatically skipped, so this command
# stays green on headless / Linux / fresh machines.
python -m pytest tests/ -W ignore

# Add --include-vba to also run the Access-COM suite (Windows + Office
# only; ~10 minutes for the full battery).  Boots Access for each
# function-scoped fixture, so it's slow but covers the real VBA path.
python -m pytest tests/ -W ignore --include-vba

# A single VBA test file (when you don't want all of them):
python -m pytest tests/test_vba_export.py -v -W ignore --include-vba
```

### After each CBDB `.mdb` refresh

```bash
python analysis/dump_metadata.py        # refreshes analysis/dump/*.json
python analysis/dump_vba.py             # refreshes analysis/dump/vba/
python analysis/discover_test_inputs.py # picks fresh fixtures
python analysis/run_all_audits.py       # human-readable audit sweep
                                         # — 21 static audits, ~6-7s
python -m pytest tests/ -W ignore       # fast suite (no Access)
python -m pytest tests/ -W ignore --include-vba   # full suite

# To regenerate the bilingual bug-report (FOUR files) AFTER a new dump
# — captures fresh demo personid hints, known-bug statuses, and per-
# issue 'marker no longer reproduces — please verify' banners wherever
# a regression marker has flipped (see policy below: a flipped marker
# is a signal to investigate, NOT an automatic fix confirmation):
python -m pytest tests/test_known_bugs.py -W ignore --no-discover-inputs \
    --json-report --json-report-file=reports/known_bugs_status.json
python reports/probe_demo_persons.py    # refresh concrete demo personids
python reports/collect_index_year_diffs.py  # refresh drift appendix
python reports/capture_screenshots.py   # refresh real-Access screenshots
python reports/generate_report.py       # rebuild all four outputs:
                                         #   reports/CBDB_Issues_Report_EN.md          (committed)
                                         #   reports/CBDB_Issues_Report_ZH-Hant.md     (committed)
                                         #   reports/CBDB_Issues_Report_EN.docx        (gitignored)
                                         #   reports/CBDB_Issues_Report_ZH-Hant.docx   (gitignored)
```

Only the `.md` siblings are committed (they're deterministic and
GitHub-renderable); `.docx` files are produced by the same run
but are gitignored and regenerated on demand.  Reviewers can browse
the markdown directly on GitHub (anchor-linked TOC + image
references resolve to `reports/screenshots/*.png`).  All four are
emitted by a single `generate_report.py` run so they never drift
out of sync.

The Word doc generator NEVER deletes issues automatically — when a
regression marker fails (= the marker no longer reproduces), the
corresponding issue gets a "⚠ please verify in person" banner
prepended, but its description / steps / fix recommendation stay
intact so nothing is lost if the test gave a false positive.

**Policy — marker failure ≠ upstream fix.** A flipped marker is a
signal to investigate; it is NOT an automatic confirmation that the
bug is fixed.  The marker can also flip because the input fixture
or Access driver changed under the test, or because the original
classification was wrong.  Only delete an issue from
`reports/generate_report.py`'s `ISSUES` dict after inspecting a
new VBA / queries dump that shows the source-level fix, or after
the maintainer explicitly confirms it.  Until then, prefer
re-classifying the issue (Dormant / Latent / Not currently
reproducible) over removing it.

`run_all_audits.py` prints a FLAGGED / CLEAN summary.  As of the
shipped dump (2026-05-02): 6 of 21 audits flag known bugs (#1-#19);
the remaining 15 act as long-term regression guards.  Use
`--ci` to exit non-zero ONLY on findings ABOVE the recorded baseline
(`analysis/audit_baseline.json`); use `--update-baseline` after
intentionally accepting a new finding or a fix.  When a CBDB
release adds new code, anything that newly flags is your investigation
queue; anything that newly drops back to clean means a marker no
longer reproduces — investigate (per the policy above) before
flipping the matching marker in `tests/test_known_bugs.py`.

---

## Plan & status

This section is the source of truth for what's done, what's in flight,
and what's intentionally deferred. **Update it whenever a milestone
moves** — `AGENTS.md` enforces this as a contributor rule.

### Coverage as of 2026-05-02

| Form | Real-VBA matrix | Real export | StoreID / RecallID | Import-list | Save-list | Notes |
|------|-----------------|-------------|---------------------|-------------|-----------|-------|
| LookAtEntry            | ✅ `test_vba_matrix.py` | ✅ CmdGIS (byte-diff) | ✅ Store; round-trip → Kinship ✅ | ✅ EntryCodes + Places | ✅ EntryCodes (3-col) | An early "Bug #3" suspected the multi-table backfill UPDATE here failed silently on >30k-row results; the 2026-05-02 re-verification found 0 NULL backfills on the original fixture (entry 36, 92,514 rows) and the issue was removed from the maintainer-facing report on 2026-05-03 as a false positive |
| LookAtStatus           | ✅ 3 fixtures | ✅ CmdGIS (struct) + ✅ CmdNeo4j (struct, 6-file set, on the matrix `status_<top_code>_unfiltered` fixture; `tests/test_vba_cmdneo4j_cross_form.py`) | ✅ Store | ✅ StatusCodes + Places | ✅ StatusCodes (3-col) | 17 023 + 4 931 rows; CmdNeo4j needs **no** driver patch — the prior skip was a false-positive copy-paste from the Pajek/Gephi Status skip (PR #127 Q3 refuted the 'same root family' claim).  CmdPajek + CmdGephi on Status remain skipped on the cross-form Pajek/Gephi test, blocked by a separate cleanup-rebind 'Object required' issue (`Form_LookAtStatus.vb:1457+1460`) that's the next driver/meta workaround line. |
| LookAtTexts            | ✅ biblcat 1 | ✅ CmdGIS (struct) | ✅ Store | ✅ TextCategories + Places | ✅ TextCategories (3-col) | 15 774 rows |
| LookAtAssociations     | ✅ 3 fixtures | ✅ CmdGIS (struct) + ✅ CmdNeo4j (struct, 8-file set, on the matrix `assoc_<top_code>_unfiltered` fixture; `tests/test_vba_cmdneo4j_cross_form.py`) | ✅ Store | ✅ Associations + Places | ✅ Associations (3-col) | 11 867 rows; CmdNeo4j unblocked by PR #116 (`c_index_addr_type_code → c_addr_type` driver rewrite for canonical Issue #23) + PR #117 (5-prefix debug-MsgBox suppress) — Issue #23 stays P1 canonical, the workarounds make the cell *testable* on the existing source, not *fixed* upstream |
| LookAtOffice           | ✅ 2 fixtures | ✅ CmdGIS (struct) | ✅ Store | ✅ Offices + PlaceOffice + PlacePeople | ✅ Offices (3-col) | 37 429 + 35 748 rows |
| LookAtPlace            | ✅ 2 fixtures | ✅ CmdGIS (struct, GISFrame→CodeFrame workaround) + ✅ CmdNeo4j (struct, 6-file set, on the matrix `place_addr_<top_addr_id>` fixture; `tests/test_vba_cmdneo4j_cross_form.py`) | ✅ Store | ✅ Places | — (no save button) | 5 962 + 3 528 rows; CmdNeo4j unblocked by PR #123 (`tRstPeople` SELECT projection extension driver rewrite for canonical Issue #24) — Issue #24 stays P1 canonical, the workaround makes the cell *testable* on the existing source, not *fixed* upstream |
| LookAtKinship          | ✅ 1 fixture | ✅ CmdGIS (struct, requery shim) + ⚠ CmdUCINet (.vna 3-section, family first coverage cell, `tests/test_vba_cmducinet_kinship.py`; **KNOWN FIXTURE-FRAGILE per Issue #22 sibling-form confirmation** — passes only because person 3211's kin network has no Han-character c_name values; a fixture that reaches a Han-name person would reproduce Issue #22's VBA error 5 here too) | ✅ Store + ✅ Recall | ✅ CmdImport | — (no save button) | 949 rows (Zhao Tingmei) |
| LookAtAssociationPairs | ✅ tiny fixture (`test_vba_matrix_hard_forms.py`, 4×5) | ✅ CmdPajek (struct) + ✅ CmdGephi (struct) + ✅ CmdNeo4j (struct, 6-file set) via 1×3 known-edged person pair (`test_vba_pajek_gephi_cross_form.py` + `test_vba_cmdneo4j_cross_form.py`) | ✅ Recall | ✅ CmdImportList | — (no save button) | `Link1stOrder` ASSOC_DATA self-join too slow at full scale; tiny fixture covers CmdQuery happy-path; CmdPajek/CmdGephi unblocked by AssociationPairs SetFocus driver patch; CmdNeo4j unblocked by AssociationPairs CmdNeo4j_Click debug-MsgBox suppress driver patch |
| LookAtNetworks         | ⏭ skipped in default matrix; ✅ small-fixture CmdRun smoke (`tests/test_vba_networks_small_fixture.py`, Cao Zhi pid=30270) | — | ⏭ Recall (CmdRun timeout) | ⏭ ImportPeople / ImportPlaces (CmdRun timeout) | — (no save button) | recursive expansion (Zhu Xi 2 471 assocs) is the matrix blocker.  PR AR-AX bisected a SEPARATE Form_Open hang under full injection: any sibling Form_LookAt* VBA module modification dirties the project, and Networks's Form_Open then deadlocks on the project-wide auto-compile interaction with its `Forms!LookAtNetworks!<sub>.Form.Recordset` self-reference.  Workaround verified: minimal-injection mode (skip the 9 sibling autodetect entries via the PR AT `skip_inject_autodetect_forms` kwarg).  PR AY landed the small-fixture smoke; matrix unskip needs either a per-form minimal-injection refactor or a deeper Access-side fix |
| LookAtGroupData        | ✅ tiny fixture (`test_vba_matrix_hard_forms.py`, c_person_id=1, 2 entries / 2 statuses) | — | ✅ Recall | ✅ CmdImport | — (no save button) | recursion; tiny fixture covers backfill on `ZZ_SCRATCH_IMPORT_PEOPLE` |

**Latest matrix run**: `12 passed, 3 skipped in 110.22s`.
**Latest Store/Recall run** (`tests/test_vba_storeid_recallid.py`): `11 passed, 1 skipped in 142.82s`.
**Latest Import-list run** (`tests/test_vba_import_lists.py`): `15 passed, 2 skipped in 142.04s`.
**Latest Save-list run** (`tests/test_vba_save_lists.py`): `5 passed in 43.31s`.
**Latest Bilingual run** (`tests/test_vba_bilingual_ui.py`): `9 passed in 145.73s`.
**Latest CmdGIS-other-forms run** (`tests/test_vba_cmdgis_other_forms.py`): `7 passed` (Status / Texts / Associations / Office / Place / Kinship via the parametrized `test_cmd_gis_produces_file` + GroupData via `test_cmd_gis_groupdata_clean_branches`, which covers Status / Office / Addr branches but explicitly excludes the Entry branch — Issue #6 P1, separately bug-pinned — and the Text branch — benign 0-files on `person_1`).
**Combined regression** (all 8 test files, 71 tests): `59 passed, 8 skipped, 4 xfailed in 837.76s`.

### Confirmed bugs

**21 documented issues** (highest assigned ID: #22 — ID #3 was retired 2026-05-03 as an early false positive), all detail in
[`reports/CBDB_Issues_Report_EN.md`](./reports/CBDB_Issues_Report_EN.md)
([中文版](./reports/CBDB_Issues_Report_ZH-Hant.md)).  Both Markdown
files are auto-generated from the `ISSUES` dict in
[`reports/generate_report.py`](./reports/generate_report.py) — that
script is the **single source of truth**, please don't paste bug
content elsewhere.  The same script can emit `.docx` siblings on
demand (gitignored; run `python reports/generate_report.py` to
produce fresh ones).

| Tier | Count | Meaning |
|------|------:|---------|
| P0 — Silent data corruption | 3 | wrong data shown, no user warning |
| P1 — Visible runtime crash  | 4 | error popup on a normal user click |
| P2 — Silent display         | 1 | user-visible bound control renders blank where data exists |
| P3 — Missing UI             | 5 | event handler exists in code but no button on the form |
| P4 — Setup                  | 1 | one-time install fix (dao360.dll on Office 2016+) |
| P5 — Dormant / latent / not currently reproducible | 7 | defect real but doesn't fire on the current dump or has no UI trigger today; **none have been verified as upstream-fixed** (#1 dormant; #4, #5, #9, #11, #12, #14 latent) |

> **Re-verifications.** Issues that earlier snapshots graded 🔴/🟡
> were re-checked end-to-end in three passes:
> - 2026-05-02 — #1, #4, #5, #14 demoted to P5 (dormant/latent)
> - 2026-05-03 — #11, #12 demoted to P5 (COM probe showed the blamed
>   controls are `Visible=False` hidden internal join-key holders,
>   not user-facing)
> - 2026-05-03 — **#3 removed entirely** as an early false positive:
>   re-verification found 0 NULL backfills on the original fixture
>   (entry 36, 92,514 rows) AND there was no upstream source-level
>   fix to point at, so per the marker-failure-≠-fix policy it was
>   treated as a testing-infrastructure / fixture / driver artifact
>   rather than a CBDB-maintainer bug
> - 2026-05-04 — **#9 demoted to P5 latent** (source-level typo at
>   `Form_LookAtEntry.vb:1425` is real, but the buggy `With` block
>   sits inside `If tRecDeleted > 0 Then` at line 1389, which is
>   gated by `ZZ_SCRATCH_ENTRY.c_inst_code > 0`; the current dump
>   has 0 of 263,454 ENTRY_DATA rows with that, so the branch is
>   unreachable for any LookAtEntry fixture and produces no popup
>   — verified end-to-end via SQL pre-image + real Access COM in
>   [`analysis/issue9_neo4j_institutioncodes_reverification.md`](./analysis/issue9_neo4j_institutioncodes_reverification.md))
>
> The classifier lives in [`analysis/reverify_all_issues.py`](./analysis/reverify_all_issues.py)
> and is cross-checked against `generate_report.py` so the report
> can't drift back.

**Regression coverage**: `tests/test_known_bugs.py` has a CI marker
for every issue — when a marker stops reproducing, the
corresponding test fails with a "Bug #N marker no longer reproduces
(investigate upstream fix vs. fixture/driver change vs.
misclassification before flipping)" message.  See the marker-failure
policy near the top of this README — a flipped marker is a signal to
investigate, not an automatic fix confirmation.

### Roadmap

| # | Status | Item |
|---|--------|------|
| 1 | ✅ done | Build Access COM driver (`tests/cbdb_driver/vba_session.py`) — `Form_Timer` trigger pattern, `Me.Tag` chain pattern, `MsgBox Err.Description` rewrite for diagnostics |
| 2 | ✅ done | Cross-form matrix covering 7/10 LookAt forms (12 fixtures) |
| 3 | ✅ done | Real export byte-level test (`LookAtEntry.CmdGIS`) |
| 4 | ✅ done | 12-dimension data-integrity battery (`test_vba_integrity.py`) |
| 5 | ✅ done | Independent SQL replay layer (`tests/cbdb_replay/`) for differential checking |
| 6 | ✅ done | Auto-discovery of fresh fixtures (`analysis/discover_test_inputs.py`) |
| 7 | 🟡 partial | `tests/test_vba_matrix_hard_forms.py` (added 2026-05-02) handles 2 of the 3 forms with hand-picked tiny fixtures: LookAtGroupData (c_person_id=1, 2 entries / 2 statuses — backfill check on ZZ_SCRATCH_IMPORT_PEOPLE) and LookAtAssociationPairs (4 × 5 — CmdQuery completes, ZZ_SOCIAL_NETWORK readable). **Still open:** LookAtNetworks CmdRun times out on high-degree anchors (Zhu Xi 2 471 assocs). PR AA's Form_Open probe (`analysis/probe_lookatnetworks_form_open.py`) showed `Form_Open` itself opens cleanly in ~2 s under all 4 tested configurations; the earlier wording that blamed `Form_Open` was a misdiagnosis (root cause was a separate `LinkListInit` popup on the probe). The actual blocker is `CmdRun` network expansion, not `Form_Open`; smaller-fixture experiment is queued for the morning of 2026-05-04 |
| 8 | 🟡 partial | First slice landed: `tests/test_vba_cmdgis_other_forms.py` extends real-`CmdGIS` coverage from LookAtEntry (`test_vba_export.py`, byte-diff) to Status / Texts / Associations / Office / Place / Kinship (6 added). Uses **structural** assertions rather than byte-diff. **Subsequent slices** (added 2026-05-02): `test_vba_cmdguess_cross_form.py` (Kinship + Office), `test_vba_pajek_gephi_cross_form.py` (7 passing — 5 original + AssociationPairs CmdPajek + CmdGephi via 1×3 known-edged pair, unblocked by the AssociationPairs SetFocus driver patch; 2 Status skipped), `test_vba_cmdgispeople_office.py`, `test_vba_cmdneo4j_cross_form.py` (8 passing — Entry / Texts / Office / Kinship / AssociationPairs / Associations / Place / Status; 0 skipped; **LookAtStatus promoted from skip 2026-05-08** with the matrix `status_<top_code>_unfiltered` fixture; the prior skip was a false-positive copy-paste from the Pajek/Gephi cross-form Status skip — PR #127's driver/meta probe directly refuted the 'same root family' claim by showing CmdNeo4j on Status runs cleanly (6 files, 0 :ERR, 0 watchdog dialogs) while CmdPajek + CmdGephi fail with VBA 424 'Object required' from the cleanup-rebind at `Form_LookAtStatus.vb:1457+1460`.  **No driver patch needed** for Status × CmdNeo4j (unlike Place / Associations); per-shape pinning asserts exactly 6 files (People, PeopleStatus, Places, PeoplePlaces, PersonPlaceCodes, StatusCodes); classifier gained 1 new single-col shape (StatusCode → StatusCodes code-table) + 1 new 2-col disambiguator ((NameID, StatusCode) → PeopleStatus).  Status × CmdPajek + CmdGephi remain skipped on the cross-form Pajek/Gephi test, blocked by the same cleanup-rebind 'Object required' issue — that's a separate driver/meta workaround verification line.  **LookAtPlace promoted from skip 2026-05-08** with the matrix `place_addr_<top_addr_id>` fixture, unblocked by PR #123 (`_rewrite_place_cmdneo4j_trstpeople_projection` extends the `tRstPeople` SELECT projection inside `Form_LookAtPlace.CmdNeo4j_Click` to add `DYNASTIES.c_dynasty / c_dynasty_chn / BIOG_MAIN.c_female` so the downstream `!c_dynasty` / `!c_female` field reads no longer fire JET 3265 — canonical Issue #24); per-shape pinning asserts exactly 6 files (People, PeoplePlaces, Places, PersonPlaceRelations, PersonPlaceRelCodes, IndexAddrTypeCodes) plus chain-completion via ZZ_TEST_DEBUG ENTER+DONE markers and no `:ERR` rows; classifier gained 3 new shapes (PersonID, PersonPlaceRelCode, IndexAddrTypeCode).  Issue #24 stays P1 canonical — workaround makes the cell *testable* on existing source, not *fixed* upstream.  Unlike Associations (#116 → #117 → #118), Place needed only the single SELECT-projection rewrite — `Form_LookAtPlace.CmdNeo4j_Click` has no concat-form debug MsgBox layer.  **LookAtAssociations promoted from skip 2026-05-07** with the matrix `assoc_<top_code>_unfiltered` fixture, unblocked by two driver workarounds for canonical Issue #23 (P1): PR #116 rewrites `c_index_addr_type_code → c_addr_type` (target column name typo) and PR #117 suppresses 5 concat-form debug MsgBoxes; per-shape pinning asserts exactly 8 files (People, Places, PeoplePlaces, PeopleAssociations, AssociationCodes, KinshipCodes, OccasionCodes, TopicCodes) plus chain-completion via ZZ_TEST_DEBUG ENTER+DONE markers and no `:ERR` rows; classifier gained 2 new shapes (OccasionCode, TopicCode) + 1 new 2-col disambiguator ((AssociationCode, AssociationTypeID) for the 4-col Associations variant).  Issue #23 stays P1 canonical — workarounds make the cell *testable* on existing source, not *fixed* upstream.  **LookAtAssociationPairs promoted from skip 2026-05-07** with the 1×3 known-edged person pair (same fixture as the Pajek/Gephi cross-form test), unblocked by the AssociationPairs CmdNeo4j_Click debug-MsgBox suppress driver patch (PR #109); per-shape pinning asserts exactly 6 files (People, Places, PeoplePlaces, PeopleAssociations, AssociationCodes, KinshipCodes) plus chain-completion via ZZ_TEST_DEBUG ENTER+DONE markers and no `:ERR` rows; classifier gained 3 new shapes (Person1_ID, AssociationCode, KinshipCode).  **LookAtEntry promoted from skip 2026-05-04** with the c_entry_code=101 matrix fixture from the Issue #9 reverification probe, plus a structural assertion that the file set is exactly {People, PeopleEntry, Places, PeoplePlaces, PersonPlaceCodes, EntryCodes, AssocCodes} AND no InstitutionCodes file — pins the Issue #9 LATENT-gate.  Surfaced **3 new bugs** in the CmdNeo4j family: Bugs #7 / #8 / #9). Driver gained: `patch_filedialog` v8 with directory mode (counter-suffixed unique file per call) for multi-dialog subs like CmdNeo4j. **Still open:** CmdKML across all forms; CmdUCINet first coverage cell landed on LookAtKinship (`tests/test_vba_cmducinet_kinship.py`, .vna 3-section family shape established by probe `investigate/cmducinet-family-shape`) **but now KNOWN FIXTURE-FRAGILE per Issue #22's runtime-confirmed Kinship sibling-form** — passes only because matrix-supplied person 3211's kin network has no Han-character c_name values; a fixture that reaches a Han-name person would reproduce Issue #22's VBA error 5 here too (sibling probe `investigate/kinship-cmducinet-sibling-risk` demonstrated this directly); Associations CmdUCINet directly canonicalized as Issue #22 P1 (mid-write VBA error 5 reproduced + static + runtime markers), Place CmdUCINet still gap (uses ADO Stream — needs separate per-form manifest), other forms unchanged |
| 9 | ✅ done | `tests/conftest.py::pytest_configure` re-runs `analysis/discover_test_inputs.py` automatically at session start when `analysis/dump/test_inputs.json` is missing or older than `data/CBDB_BJ_User.mdb`. Opt-out via `pytest --no-discover-inputs` |
| 10 | 🟡 partial | First slice landed: `tests/test_vba_pickers_smoke.py` opens each of the 10 `frmPick*` pickers and verifies it loads + has a commit/cancel button (CmdSelect/CmdSelectAll/CmdOK + CmdCancel/CmdExit/CmdClose). 9 pass, 1 skipped (`frmPickTEXT_BIBLCAT` is orphan — 0 callers anywhere in VBA / macros, end users can't reach it). Smoke-only — doesn't drive Treeview/ListBox interaction yet; that's depth-pass work. **Still open for depth pass:** drive a Select+Commit through e.g. `frmPickEntry_multi` and verify the resulting `ZZ_ENTRY_CODE` table contents match the picker's selected items |
| 11 | ✅ done | Bilingual UI test (`tests/test_vba_bilingual_ui.py`) — for each of the 9 forms with the standard `CmdFanti` / `CmdJianti` toggle pair (Networks uses different names + CmdRun expansion timeout makes downstream interactive flows untestable), opens the form, drives one Fanti round-trip to force captions through `changeDisplayLanguage` (so the FormLabels-derived state, not design-time captions, is the baseline), then asserts each toggle changes ≥5 captions and the second toggle restores the baseline exactly. Latest: `9 passed in 145.73s`. **Caveat surfaced:** a few forms (LookAtPlace, LookAtGroupData) ship with design-time captions that don't match `FormLabels` (e.g. `LblFrom = "  From"` with leading spaces vs `FormLabels.c_fanti = "From"`); the design-time text is replaced on the first `changeDisplayLanguage` call. Pre-existing CBDB UX nit, documented in the test |
| 12 | ✅ done | `tests/test_index_year_xcheck.py` + `analysis/download_hf_sqlite.py` cross-check the User MDB's `BIOG_MAIN` against the weekly cbdb-online-main-server SQLite snapshot at <https://huggingface.co/datasets/cbdb/cbdb-sqlite/blob/main/latest.zip>. Compares 4 fields per `c_personid` — derived (`c_index_year`, `c_index_addr_id`) and source (`c_birthyear`, `c_deathyear`) — with 0.5 % / 0.1 % thresholds. **The two sides are independent implementations**: SQLite's `c_index_year` ← PHP [`IndexYearRebuildService.php`](https://github.com/cbdb-project/cbdb-online-main-server), `c_index_addr_id` ← `IndexAddressRebuildService.php` (same repo); User MDB-side: `c_index_addr_id` ← `analysis/dump/vba/Form_frmIndexAddr.vb` in the front-end mdb; `c_index_year` ← **37 saved QueryDefs `BM IY Rule …` in the linked-tables backend** `data/CBDB_<YYYYMMDD>_DATA.mdb` (extracted by `analysis/dump_data_mdb_algorithms.py` → `analysis/dump_data/querydefs_index/*.sql`; PR G mistakenly said the c_index_year rebuild was missing). [`analysis/classify_index_drift.py`](./analysis/classify_index_drift.py) runs the full per-personid taxonomy → [`reports/index_drift_classification.json`](./reports/index_drift_classification.json).  [`analysis/compare_index_year_rules.py`](./analysis/compare_index_year_rules.py) goes one level deeper — pairs the 37 BM IY rules with PHP methods (PHP pinned at commit `a642f7a`, snapshotted at [`analysis/php_source/IndexYearRebuildService.php`](./analysis/php_source/IndexYearRebuildService.php)).  PR N corrected the comparator to read the **runtime** Access source (`GetBirthIndexYearSQL` inside `frmBaseMaintenance`, dumped by PR M) instead of the vestigial 37 `BM IY Rule` QueryDefs PR I had used.  Updated [`analysis/index_year_rule_comparison.md`](./analysis/index_year_rule_comparison.md) / [`.json`](./analysis/index_year_rule_comparison.json) pairs by emitted `c_index_year_type_code`: **22 matched, 8 matched_minor_diff, 0 logic_diff**, 3 access-only (concubine wife variants 31/32/33), 0 php-only.  PR I's `+N`/`-N` sign-flip flag was an artefact of comparing the wrong Access source.  Closest thing to a real rule-level divergence: off-by-1 (Rule 29) / off-by-3 (Rule 30) on deathyear-default offsets — Access `c_deathyear-64`/`c_deathyear-53` vs PHP `-63`/`-56`.  **None confirmed as bugs.**  [`analysis/classify_index_year_drift_by_rule.py`](./analysis/classify_index_year_drift_by_rule.py) takes the rule-level findings and walks the 69 year-only diffs from PR G per row → [`reports/index_year_drift_rule_classification.json`](./reports/index_year_drift_rule_classification.json): 19 `php_did_not_compute`, 7 `access_did_not_compute`, 5 `iteration_order_diff`, 14 `consistent_within_rule`, 1 `php_returned_sentinel`, 5 `candidate_algorithm_divergence`, 18 `unclassified`.  ~74 % into named evidence categories.  [`analysis/triage_index_year_drift_groups.py`](./analysis/triage_index_year_drift_groups.py) does a deeper triage on the leftover buckets → [`reports/index_year_drift_rule_groups.json`](./reports/index_year_drift_rule_groups.json): 5 named groups for `consistent_within_rule` (all flagged as `candidate_same_rule_tie_break_or_aggregation_diff`; a recurring **diff = −20 across Rules 11 / 13 / 15 / 19** is the standout pattern, wanting a runtime-rule / phase-order triage against `GetBirthIndexYearSQL`), 18/18 `unclassified` rows triaged (17 of them `blocked_by_runtime_priority_triage_pending` — PR X renamed this from the older `blocked_by_missing_frmBaseMaintenance_vba`, which became stale once PR M dumped `frmBaseMaintenance`), and 6 named groups for `php_did_not_compute` (biggest: 7 × `access_tcode='05'` jinshi-entry gap, `candidate_php_entry_code_mapping_gap`).  After K2 (and the PR X re-derivation against PR N), the 17 leftover rows are blocked on **runtime priority / iteration-order triage**, not on missing source — that was PR M's job.  **Nothing yet labelled as a confirmed bug.**  [`analysis/classify_index_addr_drift.py`](./analysis/classify_index_addr_drift.py) does the same kind of per-row analysis for the 488 c_index_addr diffs (478 `index_addr_only_diff` + 10 `index_both_diff`), against PHP `IndexAddressRebuildService.php` (pinned `e31fba7` → [`analysis/php_source/IndexAddressRebuildService.php`](./analysis/php_source/IndexAddressRebuildService.php)) → [`reports/index_addr_drift_classification.json`](./reports/index_addr_drift_classification.json).  Headline: **412 `mdb_stale_index_addr`** (SQLite matches the rank+MAX(c_sequence) recompute over its data, User MDB doesn't — strong signal that the User MDB shipped with a stale `c_index_addr_id` that was never re-run after BIOG_ADDR_DATA was updated; PHP re-runs weekly), 47 `mdb_value_php_null`, 10 `same_candidates_diff_winner` (only candidate algorithm-divergence rows, all in addr_type=1), 10 `both_stale_recompute_mismatch`, 6 `both_sides_match_recomputed`, 2 `sqlite_stale_index_addr`, 1 `mdb_null_php_value`, 0 `unclassified`.  Verified separately that the rank table (BIOG_ADDR_CODES) is identical between the two sides for all 22 addr_types.  [`analysis/dump_data_mdb_vba.py`](./analysis/dump_data_mdb_vba.py) (PR M) used `Access.Application.SaveAsText` to dump `frmBaseMaintenance` + the 4 DATA-mdb modules → [`analysis/dump_data/vba/Form_frmBaseMaintenance.vb`](./analysis/dump_data/vba/Form_frmBaseMaintenance.vb).  Reading them rewrote two earlier conclusions (full write-up in [`analysis/index_drift_algorithm_notes.md`](./analysis/index_drift_algorithm_notes.md) § "Maintenance trigger path"): (a) `CmdIndexYear_Click` calls **`GetBirthIndexYearSQL`** — inline VBA matching PHP's `-N` offsets and `ENTRY_CODE_TYPE_REL` joins, NOT the 37 `BM IY Rule` QueryDefs from PR H — so PR I's sign-flip "logic_diff" flag was largely a **false alarm**; (b) `CmdIndexAddress_Click` does **NOT** explicitly `MAX(c_sequence)`-aggregate the way PHP does, so when a person has multiple BIOG_ADDR_DATA rows of the same addr_type Access picks whichever JET surfaces — a **candidate algorithmic divergence** on top of the maintenance-cadence issue.  Operational mitigation for the 412 `mdb_stale_index_addr`: add a release-checklist step that runs `CmdIndexYear` then `CmdIndexAddress` on the DATA mdb before shipping a new User MDB.  PR S then deep-dived the 10 `same_candidates_diff_winner` rows ([`analysis/deep_dive_addr_same_candidates.py`](./analysis/deep_dive_addr_same_candidates.py) → [`reports/index_addr_same_candidates_deep_dive.json`](./reports/index_addr_same_candidates_deep_dive.json)): **10/10 have MAX(c_sequence) ties** — both sides' algorithms pick non-deterministically when multiple BIOG_ADDR_DATA rows share the same max sequence.  Neither side is "wrong"; candidate mitigation is to add an explicit secondary tie-break (e.g. MIN(c_addr_id)) to both implementations. Latest run on 657 245 common personids: **656 682 exact match (99.914 %)**; net diffs **563** = 16 attributable to source drift (birthyear/deathyear) + 547 in named cause buckets (`index_year_only_diff` 59, `index_addr_only_diff` 478, `index_both_diff` 10). PR Y then upgraded the appendix from per-bucket counts to a **cause analysis** — for each bucket, what likely explains it, what algorithm + row evidence supports the explanation, confidence level, and the next concrete investigation step. Outputs: [`analysis/index_drift_cause_analysis.md`](./analysis/index_drift_cause_analysis.md) (prose) + [`reports/index_drift_cause_summary.json`](./reports/index_drift_cause_summary.json) (machine-readable). Headline framing: cause buckets now exist for every diff; some are high-confidence (e.g. 412 `mdb_stale_index_addr` → release-process step; 10 `same_candidates_diff_winner` → secondary tie-break per PR S), some are medium (coverage-gap candidates), some remain blocked on runtime priority/iteration-order triage (17 rows). Nothing yet labelled a confirmed CBDB bug. The 0.5 % / 0.1 % thresholds are coarse guards against a wholesale jump; they say nothing about whether any specific row is a bug |
| 13 | ✅ done | Import-list buttons (`tests/test_vba_import_lists.py`) — covers all 11 unique button names across 8 forms (15 of 17 tests pass; the 2 LookAtNetworks ones are skipped for the same Form_Open hang as items 7/15). Drives each via `Form_Timer`, points at a fixture file (whose delimiter / column count matches the saved `MSysIMEXSpecs`), and asserts: (a) the target `ZZ_SCRATCH_*` table contains exactly the valid IDs, (b) `InputErrorList` contains exactly the invalid IDs. The `gUse*` global side-effect is documented per spec but not asserted — an early inject-based reader caused JET re-entrancy hangs in matrix CmdQuery; the table-shape assertion is the meaningful contract. Driver gained: `patch_filedialog` now also handles the `With dlgX` block's `If .Show = -1 Then` (used in import subs) |
| 14 | ✅ done | Save-list buttons (`tests/test_vba_save_lists.py`) — covers all 5 `CmdSave*_Click` handlers. Pre-populates the source `ZZ_SCRATCH_<X>` table directly (skips CmdQuery), patches `FileDialog(msoFileDialogSaveAs)` via the existing `vba.patch_filedialog`, fires the button, and asserts the resulting tab-separated UTF-8 (BOM-stripped) file contains exactly the seeded IDs — and for the 3-column specs (Entry, Associations, Office, Status, TextCategories) also that the desc / desc_chn fields match an INNER JOIN against the codes table. Latest run: `5 passed in 43.31s`. **Important:** several Form_Open handlers wipe their picker scratch table on form load; the test seeds AFTER `open_form` to avoid this |
| 15 | ✅ done | `CmdStoreID` / `CmdRecallID` round-trip (`tests/test_vba_storeid_recallid.py`) — covers all 7 query-runnable forms for Store, 3 of 4 forms for Recall (Networks CmdRun times out on high-degree anchors — same family as the matrix Networks skip; PR AA: Form_Open is fine), plus an end-to-end Entry → Kinship round-trip. Driver gained: `MsgBox "literal"` neutralizer in `_inject_autodetect`; chain+DONE block moved to *after* `Exit_<name>:` so it survives the `Resume Exit` from the form's Err handler |

---

## Architecture cheat-sheet

The driver pattern that makes this work is non-obvious; the full
landmine list lives in [`AGENTS.md`](./AGENTS.md). The minimum a
contributor needs to know:

1. `CmdQuery_Click` is `Private`, so `Application.Run "Form_X.CmdQuery_Click"`
   does **not** work. We trigger it via `Form_Timer`.
2. `Form_Timer` fires only **once** per `OpenForm`. To chain
   `CmdQuery → CmdGIS` we read `Me.Tag` and call the next sub from
   inside `CmdQuery_Click` itself (an autodetect-injected post-body
   block).
3. JET cache is incoherent across pyodbc and Access — write through
   pyodbc, then `DBEngine.Idle 8 + RefreshDatabaseWindow` before
   triggering VBA.
4. Use `DispatchEx` (not `Dispatch`) and skip `DoCmd.Close /
   CloseCurrentDatabase / Quit` — `taskkill` is the only reliable
   shutdown after a heavy `CmdQuery_Click`.
5. The autodetect injection rewrites `MsgBox Err.Description` into
   `INSERT INTO ZZ_TEST_DEBUG (msg) VALUES ('<form>:ERR ...')` so a
   hidden modal dialog can never silently block the COM thread again.

---

## Contributing

PRs welcome. Before submitting:

- `python -m pytest tests/ -W ignore` passes locally
- if you change anything that moves a roadmap item, **update
  the [Plan & status](#plan--status) table in this README** in the
  same PR (this is enforced by `AGENTS.md`)
- new bugs go in `reports/generate_report.py` (`ISSUES` dict) with
  a regression test in `tests/test_known_bugs.py`; regenerate the
  4 report outputs in the same PR
- new landmines (Access quirks, JET behavior, COM gotchas) go in
  `AGENTS.md`

---

## License

The Python test code in this repository is released under the MIT
License — see [`LICENSE`](./LICENSE) once added.

The CBDB Access database files (`CBDB_BJ_User.mdb`, `CBDB_*_DATA.mdb`,
the HelpFile PDFs) are **not** part of this repository; obtain them
under their own terms from
[projects.iq.harvard.edu/cbdb](https://projects.iq.harvard.edu/cbdb).

---

# 中文簡介 (Traditional Chinese)

本專案是 [中國歷代人物傳記資料庫 CBDB](https://projects.iq.harvard.edu/cbdb)
**用戶端 Access 資料庫**（即 `CBDB_BJ_User.mdb`，現行版本檔名）的
**自動化回歸測試套件**。每次 CBDB 釋出新資料時，由 Python 透過
COM 自動化觸發 Access 內真實的 VBA 事件（`CmdQuery_Click` /
`CmdRun_Click` / `CmdGIS_Click`），擷取 scratch tables 與匯出檔，
與冷凍 golden 檔比對，找出資料更新引入的 query / 欄位 / 匯出 regression。

## 為什麼

CBDB 用戶端是給歷史學家用的 Windows-only Access 介面，每次資料更新可能
默默壞掉一個 query、丟失一個欄位、或讓某個 export 變空。10 個 `LookAt*`
表單、上百個 saved query、約五萬行 VBA — 人工每次都重測不可能。

此套件在 Python 端啟動 Access、觸發真正的 VBA handler，斷言：

- 結果列數 vs. 獨立 SQL replay
- 欄位集合相等（用戶最常回報的「丟欄位」bug）
- GIS / Neo4j / KML 匯出檔位元組級相等
- 跨 lookup 表的外鍵完整性

已找到的 bug 在 [`reports/CBDB_Issues_Report_ZH-Hant.md`](./reports/CBDB_Issues_Report_ZH-Hant.md)（共 21 個 issue，由 `reports/generate_report.py` 自動生成 4 份輸出）。

## 安裝

1. Windows 10/11 + Microsoft Access 2016+ + Python 3.11+
   （快速 suite 在 Linux/headless 也可跑；`--include-vba` 才需要 Office）
2. `pip install -r requirements.txt`
   （涵蓋 driver 套件 `pyodbc` / `pywin32` / `pywinauto` / `pandas`、
   測試 `pytest` / `pytest-json-report`、報告生成 `python-docx` / `Pillow` / `opencc`；
   版本範圍鎖定見 `requirements.txt`）
3. 從 [CBDB 官網](https://projects.iq.harvard.edu/cbdb/download-cbdb)
   取得 `CBDB_BJ_User.mdb`、`CBDB_<日期>_DATA.mdb`、`HelpFiles/`，
   放在 `data/` 之下（此資料夾已 gitignored，不會上傳）。
4. 第一次須在 Access VBE 中清掉斷掉的 `dao360.dll` 引用，或執行：
   ```bash
   python analysis/check_vba_refs.py
   ```

## 執行

```bash
# 快速 SQL replay + 靜態檢查（不啟 Access；~30 秒）。
# 這是 CI 安全的入口 —— 所有 test_vba_*.py（需 Windows + Access COM）
# 都會被自動跳過，在無界面 / Linux / 全新機器上也能 stable 跑綠。
python -m pytest tests/ -W ignore

# 加 --include-vba 才會跑 Access-COM 套件（僅限 Windows + Office；全套
# 約 10 分鐘）。每個 function 級 fixture 都會啟一個 Access 進程，所以慢，
# 但能覆蓋真實的 VBA 路徑。
python -m pytest tests/ -W ignore --include-vba

# 只跑某一個 VBA 測試檔案：
python -m pytest tests/test_vba_export.py -v -W ignore --include-vba
```

## 專案計畫與當前狀態

完整的計畫與覆蓋狀態在英文版的
[§ Plan & status](#plan--status) 表格中（單一真相來源，請勿在中文版重複維護以免不一致）。

簡述：
- ✅ 7/10 LookAt 表單已納入真 VBA matrix（12 fixtures，110 秒跑完）
- ✅ 1 個真實匯出位元組對比（`LookAtEntry.CmdGIS`）
- ✅ 12 維度資料完整性檢查
- ✅ 已確認 21 個 issue（最高編號為 #22；ID #3 已於 2026-05-03 退役為早期 false positive；詳見 [`reports/CBDB_Issues_Report_ZH-Hant.md`](./reports/CBDB_Issues_Report_ZH-Hant.md)；P0 ×3 / P1 ×4 / P2 ×1 / P3 ×5 / P4 ×1 / P5 dormant·latent ×7。原 Bug #3 於 2026-05-03 從報告中移除——重新驗證 0 NULL backfill 且無上游源碼修復證據，視為早期 false positive。Issue #9 於 2026-05-04 從 P0 降為 P5 latent——`Form_LookAtEntry.vb:1425` 的 source-level typo 雖然真實存在，但被第 1389 行的 `If tRecDeleted > 0 Then` gate 擋下；當前 dump 263,454 筆 ENTRY_DATA 中 `c_inst_code > 0` 為 0,所以任何 LookAtEntry fixture 都進不到那個分支、不會 popup（端到端 SQL + Access COM 證據見 `analysis/issue9_neo4j_institutioncodes_reverification.md`）。新增 Issue #20（GIS 匯出嵌入 tab／BOM 造成欄位錯位），證據 + probe 腳本見 `analysis/gis_status_embedded_delim_root_cause.md` / `analysis/probe_status_gis_embedded_delim.py` / `analysis/probe_status_gis_export_bytes.py` / `tests/test_addr_codes_embedded_delim.py`）
- 🟡 剩 1 個表單（LookAtNetworks）因 Form_Open hang 暫跳過；其餘 2 個（LookAtAssociationPairs / LookAtGroupData）已用 tiny fixture 覆蓋（`test_vba_matrix_hard_forms.py`）— roadmap 第 7 項
- 🟡 其他匯出按鈕大多已部分涵蓋：CmdGIS（6 個 form, `test_vba_cmdgis_other_forms.py`）、CmdGUESS（Kinship+Office, `test_vba_cmdguess_cross_form.py`）、CmdPajek/Gephi（`test_vba_pajek_gephi_cross_form.py`，7 過 — 原 5 過 + AssociationPairs 的 CmdPajek/CmdGephi 透過 1×3 known-edged pair 解開（依賴 SetFocus driver patch）/ 2 Status skip）、CmdGISPeople（Office）、CmdNeo4j（8 過 / 0 skip — Entry/Texts/Office/Kinship/AssociationPairs/Associations/Place/Status 全部通過。Status 於 2026-05-08 從 skip 提升為 covered，使用 matrix `status_<top_code>_unfiltered` fixture；先前的 skip 為 false-positive copy-paste（從 Pajek/Gephi cross-form Status skip 抄過來），PR #127 driver/meta probe 直接反駁了「same root family」的判斷 — CmdNeo4j 在 Status 上跑得乾淨（6 files、0 :ERR、0 watchdog dialogs），而 CmdPajek + CmdGephi 仍因 cleanup-rebind 在 `Form_LookAtStatus.vb:1457+1460` 觸發 VBA 424 'Object required'。**Status × CmdNeo4j 不需要 driver patch**（與 Place / Associations 不同）；per-shape pin 鎖死 6 個檔（People / PeopleStatus / Places / PeoplePlaces / PersonPlaceCodes / StatusCodes）；classifier 新增 1 個單列 shape（StatusCode → StatusCodes code-table）+ 1 個 2-col disambiguator（(NameID, StatusCode) → PeopleStatus）。Status × CmdPajek + CmdGephi 仍在 cross-form Pajek/Gephi test 上 skipped，受同一 cleanup-rebind 'Object required' issue 阻擋 —— 那是獨立的 driver/meta workaround verification line。Place 於 2026-05-08 從 skip 提升為 covered，使用 matrix `place_addr_<top_addr_id>` fixture，仰賴 PR #123（`_rewrite_place_cmdneo4j_trstpeople_projection` 把 `Form_LookAtPlace.CmdNeo4j_Click` 中 `tRstPeople` 的 SELECT projection 擴成含 `DYNASTIES.c_dynasty / c_dynasty_chn / BIOG_MAIN.c_female`，下游 `!c_dynasty` / `!c_female` 欄位讀取不再觸發 JET 3265 — canonical Issue #24）；per-shape pin 鎖死 6 個檔（People / PeoplePlaces / Places / PersonPlaceRelations / PersonPlaceRelCodes / IndexAddrTypeCodes）；classifier 新增 3 個 shape（PersonID / PersonPlaceRelCode / IndexAddrTypeCode）。Issue #24 仍維持 P1 canonical —— workaround 讓此 cell *可測試*，並非 *已修復*。與 Associations（#116 → #117 → #118）不同，Place 只需這一個 SELECT-projection rewrite —— `Form_LookAtPlace.CmdNeo4j_Click` 沒有 concat-form debug MsgBox 層。Entry 於 2026-05-04 從 skip 提升為 covered，固定 Issue #9 LATENT-gate；AssociationPairs 於 2026-05-07 從 skip 提升為 covered，使用 1×3 known-edged pair（與 Pajek/Gephi 同一 fixture），仰賴 PR #109 的 CmdNeo4j_Click debug-MsgBox suppress driver patch；classifier 新增 3 個 shape：Person1_ID / AssociationCode / KinshipCode。Associations 也於 2026-05-07 從 skip 提升為 covered，使用 matrix `assoc_<top_code>_unfiltered` fixture，仰賴 PR #116（c_addr_type rewrite 解掉 canonical Issue #23 的 target-column typo）+ PR #117（5-prefix debug-MsgBox suppress）；per-shape pin 鎖死 8 個檔（People / Places / PeoplePlaces / PeopleAssociations / AssociationCodes / KinshipCodes / OccasionCodes / TopicCodes）；classifier 新增 2 個 shape（OccasionCode / TopicCode）+ 1 個 2-col disambiguator。Issue #23 仍維持 P1 canonical —— workarounds 讓此 cell *可測試*，並非 *已修復*。順手挖出 Bug #7/#8/#9）。CmdKML 仍未測；CmdUCINet 的 family first coverage cell 已落地於 LookAtKinship（`tests/test_vba_cmducinet_kinship.py`，.vna 三段結構由 probe `investigate/cmducinet-family-shape` 確立）**但已知 fixture-fragile，依 Issue #22 的 runtime-confirmed Kinship sibling form** —— 通過僅是因為 matrix 提供的 person 3211 親屬網絡剛好沒有 Han 字符 c_name；換成能觸達 Han 名字的 fixture 就會在同一處重現 Issue #22 的 VBA error 5（sibling probe `investigate/kinship-cmducinet-sibling-risk` 已直接示範）；Associations CmdUCINet 已直接 canonicalize 為 Issue #22 P1（匯出途中 VBA error 5 已重現 + static + runtime marker 雙向釘死），Place CmdUCINet 仍 gap（使用 ADO Stream，需要單獨 per-form 處理）— roadmap 第 8 項
- ✅ pytest 啟動時自動偵測 `test_inputs.json` 是否過時並重跑 `discover_test_inputs.py`（`pytest --no-discover-inputs` 可關閉）— roadmap 第 9 項
- ✅ 比對 User MDB 的 `index year` / `index address` 算法與 [`cbdb-online-main-server`](https://github.com/cbdb-project/cbdb-online-main-server) — roadmap 第 12 項已成。**兩邊是兩套獨立的實作**：SQLite 快照中的 `c_index_year` ← PHP `IndexYearRebuildService.php`、`c_index_addr_id` ← `IndexAddressRebuildService.php`；User MDB 中 `c_index_addr_id` ← `analysis/dump/vba/Form_frmIndexAddr.vb`，而 `c_index_year` 在出貨的 User MDB 裡找不到對應的 rebuild VBA（疑在 Admin MDB，詳見 `analysis/index_drift_algorithm_notes.md`）。`analysis/classify_index_drift.py` 對全部共有的 personid 做完整分類，結果寫到 `reports/index_drift_classification.json`。最新一次跑：657 245 個共有 personid 中 656 682 完全一致（99.914 %）；淨差異 563 筆 = 16 筆可歸因於 birthyear/deathyear 的源資料漂移 + 547 筆待追查（`index_year_only_diff` 59、`index_addr_only_diff` 478、`index_both_diff` 10）。待追查的桶**不**等於 bug —— 可能是 PHP↔VBA 演算法差異，也可能是 evidence 表（BIOG_ADDR_DATA / ENTRY_DATA / NIAN_HAO 等）裡的漂移，那是我們本層沒比較到的。0.5 % / 0.1 % 的門檻只能擋整體突然劇烈變動，無法判定 547 筆是不是 bug。詳見 AGENTS.md「Index-year cross-check」段 + report 末尾分類匯總
- ✅ Import-list 按鈕（`tests/test_vba_import_lists.py`，11 種按鈕跨 8 個表單，15 passed + 2 skipped；只有 LookAtNetworks 兩項因 Form_Open 卡住而 skip）— roadmap 第 13 項
- ✅ Save-list 按鈕（`tests/test_vba_save_lists.py`，5 個 `CmdSave*_Click` 全部覆蓋；位元組級對比 + 對碼表 INNER JOIN 比對 desc/desc_chn）— roadmap 第 14 項
- ✅ `CmdStoreID` / `CmdRecallID` 跨 form round-trip 測試（`tests/test_vba_storeid_recallid.py`，11 passed + 1 skipped；含 Entry → Kinship 完整 round-trip）— roadmap 第 15 項

## 貢獻

詳見 [`AGENTS.md`](./AGENTS.md) 中的 landmine 與 driver 模式說明。
**任何改動如果讓 [Plan & status](#plan--status) 中的項目進度改變，
請在同一個 PR 中同步更新該表格** —
這是 `AGENTS.md` 強制執行的規則。
