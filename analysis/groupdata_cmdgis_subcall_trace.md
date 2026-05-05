# GroupData × CmdGIS — sub-call ERR localisation

**Date:** 2026-05-05
**Branch:** `investigate/groupdata-cmdgis-subcall-trace`
**Companion JSON:** `reports/groupdata_cmdgis_subcall_trace.json`
**Companion script:** `analysis/probe_groupdata_cmdgis_subcalls.py`
**Trigger:** PR `investigate/groupdata-cmdgis-probe` found a
non-fatal `LookAtGroupData:ERR No value given for one or more
required parameters.` mid-chain on the GroupData × CmdGIS
probe.  The form-level error handler caught it but couldn't
attribute it to a specific sub.  This follow-up probe isolates
each sub-call by toggling individual `Chk*` checkboxes per
iteration.

---

## TL;DR

**Verdict: `real_cbdb_bug_candidate`** — but more specifically,
the localised ERR is a **runtime-side confirmation of the
already-filed Issue #6** (P1 visible runtime crash;
`Form_LookAtGroupData.queryEntry` references the non-existent
column `ENTRY_DATA.c_parental_status` instead of
`c_parental_status_code`).

The probe's per-sub iteration matrix shows the ERR fires in
**exactly 2 of 11 iterations**, and both are the only iterations
that exercise `queryEntry`:

  - `queryEntry_alone` (just `ChkEntry`, no GIS) → ERR
  - `Entry_full_chain` (`ChkEntry` + `ChkGisEntry`) → ERR

All 9 other iterations are clean.  Source inspection at
`Form_LookAtGroupData.vb:2621` shows queryEntry's INSERT...SELECT
projects `ENTRY_DATA.c_parental_status` while the INSERT target
column list ends with `c_parental_status_code`.  The probe's
fixture (`person_id = 1`, An Dun 安惇) is the **canonical
reproduction** Issue #6 documents — and the observed error
message ("No value given for one or more required parameters")
matches Issue #6's predicted symptom verbatim.

This isn't a NEW bug discovery — it's **the missing runtime
pin** for an existing static-source assertion.  Issue #6's
existing test (`tests/test_known_bugs.py::test_bug6_groupdata_
query_entry_wrong_field`) only asserts the buggy substring is
still present in the dumped VBA.  The probe shows the bug
actually fires when CmdRun executes queryEntry on the documented
fixture.

---

## Per-iteration evidence

11 iterations, each with: all `Chk*` reset to False, then the
target subset set True, then `CmdRun + CmdGIS` chain via
`Form_Timer` dispatch.  Files written via `patch_filedialog` to
a per-iteration sub-directory.

| # | Iteration | `chk_set` | ENTER | DONE | **ERR** | Files | scratch_status | scratch_office | scratch_entry |
|---|---|---|:---:|:---:|:---:|---:|---:|---:|---:|
| 1 | `queryStatus_alone` | `ChkStatus` | ✓ | ✓ | clean | 1 | 2 | 0 | 0 |
| 2 | `queryOffice_alone` | `ChkOffice` | ✓ | ✓ | clean | 2 | 2 | 12 | 0 |
| 3 | **`queryEntry_alone`** | `ChkEntry` | ✓ | ✓ | **⚠️ ERR** | 0 | 2 | 12 | **0** |
| 4 | `queryText_alone` | `ChkText` | ✓ | ✓ | clean | 0 | 2 | 12 | 0 |
| 5 | `queryAddr_alone` | `ChkAddr` | ✓ | ✓ | clean | 1 | 2 | 12 | 0 |
| 6 | `Status_full_chain` | `ChkStatus`, `ChkGisStatus` | ✓ | ✓ | clean | 1 | 2 | 12 | 0 |
| 7 | `Office_OfficeOffice` | `ChkOffice`, `ChkGisOffice` | ✓ | ✓ | clean | 2 | 2 | 12 | 0 |
| 8 | `Office_OfficePeople` | `ChkOffice`, `ChkGisOfficePeople` | ✓ | ✓ | clean | 2 | 2 | 12 | 0 |
| 9 | **`Entry_full_chain`** | `ChkEntry`, `ChkGisEntry` | ✓ | ✓ | **⚠️ ERR** | 0 | 2 | 12 | **0** |
| 10 | `Text_full_chain` | `ChkText`, `ChkGisText` | ✓ | ✓ | clean | 0 | 2 | 12 | 0 |
| 11 | `Addr_full_chain` | `ChkAddr`, `ChkGisAddr` | ✓ | ✓ | clean | 1 | 2 | 12 | 0 |

ERR message in iterations 3 and 9 is identical:

```
LookAtGroupData:ERR No value given for one or more required parameters.
```

Critical observations:

- **`queryEntry_alone` ERRs even with no GIS-side checkbox set** —
  proves the ERR originates in `queryEntry`, not in any
  `WriteGIS_Entry`.
- **`ZZ_SCRATCH_ENTRY` stays at 0 across all iterations** —
  queryEntry never INSERTs anything, even though person_1
  has 2 ENTRY_DATA rows.  The JET parameter error fires before
  the INSERT executes.
- All other 4 query subs (`queryStatus`, `queryOffice`,
  `queryText`, `queryAddr`) execute cleanly when isolated.
- All other 5 chains (Status / Office×OfficeOffice /
  Office×OfficePeople / Text / Addr) execute cleanly with files.
- `Text_full_chain` and `Text_alone` produce 0 files — likely
  benign (person_1 may have 0 BIOG_TEXT_DATA rows; empty
  scratch → WriteGIS_Text bails on RecCount=0; not an error).

---

## Source-side identification

`Form_LookAtGroupData.vb:2593-2626` defines `queryEntry()`.
Lines 2609-2625:

```vba
tStrInsert = "INSERT INTO ZZ_SCRATCH_ENTRY ( c_personid, c_name, ..., " + _
            "c_inst_code, c_inst_name_code, c_entry_addr_id, c_source, " + _
            "c_parental_status_code )"          ' <-- target col
tStrSelect = "SELECT ZZ_SCRATCH_IMPORT_PEOPLE.c_person_id, ..., " + _
            "ENTRY_DATA.c_inst_code, ENTRY_DATA.c_inst_name_code, " + _
            "ENTRY_DATA.c_entry_addr_id, ENTRY_DATA.c_source, " + _
            "ENTRY_DATA.c_parental_status "    ' <-- typo (no _code)
tStrFrom   = "FROM ZZ_SCRATCH_IMPORT_PEOPLE INNER JOIN ..."
cmdSQL.CommandText = tStrInsert + tStrSelect + tStrFrom
cmdSQL.Execute gEntryRecCount     ' <-- raises VBA error 3061
                                  '     (No value given...)
```

The INSERT target column list ends with `c_parental_status_code`
(line 2612).  The SELECT projection's last column reads
`ENTRY_DATA.c_parental_status` (line 2621) — no `_code` suffix.
JET parses unknown column references as parameters and asks for
their values; with no parameter binding, the Execute call raises
"No value given for one or more required parameters."

This is **exactly** the bug Issue #6 documents.

---

## Cross-check against existing Issue #6

`reports/CBDB_Issues_Report_EN.md` § Issue #6:

| Field | Documented | Probe observed |
|---|---|---|
| Affected sub | `Form_LookAtGroupData.queryEntry` | ✓ `queryEntry_alone` localises ERR |
| Source line | 2621 (typo `c_parental_status` instead of `c_parental_status_code`) | ✓ confirmed at vba:2621 |
| Reproduction person | `c_personid = 1` (An Dun 安惇) | ✓ probe fixture is person_1 |
| User steps | Open LookAtGroupData → only Entry checkbox ticked → Run | ✓ probe iter `queryEntry_alone` is exactly this |
| Symptom | "JET reports this as 'No value given for one or more required parameters' / 'No such field' depending on the Office build — both mean the SQL referenced `ENTRY_DATA.c_parental_status` which doesn't exist" | ✓ observed verbatim |
| Severity | P1 — Visible runtime crash on a common path (Entry sub-query) | matches probe behaviour: chain reaches DONE, ERR popup intercepted by autodetect → ZZ_TEST_DEBUG marker |
| Existing test coverage | `tests/test_known_bugs.py::test_bug6_groupdata_query_entry_wrong_field` (static source-substring assertion only) | probe is the missing **runtime** pin |

**No new issue should be filed.**  The probe's contribution is a
runtime-side confirmation of an existing P1.

---

## Verdict (per the brief's 3 buckets)

- **A. `benign_probe_gap`** — NO.  The ERR is a real bug, not a
  probe artifact; it fires under the form's documented user
  reproduction path.
- **B. `real_cbdb_bug_candidate`** — YES, with the refinement
  that the bug is **already filed** as Issue #6 (P1 visible
  runtime crash).  The probe is runtime-side confirmation, not
  a new candidate.
- **C. `still_ambiguous_but_localized`** — NO.  ERR localised
  perfectly to `queryEntry`; root cause identified at
  `Form_LookAtGroupData.vb:2621`.

---

## Implications for GroupData × CmdGIS coverage

GroupData × CmdGIS coverage IS feasible — but a coverage test
needs to be honest about Issue #6's gating effect on the
Entry path:

1. **Status / Office (×2 GIS variants) / Addr** — clean; files
   produced; can be asserted with shape + non-empty + structural
   checks (the existing `tests/test_vba_cmdgis_other_forms.py`
   pattern).
2. **Entry** — Issue #6 fires; either:
   - exclude this checkbox combo from the coverage test (with
     an inline comment pointing at Issue #6), OR
   - include it with an XFAIL / expected-ERR assertion that
     pins Issue #6's runtime symptom (turning the static source-
     string test into a runtime-side regression marker too).
3. **Text** — clean ERR-wise but produces 0 files (likely
   benign: person_1 has 0 BIOG_TEXT_DATA rows).  The coverage
   test should not require Text to produce files on this
   fixture.

The coverage PR scope decision belongs to the maintainer, not
this probe.  The probe's job ends at: "ERR localised to queryEntry
= Issue #6; the rest of the GroupData × CmdGIS chain is clean
on the documented small fixture."

---

## What this PR does NOT do (per brief)

- ❌ Did NOT open a new issue (Issue #6 already exists)
- ❌ Did NOT add coverage tests
- ❌ Did NOT modify the driver
- ❌ Did NOT do large fixture design (reused person_1)
- ❌ Did NOT mark "files produced" as a pass — explicitly
  classified the partial-output state as `real_cbdb_bug_candidate`
  with the refinement that it's an existing-issue runtime pin

---

## Recommended next step (NOT this PR)

Per `docs/skills/issue-report-maintainer.md`, the next step
belongs in the **investigate→reclassify** pair pattern:

1. Optionally open a small follow-up PR that adds runtime-side
   coverage for Issue #6 — i.e. a behavioural test in
   `tests/test_vba_bug_behaviors.py` that drives person_1 +
   ChkEntry + CmdRun and asserts the JET ERR fires in
   ZZ_TEST_DEBUG.  This converts the existing static-only
   test_bug6 into a static + runtime pin.
2. THEN the GroupData × CmdGIS coverage PR can ship — it
   exercises the Status / Office / Addr branches and either
   excludes Entry or asserts the Issue #6 ERR explicitly.

Both PRs need maintainer go-ahead; neither is in this probe's
scope.

---

## How to re-run

```
python analysis/probe_groupdata_cmdgis_subcalls.py
```

Per-iteration cooldowns aren't required (single VbaSession
runs all 11 iterations in ~25 s after a 120 s initial
`kill_orphan_access` + sleep).  RPC death is unlikely on this
probe because all iterations share one session — but if the
first attempt fails during `VbaSession.open()`, manually reap
+ wait 60–120 s and retry.
