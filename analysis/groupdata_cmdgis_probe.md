# GroupData × CmdGIS — probe-first investigation

**Date:** 2026-05-05
**Branch:** `investigate/groupdata-cmdgis-probe` (off main `45dca39`)
**Companion JSON:** `reports/groupdata_cmdgis_probe.json`
**Companion script:** `analysis/probe_groupdata_cmdgis.py`
**Trigger:** the gap-triage classified GroupData × CmdGIS as
bucket A (small_candidate, low risk), but the AssociationPairs
probe surfaced a hidden CmdQuery SetFocus blocker for what was
also rated low risk.  This probe answers — without writing any
test — whether GroupData has a similar hidden blocker before
any coverage PR is opened.

---

## TL;DR

**Verdict: `needs_investigation`** — GroupData × CmdGIS is **NOT
blocked the way AssociationPairs is**, but the chain emits a
non-fatal `LookAtGroupData:ERR No value given for one or more
required parameters.` mid-flight.  Files are still produced and
the chain reaches `DONE`, so coverage is mechanically feasible —
but the embedded ERR needs to be identified before shipping a
test, since it could either be:

  - a benign side-effect of the probe seeding (e.g. a sub-query
    that needs a control value the probe didn't set), OR
  - a real CBDB runtime bug in the same family as Issues #7–9
    (recordset / SELECT projection mismatch).

Until that's classified, opening a coverage PR risks pinning
either a false-positive assertion ("output is N rows!") that
breaks on the next dump, or shipping a test that silently
swallows a real bug (the AssociationPairs lesson again, in a
different shape).

**Concrete next step:** a small follow-up probe that wraps each
of CmdRun's per-checkbox sub-calls (queryStatus / queryOffice /
queryEntry / queryText / queryAddr) and each of CmdGIS's
WriteGIS_X calls in its own marker so the ERR can be traced to
a specific Sub.  THEN decide coverage scope.

---

## Probe setup

| Element | Value |
|---|---|
| Form | `LookAtGroupData` |
| Fixture person id | **1** (`An Dun` / `安惇`, Song dynasty) — same person matrix_hard_forms's `groupdata_person_1_small` uses |
| Pre-probe SQL | `SELECT COUNT(*) FROM STATUS_DATA WHERE c_personid = 1` → **2** |
| Pre-probe SQL | `SELECT COUNT(*) FROM BIOG_MAIN WHERE c_personid = 1` → **1** |
| Checkboxes set via COM | `ChkStatus = True`, `ChkGisStatus = True` |
| Trigger button | `CmdRun` (LookAtGroupData uses CmdRun, not CmdQuery) |
| Chain | `set_form_tag(form, "CmdRun,CmdGIS", out_path)` |
| Watcher | `result_table=ZZ_SCRATCH_STATUS`, timeout 180 s |
| Cooldown before probe | `kill_orphan_access()` + 60–120 s sleep (RPC death on first attempt; second attempt clean) |

`out_path` was a **single-file path** (not directory mode).  See
"Probe artifact" §4 below for what that caused.

---

## Per-question answers

### Q1 — Does CmdRun populate per-checkbox sub-queries (specifically queryStatus → ZZ_SCRATCH_STATUS)?

**YES.** End-to-end verified:

  - `ZZ_SCRATCH_IMPORT_PEOPLE` row for `c_person_id = 1` got
    backfilled with `c_name = 'An Dun'`, `c_dynasty = 'Song'`
    (matrix_hard_forms's `_check_assoc_pairs`-equivalent
    baseline check passed)
  - `ZZ_SCRATCH_STATUS` direct count = **2** (matches the 2
    STATUS_DATA rows person_1 has)
  - `click_via_timer` watcher returned `n = 2` matching the
    direct table count (no flake)

Conclusion: CmdRun's body fully executes; queryStatus dispatch
works; no upstream-CmdQuery blocker à la AssociationPairs.

### Q2 — Does CmdGIS export a file in the same session?

**YES.**  TWO files were produced:

| File | Bytes | Header (tab-separated) | Cols | Data rows |
|---|---:|---|---:|---:|
| `groupdata_status_gis.tab` | 142 | `Name\tNameChn\tSex\tIndexYear\tAddrID\tAddrName\tAddrChn\tX\tY\txy_count` | 10 | 1 |
| `groupdata_status_gis.tab.tab` | 1,277 | `Office\tOfficeChn\tFirstYear\tLastYear\tDynasty\tOfficeAddr\tOfficeAddrChn\tX\tY\txy_count` | 10 | 12 |

The chain dispatched **two** `WriteGIS_X` subs (Status + Office),
even though I only set `ChkGisStatus = True` via COM.  See
"Probe artifact" §4 for the most likely explanation (form-level
defaults / Chk*_Click side-effects).

### Q3 — Does the produced .tab pass structural checks the cross-form CmdGIS test would apply?

| Check | Pass? | Evidence |
|---|---|---|
| File exists | ✓ | both files materialised on disk |
| Non-empty | ✓ | 142 B and 1,277 B |
| Header reasonable | ✓ | both have ≥ 2 cols, contain `ID` substring (matches `_GIS_REQUIRED_COLUMNS` discrimination logic in `tests/test_vba_cmdgis_other_forms.py`) |
| First rows consistent column count | ✓ | both files: `header_n_cols == first_row_n_cols == 10` |

The probe's structural-check verdict was `True` on all 4
dimensions.

### Q4 — Are there any hidden blockers?

| Blocker class | Observed? | Detail |
|---|---|---|
| SetFocus / active form (AssociationPairs class) | **No** | CmdRun body executed normally; backfill + queryStatus both ran |
| DONE marker missing | **No** | `LookAtGroupData:DONE` present in `ZZ_TEST_DEBUG` |
| Watcher `n=0` flake | **No** | `n=2` matched direct table count |
| Export dialog hang | **No** | `patch_filedialog` redirected `Show` calls; chain didn't block |
| **`:ERR` marker mid-chain** | **YES — non-fatal** | one entry: `LookAtGroupData:ERR No value given for one or more required parameters.` |

`ZZ_TEST_DEBUG` transcript (3 entries):

```
1  LookAtGroupData:ENTER
2  LookAtGroupData:ERR No value given for one or more required parameters.
3  LookAtGroupData:DONE
```

The ERR fires BETWEEN ENTER and DONE.  Critically: **the chain
continued** after the ERR (DONE marker present, files produced).
This is a different failure-shape than AssociationPairs, where
the ERR aborted the chain before any INSERT ran.  The error
message ("No value given for one or more required parameters")
is JET's standard error for a SELECT that references an unknown
column / parameter — same bug-class as Issues #7-9 if it turns
out to be in a release-shipped code path.

The probe's marker injection wraps the FORM-level chain
(`Form_Timer` body), not individual Sub calls inside CmdRun /
CmdGIS, so this transcript can't point at WHICH Sub raised the
error.  That's the gap a follow-up probe would close.

---

## Probe artifact

§4 — The probe ran in single-file path mode, not directory mode.

When `set_form_tag(form, "CmdRun,CmdGIS", out_path)` passes a
single file path (no trailing backslash), every patched
`dlgSaveAs.Show` returns the SAME path.  WriteGIS_Status wrote
to `groupdata_status_gis.tab`; WriteGIS_Office then asked for
the path again, got the same string, and the OS / VBA path-
extension logic produced `groupdata_status_gis.tab.tab` to avoid
overwriting.

A coverage PR would use directory mode (`str(out_dir) + "\\"`)
so each `Show` call gets a fresh `f<n>.out` per the existing
pattern in `tests/test_vba_cmdneo4j_cross_form.py`.  This is a
non-issue for the investigation; flagged here so the next
implementer doesn't re-discover it.

The fact that BOTH WriteGIS_Status AND WriteGIS_Office fired
(when only `ChkGisStatus` was set via COM) suggests one of:

  - the form's design defaults `ChkGisOffice` to True
  - setting `ChkStatus_Click` event-fires (via the `_Click`
    handler at `Form_LookAtGroupData.vb:63`) propagates into
    related GIS-checkboxes
  - some other Form_Open / autodetect interaction

For the **investigation** this is informational.  For the
**coverage PR** (if it ever runs) the test should explicitly
read each Chk* control's value before the chain fires so the
test's expectations match the form's actual state.

---

## Verdict matrix

| Q | Answer |
|---|---|
| Q1 — CmdRun populates queryStatus | ✓ YES |
| Q2 — CmdGIS exports a file | ✓ YES (2 files, expected 1) |
| Q3 — Structural checks pass | ✓ YES on both files |
| Q4 — Hidden blockers? | ⚠️ Non-fatal `:ERR` mid-chain — different shape from AssociationPairs |

**Overall verdict (per script):** `needs_investigation`

**Why not `safe_to_cover_next`:** the embedded `:ERR` is unexplained.
Shipping a coverage test now would either:
  - Pin "1 row in Status export, 12 rows in Office export" as
    expected output — but those numbers depend on which Subs
    actually ran (we don't know yet whether the ERR caused a
    Sub to bail with partial output)
  - OR pin only "files exist + non-empty + structural" — which
    silently swallows the ERR (the AssociationPairs lesson:
    matrix_hard_forms's loose-pass test hid the SetFocus
    blocker for months)

**Why not `blocked_by_driver_issue`:** the chain is NOT blocked.
CmdRun fully executes; CmdGIS produces well-formed output.
The blocker class is fundamentally different from
AssociationPairs — that one was a complete activation failure;
this one is an embedded JET parameter error inside an otherwise-
working chain.

---

## Recommended next action (NOT this PR)

A small follow-up probe to identify which Sub the `:ERR` comes
from, by injecting per-Sub markers around each of:

  - `queryStatus` (already populated ZZ_SCRATCH_STATUS, so
    likely OK — but verify)
  - the **other** sub-queries CmdRun calls (`queryOffice`,
    `queryEntry`, `queryText`, `queryAddr`) — even when their
    Chk* is False, the form may still touch them; that's where
    "No value given" most plausibly fires
  - `WriteGIS_Status` and `WriteGIS_Office` (both fired in this
    probe — one of them may be the ERR source)

Once the ERR is traced to a specific Sub, the maintainer can
decide:

  - **If the Sub's ERR is benign / probe-induced** (e.g. a sub-
    query that needs a control value our probe didn't set) →
    fix the probe / coverage seed and proceed to a coverage PR
  - **If the Sub's ERR is a real CBDB bug** in a release-shipped
    code path → file as a new issue (P-tier per the maintainer's
    judgement) following the issue-report-maintainer skill, then
    optionally write a coverage test that asserts the ERR
    classification

Either way, the coverage PR comes AFTER this classification,
not in parallel with it.

---

## Constraints honoured

- ✓ Investigation artifacts only: `analysis/probe_groupdata_cmdgis.py`,
  `analysis/groupdata_cmdgis_probe.md` (this file),
  `reports/groupdata_cmdgis_probe.json`
- ✓ NOT a coverage PR — no `tests/` changes
- ✓ Did NOT touch README / reports / issue severity / driver
- ✓ Did NOT do large fixture design — reused the existing
  `groupdata_person_1_small` fixture from `matrix_hard_forms`
- ✓ Used Access COM via `VbaSession` per the brief's permission

## How to re-run

```
python analysis/probe_groupdata_cmdgis.py
```

(Requires Access COM available; reaps orphan MSACCESS first +
sleeps 45 s.  If RPC fails on first attempt, manually
`kill_orphan_access()` + sleep 60-120 s and retry.)
