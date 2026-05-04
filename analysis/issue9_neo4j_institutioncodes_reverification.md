# Issue #9 re-verification — LookAtEntry.CmdNeo4j InstitutionCodes branch

**Date:** 2026-05-04
**Source data:** `data/CBDB_BJ_User.mdb` (current dump)
**Companion JSON:** `reports/issue9_neo4j_institutioncodes_reverification.json`
**Companion script:** `analysis/investigate_issue9_neo4j_institutioncodes.py`
**Trigger:** user manual test of `c_entry_code = 101`
(recommendation / 薦舉) on LookAtEntry: clicked Neo4j, no popup,
"Finished saving to Neo4j" displayed, but no `InstitutionCodes_*.csv`
in the output folder.  This contradicts the canonical Issue #9
narrative ("popup at line 1425 from the typo `With tRstAssocCodes`").

---

## TL;DR

On the current dump:

- `ENTRY_DATA` has **263,454 rows total** and **0 of them have
  `c_inst_code > 0`** (also 0 with `c_inst_name_code > 0`).
- The InstitutionCodes branch in `Form_LookAtEntry.CmdNeo4j_Click`
  is gated by `If tRecDeleted > 0 Then` at line 1389, where
  `tRecDeleted` is the row-count of an `INSERT ... WHERE
  ZZ_SCRATCH_ENTRY.c_inst_code > 0`.  Since
  `ZZ_SCRATCH_ENTRY.c_inst_code` is copied verbatim from
  `ENTRY_DATA.c_inst_code` (lines 1645-1652), the gate evaluates
  false for **every possible LookAtEntry fixture** on this dump.
- Therefore the buggy `With tRstAssocCodes` line at 1425 is
  **unreachable on this dump**, no popup fires, no
  `InstitutionCodes_*.csv` is written, and the chain proceeds
  cleanly to "Finished saving to Neo4j".  This matches the user's
  manual observation exactly.
- The line-1425 typo (intended `tRstInstitutions`, written
  `tRstAssocCodes`) **is still a real source-level bug** — it would
  raise DAO 3021 ("No current record" on a closed Recordset) the
  moment any future dump introduces an `ENTRY_DATA` row with
  `c_inst_code > 0`.  But on the current dump it is **LATENT**, not
  user-visible.
- **Recommended reclassification:** Issue #9 is currently P1
  ("popup on a documented user step").  On the current dump it is
  better classified as a **source-level latent bug** — keep it in
  the report so the typo gets fixed, but stop describing it as a
  visible runtime popup.

---

## 1. Did the user observation reproduce? (manual)

User manual test (2026-05-04):

| Step | Observed |
|---|---|
| Pick `c_entry_code = 101` (薦舉 / recommendation) | OK |
| Click `CmdQuery` | results displayed |
| Click `CmdNeo4j` | "Finished saving to Neo4j", **no error popup** |
| Inspect output folder | other Neo4j CSVs present, **no `InstitutionCodes_*.csv`** |

The investigation script `investigate_issue9_neo4j_institutioncodes.py`
reproduces this in real Access COM (see §3 below).

---

## 2. Pure-SQL evidence: the InstitutionCodes branch is gated unreachable

### 2.1  Branch gating as actually written

```vb
' Form_LookAtEntry.vb:1375-1391
cmdSQL.CommandText = "DELETE * FROM ZZ_SCRATCH_P_TEXT"
cmdSQL.Execute tRecDeleted

tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id ) " + _
            "SELECT DISTINCT ZZ_SCRATCH_ENTRY.c_personid " + _
            "FROM ZZ_SCRATCH_ENTRY " + _
            "WHERE (((ZZ_SCRATCH_ENTRY.c_inst_code)>0))"
cmdSQL.CommandText = tQueryStr
cmdSQL.Execute tRecDeleted    ' ← row-count of the INSERT lands here

tQueryStr = "SELECT DISTINCT ZZ_SCRATCH_ENTRY.c_inst_code, ..."

If tRecDeleted > 0 Then         ' ← gates the SaveAs prompt + buggy block
    dlgSaveAs.InitialFileName = "InstitutionCodes_" + tCodeStr + ".csv"
    ...
    Set tRstInstitutions = CurrentDb.OpenRecordset(tQueryStr)
    ...
    With tRstAssocCodes         ' ← typo; intended tRstInstitutions
        .MoveFirst              ' ← would raise DAO 3021 if reached
```

### 2.2  Counts on the current dump

From `q1_entry_data_global` in the JSON:

| Quantity | Count |
|---|---:|
| `ENTRY_DATA` total rows | **263,454** |
| `ENTRY_DATA` rows with `c_inst_code > 0` | **0** |
| `ENTRY_DATA` rows with `c_inst_name_code > 0` | **0** |

### 2.3  Per-fixture pre-image

`q1_per_fixture_preimage` (CmdQuery copies `ENTRY_DATA.c_inst_code`
verbatim into `ZZ_SCRATCH_ENTRY.c_inst_code`, so an
`ENTRY_DATA WHERE c_entry_code = X AND c_inst_code > 0` count IS
the pre-image of the gating INSERT's row-count for that fixture):

| `c_entry_code` | year filter | rows in `ENTRY_DATA` | with `c_assoc_code > 0` | with `c_inst_code > 0` | with `c_inst_name_code > 0` | branch entered? |
|---:|---|---:|---:|---:|---:|---|
| 36 (jinshi / 進士) | none | 92,514 | 2 | **0** | 0 | **NO (gated out)** |
| 101 (recommendation / 薦舉) | none | 878 | 6 | **0** | 0 | **NO (gated out)** |
| 36 (jinshi) | 1100-1110 | 989 | 0 | **0** | 0 | **NO (gated out)** |

The branch cannot be entered for any fixture on this dump.

---

## 3. Real Access COM evidence

Driven by `--com` mode of the investigation script.  Each fixture
gets its own VbaSession, its own work-mdb copy, and its own
output dir (with `kill_orphan_access` between fixtures so file
locks don't leak across).

| `c_entry_code` | year filter | CmdQuery: `ZZ_SCRATCH_ENTRY` `c_inst_code > 0` | CmdNeo4j outcome | Files produced | InstitutionCodes file present? |
|---:|---|---:|---|---:|---|
| 36 | none | 0 | COM died after CmdQuery (RPC unavailable; 92,514-row CmdNeo4j too heavy for one COM session) | n/a | n/a — but pre-image already shows 0 rows would gate the branch even at full scale |
| 101 | none | 0 | **chain finished cleanly, no ERR marker** | **7** | **ABSENT** |
| 36 | 1100-1110 | 0 | **chain finished cleanly, no ERR marker** | **6** | **ABSENT** |

### 3.1  Files produced for `c_entry_code = 101` (full fixture)

`patch_filedialog` redirects every `dlgSaveAs.Show` to
`GetTestExportPath()` returning sequential `f<n>.out.csv`, so the
file-shape inference comes from the first column of each file:

| File | Bytes | First-col header | Block (per Form_LookAtEntry.vb) |
|---|---:|---|---|
| f3.out.csv | 33,333 | `nameID` | People |
| f6.out.csv | 35,458 | `NameID` | PeopleEntry |
| f9.out.csv | 20,594 | `placeID` | Places |
| f12.out.csv | 12,431 | `nameID` | PeoplePlaces |
| f15.out.csv | 197 | `personPlaceCode` | PersonPlaceCodes |
| f18.out.csv | 79 | `EntryCode` | EntryCodes |
| f21.out.csv | 89 | `AssocCode` | AssocCodes |

**No `InstitutionCode`-shape file in the set** — the branch was
silently skipped.  Also no `:ERR` marker in `ZZ_TEST_DEBUG`.
Two `:MSGBOX` markers (the form's "Finished saving to Neo4j"
informational popup, intercepted by the harness) — these are
informational, not error popups.

### 3.2  Files produced for `c_entry_code = 36` (narrowed 1100-1110)

| File | Bytes | First-col header | Block |
|---|---:|---|---|
| f3.out.csv | 37,158 | `nameID` | People |
| f6.out.csv | 41,110 | `NameID` | PeopleEntry |
| f9.out.csv | 12,155 | `placeID` | Places |
| f12.out.csv | 15,364 | `nameID` | PeoplePlaces |
| f15.out.csv | 159 | `personPlaceCode` | PersonPlaceCodes |
| f18.out.csv | 93 | `EntryCode` | EntryCodes |

One file fewer than the 101 fixture: **AssocCodes block also
skipped** (because `ZZ_SCRATCH_ENTRY.c_assoc_code > 0` count was
0 for this narrowed window — the same per-block gate-on-zero
pattern that skips InstitutionCodes for both fixtures).  This is
direct corroboration that the per-block "skip when no rows
qualify" mechanism works as designed in the surrounding code,
and the InstitutionCodes branch is silently skipped in exactly
the same way.

### 3.3  Why the full code-36 COM probe died

`CmdQuery` against `c_entry_code = 36` lands 92,514 rows in
`ZZ_SCRATCH_ENTRY`.  The follow-up CmdNeo4j chain (six per-block
SQL builds + write loops over 92k rows) reliably crashed
MSACCESS.EXE with `RPC server unavailable`.  This is an Access /
COM-driver memory or watchdog issue, **not** evidence about the
InstitutionCodes branch.  The CmdQuery evidence we did capture
already shows `c_inst_code > 0` rows = 0 even at full scale, so
the branch couldn't have fired even if the chain had completed.
The narrowed 1100-1110 fixture above gives the in-COM
end-to-end verdict for code 36.

---

## 4. Reclassification

| Question (from brief) | Answer |
|---|---|
| Does `c_entry_code = 101` actually pop up? | **No.** SQL pre-image: 0 rows would enter the branch.  COM probe: 7 files produced, InstitutionCodes absent, no `:ERR` marker.  Matches user's manual observation. |
| Does `c_entry_code = 36` pop up? | **No** at the narrowed-fixture COM scale (6 files, InstitutionCodes absent, no ERR), and the SQL pre-image (0 `c_inst_code > 0` rows) is the same at full scale.  Full-fixture COM probe couldn't be run to completion (Access RPC death on 92k-row chain), but the gating evidence is identical. |
| Is the missing `InstitutionCodes_*.csv` an error, gated-out skip, or latent bug in the branch? | **Gated-out skip** on the current dump.  The line-1425 typo (`With tRstAssocCodes` instead of `With tRstInstitutions`) is a real source-level bug, but it is unreachable on the current dump for any fixture. |
| Should Issue #9 stay P0? | **No.**  Recommend re-classifying to a **source-level latent typo**.  Current dump has no `c_inst_code > 0` rows in `ENTRY_DATA`, so the branch cannot fire under any LookAtEntry fixture; there is no user-visible runtime symptom today.  The typo will surface immediately if a future MDB drop introduces inst rows. |

> Note on `c_entry_code = 36` "popup": the SQL pre-image shows it
> ALSO has `c_inst_code > 0` count of 0, so it does not pop up
> either.  Both fixtures land in the same "no popup, missing
> InstitutionCodes file" state; the bug class is identical.

---

## 5. What changed vs the existing canonical narrative

The current `reports/CBDB_Issues_Report_EN.md` Issue #9 block
describes the bug as a runtime DAO 3265/3021 popup at line 1425
that the user can trigger today.  On the current dump, that
narrative is **wrong in the user-facing sense**.  The line-1425
typo IS real, but the gating `If tRecDeleted > 0 Then` at line
1389 prevents it from ever executing.  The correct narrative is:

1. There is a latent typo at line 1425 (`tRstAssocCodes` was
   meant to be `tRstInstitutions`).
2. It is gated out by line 1389 on this dump because no
   `ENTRY_DATA` row has `c_inst_code > 0`.
3. The user-visible symptom on this dump is **InstitutionCodes
   CSV silently absent from the export**, not a runtime popup.
4. If a future dump introduces inst rows, the typo becomes
   visible as a DAO 3021 popup.

---

## 6. Screenshot status

- `reports/screenshots/bug9_form_open.png`,
  `reports/screenshots/bug9_form_annotated.png`,
  `reports/screenshots/bug9_faux_popup.png` are now misleading
  on two grounds:
  - The form panels show empty Select Entry, suggesting "no
    entry selected", and
  - the popup shown is a faux 3265 popup, not reproducible on
    the current dump.

Recommended follow-up (not done in this PR per the brief — this
PR ships investigation artifacts only):

- If Issue #9 stays in the report as a latent source-level
  typo, the screenshots should either be removed (so readers
  don't infer a runtime popup is reproducible today) or
  replaced with:
  - A real CmdQuery screenshot showing
    `Entry Type = recommendation / 薦舉` selected and the result
    grid populated, AND
  - A screenshot of the Neo4j output folder showing
    `Finished saving to Neo4j` + the seven CSV files **without**
    `InstitutionCodes_*.csv` (i.e. the actual user-visible
    symptom: the file is silently missing).

The canonical report itself is left untouched in this PR;
re-classification + screenshot refresh belongs to a follow-up
PR after this investigation is reviewed.

---

## 7. How to re-run

```
# SQL evidence only:
python analysis/investigate_issue9_neo4j_institutioncodes.py

# SQL + real Access COM probes for all three fixtures
# (will start and kill MSACCESS.EXE several times):
python analysis/investigate_issue9_neo4j_institutioncodes.py --com
```

JSON output: `reports/issue9_neo4j_institutioncodes_reverification.json`.
COM-phase artifacts: `analysis/_issue9_reverify_neo4j_out/entry_*/`.
