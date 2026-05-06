# AssociationPairs × CmdGIS — honest negative finding

**Date:** 2026-05-05
**Branch:** `note/assocpairs-cmdgis-not-cheap-next` (artifacts only;
no code, test, driver, README, or canonical-report changes)

## Headline

`LookAtAssociationPairs × CmdGIS` is **NOT** the same class of
problem as `LookAtPlace × CmdGIS` / `LookAtKinship × CmdGIS`, and
**`_SUBFORMS_TO_REQUERY` cannot fix it on its own.**  The cell
remains `gap` / `blocked`.  It should be removed from the
"cheapest next cell" candidate list and not pursued without a
fresh, dedicated driver/meta investigation.

## Context

After the AssociationPairs SetFocus driver patch landed
(`feat/driver-patch-associationpairs-setfocus`, commit `89b46a9`)
and the AssociationPairs × CmdPajek + CmdGephi coverage PR
landed (`cover/assocpairs-pajek-gephi-1x3`, commit `4b8a927`),
the natural next-cheapest cell appeared to be
AssociationPairs × CmdGIS.  This was attempted in
`cover/assocpairs-cmdgis-1x3` (no commits, branch deleted after
this finding) with explicit reviewer authorization for a single
1-line same-class extension to `_SUBFORMS_TO_REQUERY`.

## What was tried

1. **Driver-side same-class extension** (authorized): added
   `"Form_LookAtAssociationPairs": ["ZZ_SCRATCH_PEOPLE"]` to
   `tests/cbdb_driver/vba_session.py::_SUBFORMS_TO_REQUERY`.
   Sanity check showed the entry was correctly applied (the
   chain block's `<subform>.Form.Requery\n` line was emitted
   for AssociationPairs).
2. **Test wiring**: added `LookAtAssociationPairs` to
   `_FORMS_WITH_CMDGIS_TESTABLE_HERE` in
   `tests/test_vba_cmdgis_other_forms.py` with a custom
   `_assocpairs_1x3_fixture()` (the same 1×3 known-edged pair
   that worked for the SetFocus probe and the Pajek/Gephi
   coverage PR).
3. **Defensive Chk\* resets** in the fixture controls
   (`ChkKML=0`, `ChkIncludeID=0`, `ChkDegree=0`,
   `Chk2Nodes=0`, `ChkKinship=0`).  `ChkKML.Value=True` is
   the only Chk gate inside `CmdGIS_Click` itself
   (line 2016 — calls `writeKML` and exits early), and the
   1×3 fixture's previous tests didn't need to set it.
4. **Test-side diagnostic prints** (only on the failure
   path, so passing runs aren't slowed) capturing
   `ZZ_SCRATCH_PEOPLE` / `ZZ_SOCIAL_NETWORK` underlying-table
   counts and the full `ZZ_TEST_DEBUG` transcript at the
   moment the assertion fires.

## What was observed

| Marker | Value |
|---|---|
| `ZZ_SCRATCH_PEOPLE` underlying-table count | **3 rows** (CmdQuery body ran successfully — INSERT INTO ZZ_SCRATCH_PEOPLE fired and committed) |
| `ZZ_SOCIAL_NETWORK` underlying-table count | **5 rows** (Link1stOrder / Link2ndOrder calls completed) |
| Output file `gis_LookAtAssociationPairs.tab` | **never appeared** |
| `ZZ_TEST_DEBUG` entries at assertion time | **1** — only the autodetect ENTER marker; the chain block's terminal `<form>:DONE` marker never fired |

So the body ran, the tables got populated, the chain dispatched
to CmdGIS_Click — but CmdGIS bailed somewhere between `ChkKML`
gate (line 2016) and the chain block's DONE write, without
producing the file.

## Why it isn't the Place / Kinship pattern

`Form_LookAtAssociationPairs.CmdGIS_Click` line 2023 reads
`ZZ_SCRATCH_PEOPLE.Form.Recordset.RecordCount` and bails with
`MsgBox "There are no records to save."` on `= 0`.  That's
syntactically identical to Place's `frmZZZ_PLACE.Form.Recordset.
RecordCount` check at line 1531 (and Kinship's at line 64).

But the upstream rebind shape is subtly different:

- `Form_LookAtPlace.vb:1053` (and `:1456`):
  `Set frmZZZ_PLACE.Form.Recordset = tRstPlace`
  (where `tRstPlace` was set earlier — typically already
  visited / has a populated RecordCount).
- `Form_LookAtKinship.vb:1593`:
  `Set frmZZ_SCRATCH_KIN.Form.Recordset = gRstPersonID`
  (similar — bind to an existing recordset variable).
- **`Form_LookAtAssociationPairs.vb:2000` (in the
  `Exit_CmdQuery_Click` cleanup):
  `Set ZZ_SCRATCH_PEOPLE.Form.Recordset = CurrentDb.OpenRecordset
  ("ZZ_SCRATCH_PEOPLE", dbOpenDynaset)`** — opens a
  **fresh, never-visited** dynaset inline.

For DAO `dbOpenDynaset`, `RecordCount` returns *records visited
so far*, not the total count.  A freshly-opened dynaset that has
never been iterated returns `RecordCount = 0` until something
calls `MoveLast` (or iterates).  That's likely why Place / Kinship
work via `.Form.Requery` (their bound recordsets were already
visited) but AssociationPairs doesn't.

## Why we didn't try a smarter shim

The existing `_SUBFORMS_TO_REQUERY` definition has an explicit
warning comment for `LookAtStatus`:

> Status doesn't need a requery here — its CmdQuery cleanup
> block at Exit_Run_Query already rebinds both subforms via
> `Set ZZ_SCRATCH_STATUS.Form.Recordset = CurrentDb.OpenRecordset
> (...)`.  An extra `.Form.Requery` after that rebind
> **invalidates the freshly-assigned Recordset and the
> downstream CmdPajek / CmdGephi reads `.RecordCount=0`.**

That warning matches AssociationPairs's situation exactly: the
cleanup at line 2000 is the same `Set ... = OpenRecordset(...)`
shape Status uses, and adding `.Form.Requery` to it (which is
what the dict entry I added does) likely invalidates the fresh
recordset rather than helping.  A more aggressive shim
(e.g. `<subform>.Form.Recordset.MoveLast`) or per-form-action
shape change to the dict was outside the brief's authorized
scope.

The reviewer explicitly chose **option (c)** ("accept this cell
as still gap; revert; document; move on") rather than (a)
"smarter shim" or (b) "test-side split-chain workaround".

## Conclusion

- AssociationPairs × CmdGIS **stays `gap` / `blocked`** in the
  inventory.
- It is **NOT** in the same equivalence class as
  Place / Kinship for `_SUBFORMS_TO_REQUERY`.
- It should be **removed from the "cheapest next cell"
  candidate list** in the export-gap triage; future work on
  this cell would need either a per-form-action driver shim
  expansion (e.g. `MoveLast` after `Requery` for AssocPairs
  only) or a deeper investigation into why the rebind alone
  doesn't populate the subform's RecordCount.
- Neither expansion is justified on cost-benefit right now;
  AssociationPairs's three other still-blocked cells
  (CmdNeo4j, CmdUCINet) face their own separate issues that
  this driver investigation wouldn't help.

## What was preserved from this attempt

- **Nothing in code, tests, or driver.**  Working tree was
  fully reverted before this note was written.
- **Just this note**, plus a small dated mini-refresh to
  `analysis/export_gap_triage_plan.md` and
  `reports/export_gap_triage_plan.json` reclassifying
  AssociationPairs × CmdGIS from "next cheapest cell after
  Pajek/Gephi" to "blocked, not cheap, do not pursue".

## Successor recommendation

The AssociationPairs line of work has by now produced its
high-value PRs:
1. SetFocus driver/meta-PR (`feat/driver-patch-associationpairs-setfocus`)
2. CmdPajek + CmdGephi coverage (`cover/assocpairs-pajek-gephi-1x3`)
3. CmdGIS proven NOT cheap (this note)

**Cheapest-next candidates after this note: none.**

Originally one might have expected `LookAtGroupData × CmdNeo4j`
to remain the cheap-next candidate after AssocPairs × CmdGIS
was ruled out.  But that's stale too: `probe/groupdata-
cmdneo4j` (commit `4ace85b`) and `investigate/groupdata-
cmdneo4j-tail` (commit `3bfcba8`) — both already merged to
main — confirmed mid-chain `LookAtGroupData:ERR No current
record.` (DAO 3021, an unguarded `.MoveFirst` on empty
recordset).

> **At the time this note was written, a candidate issue had
> been filed separately on branch `chore/file-issue-21`
> (commit `934f220`) pending maintainer review.**  Current
> main has since canonicalized this as **Issue #21 (P1)** in
> `reports/generate_report.py::ISSUES` via the re-spun PR
> `chore/file-issue-21-v2` (commit `bc85092`, merged to
> main; the original `chore/file-issue-21` branch was
> rejected for branch-shape failure and deleted without
> merging).  Both source-side static marker
> (`tests/test_known_bugs.py::test_bug21_groupdata_cmdneo4j
> _missing_eof_guard`) and runtime behavioural pin
> (`tests/test_vba_bug_behaviors.py::test_bug21_lookat
> _groupdata_cmdneo4j_fires_no_current_record`) are in
> place.

Either way (then-candidate or now-canonical), GroupData ×
CmdNeo4j sits on the issue / investigation line, NOT on the
coverage line.

So on the current dump, under the standing brief (no
AssocPairs, no Networks driver-meta, no CmdUCINet new family,
no investigation-first cells without a fresh brief), **no
remaining export-gap cell qualifies as cheap-next.**  The
cheapest-next list is genuinely empty.

Two natural next directions (NOT framed as "pick another
cell"):

1. **GroupData CmdNeo4j tail / empty-recordset-guard
   follow-up** — Issue #21 is now canonical on main (PR
   `chore/file-issue-21-v2`, commit `bc85092`).  Next step
   is either coordinating an upstream CBDB fix (per the
   canonical issue's `fix_en` recommendation: guard the
   `.MoveFirst` in blocks #9 and #10), OR writing a
   per-block bugfix verification probe that flips the
   existing test_bug21 markers when the upstream fix
   lands.
2. **A fresh whole-triage refresh** — re-baseline the
   export-gap queue from scratch given that two cells
   (AssocPairs × CmdGIS and GroupData × CmdNeo4j) have moved
   out of the cheap-next zone since the original 2026-05-04
   plan.

Explicitly NOT recommended: "continue closing the next
cheapest cell" — there is no next cheapest cell on this
dump under the standing brief today.
