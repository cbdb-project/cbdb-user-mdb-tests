# CBDB User MDB — Issues Report

_A respectful summary of issues uncovered during regression testing._

Dear maintainer,

Below is a summary of the issues we uncovered while building an automated regression-test suite for the CBDB User MDB. We hope this report is useful as you continue your wonderful stewardship of this dataset, and we sincerely thank you for the immense work that has gone into building it.

The issues are ordered by severity (P0 highest). Each entry includes a concise description, step-by-step user reproduction, screenshots where the issue is visible in the Access UI, and a suggested fix. None of these are urgent; they are documented so they can be addressed at the maintainer's convenience.

## Table of Contents

- [P0 — Silent data corruption](#p0--silent-data-corruption)
  - [Issue #7 — LookAtPlace.CmdNeo4j people-CSV silently fails on the first record](#issue-7--lookatplacecmdneo4j-people-csv-silently-fails-on-the-first-record)
  - [Issue #8 — LookAtNetworks.CmdNeo4j people/place CSVs silently fail on the first record](#issue-8--lookatnetworkscmdneo4j-peopleplace-csvs-silently-fail-on-the-first-record)
  - [Issue #9 — LookAtEntry.CmdNeo4j Institutions block uses the wrong recordset variable](#issue-9--lookatentrycmdneo4j-institutions-block-uses-the-wrong-recordset-variable)
- [P1 — Visible runtime crash](#p1--visible-runtime-crash)
  - [Issue #6 — LookAtGroupData ChkEntry path references a non-existent column ENTRY_DATA.c_parental_status](#issue-6--lookatgroupdata-chkentry-path-references-a-non-existent-column-entry_datac_parental_status)
  - [Issue #13 — BIOG_MAIN_2 Subform tries to open a picker form (frmPickNIAN_HAO) that doesn't exist](#issue-13--biog_main_2-subform-tries-to-open-a-picker-form-frmpicknian_hao-that-doesnt-exist)
- [P2 — Silent display](#p2--silent-display)
  - [Issue #10 — EVENT_ADDR_2 Subform address columns silently render blank (wrong ControlSource)](#issue-10--event_addr_2-subform-address-columns-silently-render-blank-wrong-controlsource)
  - [Issue #11 — EVENTS_DATA_2 Subform has a control bound to a non-existent column c_event_record_id](#issue-11--events_data_2-subform-has-a-control-bound-to-a-non-existent-column-c_event_record_id)
  - [Issue #12 — POSTED_TO_OFFICE_DATA_2 Subform appointment-type control bound to wrong column name](#issue-12--posted_to_office_data_2-subform-appointment-type-control-bound-to-wrong-column-name)
- [P3 — Missing UI](#p3--missing-ui)
  - [Issue #15 — LookAtPlace is missing its CmdGIS button (handler exists but no UI control)](#issue-15--lookatplace-is-missing-its-cmdgis-button-handler-exists-but-no-ui-control)
  - [Issue #16 — LookAtStatus is missing its CmdPajek button](#issue-16--lookatstatus-is-missing-its-cmdpajek-button)
  - [Issue #17 — LookAtStatus is missing its CmdGephi button](#issue-17--lookatstatus-is-missing-its-cmdgephi-button)
  - [Issue #18 — LookAtStatus is missing its CmdUCINet button](#issue-18--lookatstatus-is-missing-its-cmducinet-button)
  - [Issue #19 — LookAtOffice is missing its CmdGUESS button](#issue-19--lookatoffice-is-missing-its-cmdguess-button)
- [P4 — Setup](#p4--setup)
  - [Issue #2 — VBA project references the legacy dao360.dll which isn't on Office 2016+ machines](#issue-2--vba-project-references-the-legacy-dao360dll-which-isnt-on-office-2016-machines)
- [P5 — Resolved / not currently reproducible](#p5--resolved--not-currently-reproducible)
  - [Issue #1 — View_StatusData would display last-year range in the first-year column — DORMANT (no source rows trigger it on this dump)](#issue-1--view_statusdata-would-display-last-year-range-in-the-first-year-column--dormant-no-source-rows-trigger-it-on-this-dump)
  - [Issue #3 — LookAtEntry.CmdQuery backfill UPDATE — historical Bug #3, NOT reproducible on the current dump](#issue-3--lookatentrycmdquery-backfill-update--historical-bug-3-not-reproducible-on-the-current-dump)
  - [Issue #4 — LookAtPlace.CmdGIS would abort with 'Object required' — LATENT, masked by Issue #15 (no CmdGIS button on the form)](#issue-4--lookatplacecmdgis-would-abort-with-object-required--latent-masked-by-issue-15-no-cmdgis-button-on-the-form)
  - [Issue #5 — LookAtStatus.CmdPajek references a missing control AND uses three columns that don't exist](#issue-5--lookatstatuscmdpajek-references-a-missing-control-and-uses-three-columns-that-dont-exist)
  - [Issue #14 — KIN_DATA Subform's CmdPickKinRel calls a missing picker (frmPickKINSHIP_CODES) — but the host sub-form is currently an orphan (LATENT)](#issue-14--kin_data-subforms-cmdpickkinrel-calls-a-missing-picker-frmpickkinship_codes--but-the-host-sub-form-is-currently-an-orphan-latent)
- [Severity legend](#severity-legend)
- [Appendix — c_index_year / c_index_addr_id drift vs the cbdb-online-main-server snapshot (not bugs)](#appendix--c_index_year--c_index_addr_id-drift-vs-the-cbdb-online-main-server-snapshot-not-bugs)
- [Closing note](#closing-note)

## Severity legend

- P0 — Silent data corruption: data is wrong or missing without an error popup.
- P1 — Visible runtime crash: a popup appears, the operation aborts.
- P2 — Silent display: form fields render blank when they should show data.
- P3 — Missing UI: a feature exists in code but no button invokes it.
- P4 — Setup: one-time hurdle on each new install.
- P5 — Resolved / not currently reproducible: kept as historical record; we re-checked on the current dump and could not trigger the symptom.

## P0 — Silent data corruption

### Issue #7 — LookAtPlace.CmdNeo4j people-CSV silently fails on the first record

**Affected sub:** `Form_LookAtPlace.CmdNeo4j_Click`

**Severity:** P0 — Silent data corruption (export silently produces nothing)

#### Description

The People-CSV section of `LookAtPlace.CmdNeo4j_Click` (line ~322 onward) builds a recordset from a SELECT that projects only four `ZZ_SCRATCH_P_TEXT` columns, but the row-write loop reads `!c_dynasty`, `!c_dynasty_chn`, and `!c_female` from that recordset. As soon as the loop hits the first row, JET raises 'Item not found in this collection'. The error handler silences it with `MsgBox`, so the user sees a single popup, then NO files are produced for any of the downstream Neo4j export steps.

#### Steps to reproduce

1. Open **LookAtPlace**.  Pick the address picker, choose a well-attested address — for example **c_addr_id = 7213** (Kaifeng 開封) — so the resulting query has plenty of people to feed the People-CSV loop.  Click **Run Query**.
2. Once the query finishes, click the **Neo4j** export button.
3. Pick a save location at the first SaveAs prompt (the 'People file' prompt).
4. A `Run-time error 3265 — Item not found in this collection` popup appears almost immediately.
5. After clicking OK, the chosen folder is empty — no Neo4j export file was written.

#### Screenshots

![bug7_step1_annotated.png](screenshots/bug7_step1_annotated.png)

![bug7_step2_faux_popup.png](screenshots/bug7_step2_faux_popup.png)

_The popup users see (re-rendered for the report; the real popup blocks the COM thread our test driver runs in)._

#### Suggested fix

Extend the SELECT in the People-CSV branch to project the fields the loop reads, e.g. `DYNASTIES.c_dynasty`, `DYNASTIES.c_dynasty_chn`, `BIOG_MAIN.c_female` (the JOINs already expose them).

### Issue #8 — LookAtNetworks.CmdNeo4j people/place CSVs silently fail on the first record

**Affected sub:** `Form_LookAtNetworks.CmdNeo4j_Click`

**Severity:** P0 — Silent data corruption

#### Description

Same shape as Issue #7 but on a different form. Two SELECTs in `LookAtNetworks.CmdNeo4j_Click` are missing fields that the row-write loop reads:

  • `tRstPlace` SELECT (line 2458) projects 3 columns; the loop reads `!x_coord` / `!y_coord` (not projected).
  • `tRstPeoplePlace` SELECT similarly omits `c_person_id` / `c_index_addr_id` that the loop reads.

Same silent-fail symptom as Issue #7.

#### Steps to reproduce

1. Open **LookAtNetworks** (note: this form has a known opening-delay issue; please allow several seconds).
2. Run a query, then click **Neo4j**.
3. When the export reaches the People-with-Place file, the same `Item not found` popup appears, and no further files are written.

#### Suggested fix

Extend each SELECT to project every field the loop reads. For tRstPlace: add `ADDR_CODES.x_coord`, `ADDR_CODES.y_coord`. For tRstPeoplePlace: add the missing `c_person_id` / `c_index_addr_id` columns from the joined tables.

### Issue #9 — LookAtEntry.CmdNeo4j Institutions block uses the wrong recordset variable

**Affected sub:** `Form_LookAtEntry.CmdNeo4j_Click`

**Severity:** P0 — Silent data corruption (export silently produces nothing)

#### Description

Line 1415 of `Form_LookAtEntry.vb` opens the institutions recordset as `tRstInstitutions = CurrentDb.OpenRecordset(tQueryStr)`. Ten lines later, line 1425 says `With tRstAssocCodes` and the loop reads `!c_inst_code`, `!c_inst_name_code`, etc. against THAT recordset — which was bound much earlier to the AssocCodes SELECT and doesn't have `c_inst_*` columns. Same `Item not found` symptom; InstitutionCodes file is never written.

Note: triggering this path requires entries with `c_inst_code > 0` (i.e. social-institution-bearing entries). Not every fixture reaches this With block.

#### Steps to reproduce

1. Open **LookAtEntry** with a query that produces entries with social institution codes (those are uncommon — most entries don't trigger this).
2. Click **Neo4j**, accept all the SaveAs dialogs.
3. When the export reaches the InstitutionCodes file, the same `Item not found` popup appears.

#### Suggested fix

Change `With tRstAssocCodes` on line 1425 to `With tRstInstitutions`. Single-character class of fix; the underlying recordset variable was simply mis-named.

## P1 — Visible runtime crash

### Issue #6 — LookAtGroupData ChkEntry path references a non-existent column ENTRY_DATA.c_parental_status

**Affected sub:** `Form_LookAtGroupData.queryEntry`

**Severity:** P1 — Visible crash on a common path (Entry sub-query)

#### Description

`Form_LookAtGroupData.vb` line 2621 has an INSERT INTO whose target column list ends with `c_parental_status_code` but whose SELECT projection ends with `ENTRY_DATA.c_parental_status` (no `_code` suffix). The actual column on `ENTRY_DATA` is `c_parental_status_code`; the typo means the SQL crashes with 'No such field' the moment the user checks **Entry** and clicks **Run**.

`Form_LookAtEntry.vb:1650` does the same logical query and uses the correct name, so this is a single-line drift.

#### Steps to reproduce

**Recommended demo person:** `c_personid=1` (安惇, An Dun)

Use person 1 (安惇, An Dun) as the import list (small: only 2 entry row, fast to reproduce). In LookAtGroupData, leave only the **Entry** checkbox ticked, click **Run**. _Picked by `reports/probe_demo_persons.py`; a SQL probe selected this person because their row counts genuinely satisfy the precondition the bug needs._

1. In **LookAtGroupData**, populate the import list with one entry — c_personid = 1 (An Dun 安惇) is enough; he has exactly 2 ENTRY_DATA rows so the broken queryEntry SQL will run on a tiny well-known sample.
2. Tick **only** the **Entry** checkbox (leave Status / Office / Text / Addr unchecked so the unrelated query branches don't fire).
3. Click **Run**.
4. A popup appears reporting that a field doesn't exist (JET reports this as 'No value given for one or more required parameters' / 'No such field' depending on the Office build — both mean the SQL referenced `ENTRY_DATA.c_parental_status` which doesn't exist).

#### Suggested fix

Change `ENTRY_DATA.c_parental_status` to `ENTRY_DATA.c_parental_status_code` on line 2621. One-line fix.

### Issue #13 — BIOG_MAIN_2 Subform tries to open a picker form (frmPickNIAN_HAO) that doesn't exist

**Affected sub:** `Form_BIOG_MAIN_2_Subform.c_fl_ey_notes_Click`

**Severity:** P1 — Visible crash on a user click

#### Description

When the user clicks the `c_fl_ey_notes` field on a person's biographical detail subform, `Sub c_fl_ey_notes_Click` runs `DoCmd.OpenForm "frmPickNIAN_HAO"`. There is no form named `frmPickNIAN_HAO` in the .mdb's CurrentProject.AllForms collection. Access raises 'Item not found …' and the field click does nothing useful for the user.

Likely cause: a picker form was renamed or consolidated in an earlier refactor, and this caller wasn't updated.

#### Steps to reproduce

**Recommended demo person:** `c_personid=5` (查籥, Zha Yue)

Open person 5 (查籥, Zha Yue). Their `c_fl_ey_notes` field has actual text in it (sample: '紹興二十一年進士。…'), so clicking it actually triggers the broken Sub. _Picked by `reports/probe_demo_persons.py`; a SQL probe selected this person because their row counts genuinely satisfy the precondition the bug needs._

1. Open the biographical detail form for **c_personid = 5 (Zha Yue 查籥)** — picked because his BIOG_MAIN row has a non-empty `c_fl_ey_notes` value, so the field is interactable (clicking an empty field doesn't fire the Sub).
2. On the BIOG_MAIN_2 subform, click the `c_fl_ey_notes` field — that fires the `c_fl_ey_notes_Click` Sub.
3. An `Item not found in this collection.` popup appears (because the Sub tries `DoCmd.OpenForm "frmPickNIAN_HAO"` and that form doesn't exist).

#### Suggested fix

Either restore the picker form `frmPickNIAN_HAO`, or update the caller in `Form_BIOG_MAIN_2_Subform.c_fl_ey_notes_Click` to use whichever picker form replaced it.

## P2 — Silent display

### Issue #10 — EVENT_ADDR_2 Subform address columns silently render blank (wrong ControlSource)

**Affected sub:** `EVENT_ADDR_2 Subform`

**Severity:** P2 — Silent display (address columns blank)

#### Description

On the EVENT_ADDR_2 sub-form (events with addresses), the two address controls are bound as follows:

  • `TxtAddrCHN`.ControlSource = `c_name_chn`
  • `TxtAddrPY`.ControlSource = `c_name`

But the form's RecordSource is the saved query `View_EventAddrData`, which aliases ADDR_CODES.c_name_chn as `c_event_addr_chn` and ADDR_CODES.c_name as `c_event_addr_name`. Neither `c_name` nor `c_name_chn` is in the projection, so both controls silently render blank for every row.

#### Steps to reproduce

**Recommended demo person:** `c_personid=44872` (孫才, Sun Cai)

Open person 44872 (孫才, Sun Cai). The EVENTS sub-datasheet shows 1 event row(s); 1 of them have an associated address. That's where the bound controls render blank on every row. _Picked by `reports/probe_demo_persons.py`; a SQL probe selected this person because their row counts genuinely satisfy the precondition the bug needs._

1. Open the biographical detail form for **c_personid = 44872 (Sun Cai 孫才)** — picked because he has both EVENTS_DATA and EVENTS_ADDR rows in a small enough quantity to inspect by eye.
2. Switch to the EVENT_ADDR sub-datasheet.
3. The Chinese address column and the Pinyin address column are blank on every row, even though the underlying ADDR_CODES rows actually have those fields populated.

#### Screenshots

![bug10_subform_annotated.png](screenshots/bug10_subform_annotated.png)

_EVENT_ADDR_2 in design view, annotated — TxtAddrCHN's ControlSource (`c_name_chn`) is not in the form's RecordSource projection._

#### Suggested fix

In the form designer, change `TxtAddrCHN`.ControlSource from `c_name_chn` to `c_event_addr_chn`, and `TxtAddrPY`.ControlSource from `c_name` to `c_event_addr_name` (the actual aliases in View_EventAddrData).

### Issue #11 — EVENTS_DATA_2 Subform has a control bound to a non-existent column c_event_record_id

**Affected sub:** `EVENTS_DATA_2 Subform`

**Severity:** P2 — Silent display (column blank)

#### Description

The EVENTS_DATA_2 sub-form has a control whose ControlSource is `c_event_record_id`. Neither the source table EVENTS_DATA nor the form's RecordSource (`View_EventsData`) has a column of that name — likely a stale design-time leftover from when the schema had an event-record id, or an intended `c_event_code` that was typo'd. The control silently shows blank for every row.

#### Steps to reproduce

**Recommended demo person:** `c_personid=44872` (孫才, Sun Cai)

Open person 44872 (孫才, Sun Cai). The EVENTS sub-datasheet shows 1 event row(s); 1 of them have an associated address. That's where the bound controls render blank on every row. _Picked by `reports/probe_demo_persons.py`; a SQL probe selected this person because their row counts genuinely satisfy the precondition the bug needs._

1. Open the biographical detail form for **c_personid = 44872 (Sun Cai 孫才)** — same person used for Issue #10 above; he has multiple EVENTS_DATA rows so the offending column renders on each.
2. Switch to the EVENTS sub-datasheet.
3. The control bound to `c_event_record_id` is blank for every row (because neither EVENTS_DATA nor View_EventsData has that column).

#### Screenshots

![bug11_subform_annotated.png](screenshots/bug11_subform_annotated.png)

_EVENTS_DATA_2 in design view, annotated._

#### Suggested fix

Decide what was intended. If the column is no longer needed, remove the control. If it should map to `c_event_code`, fix the ControlSource to that name. If the schema needs a real event-record-id column, add it to EVENTS_DATA AND project it in View_EventsData.

### Issue #12 — POSTED_TO_OFFICE_DATA_2 Subform appointment-type control bound to wrong column name

**Affected sub:** `POSTED_TO_OFFICE_DATA_2 Subform`

**Severity:** P2 — Silent display (column blank)

#### Description

The control `c_appt_type_code` on POSTED_TO_OFFICE_DATA_2 subform has ControlSource `c_appt_type_code`. The form's RecordSource (`View_PostingOfficeData`) projects `POSTED_TO_OFFICE_DATA.c_appt_code` (no `_type` infix). The control silently shows blank.

Looks like a renamed column the form designer didn't follow.

#### Steps to reproduce

**Recommended demo person:** `c_personid=2` (安邡, An Fang)

Open person 2 (安邡, An Fang). The POSTED-TO-OFFICE sub-datasheet shows 1 posting row(s) with non-null c_appt_code — yet the appointment-type column on every row is blank. _Picked by `reports/probe_demo_persons.py`; a SQL probe selected this person because their row counts genuinely satisfy the precondition the bug needs._

1. Open the biographical detail form for **c_personid = 2 (An Fang 安邡)** — picked because he has a small number of POSTED_TO_OFFICE_DATA rows, all with non-NULL `c_appt_code`, so the column we want to inspect actually has source data.
2. Switch to the POSTED_TO_OFFICE sub-datasheet.
3. The appointment-type column is blank for every row, even though c_appt_code in the source table has a real value on each.

#### Screenshots

![bug12_subform_annotated.png](screenshots/bug12_subform_annotated.png)

_POSTED_TO_OFFICE_DATA_2 in design view, annotated._

#### Suggested fix

Change the control's ControlSource from `c_appt_type_code` to `c_appt_code` (the actual column projected by View_PostingOfficeData).

## P3 — Missing UI

### Issue #15 — LookAtPlace is missing its CmdGIS button (handler exists but no UI control)

**Affected sub:** `LookAtPlace`

**Severity:** P3 — Missing UI (feature unavailable to users)

#### Description

`Form_LookAtPlace.vb` defines a fully functional `CmdGIS_Click` handler — it builds and writes a GIS .tab export, identical in shape to the GIS button on Status / Texts / Associations / Office / Kinship. But LookAtPlace's form design has NO `CmdGIS` button on it. Users on Place can use Pajek / Gephi / Neo4j export but cannot use GIS export — the handler is there, just unreachable from the UI.

#### Steps to reproduce

1. Open **LookAtPlace**.
2. Look at the export-buttons row at the bottom right.
3. There's no GIS button. Compare with LookAtStatus / LookAtAssociations / LookAtOffice etc., all of which have one.

#### Screenshots

![bug15_LookAtPlace_no_CmdGIS_annotated.png](screenshots/bug15_LookAtPlace_no_CmdGIS_annotated.png)

_LookAtPlace as it ships — no GIS button is rendered, even though `Sub CmdGIS_Click()` exists in the module._

#### Suggested fix

In LookAtPlace's form design, add a CmdGIS button next to the existing CmdPajek / CmdGephi buttons, with `OnClick = [Event Procedure]` so it invokes the existing CmdGIS_Click Sub. (Also fix Issue #4 first, otherwise the button will throw 'Object required' on the first click.)

### Issue #16 — LookAtStatus is missing its CmdPajek button

**Affected sub:** `LookAtStatus`

**Severity:** P3 — Missing UI

#### Description

Same shape as Issue #15. `Sub CmdPajek_Click()` exists in `Form_LookAtStatus.vb` (would write a Pajek .net export of the status data) but no CmdPajek button is rendered on Status's form design.

Note: even if the button is added, Issue #5 (the SQL/control defects in CmdPajek_Click itself) needs to be fixed first.

#### Steps to reproduce

1. Open **LookAtStatus**. The export-buttons row has only GIS and Neo4j; there's no Pajek button.

#### Screenshots

![bug16_LookAtStatus_no_CmdPajek_annotated.png](screenshots/bug16_LookAtStatus_no_CmdPajek_annotated.png)

#### Suggested fix

Add a CmdPajek button to LookAtStatus's design (after fixing Issue #5).

### Issue #17 — LookAtStatus is missing its CmdGephi button

**Affected sub:** `LookAtStatus`

**Severity:** P3 — Missing UI

#### Description

`Sub CmdGephi_Click()` exists in `Form_LookAtStatus.vb` but no matching button is on the form design.

#### Steps to reproduce

1. Open **LookAtStatus**. There is no Gephi export button.

#### Screenshots

![bug17_LookAtStatus_no_CmdGephi_annotated.png](screenshots/bug17_LookAtStatus_no_CmdGephi_annotated.png)

#### Suggested fix

Add a CmdGephi button to LookAtStatus's design.

### Issue #18 — LookAtStatus is missing its CmdUCINet button

**Affected sub:** `LookAtStatus`

**Severity:** P3 — Missing UI

#### Description

`Sub CmdUCINet_Click()` exists in `Form_LookAtStatus.vb` but no matching button is on the form design.

#### Steps to reproduce

1. Open **LookAtStatus**. There is no UCINet export button.

#### Screenshots

![bug18_LookAtStatus_no_CmdUCINet_annotated.png](screenshots/bug18_LookAtStatus_no_CmdUCINet_annotated.png)

#### Suggested fix

Add a CmdUCINet button to LookAtStatus's design.

### Issue #19 — LookAtOffice is missing its CmdGUESS button

**Affected sub:** `LookAtOffice`

**Severity:** P3 — Missing UI

#### Description

`Sub CmdGUESS_Click()` exists in `Form_LookAtOffice.vb` but no CmdGUESS button is on the form design. Users on Office can use GIS / GISPeople / Neo4j export but not GUESS.

#### Steps to reproduce

1. Open **LookAtOffice**. There is no GUESS export button.

#### Screenshots

![bug19_LookAtOffice_no_CmdGUESS_annotated.png](screenshots/bug19_LookAtOffice_no_CmdGUESS_annotated.png)

#### Suggested fix

Add a CmdGUESS button to LookAtOffice's design.

## P4 — Setup

### Issue #2 — VBA project references the legacy dao360.dll which isn't on Office 2016+ machines

**Affected sub:** `VBE Project References`

**Severity:** P4 — One-time setup hurdle on each new machine

#### Description

The shipped .mdb's VBA project carries a hard reference to `C:\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll`, which was the DAO 3.6 location used by Access 2003. Modern Office (2016 onward) ships `ACEDAO.DLL` instead and does NOT install the legacy DLL. On any clean modern machine, the first attempt to open any LookAt form raises 'Can't find project or library', which is opaque and scary to end users.

Severity is low because it's a one-time fix per machine, but every new install hits it.

#### Steps to reproduce

1. Install `CBDB_BJ_User.mdb` on a fresh modern Office machine.
2. Open the file. Press Alt+F11 to enter the VBE.
3. Tools → References. Notice an entry marked `MISSING: dao360.dll`.
4. Open any LookAt form. A 'Can't find project or library' error appears.

#### Suggested fix

Once, on the maintainer's machine, do:
  1. Open the .mdb in Access. Press Alt+F11.
  2. Tools → References. Untick the MISSING dao360.dll entry.
  3. Tick `Microsoft Office 16.0 Access Database Engine Object Library` (i.e. ACEDAO.DLL).
  4. Save the .mdb.

Then re-distribute the fixed file. Future end users won't need to do anything.

## P5 — Resolved / not currently reproducible

_Items in this tier are kept as historical / latent record.  They fall into three categories: (a) DORMANT — verified that current source data doesn't trigger the symptom; (b) RESOLVED — the symptom no longer occurs even though the suspect code is still present (likely fixed by some Office / JET update or a previous iteration); (c) LATENT — the source-code defect is real, but the user can't reach it because another issue (e.g. a missing UI button) blocks the path.  None of these are user-facing today; please consult before treating any of them as urgent._

### Issue #1 — View_StatusData would display last-year range in the first-year column — DORMANT (no source rows trigger it on this dump)

**Affected sub:** `View_StatusData`

**Severity:** P5 — Dormant on this dump (would be P0 if any STATUS_DATA row had both fy/ly range codes set differently)

#### Description

The saved query `View_StatusData` joins `YEAR_RANGE_CODES` twice (once aliased as `YEAR_RANGE_CODES_1` for the last-year range), but the SELECT list pulls every range field from the _1 alias. As a result, every status row displayed in the Status sub-datasheet shows the last-year range value in the first-year range column.

#### Steps to reproduce

⚠ **Currently UI-dormant on this data snapshot — see note below**

On this data snapshot the bug is **DORMANT** — STATUS_DATA has 70,761 rows, but only 13 have c_fy_range > 0 and 0 have c_ly_range > 0; 0 have both populated AND different. So no person currently surfaces the alias swap through the UI.  The SQL bug still exists; the moment a future data refresh adds a STATUS_DATA row with both fy/ly range codes set differently, the corresponding sub-datasheet line will display the wrong text.  To verify the bug today, run the SQL directly:
  SELECT c_personid, c_fy_range, c_fy_range_desc, c_ly_range, c_ly_range_desc FROM View_StatusData WHERE c_fy_range > 0 OR c_ly_range > 0;

1. Because no STATUS_DATA row in the current dump has both c_fy_range AND c_ly_range populated, this bug cannot be demonstrated through the UI today.  Verify it directly in SQL instead:
2. Open the .mdb in Access.  Press F11 to show the navigation pane, then double-click query **View_StatusData**.
3. Inspect the SELECT clause: every `c_fy_range_*` alias is pulled from `YEAR_RANGE_CODES_1`, but the FROM clause joins that alias on the LAST-year range.  That's the swap.
4. (Optional) Run `SELECT TOP 100 c_personid, c_fy_range, c_fy_range_desc, c_ly_range, c_ly_range_desc FROM View_StatusData WHERE c_fy_range > 0 OR c_ly_range > 0` in the Access query window — once a future data refresh populates both fields differently, every such row will display the wrong first-year text.

#### Suggested fix

In `View_StatusData` change `YEAR_RANGE_CODES_1.c_range AS c_fy_range_desc` and `YEAR_RANGE_CODES_1.c_range_chn AS c_fy_range_chn` to use the un-aliased `YEAR_RANGE_CODES.*` fields (which the FROM clause already joins on `c_fy_range`).

### Issue #3 — LookAtEntry.CmdQuery backfill UPDATE — historical Bug #3, NOT reproducible on the current dump

**Affected sub:** `Form_LookAtEntry.CmdQuery_Click`

**Severity:** P5 — Resolved / not reproducible on current dump

#### Description

Historical context: an earlier dump of `Form_LookAtEntry.vb` (line 1778-1789, a single UPDATE joining seven+ lookup tables to backfill `c_entry_desc` / `c_addr_name` / `c_kin_name` etc. into `ZZ_SCRATCH_ENTRY`) was reported to silently leave those columns NULL on result sets above ~30 000 rows.

**RE-VERIFIED on the current dump (2026-05-02): cannot reproduce.**  We fired CmdQuery on the same fixture (entry code 36 jinshi general, no year filter, 92,514 rows in `ZZ_SCRATCH_ENTRY`) and counted exactly **0 rows** with `c_entry_code IS NOT NULL` AND `c_entry_desc IS NULL`; same for `c_addr_id > 0` AND `c_addr_name IS NULL`.  The maintainer also confirmed the UI shows correct desc / addr columns.

The giant multi-table UPDATE statement is still in the VBA module (the source code wasn't rewritten), so structurally the SQL pattern that was suspect remains.  But its runtime behaviour now produces correct backfills on this dump — likely because of a JET / Office update changing how it schedules complex UPDATE plans, or because the original diagnosis was a false positive.

**Recommendation:** treat as resolved unless someone can produce a fresh repro on a current dump.  Verification script: `analysis/verify_bug3.py`.

#### Steps to reproduce

1. Run `python analysis/verify_bug3.py` from the repo root.
2. It opens LookAtEntry, fires CmdQuery on entry code 36 with no year filter, and reports the count of rows whose `c_entry_desc` is NULL despite a non-null `c_entry_code`.
3. On the current dump the count is 0 — the bug is no longer observable.  If a future dump regresses, this same script will report a non-zero count.

#### Suggested fix

No action required for this dump.  If a future regression is observed: split the giant multi-table UPDATE into several smaller ones — one per lookup join — matching the pattern Status / Texts / Associations already use.

### Issue #4 — LookAtPlace.CmdGIS would abort with 'Object required' — LATENT, masked by Issue #15 (no CmdGIS button on the form)

**Affected sub:** `Form_LookAtPlace.CmdGIS_Click`

**Severity:** P5 — Latent (would be P1 if Issue #15 fixed without first fixing this)

#### Description

Note: this issue is moot in the current dump because there is no CmdGIS button on LookAtPlace's design (Issue #15) — users physically cannot click it. But the underlying VBA problem remains: line 1539 of `Form_LookAtPlace.vb` reads `If GISFrame.Value = 1 Then`, and there is no control named `GISFrame` on this form (the actual encoding control is `CodeFrame`). If the missing button is ever re-added (Issue #15) without first fixing this line, every click will throw.

#### Steps to reproduce

1. (Hypothetical, after Issue #15 is fixed.) Open **LookAtPlace**.
2. Run any query.
3. Click the GIS button.
4. A `Run-time error 424 — Object required` popup appears, the export does nothing.

#### Screenshots

![bug4_step1_annotated.png](screenshots/bug4_step1_annotated.png)

![bug4_step2_annotated.png](screenshots/bug4_step2_annotated.png)

![bug4_step3_faux_popup.png](screenshots/bug4_step3_faux_popup.png)

_Re-rendered popup — exact runtime error users would see if the button were present._

#### Suggested fix

Change `GISFrame.Value` to `CodeFrame.Value` on line 1539 of `Form_LookAtPlace.vb`. Same change `CmdNeo4j_Click`, `CmdGephi_Click`, and `CmdPajek_Click` on the same form already use correctly.

### Issue #5 — LookAtStatus.CmdPajek references a missing control AND uses three columns that don't exist

**Affected sub:** `Form_LookAtStatus.CmdPajek_Click`

**Severity:** P5 — Latent (would be P1 if Issue #16 fixed without first fixing this)

#### Description

Two related defects in the same handler:

  (a) Line 2308 reads `If ChkIDs.Value Then`, but Status has no control named `ChkIDs` — only `ChkXYRef`, `ChkKML`, and `ChkSubUnits`.

  (b) Lines 2335–2338 build a SELECT that references `ZZ_SCRATCH_STATUS.c_person_id`, `c_status_id`, and `c_status_count` — none of which exist in the schema (the real columns are `c_personid`, `c_status_code`, no count column at all).

The whole sub looks copy-pasted from `LookAtAssociations.CmdPajek_Click` where these names ARE valid; the rename pass missed both spots. Like Issue #4 this is also somewhat moot because LookAtStatus has no Pajek button (Issue #16); the SQL still fails the moment the sub is invoked though, so adding the button without fixing the SQL would just expose the failure to users.

#### Steps to reproduce

1. (Hypothetical, after Issue #16 is fixed.) Open **LookAtStatus**.
2. Run a query, then click the Pajek button.
3. First: an `Object required` popup appears (the ChkIDs reference).
4. If that's worked around, the next click hits the SQL: a `No such field` error from the SELECT that references three non-existent columns.

#### Suggested fix

Two fixes:
  (a) Replace `ChkIDs.Value` with either a constant `False` (if the optional behaviour isn't needed) or add a real ChkIDs control to LookAtStatus's design.
  (b) Rewrite the SELECT to use `ZZ_SCRATCH_STATUS.c_personid` and `ZZ_SCRATCH_STATUS.c_status_code`, and either drop the count aggregate or compute it some other way (the source table doesn't have `c_status_count`).

Realistically the whole sub probably needs a thoughtful rewrite rather than spot fixes — it was clearly inherited from another form without verification.

### Issue #14 — KIN_DATA Subform's CmdPickKinRel calls a missing picker (frmPickKINSHIP_CODES) — but the host sub-form is currently an orphan (LATENT)

**Affected sub:** `Form_KIN_DATA_Subform`

**Severity:** P5 — Latent (would be P1 if `KIN_DATA Subform` were re-embedded somewhere users can reach)

#### Description

**Static defect is real, runtime trigger is not currently reachable.** The Sub `CmdPickKinRel_Click` in `Form_KIN_DATA_Subform` (line 52) calls `DoCmd.OpenForm "frmPickKINSHIP_CODES"` and references `Forms!frmPickKINSHIP_CODES!frmKINSHIP_CODES.Form!c_kincode`. Neither form exists in the current .mdb — same shape as Issue #13.

**Why LATENT.** The host sub-form `KIN_DATA Subform` (which owns the `CmdPickKinRel` button) is not contained by any active form in the current `control_inventory.json`. `BIOG_MAIN_2_Subform` (the kinship surface users actually navigate to via CBDB_Browser_2) embeds `KIN_DATA_2 Subform` instead — and that variant has no `CmdPickKinRel` button (its 15 controls are all read-only fields). The only place the embedding still appears is `Form__TMPCLP487951` (a design-time backup snapshot, not a navigable form).

Because no user-facing navigation reaches the picker button, users cannot trigger the popup from normal use. The latent code path will resurface the moment a developer re-embeds `KIN_DATA Subform` somewhere reachable, so the underlying fix is still worth applying.

#### Steps to reproduce

**Recommended demo person:** `c_personid=1` (安惇, An Dun)

Open person 1 (安惇, An Dun). The KIN_DATA sub-datasheet shows 5 kinship row(s) — click any one's kinship-code picker to trigger the broken Sub. _Picked by `reports/probe_demo_persons.py`; a SQL probe selected this person because their row counts genuinely satisfy the precondition the bug needs._

1. Verification path is **static-only** — the runtime click cannot be reproduced in the current .mdb because no parent form embeds the affected sub-form.
2. Static evidence (1): open `analysis/dump/vba/Form_KIN_DATA_Subform.vb` line 52 — confirms the Sub calls `DoCmd.OpenForm "frmPickKINSHIP_CODES"`.
3. Static evidence (2): open `analysis/dump/control_inventory.json` and search for `"frmPickKINSHIP_CODES"` as a key — absent. The picker form does not exist.
4. Reachability evidence: in the same JSON, search for `"KIN_DATA Subform"` as a `source_object` or sub-form control name — only `Form__TMPCLP487951` (a design backup) references it. `BIOG_MAIN_2_Subform` embeds `KIN_DATA_2 Subform` instead, which has no `CmdPickKinRel` button.

#### Suggested fix

Same as Issue #13: restore the picker form (or update the caller to its replacement). Even though the runtime path is not currently reachable, the static defect should be cleaned up so it doesn't resurface when `KIN_DATA Subform` is re-embedded.

## Appendix — c_index_year / c_index_addr_id drift vs the cbdb-online-main-server snapshot (not bugs)

When we compare BIOG_MAIN's `c_index_year` and `c_index_addr_id` between this User MDB and the weekly cbdb-online-main-server SQLite snapshot, a small fraction of persons disagree. We want to be very clear that these are NOT regressions — both pipelines run the same `IndexYearRebuildService.php` algorithm, but on different snapshots of source data and with different downstream decisions.

### Examples where only c_index_year disagrees

**`c_personid = 3501` — 李孝稱 (Li Xiaocheng)**

| Field | User MDB | cbdb-online-main-server snapshot |
|---|---|---|
| `c_index_year` | 1018 | 1028 |
| `c_index_addr_id` | 100658 | 100658 |
| `c_birthyear` | 0 | 0 |
| `c_deathyear` | 0 | 0 |
| `c_index_year_type_code` | 19 | 1912 |
| `c_index_year_source_id` | 19149 | 3479 |

**`c_personid = 16266` — 錢孟回 (Qian Menghui)**

| Field | User MDB | cbdb-online-main-server snapshot |
|---|---|---|
| `c_index_year` | 1004 | 992 |
| `c_index_addr_id` | 12723 | 12723 |
| `c_birthyear` | 0 | 0 |
| `c_deathyear` | 0 | 0 |
| `c_index_year_type_code` | 13 | 0512 |
| `c_index_year_source_id` | 700103 | 3035 |

**`c_personid = 16267` — 錢知雄 (Qian Zhixiong)**

| Field | User MDB | cbdb-online-main-server snapshot |
|---|---|---|
| `c_index_year` | 1034 | 1071 |
| `c_index_addr_id` | 12723 | 12723 |
| `c_birthyear` | 0 | 0 |
| `c_deathyear` | 0 | 0 |
| `c_index_year_type_code` | 1312 | 14 |
| `c_index_year_source_id` | 16266 | 16269 |

### Examples where only c_index_addr_id disagrees

**`c_personid = 1` — 安惇 (An Dun)**

| Field | User MDB | cbdb-online-main-server snapshot |
|---|---|---|
| `c_index_year` | 1042 | 1042 |
| `c_index_addr_id` | 101117 |  |
| `c_birthyear` | 1042 | 1042 |
| `c_deathyear` | 1104 | 1104 |
| `c_index_year_type_code` | 01 | 01 |
| `c_index_year_source_id` |  |  |

**`c_personid = 470` — 金君卿 (Jin Junqing)**

| Field | User MDB | cbdb-online-main-server snapshot |
|---|---|---|
| `c_index_year` | 1012 | 1012 |
| `c_index_addr_id` | 12879 | 12785 |
| `c_birthyear` | 0 | 0 |
| `c_deathyear` | 0 | 0 |
| `c_index_year_type_code` | 05 | 05 |
| `c_index_year_source_id` |  |  |

**`c_personid = 481` — 周秩 (Zhou Zhi)**

| Field | User MDB | cbdb-online-main-server snapshot |
|---|---|---|
| `c_index_year` | 1043 | 1043 |
| `c_index_addr_id` | 100416 | 12785 |
| `c_birthyear` | 0 | 0 |
| `c_deathyear` | 0 | 0 |
| `c_index_year_type_code` | 05 | 05 |
| `c_index_year_source_id` |  |  |

**`c_personid = 485` — 周穜 (Zhou Tong)**

| Field | User MDB | cbdb-online-main-server snapshot |
|---|---|---|
| `c_index_year` | 1046 | 1046 |
| `c_index_addr_id` | 100416 | 12785 |
| `c_birthyear` | 0 | 0 |
| `c_deathyear` | 0 | 0 |
| `c_index_year_type_code` | 05 | 05 |
| `c_index_year_source_id` |  |  |

**`c_personid = 562` — 范沖 (Fan Chong)**

| Field | User MDB | cbdb-online-main-server snapshot |
|---|---|---|
| `c_index_year` | 1067 | 1067 |
| `c_index_addr_id` | 100658 | 13292 |
| `c_birthyear` | 1067 | 1067 |
| `c_deathyear` | 1141 | 1141 |
| `c_index_year_type_code` | 01 | 01 |
| `c_index_year_source_id` |  |  |

### Examples where the SOURCE data itself differs (birthyear / deathyear)

**`c_personid = 1455` — 沈邈 (Shen Miao)**

| Field | User MDB | cbdb-online-main-server snapshot |
|---|---|---|
| `c_index_year` | 1001 | 1008 |
| `c_index_addr_id` | 12887 | 12887 |
| `c_birthyear` | 1001 | 0 |
| `c_deathyear` | 1047 | 0 |
| `c_index_year_type_code` | 01 | 05 |
| `c_index_year_source_id` |  |  |

**`c_personid = 19149` — 李孝基 (Li Xiaoji)**

| Field | User MDB | cbdb-online-main-server snapshot |
|---|---|---|
| `c_index_year` | 1016 | 1026 |
| `c_index_addr_id` | 100658 | 100658 |
| `c_birthyear` | 1016 | 0 |
| `c_deathyear` | 1076 | 0 |
| `c_index_year_type_code` | 01 | 11 |
| `c_index_year_source_id` |  | 41030 |

## Closing note

Thank you for taking the time to read this report. None of the items above is urgent; we hope having them all in one place makes it easy to address them at your own pace.

If any of the descriptions or suggested fixes are unclear, we would be glad to discuss further. The corresponding regression tests in this repository will automatically flip from PASS to FAIL the moment any issue is fixed in the source dump — so you can use them as a confirmation signal.
