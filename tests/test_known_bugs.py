"""
Regression tests for known bugs in CBDB_BJ_User.mdb (current as of 2026-04-30).

Each test here documents a discrete bug and verifies the buggy behaviour.
When the marker stops reproducing, the corresponding test will start
FAILING — that's the signal to investigate **why** before flipping the
assertion: the candidates are (a) upstream actually patched the source
.mdb / VBA, (b) the input fixture or Access driver behaviour changed
out from under the test, (c) the original bug was misclassified.  Only
(a) justifies marking the issue as fixed in `reports/generate_report.py`'s
`ISSUES` dict, and only after inspecting the new VBA / queries dump
or hearing from the maintainer.  The failure messages below
deliberately use "no longer reproduces" rather than "FIXED".

Bugs documented:
  1. View_StatusData: c_fy_range_desc / c_fy_range_chn pull from the
     YEAR_RANGE_CODES_1 alias, which the FROM clause joined on
     c_ly_range — so the displayed first-year range value is actually
     the last-year range value.
  2. CBDB_BJ_User.mdb ships with a broken VBA reference to
     C:\\Program Files\\Common Files\\Microsoft Shared\\DAO\\dao360.dll
     which causes "Can't find project or library" errors when forms
     try to open on machines without legacy DAO 3.6 installed.
"""
from __future__ import annotations

import json
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent


def test_bug_view_statusdata_fy_alias_swap():
    """View_StatusData has c_fy_range_desc / c_fy_range_chn pulled from
    the wrong YEAR_RANGE_CODES alias."""
    queries = json.loads(
        (REPO / "analysis/dump/queries.json").read_text(encoding="utf-8")
    )
    sql = next(q["sql"] for q in queries if q["name"] == "View_StatusData")

    # The alias YEAR_RANGE_CODES_1 was joined for c_ly_range:
    assert "AS YEAR_RANGE_CODES_1 ON STATUS_DATA.c_ly_range" in sql
    # but is incorrectly used for c_fy_range_desc/c_fy_range_chn:
    assert "YEAR_RANGE_CODES_1.c_range AS c_fy_range_desc" in sql, (
        "Bug marker no longer reproduces (investigate upstream fix vs. fixture/driver change vs. misclassification before flipping) — first-year range now pulls from the "
        "correct alias. Update this test to assert the corrected SQL."
    )
    assert "YEAR_RANGE_CODES_1.c_range_chn AS c_fy_range_chn" in sql


def test_bug_view_statusdata_fy_value_equals_ly_value(ro_conn):
    """As long as the bug above stands, every row in View_StatusData where
    c_fy_range and c_ly_range differ should display c_fy_range_desc ==
    c_ly_range_desc (i.e. the BUG is observable in actual data)."""
    cur = ro_conn.cursor()
    cur.execute("""
        SELECT TOP 50 c_personid, c_fy_range_desc, c_ly_range_desc
        FROM View_StatusData
        WHERE c_fy_range_desc IS NOT NULL OR c_ly_range_desc IS NOT NULL
    """)
    rows = cur.fetchall()
    cur.close()
    if not rows:
        pytest.skip("no STATUS_DATA rows with range descriptions")
    # All rows should have fy_desc == ly_desc (because both pull from
    # the same source column in the buggy SQL).
    mismatches = [r for r in rows if r.c_fy_range_desc != r.c_ly_range_desc]
    assert not mismatches, (
        "Bug marker no longer reproduces (investigate upstream fix vs. fixture/driver change vs. misclassification before flipping) — found rows where c_fy_range_desc != "
        "c_ly_range_desc. Update this test to remove the regression "
        f"check. ({len(mismatches)} rows differ)"
    )


# NOTE: There used to be a `test_bug3_lookat_entry_cmdquery_backfill_*`
# here, asserting the giant multi-table UPDATE in
# Form_LookAtEntry.CmdQuery_Click was still structurally present.
# Bug #3 was removed from the documented ISSUES set on 2026-05-03 —
# see PR E.  Re-verification on the current dump found 0 NULL
# backfills out of 92,514 rows on the original fixture, AND there's
# no upstream source-level fix to point at, so we treated the
# original report as a false positive (testing infrastructure /
# fixture / driver) rather than a CBDB-maintainer bug.  The
# structural marker test would only have asserted "the legacy SQL
# pattern is still here", which is not a useful guard now that the
# issue isn't claimed.


def test_bug4_lookat_place_cmdgis_references_nonexistent_gisframe():
    """Bug #4 — Form_LookAtPlace.CmdGIS_Click references `GISFrame.
    Value` but the form has no `GISFrame` control (only `CodeFrame`).
    The driver works around this via _PER_FORM_CMDGIS_PATCHES.
    """
    vba_path = (REPO / "analysis" / "dump" / "vba"
                / "Form_LookAtPlace.vb")
    body = vba_path.read_bytes().decode("utf-8")
    assert "GISFrame.Value" in body, (
        "Bug #4 marker no longer reproduces (investigate upstream fix vs. fixture/driver change vs. misclassification before flipping) — `GISFrame.Value` no longer "
        "appears in Form_LookAtPlace.vb.  Drop the workaround in "
        "tests/cbdb_driver/vba_session.py::_PER_FORM_CMDGIS_PATCHES "
        "and flip this assertion."
    )


def test_bug5_lookat_status_cmdpajek_references_nonexistent_chkids():
    """Bug #5 — Form_LookAtStatus.CmdPajek_Click references
    `ChkIDs.Value` but Status has no `ChkIDs` control.  AND the SQL
    inside CmdPajek references three columns that don't exist on
    ZZ_SCRATCH_STATUS (`c_person_id`, `c_status_id`, `c_status_count`).
    """
    vba_path = (REPO / "analysis" / "dump" / "vba"
                / "Form_LookAtStatus.vb")
    body = vba_path.read_bytes().decode("utf-8")
    assert "ChkIDs.Value" in body, (
        "Bug #5 (ChkIDs side) marker no longer reproduces "
        "(investigate upstream fix vs. fixture/driver change vs. "
        "misclassification before flipping) — `ChkIDs.Value` is "
        "gone from Form_LookAtStatus."
    )
    # The SQL bug is independent — flip when both halves clear.
    assert "ZZ_SCRATCH_STATUS.c_status_count" in body, (
        "Bug #5 (SQL side) marker no longer reproduces (investigate upstream fix vs. fixture/driver change vs. misclassification before flipping) — the ZZ_SCRATCH_STATUS"
        ".c_status_count reference is gone.  The CmdPajek SQL was "
        "copy-pasted from LookAtAssociations and the column names "
        "weren't updated to match Status's schema.  Real fix needs a "
        "rewrite, not just a workaround."
    )


def test_bug8_lookat_networks_cmdneo4j_select_missing_xy():
    """Bug #8 — Form_LookAtNetworks.CmdNeo4j_Click builds tRstPlace
    from a SELECT projecting only c_index_addr_id / c_index_addr_name
    / c_index_addr_chn, but the loop reads `!x_coord` / `!y_coord`.
    Same pattern as Bug #7 (LookAtPlace.CmdNeo4j) and Bug #9
    (LookAtEntry.CmdNeo4j) — the CmdNeo4j family-of-three.
    """
    vba_path = (REPO / "analysis" / "dump" / "vba"
                / "Form_LookAtNetworks.vb")
    body = vba_path.read_bytes().decode("utf-8")
    # The buggy SELECT for the People-Place file has 3 cols and joins
    # ADDR_CODES but doesn't project x_coord / y_coord.
    buggy_select = (
        "SELECT DISTINCT BIOG_MAIN.c_index_addr_id, "
        "ADDR_CODES.c_name AS c_index_addr_name, "
        "BIOG_MAIN.c_name_chn AS c_index_addr_chn"
    )
    assert buggy_select in body, (
        "Bug #8 marker no longer reproduces (investigate upstream fix vs. fixture/driver change vs. misclassification before flipping) — the LookAtNetworks.CmdNeo4j "
        "tRstPlace SELECT no longer matches the buggy 3-col form. "
        "Flip this assertion."
    )


def test_bug6_groupdata_query_entry_wrong_field():
    """Bug #6 (reports/CBDB_Issues_Report_EN.md) — Form_LookAtGroupData.queryEntry projects
    `ENTRY_DATA.c_parental_status` (no `_code` suffix), but the actual
    schema column is `c_parental_status_code`.  The INSERT...SELECT
    statement must reference `ENTRY_DATA.c_parental_status_code` to be
    valid.  Until CBDB fixes this, the SQL string in the dump should
    contain the buggy form.  When fixed, this assertion fires.
    """
    vba_path = (REPO / "analysis" / "dump" / "vba"
                / "Form_LookAtGroupData.vb")
    body = vba_path.read_bytes().decode("utf-8")
    assert "ENTRY_DATA.c_parental_status " in body, (
        "Bug #6 marker no longer reproduces (investigate upstream fix vs. fixture/driver change vs. misclassification before flipping) — `ENTRY_DATA.c_parental_status` "
        "no longer appears in Form_LookAtGroupData.vb (likely renamed "
        "to `c_parental_status_code`).  Update this test to assert the "
        "corrected form."
    )
    # The corrected form should be `c_parental_status_code`.  In the
    # buggy state both the INSERT col list AND the SELECT projection
    # mention `c_parental_status_code` once each (target-col mention
    # is fine; the bug is the SELECT side missing the suffix).
    n_buggy = body.count("ENTRY_DATA.c_parental_status ")
    assert n_buggy >= 1, (
        f"expected at least 1 buggy `ENTRY_DATA.c_parental_status ` "
        f"reference; found {n_buggy}"
    )



def test_bug21_groupdata_cmdneo4j_missing_eof_guard():
    """Bug #21 (reports/CBDB_Issues_Report_EN.md) —
    `Form_LookAtGroupData.CmdNeo4j_Click` opens a recordset on
    `ZZ_SCRATCH_ENTRY` (block #9 PeopleEntry, ~line 1245) and
    immediately calls `.MoveFirst` without first checking `.EOF`.
    When `ZZ_SCRATCH_ENTRY` is empty (a normal user state — the
    queried person has no Entry data, OR ChkEntry was left
    unticked), `.MoveFirst` raises DAO 3021 'No current record'
    and the entire Neo4j export chain aborts mid-way.

    Distinct from Issue #6 — Issue #6 is a column-typo upstream
    in `queryEntry` that prevents `ZZ_SCRATCH_ENTRY` from being
    populated at all when ChkEntry is on.  Issue #21 is the
    independent downstream missing-guard bug that fires whenever
    `ZZ_SCRATCH_ENTRY` is empty for any reason.

    Robust regression marker: anchor on the
    `OpenRecordset("ZZ_SCRATCH_ENTRY"` call (which is unique in
    the form's VBA body to the PeopleEntry block), find the
    immediately-following `.MoveFirst`, and assert that the
    window between them contains NO `.EOF` / `.RecordCount`
    guard pattern.  When CBDB fixes this — by inserting any of
    `If Not .EOF Then`, `If .RecordCount > 0 Then`, etc. between
    OpenRecordset and .MoveFirst — this test will fire and
    require a maintainer to update both the test and Issue #21
    in `reports/generate_report.py`.
    """
    vba_path = (REPO / "analysis" / "dump" / "vba"
                / "Form_LookAtGroupData.vb")
    body = vba_path.read_bytes().decode("utf-8")

    # Anchor 1: the OpenRecordset call on ZZ_SCRATCH_ENTRY is
    # unique to block #9 PeopleEntry in the CmdNeo4j_Click body.
    open_marker = 'CurrentDb.OpenRecordset("ZZ_SCRATCH_ENTRY"'
    open_idx = body.find(open_marker)
    assert open_idx >= 0, (
        "Bug #21 marker no longer reproduces (investigate "
        "upstream fix vs. fixture/driver change vs. "
        "misclassification before flipping) — anchor "
        f"`{open_marker}` not found in Form_LookAtGroupData.vb. "
        "The PeopleEntry block in CmdNeo4j_Click was likely "
        "refactored.  Update this test to assert the corrected "
        "form, and update Issue #21 in "
        "reports/generate_report.py accordingly."
    )

    # Anchor 2: the .MoveFirst that follows the OpenRecordset.
    # In the buggy form these are ~2 lines apart (with a `With
    # tRstPeopleEntry` between).  Cap the search to 800 chars so
    # we don't accidentally pick up a different block's
    # .MoveFirst if the PeopleEntry block is removed but later
    # blocks remain.
    move_idx = body.find(".MoveFirst", open_idx)
    assert 0 <= move_idx - open_idx <= 800, (
        "Bug #21 marker no longer reproduces — could not locate "
        "the PeopleEntry block's .MoveFirst within 800 chars "
        f"after `{open_marker}`.  The block was likely "
        "refactored.  Investigate upstream fix vs. fixture/"
        "driver change vs. misclassification before flipping; "
        "update Issue #21 in reports/generate_report.py."
    )

    # The window between OpenRecordset and .MoveFirst is where
    # any reasonable .EOF / .RecordCount guard would live.
    # Enumerate the guard patterns a fix would plausibly use.
    window = body[open_idx:move_idx]
    fix_patterns = (
        "If Not .EOF",
        "If Not tRstPeopleEntry.EOF",
        "If .EOF Then",
        "If tRstPeopleEntry.EOF Then",
        ".RecordCount >",
        ".RecordCount <>",
        "RecordCount > 0",
        "RecordCount <> 0",
    )
    found_guards = [p for p in fix_patterns if p in window]
    assert not found_guards, (
        "Bug #21 marker no longer reproduces (investigate "
        "upstream fix vs. fixture/driver change vs. "
        "misclassification before flipping) — found protective "
        f"guard pattern(s) {found_guards} between "
        f"`{open_marker}` and the immediately-following "
        ".MoveFirst.  This suggests the PeopleEntry block now "
        "guards against an empty ZZ_SCRATCH_ENTRY (Issue #21 "
        "was fixed upstream).  Update this test to assert the "
        "corrected form, and update Issue #21 in "
        "reports/generate_report.py accordingly.  Note: per "
        "Issue #21's fix recommendation, a complete fix should "
        "add the same guard to ALL 11 blocks of "
        "CmdNeo4j_Click; consider extending this test to also "
        "verify the EntryCode block (around line 1385)."
    )


def test_bug22_associations_cmducinet_createtextfile_no_unicode_arg():
    """Bug #22 (reports/CBDB_Issues_Report_EN.md) —
    `Form_LookAtAssociations.CmdUCINet_Click` writes the .vna
    export via `Scripting.FileSystemObject.CreateTextFile
    (tFileName, True)` (line ~2575) WITHOUT the 3rd Unicode
    argument.  The 3rd arg defaults to FALSE → file opens in
    cp1252 / system ANSI → `tVNA.WriteLine` raises VBA error
    5 ('Invalid procedure call or argument') when `c_name`
    contains a non-cp1252 character with no FSO substitution
    (CJK Han ideographs in particular).  The export aborts
    mid-`*node properties`; the partial .vna file on disk is
    unusable.

    Sister test in `tests/test_vba_bug_behaviors.py::test_bug22
    _associations_cmducinet_fires_invalid_procedure_call` pins
    the runtime symptom on the verified person-437 fixture.
    Don't replace the static test — both source-side and
    runtime-side guards should remain.

    Robust regression marker: locate the CmdUCINet_Click
    body's `CreateTextFile` call and assert it has only 2
    arguments (no Unicode flag).  When CBDB upstream fixes
    this — by adding `, True` as the 3rd argument — this
    test will fire and require a maintainer to update both
    the test and Issue #22 in `reports/generate_report.py`.

    **Extended scope (Kinship sibling form):** As of the
    Kinship sibling-risk probe (commit 154bb4b on main),
    `Form_LookAtKinship.CmdUCINet_Click` is **runtime-
    confirmed** to have the SAME bug-family pattern at line
    ~2510.  This marker now asserts the buggy 2-arg
    `CreateTextFile(tFileName, True)` pattern is present in
    BOTH `Form_LookAtAssociations.vb::CmdUCINet_Click` AND
    `Form_LookAtKinship.vb::CmdUCINet_Click`, and the fixed
    3-arg form (with Unicode = True) is in NEITHER.  An
    upstream fix should add the 3rd arg to BOTH call sites
    in the same patch (per Issue #22's fix recommendation).

    The existing Kinship × CmdUCINet coverage test
    (`tests/test_vba_cmducinet_kinship.py`) is now KNOWN
    fixture-fragile — it passes only because matrix-
    supplied person 3211's kin network happens to contain
    no Han-character c_name values.  Documented in that
    test's docstring + the inventory manifest's notes.
    """
    # Both forms share the same 2-arg CreateTextFile pattern
    # inside their respective CmdUCINet_Click subs.  Iterate
    # both and apply identical assertions.  When CBDB
    # upstream fixes either / both — by adding `, True` as
    # the 3rd argument — this test fires per-form and
    # requires a maintainer to update both the test and
    # Issue #22 in reports/generate_report.py.
    forms = [
        ("Form_LookAtAssociations", "Associations"),
        ("Form_LookAtKinship", "Kinship"),
    ]
    create_call_2arg = (
        'Set tVNA = tFileSystem.CreateTextFile(tFileName, True)'
    )
    create_call_3arg = (
        'Set tVNA = tFileSystem.CreateTextFile(tFileName, '
        'True, True)'
    )

    for form_module, form_label in forms:
        vba_path = (REPO / "analysis" / "dump" / "vba"
                    / f"{form_module}.vb")
        body = vba_path.read_bytes().decode("utf-8")

        # Locate the CmdUCINet_Click sub for this form.
        sub_marker = "Private Sub CmdUCINet_Click()"
        sub_start = body.find(sub_marker)
        assert sub_start >= 0, (
            "Bug #22 marker no longer reproduces "
            f"({form_label} side) (investigate upstream fix "
            "vs. fixture/driver change vs. misclassification "
            "before flipping) — `Private Sub "
            f"CmdUCINet_Click()` not found in "
            f"{form_module}.vb.  The whole CmdUCINet "
            "handler may have been refactored or removed."
        )
        sub_end = body.find("End Sub", sub_start)
        assert sub_end > sub_start, (
            f"couldn't find End Sub for "
            f"{form_label}.CmdUCINet_Click"
        )
        sub_body = body[sub_start:sub_end]

        # Positive check: the buggy 2-arg call is present.
        assert create_call_2arg in sub_body, (
            "Bug #22 marker no longer reproduces "
            f"({form_label} side) (investigate upstream fix "
            "vs. fixture/driver change vs. misclassification "
            f"before flipping) — the buggy 2-argument call "
            f"`{create_call_2arg}` is no longer present in "
            f"`{form_module}.CmdUCINet_Click`.  The fix "
            "likely added a 3rd arg "
            "(`CreateTextFile(tFileName, True, True)` for "
            "Unicode/UTF-16LE) — flip this assertion for "
            f"{form_label}, AND check whether the other "
            "form was also fixed (a single upstream patch "
            "should fix both per Issue #22's fix_en), AND "
            "if both are fixed flip the runtime pin in "
            "tests/test_vba_bug_behaviors.py + update Issue "
            "#22 in reports/generate_report.py."
        )

        # Negative check: the corrected 3-arg form should
        # NOT yet appear in this sub.
        assert create_call_3arg not in sub_body, (
            "Bug #22 marker partially reproduces "
            f"({form_label} side) but the 3-arg Unicode "
            "form `CreateTextFile(tFileName, True, True)` "
            "is ALSO present.  This is contradictory — "
            "investigate before flipping (perhaps two "
            "CreateTextFile calls now, one fixed and one "
            f"not).  Update Issue #22 to describe the new "
            f"state in {form_label}."
        )


def test_bug7_lookat_place_cmdneo4j_select_missing_dynasty_female():
    """Bug #7 (reports/CBDB_Issues_Report_EN.md) — Form_LookAtPlace.CmdNeo4j builds a
    People-CSV from a SELECT that omits c_dynasty / c_dynasty_chn /
    c_female, but the loop reads them.  Detected by
    `analysis/audit_recordset_sql_projection.py`.  This regression
    test just confirms the buggy SQL string still matches the audit's
    expectation."""
    vba_path = (REPO / "analysis" / "dump" / "vba"
                / "Form_LookAtPlace.vb")
    body = vba_path.read_bytes().decode("utf-8")
    # The buggy SELECT has exactly 4 ZZ_SCRATCH_P_TEXT.* columns and
    # no DYNASTIES / BIOG_MAIN columns in the projection (they're
    # joined but not selected).
    buggy_select = (
        "SELECT DISTINCT ZZ_SCRATCH_P_TEXT.c_person_id, "
        "ZZ_SCRATCH_P_TEXT.c_name, ZZ_SCRATCH_P_TEXT.c_name_chn, "
        "ZZ_SCRATCH_P_TEXT.c_index_year"
    )
    assert buggy_select in body, (
        "Bug #7 marker no longer reproduces (investigate upstream fix vs. fixture/driver change vs. misclassification before flipping) — the LookAtPlace.CmdNeo4j SELECT "
        "no longer matches the buggy 4-col projection.  Re-run "
        "`analysis/audit_recordset_sql_projection.py` to confirm; if "
        "clean, flip this assertion."
    )


def test_bug9_lookat_entry_cmdneo4j_with_wrong_var():
    """Bug #9 — Form_LookAtEntry.CmdNeo4j has `With tRstAssocCodes`
    where `With tRstInstitutions` was intended.

    Reclassified 2026-05-04 from P0 to P5 latent: the typo on line
    1425 is real, but the entire SaveAs prompt + buggy `With` block
    sit inside the gate `If tRecDeleted > 0 Then` at line 1389.
    `tRecDeleted` is the row count of an
    `INSERT INTO ZZ_SCRATCH_P_TEXT … WHERE
    ZZ_SCRATCH_ENTRY.c_inst_code > 0`, and on the current dump
    0 of 263,454 ENTRY_DATA rows have c_inst_code > 0 — so the
    branch is unreachable from any LookAtEntry fixture.

    This test is therefore a SOURCE-level marker only: it pins
    that the typo is still in the dumped VBA AND that the gating
    condition still holds (so the issue stays correctly classified
    as latent).  If the gate condition flips (some future MDB drop
    introduces a c_inst_code > 0 row), the LATENT-gate assertion
    below will fail and force the maintainer to re-promote Issue #9
    back to P1.
    """
    vba_path = (REPO / "analysis" / "dump" / "vba"
                / "Form_LookAtEntry.vb")
    body = vba_path.read_bytes().decode("utf-8")
    # The pattern: a `Set tRstInstitutions = ...` line, then a few
    # lines later `With tRstAssocCodes` followed by `!c_inst_code`
    # within a few lines.  Loose check via two substrings being close.
    set_idx = body.find("Set tRstInstitutions = CurrentDb.OpenRecordset")
    if set_idx < 0:
        pytest.fail("setup line `Set tRstInstitutions = ...` missing — "
                    "bug context changed; review.")
    tail = body[set_idx:set_idx + 3000]
    assert ("With tRstAssocCodes" in tail
            and "!c_inst_code" in tail), (
        "Bug #9 source-level marker no longer reproduces "
        "(investigate upstream fix vs. fixture/driver change vs. "
        "misclassification before flipping) — `With tRstAssocCodes` "
        "immediately following the tRstInstitutions OpenRecordset is "
        "no longer present (likely renamed to `With "
        "tRstInstitutions`).  Flip this assertion AND remove the "
        "Issue #9 entry from generate_report.py."
    )

    # Also pin the gate: ENTRY_DATA.c_inst_code must still be 0
    # on the current dump for Issue #9 to remain LATENT.  If this
    # fails, the maintainer must re-promote Issue #9 to P1 and
    # restore behavioural / popup coverage.
    user_mdb = REPO / "data" / "CBDB_BJ_User.mdb"
    if not user_mdb.exists():
        pytest.skip(f"{user_mdb} not present — gate-check skipped")
    import pyodbc
    conn = pyodbc.connect(
        "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={user_mdb};", autocommit=True, readonly=True)
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM ENTRY_DATA "
                "WHERE c_inst_code > 0")
    n_inst = int(cur.fetchone()[0])
    cur.execute("SELECT COUNT(*) FROM ENTRY_DATA "
                "WHERE c_inst_name_code > 0")
    n_inst_name = int(cur.fetchone()[0])
    assert n_inst == 0 and n_inst_name == 0, (
        f"Bug #9 LATENT-gate flipped: ENTRY_DATA rows with "
        f"c_inst_code > 0 = {n_inst}, c_inst_name_code > 0 = "
        f"{n_inst_name}.  The InstitutionCodes branch is no longer "
        f"unreachable.  Re-promote Issue #9 to P1 in "
        f"reports/generate_report.py and restore behavioural / "
        f"popup coverage; the source-level typo at line 1425 is "
        f"now user-visible."
    )


def test_bug10_event_addr_subform_uses_unaliased_columns():
    """Bug #10 (candidate, found by audit_subform_control_sources):
    EVENT_ADDR_2 Subform's TxtAddrCHN/TxtAddrPY controls have
    ControlSource `c_name_chn` / `c_name`, but its RecordSource
    View_EventAddrData aliases ADDR_CODES.c_name AS c_event_addr_name
    and ADDR_CODES.c_name_chn AS c_event_addr_chn.  The address
    columns silently render blank in the events-with-address sub.
    """
    import json
    inv = json.loads((REPO / "analysis" / "dump"
                       / "control_inventory.json").read_text(encoding="utf-8"))
    form = inv.get("EVENT_ADDR_2 Subform", {})
    cs_by_name = {c["name"]: c.get("control_source", "")
                   for c in form.get("controls", [])}
    assert cs_by_name.get("TxtAddrCHN") == "c_name_chn", (
        "Bug #10 marker no longer reproduces (investigate upstream fix vs. fixture/driver change vs. misclassification before flipping) — TxtAddrCHN no longer points at the "
        "wrong column (was 'c_name_chn', should be 'c_event_addr_chn')."
    )
    assert cs_by_name.get("TxtAddrPY") == "c_name", (
        "Bug #10 marker no longer reproduces (investigate upstream fix vs. fixture/driver change vs. misclassification before flipping) — TxtAddrPY no longer points at "
        "'c_name'."
    )


def test_bug11_events_data_subform_references_missing_column():
    """Bug #11 (candidate): EVENTS_DATA_2 Subform has a control with
    ControlSource `c_event_record_id`, but neither the EVENTS_DATA
    table nor View_EventsData (the form's RecordSource) has that
    column."""
    import json
    inv = json.loads((REPO / "analysis" / "dump"
                       / "control_inventory.json").read_text(encoding="utf-8"))
    form = inv.get("EVENTS_DATA_2 Subform", {})
    cs_values = {c.get("control_source", "")
                  for c in form.get("controls", [])}
    assert "c_event_record_id" in cs_values, (
        "Bug #11 marker no longer reproduces (investigate upstream fix vs. fixture/driver change vs. misclassification before flipping) — no control on EVENTS_DATA_2 references "
        "the non-existent `c_event_record_id` column anymore."
    )
    # Also assert it really doesn't exist on EVENTS_DATA.
    tables = json.loads((REPO / "analysis" / "dump"
                          / "tables.json").read_text(encoding="utf-8"))
    events_data = next((t for t in tables
                         if t.get("name") == "EVENTS_DATA"), {})
    cols = {c.get("name") for c in events_data.get("columns", [])}
    assert "c_event_record_id" not in cols, (
        "Bug #11 marker no longer reproduces (investigate upstream fix vs. fixture/driver change vs. misclassification before flipping) — EVENTS_DATA gained the "
        "`c_event_record_id` column, so the control now resolves."
    )


def test_bug12_posting_office_subform_wrong_appt_column():
    """Bug #12 (candidate): POSTED_TO_OFFICE_DATA_2 Subform has a
    control `c_appt_type_code` with ControlSource `c_appt_type_code`,
    but View_PostingOfficeData projects `c_appt_code` (no `_type`
    infix)."""
    import json
    inv = json.loads((REPO / "analysis" / "dump"
                       / "control_inventory.json").read_text(encoding="utf-8"))
    form = inv.get("POSTED_TO_OFFICE_DATA_2 Subform", {})
    cs_values = {c.get("control_source", "")
                  for c in form.get("controls", [])}
    assert "c_appt_type_code" in cs_values, (
        "Bug #12 marker no longer reproduces (investigate upstream fix vs. fixture/driver change vs. misclassification before flipping) — no control on POSTED_TO_OFFICE_DATA_2 "
        "references the non-projected `c_appt_type_code` column."
    )


def test_bug13_biog_main_2_subform_picks_missing_form():
    """Bug #13 (candidate, found by audit_cross_form_references):
    Form_BIOG_MAIN_2_Subform's `c_fl_ey_notes_Click` does
    `DoCmd.OpenForm "frmPickNIAN_HAO"`, but no form named
    `frmPickNIAN_HAO` exists in the .mdb.  Clicking that field on
    the person-detail subform throws 'Item not found in this
    collection.' instead of opening a picker."""
    body = ((REPO / "analysis" / "dump" / "vba"
              / "Form_BIOG_MAIN_2_Subform.vb")
             .read_bytes().decode("utf-8"))
    assert 'frmPickNIAN_HAO' in body, (
        "Bug #13 marker no longer reproduces (investigate upstream fix vs. fixture/driver change vs. misclassification before flipping) — `frmPickNIAN_HAO` no longer "
        "referenced in Form_BIOG_MAIN_2_Subform.  Re-run "
        "`analysis/audit_cross_form_references.py`."
    )
    # Confirm the form really doesn't exist.
    import json
    inv = json.loads((REPO / "analysis" / "dump"
                       / "control_inventory.json").read_text(encoding="utf-8"))
    keys_lc = {k.lower() for k in inv.keys()}
    assert "frmpicknian_hao" not in keys_lc, (
        "Bug #13 marker no longer reproduces (investigate upstream fix vs. fixture/driver change vs. misclassification before flipping) — `frmPickNIAN_HAO` was added to the "
        ".mdb.  The cross-form reference now resolves."
    )


def test_bug14_kin_data_subform_picks_missing_form():
    """Bug #14 (candidate): Form_KIN_DATA_Subform similarly
    references `frmPickKINSHIP_CODES`, which doesn't exist in the
    .mdb.  The kinship-code picker on the kinship subform fails."""
    body = ((REPO / "analysis" / "dump" / "vba"
              / "Form_KIN_DATA_Subform.vb")
             .read_bytes().decode("utf-8"))
    assert 'frmPickKINSHIP_CODES' in body, (
        "Bug #14 marker no longer reproduces (investigate upstream fix vs. fixture/driver change vs. misclassification before flipping) — `frmPickKINSHIP_CODES` no longer "
        "referenced in Form_KIN_DATA_Subform."
    )
    import json
    inv = json.loads((REPO / "analysis" / "dump"
                       / "control_inventory.json").read_text(encoding="utf-8"))
    keys_lc = {k.lower() for k in inv.keys()}
    assert "frmpickkinship_codes" not in keys_lc, (
        "Bug #14 marker no longer reproduces (investigate upstream fix vs. fixture/driver change vs. misclassification before flipping) — `frmPickKINSHIP_CODES` form added."
    )


def test_bugs_15_to_19_orphan_export_handlers():
    """Bugs #15-#19 (candidate, found by audit_orphan_event_handlers):
    several LookAt forms have export-button event handlers
    (CmdGIS_Click / CmdPajek_Click / CmdGephi_Click / CmdUCINet_Click /
    CmdGUESS_Click) defined in VBA but **no matching button on the
    form design**.  End users in the UI never see those export options
    even though the underlying code is fully functional (and our
    timer-trigger tests exercise it).

    Five distinct gaps:
      Bug #15: LookAtPlace.CmdGIS — no GIS button on Place
      Bug #16: LookAtStatus.CmdPajek — no Pajek button on Status
      Bug #17: LookAtStatus.CmdGephi — no Gephi button on Status
      Bug #18: LookAtStatus.CmdUCINet — no UCINet button on Status
      Bug #19: LookAtOffice.CmdGUESS — no GUESS button on Office
    """
    import json
    inv = json.loads((REPO / "analysis" / "dump"
                       / "control_inventory.json").read_text(encoding="utf-8"))

    def _has_ctl(form: str, ctl: str) -> bool:
        info = inv.get(form, {})
        ctls = {(c.get("name") or "").lower()
                for c in info.get("controls", [])}
        return ctl.lower() in ctls

    def _has_sub(form: str, sub: str) -> bool:
        body = ((REPO / "analysis" / "dump" / "vba"
                  / f"Form_{form}.vb").read_bytes()
                 .decode("utf-8"))
        return f"Sub {sub}(" in body

    cases = [
        ("LookAtPlace", "CmdGIS", 15),
        ("LookAtStatus", "CmdPajek", 16),
        ("LookAtStatus", "CmdGephi", 17),
        ("LookAtStatus", "CmdUCINet", 18),
        ("LookAtOffice", "CmdGUESS", 19),
    ]
    for form, ctl, bug_num in cases:
        # Sub still exists in VBA module.
        assert _has_sub(form, f"{ctl}_Click"), (
            f"Bug #{bug_num} marker no longer reproduces (investigate upstream fix vs. fixture/driver change vs. misclassification before flipping) — {form}.{ctl}_Click no "
            f"longer exists in the VBA module."
        )
        # Control still missing from form design.
        assert not _has_ctl(form, ctl), (
            f"Bug #{bug_num} marker no longer reproduces (investigate upstream fix vs. fixture/driver change vs. misclassification before flipping) — {form} now has a {ctl} "
            f"button on its design.  Drop the relevant skip in the "
            f"cross-form export test or expand fixture coverage."
        )


def test_bug_dao_reference_broken_in_user_mdb():
    """The shipped .mdb references DAO 3.6 (C:\\Program Files\\Common
    Files\\Microsoft Shared\\DAO\\dao360.dll) which is not installed on
    modern Office. This causes "Can't find project or library" when
    opening any form via Access COM automation.

    NOTE: This is documented in the analysis/check_vba_refs.py output
    (see analysis/dump for the captured snapshot).
    """
    # We can't easily check this without launching Access; treat as
    # documentation. Ensure the analysis script exists so the user can
    # re-confirm.
    chk = REPO / "analysis" / "check_vba_refs.py"
    assert chk.exists(), (
        "analysis/check_vba_refs.py missing -- can't verify the broken "
        "DAO reference. Re-pull from the project root."
    )
