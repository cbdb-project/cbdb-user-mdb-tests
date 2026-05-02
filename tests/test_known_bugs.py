"""
Regression tests for known bugs in CBDB_BJ_User.mdb (current as of 2026-04-30).

Each test here documents a discrete bug and verifies the buggy behaviour.
When a bug is fixed in the .mdb, the corresponding test will start FAILING —
that's the signal to flip its assertion to expect the corrected behaviour.

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
        "Bug appears to be FIXED — first-year range now pulls from the "
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
        "Bug appears to be FIXED — found rows where c_fy_range_desc != "
        "c_ly_range_desc. Update this test to remove the regression "
        f"check. ({len(mismatches)} rows differ)"
    )


def test_bug3_lookat_entry_cmdquery_backfill_silent_fail():
    """Bug #3 — LookAtEntry.CmdQuery_Click's backfill UPDATE silently
    fails on multi-table joins when the result set is large enough,
    leaving c_entry_desc / c_addr_name etc. NULL.  Confirmed only on
    LookAtEntry (Status / Texts / Associations all backfill correctly
    at similar row counts).  Real fix is for CBDB to split the big
    UPDATE into smaller ones (one per lookup table).

    This regression test is structural: it confirms the giant 7+
    table JOIN UPDATE is still in the dump.  When CBDB rewrites it
    into smaller UPDATEs, this assertion fires.  The behavioural
    side of the bug is tested by `test_vba_matrix.py::
    test_vba_full_matrix[top_entry_code_36_unfiltered]` which
    is currently expected-to-fail (xfail-style).
    """
    vba_path = (REPO / "analysis" / "dump" / "vba"
                / "Form_LookAtEntry.vb")
    body = vba_path.read_bytes().decode("utf-8")
    # The buggy UPDATE chain joins 6+ tables before SET.
    # Look for the distinctive run of LEFT JOINs in CmdQuery_Click.
    assert ("KINSHIP_CODES" in body
            and "SOCIAL_INSTITUTION_NAME_CODES" in body
            and "BIOG_MAIN_1" in body), (
        "Bug #3 may be FIXED — the giant multi-table UPDATE in "
        "Form_LookAtEntry.CmdQuery_Click no longer references "
        "KINSHIP_CODES / BIOG_MAIN_1 / SOCIAL_INSTITUTION_NAME_CODES "
        "together.  Re-run the matrix test "
        "[top_entry_code_36_unfiltered] to verify; if it now passes, "
        "flip this assertion."
    )


def test_bug4_lookat_place_cmdgis_references_nonexistent_gisframe():
    """Bug #4 — Form_LookAtPlace.CmdGIS_Click references `GISFrame.
    Value` but the form has no `GISFrame` control (only `CodeFrame`).
    The driver works around this via _PER_FORM_CMDGIS_PATCHES.
    """
    vba_path = (REPO / "analysis" / "dump" / "vba"
                / "Form_LookAtPlace.vb")
    body = vba_path.read_bytes().decode("utf-8")
    assert "GISFrame.Value" in body, (
        "Bug #4 appears to be FIXED — `GISFrame.Value` no longer "
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
        "Bug #5 (ChkIDs side) appears to be FIXED in Form_LookAtStatus."
    )
    # The SQL bug is independent — flip when both halves clear.
    assert "ZZ_SCRATCH_STATUS.c_status_count" in body, (
        "Bug #5 (SQL side) appears to be FIXED — the ZZ_SCRATCH_STATUS"
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
        "Bug #8 appears to be FIXED — the LookAtNetworks.CmdNeo4j "
        "tRstPlace SELECT no longer matches the buggy 3-col form. "
        "Flip this assertion."
    )


def test_bug6_groupdata_query_entry_wrong_field():
    """Bug #6 (findings.md) — Form_LookAtGroupData.queryEntry projects
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
        "Bug #6 appears to be FIXED — `ENTRY_DATA.c_parental_status` "
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


def test_bug7_lookat_place_cmdneo4j_select_missing_dynasty_female():
    """Bug #7 (findings_en.md) — Form_LookAtPlace.CmdNeo4j builds a
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
        "Bug #7 appears to be FIXED — the LookAtPlace.CmdNeo4j SELECT "
        "no longer matches the buggy 4-col projection.  Re-run "
        "`analysis/audit_recordset_sql_projection.py` to confirm; if "
        "clean, flip this assertion."
    )


def test_bug9_lookat_entry_cmdneo4j_with_wrong_var():
    """Bug #9 — Form_LookAtEntry.CmdNeo4j has `With tRstAssocCodes`
    where `With tRstInstitutions` was intended (the `!c_inst_*` reads
    that follow only make sense against the institutions recordset).
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
        "Bug #9 appears to be FIXED — `With tRstAssocCodes` immediately "
        "following the tRstInstitutions OpenRecordset is no longer "
        "present (likely renamed to `With tRstInstitutions`).  Flip "
        "this assertion."
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
        "Bug #10 may be FIXED — TxtAddrCHN no longer points at the "
        "wrong column (was 'c_name_chn', should be 'c_event_addr_chn')."
    )
    assert cs_by_name.get("TxtAddrPY") == "c_name", (
        "Bug #10 may be FIXED — TxtAddrPY no longer points at "
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
        "Bug #11 may be FIXED — no control on EVENTS_DATA_2 references "
        "the non-existent `c_event_record_id` column anymore."
    )
    # Also assert it really doesn't exist on EVENTS_DATA.
    tables = json.loads((REPO / "analysis" / "dump"
                          / "tables.json").read_text(encoding="utf-8"))
    events_data = next((t for t in tables
                         if t.get("name") == "EVENTS_DATA"), {})
    cols = {c.get("name") for c in events_data.get("columns", [])}
    assert "c_event_record_id" not in cols, (
        "Bug #11 may be FIXED — EVENTS_DATA gained the "
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
        "Bug #12 may be FIXED — no control on POSTED_TO_OFFICE_DATA_2 "
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
        "Bug #13 may be FIXED — `frmPickNIAN_HAO` no longer "
        "referenced in Form_BIOG_MAIN_2_Subform.  Re-run "
        "`analysis/audit_cross_form_references.py`."
    )
    # Confirm the form really doesn't exist.
    import json
    inv = json.loads((REPO / "analysis" / "dump"
                       / "control_inventory.json").read_text(encoding="utf-8"))
    keys_lc = {k.lower() for k in inv.keys()}
    assert "frmpicknian_hao" not in keys_lc, (
        "Bug #13 may be FIXED — `frmPickNIAN_HAO` was added to the "
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
        "Bug #14 may be FIXED — `frmPickKINSHIP_CODES` no longer "
        "referenced in Form_KIN_DATA_Subform."
    )
    import json
    inv = json.loads((REPO / "analysis" / "dump"
                       / "control_inventory.json").read_text(encoding="utf-8"))
    keys_lc = {k.lower() for k in inv.keys()}
    assert "frmpickkinship_codes" not in keys_lc, (
        "Bug #14 may be FIXED — `frmPickKINSHIP_CODES` form added."
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
