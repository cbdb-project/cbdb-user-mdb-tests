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
