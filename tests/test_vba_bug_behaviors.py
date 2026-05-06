"""Behavioral verification tests for Bugs #7-#19.

The static auditors that found these bugs (PRs #28 / #38 / #43 / #47)
prove the buggy *code* exists.  These tests prove the buggy *runtime
behavior* matches what the audit predicted — fire the relevant Sub
in real Access via the test driver, then assert the expected silent
failure mode appears in `ZZ_TEST_DEBUG`.

Why bother (beyond the static audits)?
  - Confirms the bug is genuinely user-reachable, not theoretical.
  - When CBDB attempts a fix, this test moves from "asserts ERR
    appears" to "asserts ERR is gone" — flipping a behavioural
    sentinel rather than a code-grep one.
  - Documents the exact error message users see, which the static
    audit can only guess at.

Each bug gets one test; we share one VBA fixture across the suite
to amortise the ~12 s Access startup.

What we DON'T do here:
  - Bug #11 / #12 / #15-#19 are design-time-only (control source
    misnamed / handler with no button) — they have no runtime
    error to observe.  test_known_bugs.py already covers them
    structurally.
  - Bug #13 / #14 trigger only when the user clicks a specific
    sub-form field; reproducing that needs a parent-form context
    we don't currently have a fixture for.  Static markers in
    test_known_bugs cover them.
"""
from __future__ import annotations

from pathlib import Path

import pytest

from cbdb_driver.vba_session import VbaSession, make_fixture
from cbdb_driver.form_specs import LOOKATPLACE, LOOKATENTRY
from test_vba_matrix_all_forms import _all_fixtures, CrossFixture, SRC


WORK = Path(__file__).resolve().parent.parent / "analysis" / "_bug_behavior_test_copy.mdb"


@pytest.fixture(scope="function")
def vba():
    yield from make_fixture(SRC, WORK)


def _seed(vba: VbaSession, fx: CrossFixture) -> None:
    spec = fx.spec
    vba.open_form(spec.name)
    if spec.name == "LookAtPlace":
        try:
            vba.set_control("LookAtPlace", "TabPlaces", 0)
        except Exception as e:
            print(f"  warn TabPlaces=0: {e}")
    for ctl, val in (fx.controls or {}).items():
        try:
            vba.set_control(spec.name, ctl, val)
        except Exception as e:
            print(f"  warn setting {ctl}={val!r}: {e}")
    if fx.picker_ids and spec.picker_table:
        vba.set_picker_codes(spec.picker_table, fx.picker_ids,
                             column=spec.picker_column)
    if fx.addr_ids:
        vba.set_picker_addrs(fx.addr_ids)


def _fixture_for(form: str) -> CrossFixture:
    for fx in _all_fixtures():
        if fx.spec.name == form:
            return fx
    pytest.skip(f"no matrix fixture for {form}")


def _read_debug_log(vba: VbaSession) -> list[str]:
    cur = vba.conn.cursor()
    cur.execute("SELECT msg FROM ZZ_TEST_DEBUG ORDER BY id ASC")
    rows = [str(r[0]) for r in cur.fetchall()]
    cur.close()
    return rows


def _chain_via_tag(vba: VbaSession, form: str, *,
                    chain: str, target_table: str = "",
                    timeout: int = 90) -> list[str]:
    """Fire CmdQuery → chain target via Form.Tag, return ZZ_TEST_DEBUG."""
    out_dir = WORK.parent / "_bug_behavior_out"
    out_dir.mkdir(exist_ok=True)
    out_path = str(out_dir) + "\\"  # directory mode for multi-file
    vba.patch_filedialog(form)
    vba.set_form_tag(form, chain, out_path)
    vba.click_via_timer(form, ctl=chain.split(",")[0],
                         result_table=target_table or None,
                         timeout=timeout)
    return _read_debug_log(vba)


def test_bug4_lookat_place_cmdgis_fires_object_required(monkeypatch):
    """Bug #4: Form_LookAtPlace.CmdGIS_Click references the
    non-existent control `GISFrame` (the actual control on Place is
    `CodeFrame`).  Driver autopatch via `_PER_FORM_CMDGIS_PATCHES`
    rewrites `GISFrame.Value` → `CodeFrame.Value` so the integration
    tests pass; this deep test temporarily disables that patch and
    confirms the un-patched code still throws 'Object required'.
    """
    from cbdb_driver.form_specs import LOOKATPLACE
    # Empty out the per-form patches dict for the duration of this
    # test only (monkeypatch undoes after teardown).
    monkeypatch.setattr(VbaSession, "_PER_FORM_CMDGIS_PATCHES", {})

    work_unpatched = WORK.parent / "_bug4_unpatched.mdb"
    if work_unpatched.exists():
        try:
            work_unpatched.unlink()
        except PermissionError:
            import time
            time.sleep(1); work_unpatched.unlink()
    vba_local = VbaSession(SRC, work_unpatched)
    vba_local.open()
    try:
        spec = LOOKATPLACE
        fx = next((f for f in _all_fixtures()
                    if f.spec.name == "LookAtPlace"), None)
        assert fx is not None, "no LookAtPlace fixture in matrix"
        _seed(vba_local, fx)
        msgs = _chain_via_tag(vba_local, "LookAtPlace",
                                chain="CmdQuery,CmdGIS",
                                target_table="ZZ_SCRATCH_PLACE")
        print(f"\nDEBUG log: {msgs}", flush=True)
        err_msgs = [m for m in msgs if ":ERR " in m]
        assert any("Object required" in m for m in err_msgs), (
            f"Bug #4 marker no longer reproduces (investigate upstream fix vs. fixture/driver change vs. misclassification before flipping) — un-patched LookAtPlace.CmdGIS "
            f"no longer raises 'Object required'.  err_msgs={err_msgs}"
        )
    finally:
        vba_local.close()


def test_bug5_lookat_status_cmdpajek_sql_fires_field_error(vba: VbaSession):
    """Bug #5 (SQL side): CmdPajek_Click body has a SELECT that
    references columns missing from ZZ_SCRATCH_STATUS.  Trigger by:
      1. Pre-populating ZZ_SCRATCH_STATUS via pyodbc with a minimal
         row so the early `RecordCount = 0` bail is skipped.
      2. Open LookAtStatus (Form_Open wipes IMPORT_PEOPLE but NOT
         the SCRATCH_STATUS data we just inserted IF we reinsert
         after open).
      3. Direct timer-fire `CmdPajek` (no chain — Status's CmdQuery
         cleanup interferes with chain dispatch).
    """
    from cbdb_driver.form_specs import LOOKATSTATUS
    spec = LOOKATSTATUS
    # Open the form first — Form_Open wipes ZZ_SCRATCH_STATUS.
    vba.open_form(spec.name)
    # Seed via Access's OWN connection (not pyodbc) — pyodbc writes
    # are invisible to Access's cached subform recordset until much
    # later, defeating the test's Requery.  CurrentDb.Execute writes
    # are visible immediately.
    db = vba.app.CurrentDb()
    db.Execute(
        "INSERT INTO ZZ_SCRATCH_STATUS "
        "(c_personid, c_sequence, c_status_code) VALUES (4, 1, 1)"
    )
    db.Execute(
        "INSERT INTO ZZ_SCRATCH_P_STATUS (c_person_id, c_addr_id) "
        "VALUES (4, 100658)"
    )
    f = vba.app.Forms(spec.name)
    try:
        f.Controls("ZZ_SCRATCH_STATUS").Form.Requery()
        f.Controls("ZZ_SCRATCH_P_STATUS").Form.Requery()
    except Exception as e:
        print(f"  warn requery: {e}")

    # Direct timer-fire CmdPajek (NO CmdQuery chain).
    vba.click_via_timer(spec.name, ctl="CmdPajek",
                         result_table=None, timeout=60)
    msgs = _read_debug_log(vba)
    print(f"\nDEBUG log: {msgs}", flush=True)
    err_msgs = [m for m in msgs if ":ERR " in m]
    # Bug #5 covers a cluster: ChkIDs control + 3 missing SQL columns +
    # potentially other knock-on issues from the broader copy-paste
    # mistake.  ANY ERR fired by CmdPajek under a fresh ZZ_SCRATCH_STATUS
    # row is the bug — a clean form would complete and write a .net file.
    # We check err is non-empty AND specific to the family.
    assert err_msgs, (
        f"Bug #5 marker no longer reproduces (investigate upstream fix vs. fixture/driver change vs. misclassification before flipping) — LookAtStatus.CmdPajek now completes "
        f"without errors.  msgs={msgs}"
    )
    expected_kinds = (
        "object required",       # ChkIDs control side
        "field",                 # c_person_id / c_status_id / c_status_count
        "c_person_id",
        "c_status_id",
        "c_status_count",
        "no such",
    )
    matched = any(any(k in m.lower() for k in expected_kinds)
                   for m in err_msgs)
    assert matched, (
        f"Bug #5 fired an unexpected error type — investigate.  "
        f"err_msgs={err_msgs}"
    )


def test_bug6_lookat_groupdata_query_entry_fires_no_such_field(vba: VbaSession):
    """Bug #6 (P1) — runtime-side pin.

    `Form_LookAtGroupData.queryEntry()` at vba:2593-2625.  The
    INSERT target column list (line 2612) ends with
    `c_parental_status_code`; the SELECT projection (line 2621)
    ends with `ENTRY_DATA.c_parental_status` — no `_code` suffix.
    JET parses the unknown column as an unbound parameter and
    raises VBA error 3061 ("No value given for one or more
    required parameters") before any INSERT executes, so
    `ZZ_SCRATCH_ENTRY` stays at 0.

    Sister test in `tests/test_known_bugs.py::test_bug6_groupdata
    _query_entry_wrong_field` pins the source-string substring;
    this test pins the runtime symptom on the documented
    reproduction (person 1 / An Dun 安惇 / only ChkEntry checked
    / click Run).  Don't replace the static test — both the
    source typo and the runtime crash should be guarded.

    Localisation evidence:
      analysis/groupdata_cmdgis_subcall_trace.md
      reports/groupdata_cmdgis_subcall_trace.json
    The 11-iteration sub-isolation probe found ERR fires in
    EXACTLY the 2 iterations that exercise queryEntry
    (queryEntry_alone + Entry_full_chain) and clean in the
    other 9.

    Fixture: matrix_hard_forms's groupdata_person_1_small —
    person 1 has 2 ENTRY_DATA rows so queryEntry's INSERT...
    SELECT genuinely tries to run against real data (not a
    "0 rows so SQL no-ops" false negative).
    """
    from cbdb_driver.form_specs import LOOKATGROUPDATA
    spec = LOOKATGROUPDATA
    fx = CrossFixture(
        name="bug6_groupdata",
        spec=spec,
        picker_ids=[1],
        controls={"ChkEntry": True},
        expected_min_rows=1,
        source_sql=None,
    )
    _seed(vba, fx)
    msgs = _chain_via_tag(vba, "LookAtGroupData",
                           chain="CmdRun",
                           target_table="")
    print(f"\nDEBUG log: {msgs}", flush=True)

    # ---- Assertion 1: ZZ_TEST_DEBUG carries a LookAtGroupData:ERR
    err_msgs = [m for m in msgs if "LookAtGroupData:ERR" in m]
    assert err_msgs, (
        f"Bug #6 marker no longer reproduces (investigate upstream "
        f"fix vs. fixture/driver change vs. misclassification "
        f"before flipping) — LookAtGroupData.CmdRun with ChkEntry "
        f"didn't raise any :ERR marker.  Full transcript: {msgs}"
    )

    # ---- Assertion 2: error text matches the JET 3061 family
    # ("No value given for one or more required parameters").
    # That's the exact wording the probe observed and what
    # Issue #6 documents.  Some Office builds / ODBC paths
    # surface the same root cause as "Could not find field
    # 'c_parental_status'" — accept either, but the parameter-
    # phrasing is the canonical one on the current dump.
    expected_signatures = (
        "no value given for one or more required parameters",
        "could not find field",
        "c_parental_status",       # the typo'd column name itself
    )
    matched_sigs = [
        sig for sig in expected_signatures
        if any(sig in m.lower() for m in err_msgs)
    ]
    assert matched_sigs, (
        f"Bug #6 ERR fired but its text doesn't match any of the "
        f"expected JET 3061 signatures {expected_signatures}.  "
        f"Investigate whether the underlying typo changed (Issue "
        f"#6's source-side fix would make `c_parental_status` go "
        f"away — flip both this test AND the static test_bug6 in "
        f"tests/test_known_bugs.py).  err_msgs={err_msgs}"
    )

    # ---- Assertion 3: ZZ_SCRATCH_ENTRY stays at 0 because the
    # JET parameter error fires BEFORE queryEntry's INSERT
    # executes.  Pinned at 0 to catch silent upstream fixes that
    # would let the INSERT proceed (in which case scratch_entry
    # would jump to person_1's 2 ENTRY_DATA rows).  Either
    # outcome — error gone OR INSERT proceeds — means the bug
    # behaviour shifted; both warrant manual review per the
    # marker-failure policy.
    cur = vba.conn.cursor()
    cur.execute("SELECT COUNT(*) FROM ZZ_SCRATCH_ENTRY")
    n_entry = int(cur.fetchone()[0] or 0)
    cur.close()
    assert n_entry == 0, (
        f"Bug #6 marker partially reproduces but ZZ_SCRATCH_ENTRY "
        f"got populated ({n_entry} rows) — the JET 3061 error "
        f"used to fire BEFORE INSERT.  If the INSERT now succeeds, "
        f"either the typo was fixed OR JET's parameter handling "
        f"changed.  Investigate per the marker-failure policy "
        f"(docs/skills/issue-report-maintainer.md) before "
        f"flipping this assertion."
    )


def test_bug21_lookat_groupdata_cmdneo4j_fires_no_current_record(
        vba: VbaSession):
    """Bug #21 (P1) — runtime-side pin.

    `Form_LookAtGroupData.CmdNeo4j_Click` has 11 dlgSaveAs.Show
    blocks, each preceded by `Set <var> = OpenRecordset(...)`
    followed unguarded by `.MoveFirst`.  When `ZZ_SCRATCH_ENTRY`
    is empty (the common case where the queried person has no
    Entry data, OR ChkEntry was left unticked), the `.MoveFirst`
    at vb line 1245 (PeopleEntry, block #9) raises DAO 3021
    'No current record' and the chain bails before writing the
    Entry-related tail blocks (PeopleEntry / EntryCode /
    InstitutionCodes).  See PR
    `investigate/groupdata-cmdneo4j-tail` (commit 3bfcba8) for
    the per-block isolation evidence.

    **Distinct from Issue #6.**  Issue #6 is JET 3061 ('No
    value given for one or more required parameters') in
    `queryEntry`, UPSTREAM of this bug.  Both can fire on the
    same Entry-enabled path, but they are different code-level
    defects: #6 is a column-typo (`ENTRY_DATA.c_parental_status`
    missing `_code` suffix); #21 is a missing `.EOF` /
    `.RecordCount > 0` guard before `.MoveFirst`.  This test
    deliberately leaves `ChkEntry` OFF so Issue #6's path is NOT
    in scope and ONLY Issue #21 (the downstream missing-guard)
    can fire.

    Sister test in
    `tests/test_known_bugs.py::test_bug21_groupdata_cmdneo4j
    _missing_eof_guard` pins the source-grep pattern (no
    `If Not .EOF Then` between OpenRecordset and .MoveFirst).
    This test pins the runtime symptom.  Don't replace the
    static test — both source-side and runtime-side should be
    guarded.

    Localisation evidence (already merged to main):
      analysis/groupdata_cmdneo4j_probe.md
      analysis/groupdata_cmdneo4j_tail_probe.md
      reports/groupdata_cmdneo4j_probe.json
      reports/groupdata_cmdneo4j_tail_probe.json
    The tail probe's iter 3 (split-then-seed) is the killer
    evidence: when one synthetic row is inserted into
    `ZZ_SCRATCH_ENTRY` before CmdNeo4j fires, the chain
    produces 10 files with no `:ERR` — proving the trigger is
    the empty source recordset, not anything else.

    Fixture: matrix_hard_forms's `groupdata_person_1_small` —
    person 1 has 2 STATUS_DATA / 2 ENTRY_DATA / ~12
    POSTED_TO_OFFICE rows.  Chk state set explicitly:
    Status / Office / Addr ON, GIS sisters ON, Entry / Text
    OFF.  Matches the GroupData × CmdGIS coverage scope; with
    ChkEntry OFF, queryEntry doesn't run and ZZ_SCRATCH_ENTRY
    stays at 0 — the precise condition Issue #21 needs to fire.
    """
    from cbdb_driver.form_specs import LOOKATGROUPDATA
    spec = LOOKATGROUPDATA
    PERSON_ID = 1

    # Picker setup
    vba.set_picker_codes(spec.picker_table, [PERSON_ID],
                          column=spec.picker_column)
    vba.open_form(spec.name)

    # All-Chk*-reset (matches the all-Chk*-reset-first pattern
    # proven by tests/test_vba_cmdgis_other_forms.py
    # ::test_cmd_gis_groupdata_clean_branches).  Without this,
    # Form_Open defaults can leave ChkEntry True and pull the
    # test into Issue #6's path instead.
    all_chk = (
        "ChkStatus", "ChkOffice", "ChkEntry", "ChkText",
        "ChkAddr",
        "ChkGisStatus", "ChkGisOffice", "ChkGisOfficePeople",
        "ChkGisEntry", "ChkGisText", "ChkGisAddr",
    )
    for c in all_chk:
        try:
            vba.set_control(spec.name, c, False)
        except Exception as e:
            print(f"  warn reset {c}: {e}")

    # Enable the clean-branches set: Status / Office / Addr +
    # GIS sisters.  ChkEntry deliberately OFF.
    for c in ("ChkStatus", "ChkOffice", "ChkAddr",
              "ChkGisStatus", "ChkGisOffice", "ChkGisAddr"):
        try:
            vba.set_control(spec.name, c, True)
        except Exception as e:
            print(f"  warn enable {c}: {e}")

    # Chain CmdRun -> CmdNeo4j via Form.Tag.  The autodetect-
    # injected chain block dispatches CmdNeo4j after CmdRun's
    # body completes.
    msgs = _chain_via_tag(vba, "LookAtGroupData",
                           chain="CmdRun,CmdNeo4j",
                           target_table="")
    print(f"\nDEBUG log: {msgs}", flush=True)

    # ---- Assertion 1: ZZ_SCRATCH_ENTRY stays at 0 (the
    # precondition that distinguishes Issue #21 from Issue #6).
    cur = vba.conn.cursor()
    cur.execute("SELECT COUNT(*) FROM ZZ_SCRATCH_ENTRY")
    n_entry = int(cur.fetchone()[0] or 0)
    cur.close()
    assert n_entry == 0, (
        f"Bug #21 precondition broke — ZZ_SCRATCH_ENTRY = "
        f"{n_entry} but should be 0 with ChkEntry off.  Either "
        f"Form_Open default changed (now sets ChkEntry True), "
        f"or queryStatus / queryOffice / queryAddr now writes "
        f"to ZZ_SCRATCH_ENTRY (unlikely).  This invalidates the "
        f"Issue #21 reproduction (Issue #21 needs an empty "
        f"feeder; if ChkEntry is implicitly enabled, what we "
        f"reproduced is Issue #6 instead).  Investigate before "
        f"flipping."
    )

    # ---- Assertion 2: at least one LookAtGroupData:ERR fired
    err_msgs = [m for m in msgs if "LookAtGroupData:ERR" in m]
    assert err_msgs, (
        f"Bug #21 marker no longer reproduces (investigate "
        f"upstream fix vs. fixture/driver change vs. "
        f"misclassification before flipping) — chain "
        f"CmdRun→CmdNeo4j with ZZ_SCRATCH_ENTRY=0 didn't raise "
        f"any :ERR marker.  Either the unguarded `.MoveFirst` "
        f"now has a guard (CBDB upstream fix), the chain didn't "
        f"reach the PeopleEntry block, or the test infra "
        f"silenced the error.  Full transcript: {msgs}"
    )

    # ---- Assertion 3: error text matches the DAO 3021 family
    # ('no current record').  Critically NOT the Issue #6
    # family ('no value given for required parameters' / 'could
    # not find field' / 'c_parental_status') — those signatures
    # would mean we're actually reproducing Issue #6, not Issue
    # #21.
    err_blob = " | ".join(err_msgs).lower()
    assert "no current record" in err_blob, (
        f"Bug #21 ERR fired but its text doesn't match DAO 3021 "
        f"('No current record').  Expected DAO 3021 (unguarded "
        f"`.MoveFirst` on empty recordset).  err_msgs={err_msgs}.  "
        f"If this is now an Issue #6-class JET 3061, the fixture "
        f"has drifted and ChkEntry is implicitly enabled."
    )

    issue_6_signatures = (
        "no value given for one or more required parameters",
        "could not find field",
        "c_parental_status",
    )
    issue_6_match = [s for s in issue_6_signatures
                     if s in err_blob]
    assert not issue_6_match, (
        f"Bug #21 reproduction is contaminated by Issue #6 "
        f"signatures {issue_6_match} — meaning ChkEntry-side "
        f"queryEntry fired (which would only happen if "
        f"ChkEntry was implicitly enabled despite the explicit "
        f"reset above).  Investigate the all-Chk*-reset step.  "
        f"err_msgs={err_msgs}"
    )

    print(f"\nBug #21 runtime pin OK: ZZ_SCRATCH_ENTRY={n_entry}, "
          f"DAO 3021 ('no current record') observed in "
          f"{len(err_msgs)} :ERR marker(s); no Issue #6 "
          f"contamination.", flush=True)


def test_bug7_lookat_place_cmdneo4j_fires_item_not_found(vba: VbaSession):
    """Bug #7: LookAtPlace.CmdNeo4j hits 'Item not found in this
    collection' on the first row of the People-CSV loop because
    the SELECT doesn't project c_dynasty / c_dynasty_chn / c_female
    that the loop reads."""
    fx = _fixture_for("LookAtPlace")
    _seed(vba, fx)
    msgs = _chain_via_tag(vba, "LookAtPlace",
                           chain="CmdQuery,CmdNeo4j",
                           target_table="ZZ_SCRATCH_PLACE")
    print(f"\nDEBUG log: {msgs}", flush=True)
    err_msgs = [m for m in msgs if ":ERR " in m]
    assert any("Item not found" in m for m in err_msgs), (
        f"Bug #7 marker no longer reproduces (investigate upstream fix vs. fixture/driver change vs. misclassification before flipping) — LookAtPlace.CmdNeo4j no longer "
        f"raises 'Item not found' at runtime.  err_msgs={err_msgs}"
    )


def test_bug8_lookat_networks_skipped_cmdrun_timeout():
    """Bug #8 deep verification needs LookAtNetworks CmdRun to
    complete, which currently times out on high-degree anchors
    (Zhu Xi 2471 assocs).  PR AA's probe showed Form_Open itself
    opens cleanly in ~2 s — the blocker is CmdRun network
    expansion, not Form_Open.  The static marker in
    test_known_bugs covers the code-side; this placeholder
    documents why we don't have runtime coverage until a smaller
    fixture is blessed."""
    pytest.skip("LookAtNetworks CmdRun times out on high-degree "
                "anchors (PR AA: Form_Open is fine).  Bug #8 "
                "covered statically by test_known_bugs.py.")


def test_bug9_lookat_entry_cmdneo4j_with_institutions_fixture(vba: VbaSession):
    """Bug #9: LookAtEntry.CmdNeo4j has `With tRstAssocCodes` where
    `With tRstInstitutions` was intended.

    Reclassified 2026-05-04 to P5 latent.  The buggy block sits
    inside `If tRecDeleted > 0 Then` at Form_LookAtEntry.vb:1389,
    where tRecDeleted is the count of rows inserted into
    ZZ_SCRATCH_P_TEXT with `WHERE c_inst_code > 0` — i.e. only
    entries with a social-institution code reach the With block.
    On the current dump 0 of 263,454 ENTRY_DATA rows have
    c_inst_code > 0, so the buggy block is unreachable from any
    LookAtEntry fixture (full re-verification:
    analysis/issue9_neo4j_institutioncodes_reverification.md).

    This behavioural test pins that property — for a normal user
    fixture (entry_code=36 jinshi), CmdNeo4j completes without
    error AND without reaching the buggy With block.  If a future
    MDB drop introduces ENTRY_DATA rows with c_inst_code > 0,
    test_known_bugs.test_bug9 will fail at the LATENT-gate
    assertion and force re-promotion to P1; this test would still
    pass for fixtures that don't pull in inst-bearing entries."""
    spec = LOOKATENTRY
    fx = CrossFixture(
        name="bug9_entry",
        spec=spec,
        picker_ids=[36],
        controls={"FrameYears": 1},
        expected_min_rows=1,
        source_sql=None,
    )
    _seed(vba, fx)
    msgs = _chain_via_tag(vba, "LookAtEntry",
                           chain="CmdQuery,CmdNeo4j",
                           target_table="ZZ_SCRATCH_ENTRY")
    print(f"\nDEBUG log: {msgs}", flush=True)
    err_msgs = [m for m in msgs if ":ERR " in m]
    # Confirm the run reached its DONE marker AND no Item-not-found
    # appeared — i.e. the buggy block is reachable only with
    # institution-bearing entries.
    assert any("LookAtEntry:DONE" in m for m in msgs), (
        f"CmdNeo4j chain didn't reach DONE marker; msgs={msgs}"
    )
    item_not_found = [m for m in err_msgs if "Item not found" in m]
    assert not item_not_found, (
        f"Bug #9 fired even with non-institution fixture — that's "
        f"unexpected; the buggy With block should be guarded by "
        f"`If tRecDeleted > 0`.  err_msgs={err_msgs}"
    )
