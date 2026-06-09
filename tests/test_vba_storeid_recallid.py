"""
Cross-form CmdStoreID_Click / CmdRecallID_Click round-trip tests
(roadmap item 15).

`CmdStoreID_Click` lives on every LookAt form.  After CmdQuery has
populated the form's primary scratch table (e.g. ZZ_SCRATCH_ENTRY),
clicking it copies DISTINCT person ids into the cross-session table
`ZZ_STORE_PERSON_ID` and updates `PersonIDSource.SourceForm` to a
form-specific label (e.g. 'Entry'), so that a subsequent form open
can recall the same set via `CmdRecallID_Click`.

`CmdRecallID_Click` lives on the four "input-list driven" forms
(Kinship, Networks, GroupData, AssociationPairs).  It reads
`ZZ_STORE_PERSON_ID` and copies into `ZZ_SCRATCH_IMPORT_PEOPLE`,
which the form's CmdRun_Click then consumes.

What this file tests:
  - test_storeid_matches_scratch:  per form, run CmdQuery on a known
    fixture, then chain CmdStoreID, then assert the contents of
    ZZ_STORE_PERSON_ID equal DISTINCT(scratch.<personid_col>).  Also
    check PersonIDSource.SourceForm == form's label (where applicable).
  - test_recallid_repopulates_target:  per recall-form, pre-populate
    ZZ_STORE_PERSON_ID, click CmdRecallID, assert
    ZZ_SCRATCH_IMPORT_PEOPLE.c_person_id matches the seeded set.
  - test_storeid_recallid_roundtrip:  query LookAtEntry → CmdStoreID,
    then open LookAtKinship → CmdRecallID, assert the Kinship input
    table matches the original Entry result.  Exercises the cross-form
    handoff that was the user-visible reason the buttons exist.

How VBA blocking is handled:
  - The injected `MsgBox "literal"` neutralizer in
    cbdb_driver.vba_session._inject_autodetect rewrites the
    "Person IDs successfully stored." popup to a debug INSERT.
  - The yes/no `If MsgBox(...) = vbNo Then` prompts are sidestepped
    by pre-cleaning ZZ_STORE_PERSON_ID and ZZ_SCRATCH_IMPORT_PEOPLE
    so the gating `DCount > 0` branch is never entered.
"""
from __future__ import annotations

import json

import pytest

from cbdb_driver.vba_session import VbaSession, make_fixture
from cbdb_driver.form_specs import (
    ALL_SPECS, FormSpec, LOOKATENTRY,
)

from test_vba_matrix_all_forms import (
    SRC, WORK, INPUTS, CrossFixture, _all_fixtures,
)


@pytest.fixture(scope="function")
def vba():
    yield from make_fixture(SRC, WORK)


# Forms whose CmdQuery currently runs in the matrix.  StoreID test
# requires a working query to first populate the source scratch table.
_STOREID_FORMS_WITH_QUERY = {
    "LookAtEntry", "LookAtStatus", "LookAtTexts",
    "LookAtAssociations", "LookAtOffice", "LookAtPlace",
    "LookAtKinship",
}


def _entry_fixture() -> CrossFixture | None:
    """LookAtEntry isn't in test_vba_matrix_all_forms._all_fixtures
    (Entry has its own dedicated test_vba_matrix.py).  Build one here
    using the same top-entry-code pattern as that file."""
    if not INPUTS.exists():
        return None
    inputs = json.loads(INPUTS.read_text(encoding="utf-8"))
    le_data = inputs.get("lookatentry") or {}
    top = le_data.get("top_entry_codes") or []
    if not top:
        return None
    c = int(top[0]["c_entry_code"])
    return CrossFixture(
        name=f"entry_{c}_unfiltered",
        spec=LOOKATENTRY,
        picker_ids=[c],
        controls={"FrameYears": 0,
                  "TxtEntryDesc": top[0]["c_entry_desc"],
                  "TxtTypeCode": "N/A"},
        expected_min_rows=1000,
    )


def _storeid_fixtures() -> list[CrossFixture]:
    """One fixture per form with both (a) a known-working CmdQuery and
    (b) a CmdStoreID button.  Picks the first matrix fixture per form,
    plus the inline LookAtEntry fixture."""
    by_form: dict[str, CrossFixture] = {}
    entry = _entry_fixture()
    if entry is not None:
        by_form["LookAtEntry"] = entry
    for fx in _all_fixtures():
        name = fx.spec.name
        if name not in _STOREID_FORMS_WITH_QUERY:
            continue
        if fx.spec.storeid_source_table is None:
            continue
        if name not in by_form:
            by_form[name] = fx
    return list(by_form.values())


def _recallid_specs() -> list[FormSpec]:
    return [s for s in ALL_SPECS.values() if s.recallid_target_table]


def _seed_query_inputs(vba: VbaSession, fx: CrossFixture) -> None:
    """Open form, set controls, populate pickers — same as the matrix
    test's steps 1+2.  Form must be open before pickers are populated
    (Form_LookAtOffice's Form_Open wipes ZZ_OFFICE_CODE)."""
    spec = fx.spec
    vba.open_form(spec.name)
    for ctl, val in fx.controls.items():
        try:
            vba.set_control(spec.name, ctl, val)
        except Exception as e:
            print(f"  warn setting {ctl}={val!r}: {e}")
    if fx.picker_ids and spec.picker_table:
        vba.set_picker_codes(spec.picker_table, fx.picker_ids,
                             column=spec.picker_column)
    if fx.addr_ids:
        vba.set_picker_addrs(fx.addr_ids)


def _distinct_personids(vba: VbaSession, table: str, col: str,
                        where: str = "") -> set[int]:
    sql = f"SELECT DISTINCT [{col}] FROM [{table}]"
    if where:
        sql += f" {where}"
    cur = vba.conn.cursor()
    cur.execute(sql)
    out = {int(r[0]) for r in cur.fetchall() if r[0] is not None}
    cur.close()
    return out


# ---------------------------------------------------------------- StoreID
@pytest.mark.parametrize(
    "fx", _storeid_fixtures(), ids=lambda f: f.name,
)
def test_storeid_matches_scratch(vba: VbaSession, fx: CrossFixture):
    """For each form that has both CmdQuery and CmdStoreID:

    1. Pre-clean ZZ_STORE_PERSON_ID (otherwise CmdStoreID hits the
       "replace?" yes/no MsgBox).
    2. Open form, populate pickers, fire CmdQuery → CmdStoreID via the
       Form_Timer chain.
    3. Read the source scratch table and ZZ_STORE_PERSON_ID.
    4. Assert ZZ_STORE_PERSON_ID == DISTINCT(scratch.<personid_col>),
       filtered by storeid_source_filter when set (AssociationPairs).
    5. Assert PersonIDSource.SourceForm == personid_source_label
       (skipped for forms whose handler omits the UPDATE).
    """
    spec = fx.spec
    # 1. clear any leftover stored ids
    vba.exec_sql("DELETE FROM ZZ_STORE_PERSON_ID")

    # 2. seed inputs + chain CmdQuery → CmdStoreID via Form_Timer
    _seed_query_inputs(vba, fx)
    vba.click_chain_via_timer(
        spec.name, [spec.cmd_name, "CmdStoreID"], sleep_after=2.0,
    )

    # CmdQuery_Click writes a "<short>:DONE" marker after the chain
    # block fires.  Wait for it to make sure CmdStoreID has run.
    short = spec.name
    ok = vba._wait_for_done(short)  # inherit DEFAULT_VBA_TIMEOUT (B5)
    assert ok, f"[{spec.name}] DONE marker for {short} not seen"

    # 3. compare the two sets
    src_filter = spec.storeid_source_filter or ""
    src_ids = _distinct_personids(
        vba, spec.storeid_source_table, spec.storeid_source_col,
        where=src_filter,
    )
    store_ids = _distinct_personids(vba, "ZZ_STORE_PERSON_ID", "c_personid")
    assert src_ids, (
        f"[{spec.name}] source {spec.storeid_source_table} is empty after "
        f"CmdQuery — query fixture broken?"
    )
    assert store_ids == src_ids, (
        f"[{spec.name}] ZZ_STORE_PERSON_ID set != DISTINCT("
        f"{spec.storeid_source_table}.{spec.storeid_source_col}). "
        f"Only in scratch: {sorted(src_ids - store_ids)[:5]}; "
        f"only in store: {sorted(store_ids - src_ids)[:5]}"
    )

    # 4. PersonIDSource label
    if spec.personid_source_label is not None:
        cur = vba.conn.cursor()
        cur.execute(
            "SELECT SourceForm FROM PersonIDSource WHERE LineNum = 1"
        )
        row = cur.fetchone()
        cur.close()
        assert row is not None, "PersonIDSource has no LineNum=1 row"
        assert str(row[0]) == spec.personid_source_label, (
            f"[{spec.name}] PersonIDSource.SourceForm = {row[0]!r}; "
            f"expected {spec.personid_source_label!r}"
        )

    print(f"[{spec.name}] StoreID OK — {len(store_ids)} ids stored")


# ---------------------------------------------------------------- RecallID
def _recallid_marks(spec: FormSpec):
    """LookAtNetworks Form_Open binds two large subforms
    (`ZZ_SOCIAL_NETWORK`, `ZZ_SOCIAL_NETWORK_AGGREGATED`) to recordsets
    that load the entire prior session's network data; under
    DispatchEx-driven Access this open call hangs indefinitely.  The
    matrix test already skips Networks for the same family of reasons
    (see test_vba_matrix_all_forms `_xfail_marks`); the recall test
    inherits the skip until we've sorted out a smaller bound or a
    pre-open subform-disconnect.  AssociationPairs / Kinship /
    GroupData all open cleanly."""
    if spec.name == "LookAtNetworks":
        return pytest.mark.skip(
            reason="LookAtNetworks CmdRun times out on high-degree "
                   "anchors — same root family as matrix Networks "
                   "CmdRun timeout (PR AA: Form_Open opens fine).  "
                   "See reports/CBDB_Issues_Report_EN.md."
        )
    return ()


@pytest.mark.parametrize(
    "spec",
    [pytest.param(s, marks=_recallid_marks(s)) for s in _recallid_specs()],
    ids=lambda s: s.name,
)
def test_recallid_repopulates_target(vba: VbaSession, spec: FormSpec):
    """For each form with a CmdRecallID button:

    1. Pre-populate ZZ_STORE_PERSON_ID with a known set of person ids.
    2. Pre-clean ZZ_SCRATCH_IMPORT_PEOPLE so the "replace?" prompt
       isn't triggered.
    3. Open form, fire CmdRecallID via Form_Timer.
    4. Assert ZZ_SCRATCH_IMPORT_PEOPLE.c_person_id == seeded set.
    """
    # Use a small known-existing set of person ids (Wang Anshi 1762,
    # Sima Guang 1294, Zhao Tingmei 526 — all referenced elsewhere in
    # the test suite as canonical fixtures).
    seed = [1762, 1294, 526]

    # 1+2. seed + clean
    vba.exec_sql("DELETE FROM ZZ_STORE_PERSON_ID")
    for pid in seed:
        vba.exec_sql(
            f"INSERT INTO ZZ_STORE_PERSON_ID (c_personid) VALUES ({pid})"
        )
    vba.exec_sql(f"DELETE FROM [{spec.recallid_target_table}]")
    vba._refresh_access_cache()

    # 3. open form, fire CmdRecallID via timer.  CmdRecallID has no
    # autodetect-injected DONE marker (it's not CmdQuery/CmdRun), so
    # we just sleep then poll the target table for the seeded count.
    vba.open_form(spec.name)
    vba.click_chain_via_timer(
        spec.name, ["CmdRecallID"], sleep_after=3.0,
    )

    # 4. assert
    got = _distinct_personids(
        vba, spec.recallid_target_table, "c_person_id"
    )
    expected = set(seed)
    assert got == expected, (
        f"[{spec.name}] {spec.recallid_target_table} = {sorted(got)}; "
        f"expected {sorted(expected)}"
    )
    print(f"[{spec.name}] RecallID OK — {len(got)} ids restored")


# ---------------------------------------------------------------- Round-trip
def test_storeid_recallid_roundtrip_entry_to_kinship(vba: VbaSession):
    """End-to-end user workflow:
       LookAtEntry → CmdQuery → CmdStoreID
       (close form)
       LookAtKinship → CmdRecallID
       Assert ZZ_SCRATCH_IMPORT_PEOPLE matches LookAtEntry result.

    This is the only thing CmdStoreID/CmdRecallID exist for: "I just
    queried a set of people on form A, now run them through form B."
    """
    entry_fx = _entry_fixture()
    if entry_fx is None:
        pytest.skip("test_inputs.json missing — run discover_test_inputs.py")
    entry = entry_fx.spec

    # Phase 1: query Entry, store
    vba.exec_sql("DELETE FROM ZZ_STORE_PERSON_ID")
    vba.exec_sql("DELETE FROM ZZ_SCRATCH_IMPORT_PEOPLE")
    _seed_query_inputs(vba, entry_fx)
    vba.click_chain_via_timer(
        entry.name, [entry.cmd_name, "CmdStoreID"], sleep_after=2.0,
    )
    ok = vba._wait_for_done(entry.name)  # inherit DEFAULT_VBA_TIMEOUT (B5)
    assert ok, "Entry DONE marker not seen — chain didn't complete"

    # Capture what was stored.
    stored = _distinct_personids(vba, "ZZ_STORE_PERSON_ID", "c_personid")
    assert stored, "Entry CmdStoreID produced no stored ids"

    try:
        vba.close_form(entry.name)
    except Exception:
        pass

    # Phase 2: Kinship form, recall
    kin = ALL_SPECS["LookAtKinship"]
    vba.open_form(kin.name)
    vba.click_chain_via_timer(
        kin.name, ["CmdRecallID"], sleep_after=3.0,
    )

    recalled = _distinct_personids(
        vba, kin.recallid_target_table, "c_person_id"
    )
    assert recalled == stored, (
        f"Round-trip mismatch — stored {len(stored)} ids on Entry, "
        f"recalled {len(recalled)} ids on Kinship.  "
        f"Lost: {sorted(stored - recalled)[:5]}; "
        f"extra: {sorted(recalled - stored)[:5]}"
    )
    print(f"Round-trip OK — {len(stored)} ids survived Entry→Kinship")
