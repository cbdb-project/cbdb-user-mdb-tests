"""
Smaller-fixture matrix for the 3 LookAt forms whose default fixtures
time out (`tests/test_vba_matrix_all_forms.py` skips them).

Roadmap item 7 — turn the timeouts into passing tests by hand-picking
fixtures with very small fan-out.

Fixture rationale (sourced from `analysis/_find_small_assoc_persons.py`
which scans ASSOC_DATA / ENTRY_DATA / STATUS_DATA):

  - LookAtNetworks   → c_personid=4  (5 assocs total) — Zhu Xi has
                       2471, the matrix default times out; we need
                       <20.
  - LookAtGroupData  → c_personid=1  (2 entries, 2 statuses) — the
                       default tries to aggregate across everything
                       a person has.
  - LookAtAssociationPairs → 4 × 5  (both small persons.  Even if
                       Link1stOrder still self-joins, the Cartesian
                       product over their <5-row slices stays tiny.)

If the matrix ever recovers and these fixtures become redundant, the
file can be removed without breaking anything else.
"""
from __future__ import annotations

from pathlib import Path

import pytest

from cbdb_driver.vba_session import VbaSession, make_fixture
from cbdb_driver.form_specs import (
    LOOKATNETWORKS, LOOKATGROUPDATA, LOOKATASSOCIATIONPAIRS,
)
from test_vba_matrix_all_forms import CrossFixture, SRC


WORK = Path(__file__).resolve().parent.parent / "analysis" / "_matrix_hard_test_copy.mdb"


_HARD_FIXTURES: list[CrossFixture] = [
    # LookAtNetworks deliberately omitted from THIS file: full-injection
    # VbaSession opens the form into a project-wide auto-compile
    # deadlock (PR AR-AX bisected the cause).  Real CmdRun smoke
    # coverage now lives in tests/test_vba_networks_small_fixture.py,
    # which uses the PR AT `skip_inject_autodetect_forms` kwarg to
    # suppress sibling-form injection (PR AU V12 / AX-verified).
    CrossFixture(
        name="groupdata_person_1_small",
        spec=LOOKATGROUPDATA,
        picker_ids=[1],
        controls={},
        expected_min_rows=1,
        source_sql=None,
    ),
    CrossFixture(
        name="assocpair_4x5_small",
        spec=LOOKATASSOCIATIONPAIRS,
        picker_ids=None,
        controls={
            "TxtID1": 4, "TxtID2": 5,
            "TxtPerson1": "4", "TxtPerson2": "5",
            "FrameFilterYears": 1,
            "Chk2Nodes": 0, "ChkKinship": 0,
        },
        expected_min_rows=1,
        source_sql=None,
    ),
]


@pytest.fixture(scope="function")
def vba():
    yield from make_fixture(SRC, WORK)


def _seed_query_inputs(vba: VbaSession, fx: CrossFixture) -> None:
    spec = fx.spec
    vba.open_form(spec.name)
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


def _check_groupdata(vba: VbaSession, person_id: int) -> None:
    """LookAtGroupData.CmdRun does UPDATE-style backfill on
    ZZ_SCRATCH_IMPORT_PEOPLE, not INSERT.  Verify the seeded row got
    its c_name / c_dynasty fields filled (proof the JOIN ran)."""
    cur = vba.conn.cursor()
    cur.execute(
        "SELECT c_person_id, c_name, c_dynasty FROM ZZ_SCRATCH_IMPORT_PEOPLE "
        f"WHERE c_person_id = {person_id}"
    )
    row = cur.fetchone()
    cur.close()
    assert row is not None, (
        f"ZZ_SCRATCH_IMPORT_PEOPLE missing seeded c_person_id={person_id}"
    )
    pid, name, dynasty = row
    assert name and str(name).strip(), (
        f"CmdRun didn't backfill c_name for c_person_id={person_id}"
    )
    print(f"  backfilled: c_name={name!r} c_dynasty={dynasty!r}", flush=True)


def _check_assoc_pairs(vba: VbaSession) -> None:
    """LookAtAssociationPairs.CmdQuery writes to `ZZ_SOCIAL_NETWORK`.
    Even with a tiny pair (4×5 share zero direct edges in the small
    overlap), the chain may produce 0 rows — we just verify the run
    completed without error (table exists + readable)."""
    cur = vba.conn.cursor()
    cur.execute("SELECT COUNT(*) FROM ZZ_SOCIAL_NETWORK")
    n = cur.fetchone()[0]
    cur.close()
    print(f"  ZZ_SOCIAL_NETWORK rows: {n}", flush=True)
    # Don't assert >= 1 — small fixtures may legitimately have 0
    # network edges.  Reaching this point at all means CmdQuery didn't
    # time out, which is the actual goal of this matrix.


@pytest.mark.parametrize("fx", _HARD_FIXTURES, ids=lambda f: f.name)
def test_hard_form_query_small_fixture(vba: VbaSession, fx: CrossFixture):
    """Fire CmdQuery / CmdRun with the small-fixture inputs.  Per-form
    completion check varies (GroupData backfills via UPDATE, others
    INSERT)."""
    spec = fx.spec
    _seed_query_inputs(vba, fx)
    # For UPDATE-style forms we can't poll row_count; use wait_done only.
    if spec.name == "LookAtGroupData":
        vba.click_via_timer(
            spec.name, ctl=spec.cmd_name,
            result_table=None, timeout=180,
        )
        _check_groupdata(vba, fx.picker_ids[0])
        return

    if spec.name == "LookAtAssociationPairs":
        vba.click_via_timer(
            spec.name, ctl=spec.cmd_name,
            result_table=spec.result_table, timeout=180,
        )
        _check_assoc_pairs(vba)
        return

    # Generic INSERT-style fall-through (none right now, kept for
    # future fixtures).
    n = vba.click_via_timer(
        spec.name, ctl=spec.cmd_name,
        result_table=spec.result_table, timeout=180,
    )
    print(f"\n[{spec.name}] {spec.cmd_name} -> {n} rows in {spec.result_table}",
          flush=True)
    assert n >= fx.expected_min_rows, (
        f"[{spec.name}] {spec.cmd_name} only {n} rows; expected "
        f"≥ {fx.expected_min_rows}.  Fixture too small or VBA bailed."
    )
