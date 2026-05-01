"""
Differential tests: real VBA `CmdQuery_Click` vs Python SQL replay.

For each fixture:
  1. Use VbaSession to fire the actual VBA handler in Access via
     pywinauto.click_input(). This populates ZZ_SCRATCH_<XXX>.
  2. Read ZZ_SCRATCH_<XXX> from the .mdb (the VBA's output).
  3. Run cbdb_replay.<form>.run(...) with the same inputs.
  4. Compare DISTINCT primary identity sets.
  5. Difference => bug in EITHER the VBA OR our Python replay.

This is the only test pattern in the suite that can find latent VBA
bugs.  The SQL-replay-vs-golden tests can't (Python replay is derived
from the same VBA).

Cost: ~12-15s per test (Access startup + form load).  For now we
parametrize sparingly. Extend by adding rows to FIXTURES.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from cbdb_driver.vba_session import VbaSession, make_fixture
from cbdb_replay import lookatentry as le


ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK = ROOT / "analysis" / "_diff_test_copy.mdb"


@pytest.fixture(scope="function")
def vba():
    yield from make_fixture(SRC, WORK)


# ----------------------------------------------------------------------
# LookAtEntry differential fixtures
#
# Each row: (id, picker_codes, picker_addrs, controls, py_inputs)
# - picker_codes: rows for ZZ_SCRATCH_ENTRY_CODE
# - picker_addrs: rows for ZZ_SCRATCH_ADDR
# - controls: dict of Control name → value
# - py_inputs: lookatentry.EntryQueryInputs for the Python replay
# ----------------------------------------------------------------------

LE_FIXTURES = [
    (
        "yin_general_kaifeng_900_1100_indexyears",
        [118],                        # ZZ_SCRATCH_ENTRY_CODE
        [100658],                     # ZZ_SCRATCH_ADDR (Kaifeng)
        {                              # form controls
            "TxtFromYear": 900,
            "TxtToYear": 1100,
            "FrameYears": 2,
            "TxtEntryDesc": "yin privilege: general",
            "TxtTypeCode": "N/A",
        },
        le.EntryQueryInputs(
            entry_codes=[118],
            addr_ids=[100658],
            addr_field="person",
            year_mode="index",
            from_year=900, to_year=1100,
        ),
    ),
    (
        "jinshi_general_kaifeng_900_1100_entryyears",
        [36],                          # examination: jinshi (general)
        [100658],
        {
            "TxtFromYear": 900,
            "TxtToYear": 1100,
            "FrameYears": 1,            # entry years
            "TxtEntryDesc": "examination: jinshi (general)",
            "TxtTypeCode": "N/A",
        },
        le.EntryQueryInputs(
            entry_codes=[36],
            addr_ids=[100658],
            addr_field="person",
            year_mode="entry",
            from_year=900, to_year=1100,
        ),
    ),
]


@pytest.mark.parametrize("case_id,codes,addrs,controls,py_inp",
                         LE_FIXTURES,
                         ids=[r[0] for r in LE_FIXTURES])
def test_lookatentry_vba_vs_python(vba: VbaSession,
                                    case_id, codes, addrs, controls, py_inp):
    """Differential test: VBA result should match Python replay result.

    Difference between the two sets of c_personid (with multiplicity)
    means EITHER the VBA OR our Python replay has a bug.  Investigate
    the diff to determine which.
    """
    # ----- run VBA path -----
    vba.set_picker_codes("ZZ_SCRATCH_ENTRY_CODE", codes,
                         column="c_entry_code")
    vba.set_picker_addrs(addrs)
    vba.open_form("LookAtEntry")
    for ctl, val in controls.items():
        vba.set_control("LookAtEntry", ctl, val)
    n_vba = vba.click_button_and_wait_table(
        "Run Query", form="LookAtEntry",
        result_table="ZZ_SCRATCH_ENTRY",
        force_enable_ctl="CmdQuery",
        timeout=30,
    )
    assert n_vba > 0, f"[{case_id}] VBA produced 0 rows"
    df_vba = vba.read("ZZ_SCRATCH_ENTRY")

    # ----- run Python replay path (using vba's pyodbc connection) -----
    df_py = le.run(vba.conn, py_inp)

    # ----- compare distinct person sets -----
    set_vba = set(df_vba["c_personid"].dropna().astype(int).tolist())
    set_py = set(df_py["c_personid"].dropna().astype(int).tolist())
    only_vba = set_vba - set_py
    only_py = set_py - set_vba
    print(f"\n[{case_id}] VBA={len(set_vba)} Py={len(set_py)} "
          f"VBA-only={len(only_vba)} Py-only={len(only_py)}")
    if only_vba or only_py:
        msg = (f"[{case_id}] VBA and Python replay disagree:\n"
               f"  in VBA but not Py: {sorted(only_vba)[:20]}"
               f"{'...' if len(only_vba) > 20 else ''}\n"
               f"  in Py but not VBA: {sorted(only_py)[:20]}"
               f"{'...' if len(only_py) > 20 else ''}\n"
               f"  total VBA: {len(set_vba)}, total Py: {len(set_py)}")
        # We assert hard: any disagreement is a real signal
        assert False, msg


def test_lookatentry_post_query_button_states(vba: VbaSession):
    """Independent of differential: just verify that after a successful
    query, the export buttons toggle Enabled correctly per the VBA's
    own state-machine logic."""
    vba.set_picker_codes("ZZ_SCRATCH_ENTRY_CODE", [118],
                         column="c_entry_code")
    vba.set_picker_addrs([100658])
    vba.open_form("LookAtEntry")
    for ctl, val in {
        "TxtFromYear": 900, "TxtToYear": 1100, "FrameYears": 2,
        "TxtEntryDesc": "yin privilege: general", "TxtTypeCode": "N/A",
    }.items():
        vba.set_control("LookAtEntry", ctl, val)

    # Pre-query: all export buttons should be disabled
    assert vba.get_control_property("LookAtEntry", "CmdGIS", "Enabled") is False
    assert vba.get_control_property("LookAtEntry", "CmdNeo4j", "Enabled") is False
    assert vba.get_control_property("LookAtEntry", "CmdStoreID", "Enabled") is False

    # Run the query
    n = vba.click_button_and_wait_table(
        "Run Query", form="LookAtEntry",
        result_table="ZZ_SCRATCH_ENTRY",
        force_enable_ctl="CmdQuery",
        timeout=30,
    )
    assert n > 0

    # Post-query: all export buttons should be enabled
    assert vba.get_control_property("LookAtEntry", "CmdGIS", "Enabled") is True
    assert vba.get_control_property("LookAtEntry", "CmdNeo4j", "Enabled") is True
    assert vba.get_control_property("LookAtEntry", "CmdStoreID", "Enabled") is True
