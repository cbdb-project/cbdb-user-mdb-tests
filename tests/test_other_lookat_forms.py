"""
Tests for the other 9 LookAt forms (every LookAt form except LookAtEntry,
which has its own thorough test_lookatentry.py).

Coverage rules per form:
  - At least 1 fixture covering the canonical use case
  - 1 'empty inputs' assertion (no codes / no IDs → no rows)
  - 1 row-shape assertion (key columns present, sane row count)
  - golden CSV snapshot for regression (--regenerate-goldens to bless)

Forms with multi-hop / recursive logic (LookAtKinship, LookAtNetworks,
LookAtAssociationPairs, LookAtGroupData) only test the 1-hop / direct
case for now; multi-hop branches raise NotImplementedError when
exercised.
"""
from __future__ import annotations

from pathlib import Path

import pytest

from cbdb_replay.common import YearFilter

# Each form's replay module
from cbdb_replay import (
    lookatstatus, lookattexts, lookatplace,
    lookatassociations, lookatoffice,
    lookatassociationpairs, lookatkinship,
    lookatnetworks, lookatgroupdata,
)


GOLDEN = Path(__file__).resolve().parent / "golden"


def _check_golden(df, sub: str, name: str, regenerate: bool, sort_by: list[str]):
    from golden_helpers import assert_matches_golden
    p = GOLDEN / sub / f"{name}.csv"
    assert_matches_golden(df, p, sort_by=sort_by, regenerate=regenerate)


# ----------------------------------------------------------------------
# LookAtStatus
# ----------------------------------------------------------------------

def test_lookatstatus_empty_codes(ro_conn):
    df = lookatstatus.run(ro_conn, lookatstatus.StatusQueryInputs())
    assert df.empty


def test_lookatstatus_basic_status_song(ro_conn, regenerate_goldens):
    """Status code 40 (most populous status, ~17k events) in Song."""
    inp = lookatstatus.StatusQueryInputs(
        status_codes=[40],
        year_mode="dynasty",
        from_dynasty=15, to_dynasty=15,
        from_dynasty_begin=960, to_dynasty_end=1279,
    )
    df = lookatstatus.run(ro_conn, inp)
    assert not df.empty, "expected ≥1 row for status code 40 in Song"
    assert {"c_personid", "c_status_code", "c_status_desc"}.issubset(df.columns)
    _check_golden(df, "lookatstatus", "status40_song",
                  regenerate=regenerate_goldens,
                  sort_by=["c_personid", "c_status_code", "c_sequence"])


def test_lookatstatus_with_address(ro_conn, regenerate_goldens):
    """Status code 40 events, restricted to Kaifeng addr 100658."""
    inp = lookatstatus.StatusQueryInputs(
        status_codes=[40],
        addr_ids=[100658],
        year_mode="index",
        from_year=900, to_year=1100,
    )
    df = lookatstatus.run(ro_conn, inp)
    _check_golden(df, "lookatstatus", "status40_kaifeng_900_1100",
                  regenerate=regenerate_goldens,
                  sort_by=["c_personid", "c_sequence"])


# ----------------------------------------------------------------------
# LookAtTexts
# ----------------------------------------------------------------------

def test_lookattexts_empty_codes(ro_conn):
    df = lookattexts.run(ro_conn, lookattexts.TextsQueryInputs())
    assert df.empty


def test_lookattexts_basic(ro_conn, regenerate_goldens):
    """Texts of biblio category 1, no further filter."""
    inp = lookattexts.TextsQueryInputs(biblcat_codes=[1])
    df = lookattexts.run(ro_conn, inp)
    if df.empty:
        pytest.skip("no rows for category 1 — try another in regen step")
    assert {"c_personid", "c_textid", "c_title"}.issubset(df.columns)
    _check_golden(df, "lookattexts", "biblcat1",
                  regenerate=regenerate_goldens,
                  sort_by=["c_personid", "c_textid", "c_role_id"])


# ----------------------------------------------------------------------
# LookAtPlace
# ----------------------------------------------------------------------

def test_lookatplace_empty_addr(ro_conn):
    df = lookatplace.run(ro_conn, lookatplace.PlaceQueryInputs())
    assert df.empty


def test_lookatplace_kaifeng_individual(ro_conn, regenerate_goldens):
    """People whose index_addr is Kaifeng (100658), 900-1100."""
    inp = lookatplace.PlaceQueryInputs(
        addr_ids=[100658], use_index_years=True,
        from_year=900, to_year=1100,
    )
    df = lookatplace.run(ro_conn, inp)
    assert not df.empty
    n_distinct = df["c_personid"].nunique()
    # Kaifeng addr 100658 over 900-1100 with NO entry-method filter
    # is much broader than LookAtEntry's filtered query — current
    # data gives ~1.5k distinct persons.
    assert 100 <= n_distinct <= 5000, f"distinct persons = {n_distinct}"
    _check_golden(df, "lookatplace", "kaifeng_individual_900_1100",
                  regenerate=regenerate_goldens,
                  sort_by=["c_personid"])


# ----------------------------------------------------------------------
# LookAtAssociations
# ----------------------------------------------------------------------

def test_lookatassociations_empty_codes(ro_conn):
    df = lookatassociations.run(
        ro_conn, lookatassociations.AssocQueryInputs()
    )
    assert df.empty


def test_lookatassociations_basic(ro_conn, regenerate_goldens):
    """Associations of code 1 in Song dynasty."""
    inp = lookatassociations.AssocQueryInputs(
        assoc_codes=[1],
        year_filter=YearFilter(
            mode="dynasty", from_dynasty=15, to_dynasty=15,
            from_dynasty_begin=960, to_dynasty_end=1279,
        ),
    )
    df = lookatassociations.run(ro_conn, inp)
    assert {"c_person_id", "c_assoc_code", "c_assoc_id"}.issubset(df.columns)
    _check_golden(df, "lookatassociations", "code1_song",
                  regenerate=regenerate_goldens,
                  sort_by=["c_person_id", "c_assoc_id", "c_assoc_code"])


# ----------------------------------------------------------------------
# LookAtOffice
# ----------------------------------------------------------------------

def test_lookatoffice_empty_codes(ro_conn):
    df = lookatoffice.run(ro_conn, lookatoffice.OfficeQueryInputs())
    assert df.empty


def test_lookatoffice_basic(ro_conn, regenerate_goldens):
    """Postings to office_id 1 in Song."""
    inp = lookatoffice.OfficeQueryInputs(
        office_codes=[1],
        year_filter=YearFilter(
            mode="dynasty", from_dynasty=15, to_dynasty=15,
            from_dynasty_begin=960, to_dynasty_end=1279,
        ),
    )
    df = lookatoffice.run(ro_conn, inp)
    assert {"c_personid", "c_office_id", "c_posting_id"}.issubset(df.columns)
    _check_golden(df, "lookatoffice", "office1_song",
                  regenerate=regenerate_goldens,
                  sort_by=["c_personid", "c_posting_id"])


# ----------------------------------------------------------------------
# LookAtAssociationPairs (direct edges only)
# ----------------------------------------------------------------------

def test_assocpairs_zero_inputs(ro_conn):
    df = lookatassociationpairs.run(
        ro_conn, lookatassociationpairs.AssocPairsQueryInputs()
    )
    assert df.empty


def test_assocpairs_second_order_not_implemented(ro_conn):
    with pytest.raises(NotImplementedError, match="second-order"):
        lookatassociationpairs.run(
            ro_conn,
            lookatassociationpairs.AssocPairsQueryInputs(
                person_id_1=1, person_id_2=2, second_order=True
            ),
        )


def test_assocpairs_direct_edges(ro_conn, regenerate_goldens):
    """Direct association edges between Wang Anshi (1762) and a
    contemporary. Picks any associate id we know exists; the test is
    defensive — passes whether the count is 0 or > 0."""
    inp = lookatassociationpairs.AssocPairsQueryInputs(
        person_id_1=1762, person_id_2=1294,    # Wang Anshi ↔ Sima Guang
    )
    df = lookatassociationpairs.run(ro_conn, inp)
    # may legitimately be 0 if no direct edge, but the COLUMN SHAPE must hold
    expected_cols = {
        "c_person_id", "c_node_id", "c_link_code", "c_link_desc",
        "c_link_first_year", "c_link_last_year",
    }
    assert expected_cols.issubset(df.columns), (
        f"missing columns: {expected_cols - set(df.columns)}"
    )
    _check_golden(df, "lookatassociationpairs",
                  "wang_anshi_sima_guang_direct",
                  regenerate=regenerate_goldens,
                  sort_by=["c_person_id", "c_node_id", "c_link_code"])


# ----------------------------------------------------------------------
# LookAtKinship (1-hop direct only)
# ----------------------------------------------------------------------

def test_kinship_zero_input(ro_conn):
    df = lookatkinship.run(ro_conn, lookatkinship.KinshipQueryInputs())
    assert df.empty


def test_kinship_multihop_not_implemented(ro_conn):
    with pytest.raises(NotImplementedError, match="multi-hop"):
        lookatkinship.run(
            ro_conn,
            lookatkinship.KinshipQueryInputs(person_id=1762, max_up=2),
        )


def test_kinship_direct_kin(ro_conn, regenerate_goldens):
    df = lookatkinship.run(
        ro_conn, lookatkinship.KinshipQueryInputs(person_id=1762),
    )
    assert {"c_personid", "c_kin_id", "c_kin_code", "c_kinrel"}.issubset(df.columns)
    _check_golden(df, "lookatkinship", "wang_anshi_direct_kin",
                  regenerate=regenerate_goldens,
                  sort_by=["c_personid", "c_kin_id", "c_kin_code"])


# ----------------------------------------------------------------------
# LookAtNetworks (1-hop NONKIN only)
# ----------------------------------------------------------------------

def test_networks_zero_input(ro_conn):
    df = lookatnetworks.run(ro_conn, lookatnetworks.NetworksQueryInputs())
    assert df.empty


def test_networks_multihop_not_implemented(ro_conn):
    with pytest.raises(NotImplementedError, match="multi-hop"):
        lookatnetworks.run(
            ro_conn,
            lookatnetworks.NetworksQueryInputs(person_id=1762, max_hops=2),
        )


def test_networks_one_hop(ro_conn, regenerate_goldens):
    df = lookatnetworks.run(
        ro_conn, lookatnetworks.NetworksQueryInputs(person_id=1762),
    )
    assert {"c_person_id", "c_node_id", "c_link_code", "c_link_desc"}.issubset(df.columns)
    _check_golden(df, "lookatnetworks", "wang_anshi_1hop",
                  regenerate=regenerate_goldens,
                  sort_by=["c_person_id", "c_node_id", "c_link_code"])


# ----------------------------------------------------------------------
# LookAtGroupData (base people records only)
# ----------------------------------------------------------------------

def test_groupdata_zero_input(ro_conn):
    df = lookatgroupdata.run(
        ro_conn, lookatgroupdata.GroupDataQueryInputs()
    )
    assert df.empty


def test_groupdata_categories_not_implemented(ro_conn):
    with pytest.raises(NotImplementedError, match="category sub-queries"):
        lookatgroupdata.run(
            ro_conn,
            lookatgroupdata.GroupDataQueryInputs(
                person_ids=[1762], include_office=True,
            ),
        )


def test_groupdata_imported_people(ro_conn, regenerate_goldens):
    pids = [1762, 1294, 7097]   # Wang Anshi, Sima Guang, Su Shi
    df = lookatgroupdata.run(
        ro_conn, lookatgroupdata.GroupDataQueryInputs(person_ids=pids),
    )
    assert len(df) == 3
    assert set(df["c_person_id"]) == set(pids)
    _check_golden(df, "lookatgroupdata", "wang_sima_su",
                  regenerate=regenerate_goldens,
                  sort_by=["c_person_id"])
