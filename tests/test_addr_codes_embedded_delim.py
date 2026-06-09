"""ADDR_CODES dirty-data regression test (PR U).

Companion to PR T's failing LookAtStatus GIS depth check (row 11476).
Root-cause evidence: analysis/gis_status_embedded_delim_root_cause.md.

This test asserts the *current state* of ADDR_CODES — that 315 rows
have a stray BOM (U+FEFF) prefix in `c_name` and `c_name_chn`.  The
JET UPDATE/INSERT pipeline mangles those BOM-prefixed values into
strings containing literal TAB characters, which then breaks
CmdGIS's tab-separated output.

If upstream cleans the BOM prefixes out of ADDR_CODES, this test
will fail and prompt re-evaluation:

  - Does the LookAtStatus GIS depth check (row 11476) now pass?
  - Are there still other code paths that introduce embedded
    delimiters from elsewhere?

The test is intentionally narrow: it does NOT replace the GIS
depth check, only documents the *upstream* dirty-data condition
that the depth check legitimately catches downstream.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pyodbc
import pytest

ROOT = Path(__file__).resolve().parent.parent
USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"

# BOM dirty-row count is a PER-BUILD data characteristic, not a constant.
# Upstream cleaned the stray U+FEFF prefixes between builds:
#   20260430 -> 315 dirty rows (Issue #20 reproduces)
#   20260602 -> 0   dirty rows (cleaned; Issue #20 dormant on this build)
# Anchor the expectation to the build so the test asserts the RIGHT thing
# per build instead of silently passing/failing on drift (plan B8).  A
# build not in this map => xfail "not calibrated for build X" — measure
# the count on that build and add it here.
_BOM_ROWS_BY_BUILD = {
    "20260430": 315,
    "20260602": 0,
}

# The one row that was reachable-and-dirty from STATUS_DATA.c_status_code=40
# (PR T's LookAtStatus fixture) on the dirty builds — flips GIS row 11476
# from a clean 9-cell line into a 10-cell line.
_KNOWN_REACHABLE_DIRTY_ADDR_ID = 702559    # Wei Shi 尉氏
_KNOWN_REACHABLE_PERSON_ID = 29619         # Ruan Fu 阮孚


def _expected_bom_rows() -> int:
    """Expected BOM-dirty row count for the build in data/, or xfail if the
    build isn't calibrated yet."""
    _ana = str(ROOT / "analysis")
    if _ana not in sys.path:
        sys.path.insert(0, _ana)
    from build_stamp import current_build  # noqa: E402
    build = current_build(ROOT)
    if build not in _BOM_ROWS_BY_BUILD:
        pytest.xfail(
            f"ADDR_CODES BOM-row count not calibrated for build {build!r}.  "
            f"Measure it (count rows where c_name contains U+FEFF) and add "
            f"'{build}': <count> to _BOM_ROWS_BY_BUILD."
        )
    return _BOM_ROWS_BY_BUILD[build]


def _open() -> pyodbc.Connection:
    return pyodbc.connect(
        "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={USER_MDB};", autocommit=True, readonly=True)


@pytest.fixture(scope="module")
def conn() -> pyodbc.Connection:
    if not USER_MDB.exists():
        pytest.skip(f"{USER_MDB} not present")
    c = _open()
    yield c
    c.close()


def test_addr_codes_bom_dirty_rows_match_build(conn) -> None:
    """ADDR_CODES BOM-prefixed row count matches the calibrated value for
    THIS build.  Dirty builds (e.g. 20260430: 315) reproduce Issue #20;
    clean builds (e.g. 20260602: 0) have it dormant.  See
    analysis/gis_status_embedded_delim_root_cause.md."""
    expected = _expected_bom_rows()
    cur = conn.cursor()
    cur.execute("SELECT c_name, c_name_chn FROM ADDR_CODES")
    n_bom_name = 0
    n_bom_name_chn = 0
    for r in cur.fetchall():
        n, nz = r
        if n is not None and "﻿" in n:
            n_bom_name += 1
        if nz is not None and "﻿" in nz:
            n_bom_name_chn += 1

    if expected == 0:
        # Clean build: the dirty-data condition (and Issue #20) is dormant.
        assert n_bom_name == 0 and n_bom_name_chn == 0, (
            f"build calibrated to 0 BOM rows but found "
            f"c_name={n_bom_name}, c_name_chn={n_bom_name_chn}.  Recalibrate "
            f"_BOM_ROWS_BY_BUILD or investigate new dirty data."
        )
    else:
        # Dirty build: require ~expected rows (±10% for incremental cleanup)
        # so the downstream LookAtStatus GIS row-11476 check stays meaningful.
        assert n_bom_name >= expected * 0.9, (
            f"ADDR_CODES.c_name BOM-prefixed row count {n_bom_name} is below "
            f"90% of the calibrated {expected} for this build.  If upstream "
            f"cleaned the data, recalibrate _BOM_ROWS_BY_BUILD and re-check "
            f"the LookAtStatus GIS test (row 11476)."
        )
        assert n_bom_name_chn >= expected * 0.9, (
            f"ADDR_CODES.c_name_chn BOM-prefixed row count {n_bom_name_chn} "
            f"is below 90% of the calibrated {expected}."
        )


def test_known_reachable_dirty_addr_present(conn) -> None:
    """addr_id 702559 (Wei Shi / 尉氏) is the row reachable from
    status_code=40 in the PR T LookAtStatus fixture — on dirty builds it
    is BOM-prefixed (producing the 10-cell GIS line at row 11476); on
    clean builds it must exist but carry NO BOM."""
    expected = _expected_bom_rows()
    cur = conn.cursor()
    cur.execute(
        "SELECT c_addr_id, c_name, c_name_chn FROM ADDR_CODES "
        f"WHERE c_addr_id = {_KNOWN_REACHABLE_DIRTY_ADDR_ID}"
    )
    rows = cur.fetchall()
    assert len(rows) == 1, (
        f"Expected exactly 1 ADDR_CODES row for c_addr_id="
        f"{_KNOWN_REACHABLE_DIRTY_ADDR_ID}, got {len(rows)}."
    )
    aid, n, nz = rows[0]

    if expected == 0:
        # Clean build: the row exists but the BOM was stripped upstream.
        assert (n is None or not n.startswith("﻿")) and \
               (nz is None or not nz.startswith("﻿")), (
            f"build calibrated clean (0 BOM rows) but c_addr_id="
            f"{_KNOWN_REACHABLE_DIRTY_ADDR_ID} still BOM-prefixed: "
            f"c_name={n!r}, c_name_chn={nz!r}.  Recalibrate _BOM_ROWS_BY_BUILD."
        )
        return

    assert n is not None and n.startswith("﻿"), (
        f"ADDR_CODES.c_name for {_KNOWN_REACHABLE_DIRTY_ADDR_ID} "
        f"no longer starts with U+FEFF: {n!r}.  If this row was "
        f"cleaned up, recalibrate _BOM_ROWS_BY_BUILD to 0 for this "
        f"build and re-check the LookAtStatus GIS row 11476 failure."
    )
    assert nz is not None and nz.startswith("﻿"), (
        f"ADDR_CODES.c_name_chn for "
        f"{_KNOWN_REACHABLE_DIRTY_ADDR_ID} no longer starts with "
        f"U+FEFF: {nz!r}."
    )

    # And cross-check the BIOG_ADDR_DATA / STATUS_DATA reachability.
    cur.execute(
        "SELECT COUNT(*) FROM BIOG_ADDR_DATA bad "
        "INNER JOIN STATUS_DATA sd ON sd.c_personid = bad.c_personid "
        "WHERE sd.c_status_code = 40 "
        f"  AND bad.c_addr_id = {_KNOWN_REACHABLE_DIRTY_ADDR_ID}"
    )
    n_reach = cur.fetchone()[0]
    assert n_reach >= 1, (
        f"addr_id {_KNOWN_REACHABLE_DIRTY_ADDR_ID} no longer "
        f"reachable from status_code=40 — the LookAtStatus GIS "
        f"row 11476 trigger may have moved.  Re-run the export "
        f"probe (analysis/probe_status_gis_export_bytes.py) to "
        f"find the new offending row."
    )
