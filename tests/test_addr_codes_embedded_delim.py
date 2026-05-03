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

from pathlib import Path

import pyodbc
import pytest

ROOT = Path(__file__).resolve().parent.parent
USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"

# Known counts as of CBDB_20260430_DATA.mdb.  When upstream fixes
# the data these will drop to 0; treat any change as a signal to
# re-evaluate the GIS depth check failure documented in
# analysis/gis_status_embedded_delim_root_cause.md.
_EXPECTED_BOM_ROWS_C_NAME = 315
_EXPECTED_BOM_ROWS_C_NAME_CHN = 315

# The one row reachable from STATUS_DATA.c_status_code=40 (PR T's
# LookAtStatus fixture).  This is the row that flips row 11476 of
# the GIS export from a clean 9-cell line into a 10-cell line.
_KNOWN_REACHABLE_DIRTY_ADDR_ID = 702559    # Wei Shi 尉氏
_KNOWN_REACHABLE_PERSON_ID = 29619         # Ruan Fu 阮孚


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


def test_addr_codes_has_known_bom_dirty_rows(conn) -> None:
    """315 rows of ADDR_CODES carry a stray BOM prefix in c_name
    and c_name_chn.  See analysis/gis_status_embedded_delim_root_
    cause.md for the full chain to silent column-misalignment in
    GIS exports."""
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

    # Don't be exact — allow ±10% drift in case a few rows get
    # cleaned up incrementally.  But require *some* dirty rows so
    # the LookAtStatus row 11476 test stays meaningful.
    assert n_bom_name >= _EXPECTED_BOM_ROWS_C_NAME * 0.9, (
        f"ADDR_CODES.c_name BOM-prefixed row count dropped from "
        f"{_EXPECTED_BOM_ROWS_C_NAME} to {n_bom_name}.  Upstream "
        f"may have cleaned the data — re-run the LookAtStatus "
        f"GIS test (row 11476) and decide if this regression test "
        f"can be removed.  See "
        f"analysis/gis_status_embedded_delim_root_cause.md."
    )
    assert n_bom_name_chn >= _EXPECTED_BOM_ROWS_C_NAME_CHN * 0.9, (
        f"ADDR_CODES.c_name_chn BOM-prefixed row count dropped "
        f"from {_EXPECTED_BOM_ROWS_C_NAME_CHN} to {n_bom_name_chn}."
    )


def test_known_reachable_dirty_addr_present(conn) -> None:
    """addr_id 702559 (Wei Shi / 尉氏) is the one BOM-dirty row
    reachable from status_code=40 in the PR T LookAtStatus
    fixture.  This is the row that produces the 10-cell line at
    GIS row 11476."""
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
    assert n is not None and n.startswith("﻿"), (
        f"ADDR_CODES.c_name for {_KNOWN_REACHABLE_DIRTY_ADDR_ID} "
        f"no longer starts with U+FEFF: {n!r}.  If this row was "
        f"cleaned up, re-check the LookAtStatus GIS row 11476 "
        f"failure — it may now pass."
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
