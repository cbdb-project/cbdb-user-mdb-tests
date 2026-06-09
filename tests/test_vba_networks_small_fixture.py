"""LookAtNetworks small-fixture real-VBA test (PR AY).

The general matrix (`tests/test_vba_matrix_all_forms.py`) skips
LookAtNetworks because its CmdRun on the default high-degree
anchor (Zhu Xi, 2 471 assocs) doesn't complete in 120 s.  Even
worse — until PR AR-AX, just opening the form via the matrix's
`VbaSession.open() + open_form()` path appeared to hang
indefinitely.

PR AR-AX bisected that hang to a real Access-side issue:
modifying ANY sibling Form_LookAt* VBA module dirties the
project, and Networks's Form_Open then deadlocks on the
project-wide auto-compile when Networks's
`Forms!LookAtNetworks!<sub>.Form.Recordset` self-reference
during Form_Open is resolved against the dirty modules.

PR AU V12 / AX confirmed the workaround: skip injection for
all 9 sibling forms (Networks's own injection still runs, so
gUsePersonID auto-detect for the picker still fires).  Networks
then opens cleanly and CmdRun completes on small anchors.

This test exercises the verified working recipe end-to-end on
Cao Zhi (c_personid=30270) — chosen because:
  - well-attested historical figure
  - small network (10 assocs / 1 kin direct, est_1hop = 10 per
    `analysis/lookatnetworks_anchor_candidates.md`)
  - PR AX measured ZZ_SOCIAL_NETWORK = 1, ZZ_SCRATCH_PEOPLE = 2
    under the minimal-control state used here

The general-matrix Networks skip remains in place — that's a
separate fix that requires either re-architecting the matrix
to use minimal injection per-form or a deeper Access-side
workaround.

Why this test is here, not in test_vba_matrix_hard_forms.py
-----------------------------------------------------------
The hard-forms matrix uses `make_fixture` which constructs
`VbaSession` with default kwargs (full injection).  This test
needs the PR AT `skip_inject_autodetect_forms` kwarg to
suppress sibling injection.  Cleanest split is its own file.
"""
from __future__ import annotations

from pathlib import Path

import pytest

from cbdb_driver.vba_session import VbaSession
from test_vba_matrix_all_forms import SRC


WORK = (
    Path(__file__).resolve().parent.parent
    / "analysis" / "_networks_small_fixture_test_copy.mdb"
)

# PR AU V12 / AX recipe: skip every sibling LookAt form's
# autodetect injection, leave only Form_LookAtNetworks injected.
SKIP_SIBLINGS: set[str] = {
    "Form_LookAtEntry",
    "Form_LookAtOffice",
    "Form_LookAtStatus",
    "Form_LookAtTexts",
    "Form_LookAtAssociations",
    "Form_LookAtPlace",
    "Form_LookAtKinship",
    "Form_LookAtAssociationPairs",
    "Form_LookAtGroupData",
}

# PR AX-measured minimal kin-only control state.  Avoids the
# gUseFilter machinery (which requires non-kin checkbox + at
# least one assoc-type filter).
MIN_CONTROL_STATE = {
    "TxtNodeDist": 1,
    "TxtMaxLoop":  0,
    "ChkKin":      -1,
    "ChkNonKin":    0,
    "ChkMale":     -1,
    "ChkFemale":   -1,
}

# Cao Zhi — PR AQ ranked-top anchor; PR AX-verified.
ANCHOR_PID = 30270
ANCHOR_NAME_CHN = "曹植"


@pytest.fixture(scope="function")
def vba_minimal_inject():
    """Yield a `VbaSession` with sibling-LookAt injection
    skipped.  Per-test fresh Access process via the existing
    VbaSession lifecycle (open → close)."""
    sess = VbaSession(
        SRC, WORK, skip_inject_autodetect_forms=SKIP_SIBLINGS,
    )
    sess.open()
    try:
        yield sess
    finally:
        try:
            sess.close()
        except Exception:
            pass


def test_cmdrun_smoke_cao_zhi(vba_minimal_inject):
    """Drive CmdRun on Cao Zhi (c_personid=30270) under minimal
    injection.  Asserts CmdRun completes within 120 s, ENTER +
    DONE markers fire, and result tables hold non-empty bounded
    output (PR AX measured exactly 1 row in ZZ_SOCIAL_NETWORK,
    2 rows in ZZ_SCRATCH_PEOPLE; we use bounded ranges so
    incidental data drift in BIOG_ADDR_DATA / KIN_DATA doesn't
    break the test)."""
    vba = vba_minimal_inject

    vba.open_form("LookAtNetworks")
    vba.set_picker_codes(
        "ZZ_SCRATCH_IMPORT_PEOPLE", [ANCHOR_PID],
        column="c_person_id",
    )
    for ctl, val in MIN_CONTROL_STATE.items():
        vba.set_control("LookAtNetworks", ctl, val)

    n_via_timer = vba.click_via_timer(
        "LookAtNetworks", ctl="CmdRun",
        result_table="ZZ_SOCIAL_NETWORK",
    )

    # CmdRun fired and returned with the result table populated.
    assert n_via_timer >= 1, (
        f"CmdRun returned {n_via_timer} rows in ZZ_SOCIAL_NETWORK "
        f"via Form_Timer poll; expected >= 1"
    )

    # Capture row counts directly so the assertion message is
    # informative if any of the bounds drift.
    row_counts: dict[str, int] = {}
    for tbl in ("ZZ_SOCIAL_NETWORK", "ZZ_SCRATCH_PEOPLE",
                 "ZZ_SOCIAL_NETWORK_AGGREGATE"):
        cur = vba.conn.cursor()
        cur.execute(f"SELECT COUNT(*) FROM [{tbl}]")
        row_counts[tbl] = int(cur.fetchone()[0])

    # Bounded sanity checks per AY brief.
    sn = row_counts["ZZ_SOCIAL_NETWORK"]
    sp = row_counts["ZZ_SCRATCH_PEOPLE"]
    assert 1 <= sn < 100, (
        f"ZZ_SOCIAL_NETWORK row count {sn} outside [1, 100); "
        f"counts={row_counts}"
    )
    assert 1 <= sp < 50, (
        f"ZZ_SCRATCH_PEOPLE row count {sp} outside [1, 50); "
        f"counts={row_counts}"
    )

    # Verify ENTER + DONE markers fired (no ERR).  Networks's
    # autodetect entry writes 'LookAtNetworks:ENTER' at top of
    # CmdRun_Click and 'LookAtNetworks:DONE' at the chain exit.
    cur = vba.conn.cursor()
    cur.execute("SELECT msg FROM ZZ_TEST_DEBUG ORDER BY id")
    msgs = [str(r[0]) for r in cur.fetchall()]
    assert "LookAtNetworks:ENTER" in msgs, (
        f"ZZ_TEST_DEBUG missing ENTER marker; got: {msgs[:10]}"
    )
    assert "LookAtNetworks:DONE" in msgs, (
        f"ZZ_TEST_DEBUG missing DONE marker; got: {msgs[:10]}"
    )
    err_msgs = [m for m in msgs if "ERR" in m]
    assert not err_msgs, (
        f"ZZ_TEST_DEBUG contains error messages: {err_msgs[:5]}"
    )
