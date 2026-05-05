"""Verification probe for the AssociationPairs `.SetFocus` driver
patch.

Background
----------
Per the export-gap triage (`analysis/export_gap_triage_plan.md`,
the B-bucket AssociationPairs entries), `Form_LookAtAssociation
Pairs.CmdQuery_Click` calls `.SetFocus` on six controls in
sequence (TxtFromYear / TxtToYear / CmdQuery in the
`gUseIndexYears` branch on lines 1620 / 1627 / 1634, then
Me.TxtID1 / Me.TxtID2 / Me.CmdQuery in the people-pair branch on
lines 1655 / 1658 / 1660).  Under the driver's headless
Form_Timer dispatch the active form is the welcome /
Navigation Pane — not LookAtAssociationPairs — so the very
first `.SetFocus` raises VBA error 2110 ("can't move the focus
to the control X.") and the handler exits via its error trap
BEFORE the body's INSERTs into ZZ_SCRATCH_PEOPLE and
ZZ_SOCIAL_NETWORK run.  The existing
`tests/test_vba_matrix_hard_forms.py::_check_assoc_pairs`
deliberately doesn't assert on the table contents (its docstring
explicitly notes that "small fixtures may legitimately have 0
network edges"), so it silently passes through the error-trap
exit.

The driver patch (this branch's change to
`tests/cbdb_driver/vba_session.py::_PER_FORM_CMDGIS_PATCHES`)
neutralises those six `.SetFocus` calls via a line-anchored
regex that comments out each standalone `<receiver>.SetFocus`
statement.  This probe is the strict positive evidence that the
patch unblocked the body — it asserts ZZ_SCRATCH_PEOPLE > 0
(immediate marker that the body's first INSERT actually ran)
AND ZZ_SOCIAL_NETWORK > 0 (the brief's specified strict
assertion).

Fixture is a known-edged pair: TxtID1=1 (An Dun 安惇),
TxtID2=3 (a different small-id person who shares one direct
ASSOC_DATA edge with person 1 on the current dump — verified
via `SELECT * FROM ASSOC_DATA WHERE c_personid=1 AND c_assoc_id
=3`).  The 4×5 pair that `matrix_hard_forms::assocpair_4x5_small`
uses doesn't share any 1st-order edges so it produces
ZZ_SCRATCH_PEOPLE > 0 but ZZ_SOCIAL_NETWORK = 0 — fine for
matrix_hard_forms (which tolerates 0 edges) but doesn't satisfy
the brief's strict ZZ_SOCIAL_NETWORK > 0 assertion.

Pre-patch baseline: ZZ_SCRATCH_PEOPLE = 0, ZZ_SOCIAL_NETWORK = 0
(both because the body never ran past the error trap).
Post-patch expectation: both > 0.
"""
from __future__ import annotations

import pytest

from cbdb_driver.vba_session import VbaSession, make_fixture
from cbdb_driver.form_specs import LOOKATASSOCIATIONPAIRS
from test_vba_matrix_all_forms import SRC

WORK = SRC.parent.parent / "analysis" / "_assocpairs_probe_copy.mdb"


@pytest.fixture(scope="function")
def vba():
    yield from make_fixture(SRC, WORK)


def test_associationpairs_cmdquery_setfocus_patch_unblocks_inserts(
        vba: VbaSession):
    """Strict positive marker for the SetFocus driver patch.

    Pre-patch (without the
    `Form_LookAtAssociationPairs` entry in
    `_PER_FORM_CMDGIS_PATCHES`): the very first `.SetFocus`
    inside CmdQuery_Click raises VBA 2110, the handler exits via
    its error trap, ZZ_SCRATCH_PEOPLE and ZZ_SOCIAL_NETWORK are
    both empty (ZZ_SCRATCH_PEOPLE was just `Delete *`'d at line
    1613 and the re-INSERT at line 1640+ never ran).

    Post-patch: the `.SetFocus` calls are commented out, the
    body runs to completion, both tables are non-empty.

    If this test fails post-patch with both tables = 0, the
    driver patch isn't being applied (check
    `_PER_FORM_CMDGIS_PATCHES`).

    If this test fails with ZZ_SCRATCH_PEOPLE > 0 but
    ZZ_SOCIAL_NETWORK = 0, the patch worked but the 4×5 fixture
    happens to have no shared associations on the current dump —
    pick a richer pair (the existing matrix_hard_forms fixture
    is the canonical small case; if it stops producing edges,
    use a known historical pair like Wang Anshi vs Sima Guang).
    """
    spec = LOOKATASSOCIATIONPAIRS

    # Person pair 1 ↔ 3 — known to share 1 direct ASSOC_DATA
    # edge on the current dump (see docstring for the SQL that
    # picked this pair).  Same control shape as the matrix
    # fixture except for the IDs.
    vba.open_form(spec.name)
    for ctl, val in (
        ("TxtID1", 1), ("TxtID2", 3),
        ("TxtPerson1", "1"), ("TxtPerson2", "3"),
        ("FrameFilterYears", 1),
        ("Chk2Nodes", 0), ("ChkKinship", 0),
    ):
        try:
            vba.set_control(spec.name, ctl, val)
        except Exception as e:
            print(f"  warn setting {ctl}={val!r}: {e}", flush=True)

    # Fire CmdQuery via Form_Timer dispatch.  Pre-patch this
    # would silently succeed (the existing matrix_hard_forms
    # _check_assoc_pairs swallows the empty result).  Post-patch
    # the body actually runs.
    n = vba.click_via_timer(
        spec.name, ctl=spec.cmd_name,
        result_table=spec.result_table, timeout=180,
    )
    print(f"\n[LookAtAssociationPairs] CmdQuery click_via_timer -> "
          f"{n} rows in {spec.result_table}", flush=True)

    # Strict positive markers.  Read both tables.
    cur = vba.conn.cursor()
    cur.execute("SELECT COUNT(*) FROM ZZ_SCRATCH_PEOPLE")
    n_people = int(cur.fetchone()[0])
    cur.execute("SELECT COUNT(*) FROM ZZ_SOCIAL_NETWORK")
    n_edges = int(cur.fetchone()[0])
    cur.close()
    print(f"[LookAtAssociationPairs] post-CmdQuery row counts: "
          f"ZZ_SCRATCH_PEOPLE={n_people}, "
          f"ZZ_SOCIAL_NETWORK={n_edges}", flush=True)

    # Marker 1 (immediate): ZZ_SCRATCH_PEOPLE must be populated.
    # CmdQuery_Click's body does `Delete * from ZZ_SCRATCH_PEOPLE`
    # at line ~1613 then INSERTs at line ~1640+.  If SetFocus
    # bailed at line 1620, this stays 0.  If the patch worked,
    # this is >= 2 (the two endpoint people from the pair, plus
    # any 1st-order linking people).
    assert n_people > 0, (
        "AssociationPairs SetFocus driver patch did NOT unblock "
        "the CmdQuery_Click body — ZZ_SCRATCH_PEOPLE is empty "
        "after click_via_timer, which means the body exited via "
        "the error trap before line 1640+'s INSERT INTO "
        "ZZ_SCRATCH_PEOPLE.  Check that the "
        "`Form_LookAtAssociationPairs` entry in "
        "`tests/cbdb_driver/vba_session.py::"
        "_PER_FORM_CMDGIS_PATCHES` is being applied at injection "
        "time (grep the patched body in the work .mdb's VBE)."
    )

    # Marker 2 (brief's strict assertion): ZZ_SOCIAL_NETWORK must
    # have at least one row.  The 4×5 pair on the current dump
    # produces some 1st-order edges via Link1stOrder("NONKIN",
    # "NONKIN").  If this assertion fires while marker 1 passes,
    # the patch worked but the fixture happened to have no
    # 1st-order edges — see docstring for a richer fixture
    # alternative.
    assert n_edges > 0, (
        "AssociationPairs SetFocus driver patch let the body run "
        f"({n_people} rows in ZZ_SCRATCH_PEOPLE) but "
        "ZZ_SOCIAL_NETWORK is empty — the 4×5 fixture has no "
        "1st-order shared associations on the current dump.  "
        "Pick a richer person pair (e.g. Wang Anshi 王安石 "
        "+ Sima Guang 司馬光) or extend Link1stOrder's "
        "k-hop expansion.  This is NOT a patch failure — the "
        "patch is working as designed; the fixture is too small."
    )
