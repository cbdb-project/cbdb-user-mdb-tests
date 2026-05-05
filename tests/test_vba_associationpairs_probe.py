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


def test_associationpairs_cmdquery_then_cmdpajek_export_chain(
        vba: VbaSession, tmp_path):
    """Downstream export smoke for the SetFocus driver patch.

    The previous test
    (`test_associationpairs_cmdquery_setfocus_patch_unblocks_
    inserts`) only proved that CmdQuery_Click's body runs to
    completion and populates ZZ_SCRATCH_PEOPLE / ZZ_SOCIAL_
    NETWORK.  This test goes one hop further and proves that an
    actual export chain (CmdQuery → CmdPajek) is now executable —
    the business goal of the patch (4 AssocPairs export gap cells:
    CmdGIS / CmdNeo4j / CmdPajek / CmdGephi).

    Pajek was picked as the **cheapest downstream representative
    smoke** for this driver patch, NOT as a substitute for
    dedicated CmdGIS / CmdGephi / CmdNeo4j coverage.  Reasons:
      - Single-file output (`.net` text format) — easiest to
        assert structurally
      - Doesn't require the full `_NEO4J_SHAPES` classifier
        (CmdNeo4j) or the GIS shape dictionary (CmdGIS)

    What this test DOES prove:
      - The headless SetFocus blocker on the CmdQuery dispatch
        path is removed (sister test
        `test_associationpairs_cmdquery_setfocus_patch_unblocks
        _inserts` proves CmdQuery body completes; this test
        proves a downstream export sub can also fire).
      - At least one downstream export chain (CmdPajek) now runs
        end-to-end on this fixture.

    What this test does NOT prove:
      - CmdNeo4j has no shape / tail / file-chain bugs of its own
        (e.g. an Issue-#21-style empty-recordset-no-EOF-guard).
      - CmdGIS has no form-specific export bug (e.g. a missing
        control reference like Bug #4 on LookAtPlace, or a
        subform-recordset rebind issue).
      - CmdGephi has no format-specific bug (e.g. a writer that
        assumes a column the SELECT doesn't project).
      - Any of the 4 export cells is "covered" in the inventory
        sense — they each still need their own coverage PR with
        proper per-form structural assertions before
        `inventory_export_coverage.py` should mark them
        `covered`.

    Strict positive markers:
      - At least 1 .net file produced
      - File is non-empty
      - File header looks like Pajek (`*Vertices` token present
        somewhere in the first 200 bytes)
      - No `LookAtAssociationPairs:ERR` marker in ZZ_TEST_DEBUG

    If this test fails post-patch with the export markers showing
    a different blocker (e.g. a cleanup-rebind issue similar to
    LookAtStatus's CmdPajek skip), the SetFocus patch is still
    correct — the export chain just has its own separate blocker
    that needs follow-up triage.  Honest negative outcomes per
    the brief: don't hard-merge if the smoke fails for an
    unrelated downstream reason.
    """
    spec = LOOKATASSOCIATIONPAIRS

    # 1. patch FileDialog so CmdPajek's dlgSaveAs.Show short-
    # circuits to a fresh f<n>.out (directory mode, trailing \).
    vba.patch_filedialog(spec.name)

    # 2. open form, set the same 1×3 known-edged pair as the
    # other test (proven to give ZZ_SOCIAL_NETWORK > 0).
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

    # 3. wire chain CmdQuery -> CmdPajek via Form.Tag, directory
    # mode (same pattern as the GroupData CmdGIS test).
    out_dir = tmp_path / "assocpairs_pajek_out"
    out_dir.mkdir()
    vba.set_form_tag(
        spec.name,
        f"{spec.cmd_name},CmdPajek",
        str(out_dir) + "\\",
    )

    # 4. fire CmdQuery via timer; chain dispatches to CmdPajek
    # after CmdQuery's body completes (which depends on the
    # SetFocus patch being applied).
    n = vba.click_via_timer(
        spec.name, ctl=spec.cmd_name,
        result_table=spec.result_table, timeout=180,
    )
    print(f"\n[LookAtAssociationPairs] CmdQuery+CmdPajek chain "
          f"-> {n} rows in {spec.result_table}", flush=True)

    # 5. Inspect ZZ_TEST_DEBUG for any :ERR marker — would
    # indicate either the SetFocus patch didn't apply OR a
    # different downstream blocker fired.
    cur = vba.conn.cursor()
    cur.execute("SELECT msg FROM ZZ_TEST_DEBUG ORDER BY id")
    msgs = [str(r[0]) for r in cur.fetchall()]
    cur.close()
    err_msgs = [m for m in msgs
                if "LookAtAssociationPairs:ERR" in m]
    print(f"[LookAtAssociationPairs] ZZ_TEST_DEBUG entries: "
          f"{len(msgs)}, ERR entries: {len(err_msgs)}",
          flush=True)
    assert not err_msgs, (
        "AssociationPairs CmdQuery+CmdPajek chain saw "
        f"LookAtAssociationPairs:ERR marker(s): {err_msgs}.  "
        "Either the SetFocus patch didn't apply (check "
        "_PER_FORM_CMDGIS_PATCHES is being injected) or the "
        "chain has a separate downstream blocker beyond the "
        "scope of this driver patch.  Honest negative outcome: "
        "do NOT hard-merge until the new blocker is understood."
    )

    # 6. Inventory output files.
    files = sorted(out_dir.glob("*"))
    print(f"[LookAtAssociationPairs] CmdPajek produced "
          f"{len(files)} files:", flush=True)
    for f in files:
        print(f"   {f.name}: {f.stat().st_size} bytes",
              flush=True)
    assert len(files) >= 1, (
        "CmdPajek produced 0 files — chain didn't reach the "
        "dlgSaveAs block.  This means CmdPajek itself bailed "
        "(e.g. RecordCount=0 from a different issue), or the "
        "Form.Tag chain didn't dispatch to CmdPajek.  Honest "
        "negative outcome: investigate before merging."
    )

    # 7. Each produced file must be non-empty.
    for f in files:
        sz = f.stat().st_size
        assert sz > 0, (
            f"CmdPajek output file {f.name} is zero bytes — "
            "the export sub ran the dlgSaveAs.Show but never "
            "wrote anything.  Suggests an empty source recordset "
            "with no .EOF guard (Issue #21-style pattern, but "
            "for a different sub)."
        )

    # 8. At least one file should look like a Pajek .net header.
    # Pajek .net files start with `*Vertices N` then list the
    # vertices, then `*Arcs` or `*Edges`.  Loose check for the
    # `*Vertices` token in the first ~200 bytes of any file.
    pajek_shaped = []
    for f in files:
        head = f.read_bytes()[:200].decode("utf-8",
                                            errors="replace")
        if "*Vertices" in head or "*vertices" in head:
            pajek_shaped.append(f.name)
    assert pajek_shaped, (
        "CmdPajek produced files but none has a `*Vertices` "
        "header — output may not be valid Pajek .net format.  "
        "Files seen: "
        f"{[(f.name, f.stat().st_size) for f in files]}.  "
        "Honest negative outcome: the SetFocus patch unblocked "
        "the chain, but the export shape is wrong — investigate "
        "before merging."
    )
    print(f"[LookAtAssociationPairs] Pajek-shaped files: "
          f"{pajek_shaped}", flush=True)
