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
        f"Bug #7 may be FIXED — LookAtPlace.CmdNeo4j no longer "
        f"raises 'Item not found' at runtime.  err_msgs={err_msgs}"
    )


def test_bug8_lookat_networks_skipped_form_open_hangs():
    """Bug #8 deep verification needs LookAtNetworks Form_Open to
    succeed, which currently hangs in this driver (driver-level issue,
    not a CBDB bug).  The static marker in test_known_bugs covers the
    code-side; this placeholder documents why we don't have runtime
    coverage."""
    pytest.skip("LookAtNetworks Form_Open hangs in this driver — "
                "Bug #8 covered statically by test_known_bugs.py.")


def test_bug9_lookat_entry_cmdneo4j_with_institutions_fixture(vba: VbaSession):
    """Bug #9: LookAtEntry.CmdNeo4j has `With tRstAssocCodes` where
    `With tRstInstitutions` was intended.

    The buggy block sits inside `If tRecDeleted > 0 Then` at line
    1390, where tRecDeleted is the count of rows inserted into
    ZZ_SCRATCH_P_TEXT with `WHERE c_inst_code > 0` — i.e. only entries
    with a social-institution code reach the With block.

    For a typical fixture (entry_code=36 jinshi general), c_inst_code
    is almost always 0, so tRecDeleted=0 and the buggy block is
    skipped.  This test verifies that property — when no institution-
    style entries exist, CmdNeo4j completes without error AND without
    reaching the bug.

    To genuinely trigger Bug #9, a fixture would need ENTRY_DATA rows
    with c_inst_code > 0 (rare).  Static `test_known_bugs.test_bug9`
    covers the code-side; a future deeper test could pre-populate
    ZZ_SCRATCH_ENTRY directly to force the With-block reach."""
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
