"""LookAtKinship × CmdUCINet — first coverage cell of the
CmdUCINet family.

**KNOWN FIXTURE-FRAGILE per Issue #22 sibling-form
confirmation (probe commit `154bb4b`).**  This test passes
on the current dump's matrix-supplied `kinship_person_3211`
fixture only because that person's kin network happens to
contain no Han-character `c_kin_name` values.  The Kinship
sibling-risk probe (`analysis/probe_kinship_cmducinet_
sibling_risk.md`) demonstrated directly that switching to
a fixture whose kin network reaches a Han-name person (e.g.
picker = pid 152930 He Jing 何淨, whose sole 1-hop kin is
pid 140733 He Mou 取) reproduces canonical Issue #22's
exact failure mode here: same `:ERR Invalid procedure call
or argument` (VBA error 5), same partial-file shape (full
`*node data` + truncated `*node properties` + missing
`*tie data`).  See Issue #22's "Affected forms" section for
the full canonical text including Kinship's runtime-
confirmed sibling-form status.

This test's `covered` status in the inventory is preserved
because the test does pass on the current matrix fixture,
but maintainers should treat this cell as Issue-#22-
vulnerable when changing fixtures or processing future
data dumps.  An upstream fix to the
`CreateTextFile(tFileName, True, True)` Unicode flag (per
Issue #22's `fix_en`) would simultaneously close
Associations AND remove the Kinship fixture-fragility.

Per the export-gap triage (post probe `investigate/cmducinet-
family-shape`, commit `4e8e0d2`, merged to main): CmdUCINet
on LookAtKinship is the cheapest first coverage candidate
because:

  - The probe ran cleanly: 3 sections written, no `:ERR`,
    13.76s end-to-end.
  - All section row counts matched the underlying scratch
    tables exactly (`*node data` 949 == `ZZ_SCRATCH_KIN`,
    `*tie data` 1260 == `ZZ_SCRATCH_KINNET`).
  - The matrix's existing `kinship_person_3211` fixture is
    sufficient — no new fixture design required.

Scope of THIS test (deliberately narrow per the brief):

  - Only `LookAtKinship × CmdUCINet`.  Associations and
    Place CmdUCINet remain `gap`; coverage for them is
    out of scope until separate per-form briefs (Associations
    is currently blocked by VBA error 5 mid-write per the
    probe; Place uses ADO Stream rather than FSO and would
    need its own per-form manifest).
  - Test-side split-fire pattern (CmdQuery via timer →
    CmdUCINet via separate timer with `wait_done=False` +
    file-poll completion).  This is **explicitly authorized**
    for THIS PR because `CmdUCINet` is NOT in
    `tests/cbdb_driver/vba_session.py::VbaSession.
    _TIMER_DISPATCH_SUBS` — the autodetect-injected chain
    block can't dispatch it.  Adding `CmdUCINet` to the
    dispatch list is a 1-line driver change deferred to a
    separate maintainer brief; this test does NOT touch the
    driver.

Family-level invariants (probe-confirmed across both probed
forms; safe to assert on Kinship):

  - File extension `.vna`
  - Encoding `cp1252` (no BOM)
  - First non-blank line starts with a `*` section marker
  - No `*tie properties` section (4th section commented out
    in source)

Per-form Kinship invariants (probe-confirmed on this exact
fixture; assertion-baseline for THIS test):

  - Section markers exactly: `*node data` → `*node properties`
    → `*tie data`, in that order
  - Header token counts: 8 / 5 / 5
  - `*node data` and `*node properties` row counts are equal
    (both = node count) AND match `ZZ_SCRATCH_KIN`
  - `*tie data` row count matches `ZZ_SCRATCH_KINNET`
  - No `LookAtKinship:ERR` in `ZZ_TEST_DEBUG`
"""
from __future__ import annotations

import time
from pathlib import Path

import pytest

from cbdb_driver.vba_session import VbaSession, make_fixture
from cbdb_driver.form_specs import LOOKATKINSHIP

from test_vba_matrix_all_forms import _all_fixtures, SRC


WORK = (Path(__file__).resolve().parent.parent
        / "analysis" / "_cmducinet_kinship_test_copy.mdb")


@pytest.fixture(scope="function")
def vba():
    yield from make_fixture(SRC, WORK)


def _kinship_matrix_fixture():
    """Pull the matrix's `kinship_person_3211` fixture so we
    don't duplicate fixture wiring.  If the matrix's fixture
    set drifts to a different person id, the assertions below
    that pin row counts to ZZ_SCRATCH_KIN / ZZ_SCRATCH_KINNET
    still hold (they're scratch-table relative, not absolute)."""
    for fx in _all_fixtures():
        if fx.spec.name == "LookAtKinship":
            return fx
    pytest.skip("no matrix fixture for LookAtKinship "
                "(test_inputs.json may be stale)")


def test_cmducinet_kinship_writes_vna_with_three_sections(
        vba: VbaSession, tmp_path):
    """LookAtKinship × CmdUCINet first coverage cell.  Drives
    CmdRun via timer (Kinship's primary populate sub is
    CmdRun, not CmdQuery), then CmdUCINet via a separate
    timer fire with file-poll completion detection.  Asserts
    the family-level + per-form invariants documented in the
    module docstring."""
    spec = LOOKATKINSHIP

    # 1. Patch the form's FileDialog so dlgSaveAs.Show
    # short-circuits to the path encoded in Form.Tag (set
    # below before each fire).
    vba.patch_filedialog(spec.name)

    # 2. Apply the matrix's kinship_person_3211 fixture.
    fx = _kinship_matrix_fixture()
    vba.open_form(spec.name)
    for ctl, val in (fx.controls or {}).items():
        try:
            vba.set_control(spec.name, ctl, val)
        except Exception as e:
            print(f"  warn setting {ctl}={val!r}: {e}")
    if fx.picker_ids and spec.picker_table:
        vba.set_picker_codes(
            spec.picker_table, fx.picker_ids,
            column=spec.picker_column)

    # 3. Fire CmdRun (Kinship's populate sub) via timer.
    # Tag value doesn't carry a chain because CmdUCINet isn't
    # in _TIMER_DISPATCH_SUBS — we'll fire it separately
    # below.
    vba.set_form_tag(spec.name, spec.cmd_name, "")
    n_scratch = vba.click_via_timer(
        spec.name, ctl=spec.cmd_name,
        result_table=spec.result_table, timeout=180,
    )
    print(f"\n[LookAtKinship] {spec.cmd_name} -> {n_scratch} "
          f"rows in {spec.result_table}", flush=True)
    assert n_scratch >= fx.expected_min_rows, (
        f"[LookAtKinship] {spec.cmd_name} only produced "
        f"{n_scratch} rows; expected ≥ {fx.expected_min_rows}. "
        f"Fixture stale or VBA filter changed."
    )

    # Capture scratch-table counts BEFORE firing CmdUCINet.
    # The post-CmdUCINet row counts are the same (CmdUCINet
    # only reads), so capturing here is sufficient.
    cur = vba.conn.cursor()
    cur.execute("SELECT COUNT(*) FROM ZZ_SCRATCH_KIN")
    n_kin = int(cur.fetchone()[0])
    cur.execute("SELECT COUNT(*) FROM ZZ_SCRATCH_KINNET")
    n_kinnet = int(cur.fetchone()[0])
    cur.close()
    print(f"[LookAtKinship] ZZ_SCRATCH_KIN={n_kin}, "
          f"ZZ_SCRATCH_KINNET={n_kinnet}", flush=True)
    assert n_kin > 0, (
        "ZZ_SCRATCH_KIN is empty after CmdRun; CmdUCINet "
        "would bail at its RecordCount=0 guard "
        "(frmZZ_SCRATCH_KIN.Form.Recordset.RecordCount = 0). "
        "Fixture drift or upstream INSERT change."
    )
    assert n_kinnet > 0, (
        "ZZ_SCRATCH_KINNET is empty after CmdRun; CmdUCINet "
        "would bail at its earlier RecordCount=0 guard "
        "(frmZZ_SCRATCH_KINNET.Form.Recordset.RecordCount = 0)."
    )

    # 4. Re-encode Form.Tag with the desired output path so
    # the patched GetTestExportPath() returns it.  Use
    # `wait_done=False` because CmdUCINet's body has no
    # autodetect-injected DONE marker (autodetect only
    # injects into CmdQuery / CmdRun bodies).
    out_path = tmp_path / "cmducinet_kinship.vna"
    vba.set_form_tag(spec.name, "CmdUCINet", str(out_path))
    vba.click_via_timer(
        spec.name, ctl="CmdUCINet",
        result_table=None, wait_done=False,
    )

    # 5. File-poll completion detection.  CmdUCINet writes
    # via Scripting.FileSystemObject.CreateTextFile —
    # synchronous on the VBA thread, so the file should
    # appear well within the 60s poll cap on this fixture.
    file_deadline = time.time() + 60
    while time.time() < file_deadline:
        if out_path.exists() and out_path.stat().st_size > 0:
            break
        time.sleep(1)
    assert out_path.exists() and out_path.stat().st_size > 0, (
        f"[LookAtKinship] CmdUCINet output {out_path} never "
        f"appeared OR is zero bytes within 60s poll.  Likely "
        f"causes: CmdUCINet bailed at its RecordCount=0 "
        f"guard (subform-recordset rebind issue), the patched "
        f"filedialog didn't engage, or CmdUCINet hit an "
        f"unobserved error."
    )
    sz = out_path.stat().st_size
    print(f"[LookAtKinship] CmdUCINet wrote {sz} bytes to "
          f"{out_path.name}", flush=True)

    # 6. No `:ERR` markers.
    cur = vba.conn.cursor()
    cur.execute(
        "SELECT msg FROM ZZ_TEST_DEBUG ORDER BY id")
    msgs = [str(r[0]) for r in cur.fetchall()]
    cur.close()
    err_msgs = [m for m in msgs if "LookAtKinship:ERR" in m]
    assert not err_msgs, (
        f"[LookAtKinship] CmdUCINet wrote a file but also "
        f"emitted :ERR markers: {err_msgs}.  Investigate "
        f"before accepting coverage — partial export with "
        f":ERR is the Associations-class blocker pattern, "
        f"not a clean Kinship case."
    )

    # 7. Family-level invariants.
    raw = out_path.read_bytes()
    assert raw[:3] != b"\xef\xbb\xbf" and raw[:2] not in (
            b"\xff\xfe", b"\xfe\xff"), (
        f"[LookAtKinship] CmdUCINet output unexpectedly has "
        f"a BOM ({raw[:3]!r}); probe established no-BOM as a "
        f"family-level invariant.  Either the FSO write "
        f"behaviour changed OR encoding semantics drifted."
    )
    # cp1252 superset of ASCII — both forms decoded cleanly
    # in the probe with cp1252.
    text = raw.decode("cp1252", errors="strict")
    lines = text.replace("\r\n", "\n").split("\n")

    # 8. Section structure: exactly the 3 expected markers,
    # in order.
    section_marker_lines = [
        (i, ln.strip()) for i, ln in enumerate(lines)
        if ln.strip().startswith("*")
    ]
    actual_markers = [m for _, m in section_marker_lines]
    expected_markers = ["*node data", "*node properties",
                        "*tie data"]
    assert actual_markers == expected_markers, (
        f"[LookAtKinship] CmdUCINet section markers differ "
        f"from probe-established baseline.\n"
        f"  expected: {expected_markers}\n"
        f"  actual:   {actual_markers}\n"
        f"If a `*tie properties` section appears, the source "
        f"may have un-commented its 4th-section block "
        f"(currently commented out in "
        f"`Form_LookAtKinship.vb` lines ~242-261)."
    )

    # 9. Per-section header token counts + data row counts.
    sections = _parse_sections(lines)
    assert set(sections.keys()) == set(expected_markers), (
        f"section parse mismatch: {sections.keys()}"
    )

    # 9a. *node data — 8-token header, row count == ZZ_SCRATCH_KIN
    nd = sections["*node data"]
    assert nd["header_n_tokens"] == 8, (
        f"[*node data] header expected 8 tokens "
        f"(`ID index_year dy_code dynasty sex x_coord "
        f"y_coord kindist`), got {nd['header_n_tokens']} — "
        f"header was: {nd['header_text']!r}"
    )
    assert nd["data_row_count"] == n_kin, (
        f"[*node data] expected {n_kin} rows "
        f"(== ZZ_SCRATCH_KIN), got {nd['data_row_count']}.  "
        f"Off-by-N or silent row-drop regression."
    )

    # 9b. *node properties — 5-token header, same row count
    # as *node data (both keyed on node id).
    np = sections["*node properties"]
    assert np["header_n_tokens"] == 5, (
        f"[*node properties] header expected 5 tokens "
        f"(`ID color shape size shortlabel`), got "
        f"{np['header_n_tokens']} — header: "
        f"{np['header_text']!r}"
    )
    assert np["data_row_count"] == n_kin, (
        f"[*node properties] expected {n_kin} rows "
        f"(== ZZ_SCRATCH_KIN, same as *node data), got "
        f"{np['data_row_count']}.  Node properties out of "
        f"sync with node data."
    )

    # 9c. *tie data — 5-token header, row count == ZZ_SCRATCH_KINNET
    td = sections["*tie data"]
    assert td["header_n_tokens"] == 5, (
        f"[*tie data] header expected 5 tokens "
        f"(`from to \"EdgeWeight\" \"edgetype\" \"edgelist\"`), "
        f"got {td['header_n_tokens']} — header: "
        f"{td['header_text']!r}"
    )
    assert td["data_row_count"] == n_kinnet, (
        f"[*tie data] expected {n_kinnet} rows "
        f"(== ZZ_SCRATCH_KINNET), got {td['data_row_count']}. "
        f"Edge export out of sync with scratch table."
    )

    print(f"[LookAtKinship] CmdUCINet OK — sections "
          f"{actual_markers}; row counts "
          f"{nd['data_row_count']} / {np['data_row_count']} "
          f"/ {td['data_row_count']} (matched scratch tables "
          f"{n_kin} / {n_kin} / {n_kinnet}); "
          f"{sz} bytes, cp1252 / no BOM",
          flush=True)


def _parse_sections(lines: list[str]) -> dict:
    """Walk lines, slot each section's header + data rows
    into a dict keyed by marker text.  CmdUCINet emits
    space-separated values; the 3rd section's header has
    quoted tokens (`from to "EdgeWeight" "edgetype"
    "edgelist"`) but Python `.split()` on whitespace will
    correctly count those as separate tokens — the quotes
    are part of the token, not delimiters.
    """
    out: dict = {}
    current: dict | None = None
    cur_marker: str | None = None
    for ln in lines:
        s = ln.strip()
        if s.startswith("*"):
            if current is not None and cur_marker is not None:
                out[cur_marker] = current
            cur_marker = s
            current = {
                "header_text": None,
                "header_n_tokens": None,
                "data_row_count": 0,
            }
            continue
        if current is not None and s:
            if current["header_text"] is None:
                current["header_text"] = s
                current["header_n_tokens"] = len(s.split())
            else:
                current["data_row_count"] += 1
    if current is not None and cur_marker is not None:
        out[cur_marker] = current
    return out
