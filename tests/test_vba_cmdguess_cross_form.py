"""
Cross-form `CmdGUESS_Click` tests (roadmap item 8, second slice).

CmdGUESS exports a `.gdf` (Graph Description Format — GUESS network
viewer's native format) to a single file.  Three forms have it:
LookAtKinship, LookAtNetworks, LookAtOffice.

Same structural-assertion approach as `test_vba_cmdgis_other_forms.py`:
- Chain CmdQuery/CmdRun → CmdGUESS via Form.Tag
- Assert output file exists + non-empty + first record begins with
  the `.gdf` `nodedef>` keyword

Skip:
- LookAtNetworks — Form_Open hangs in this driver (same family as
  the matrix Networks skip + the picker test skip).
"""
from __future__ import annotations

from pathlib import Path

import pytest

from cbdb_driver.vba_session import VbaSession, make_fixture
from cbdb_driver.form_specs import LOOKATOFFICE, LOOKATKINSHIP

from test_vba_matrix_all_forms import _all_fixtures, CrossFixture, SRC

WORK = Path(__file__).resolve().parent.parent / "analysis" / "_cmdguess_test_copy.mdb"


@pytest.fixture(scope="function")
def vba():
    yield from make_fixture(SRC, WORK)


_FORMS_WITH_CMDGUESS = {"LookAtKinship", "LookAtNetworks", "LookAtOffice"}


def _guess_fixtures() -> list[CrossFixture]:
    by_form: dict[str, CrossFixture] = {}
    for fx in _all_fixtures():
        if fx.spec.name not in _FORMS_WITH_CMDGUESS:
            continue
        by_form.setdefault(fx.spec.name, fx)
    return list(by_form.values())


def _skip_marks(fx: CrossFixture):
    if fx.spec.name == "LookAtNetworks":
        return pytest.mark.skip(
            reason="LookAtNetworks CmdRun times out on high-degree "
                   "anchors (PR AA: Form_Open is fine; matrix "
                   "Networks blocker is CmdRun expansion)."
        )
    return ()


def _seed_query_inputs(vba: VbaSession, fx: CrossFixture) -> None:
    spec = fx.spec
    vba.open_form(spec.name)
    for ctl, val in fx.controls.items():
        try:
            vba.set_control(spec.name, ctl, val)
        except Exception as e:
            print(f"  warn setting {ctl}={val!r}: {e}")
    if fx.picker_ids and spec.picker_table:
        vba.set_picker_codes(spec.picker_table, fx.picker_ids,
                             column=spec.picker_column)
    if fx.addr_ids:
        vba.set_picker_addrs(fx.addr_ids)


@pytest.mark.parametrize(
    "fx",
    [pytest.param(f, marks=_skip_marks(f)) for f in _guess_fixtures()],
    ids=lambda f: f.name,
)
def test_cmd_guess_produces_file(vba: VbaSession, fx: CrossFixture, tmp_path):
    spec = fx.spec
    vba.patch_filedialog(spec.name)
    _seed_query_inputs(vba, fx)

    out_path = tmp_path / f"guess_{spec.name}.gdf"
    vba.set_form_tag(spec.name, f"{spec.cmd_name},CmdGUESS", str(out_path))

    n = vba.click_via_timer(
        spec.name, ctl=spec.cmd_name,
        result_table=spec.result_table, timeout=120,
    )
    print(f"\n[{spec.name}] {spec.cmd_name} -> {n} scratch rows", flush=True)
    assert n >= fx.expected_min_rows, (
        f"[{spec.name}] {spec.cmd_name} only {n} rows; expected "
        f"≥ {fx.expected_min_rows}"
    )

    assert out_path.exists(), (
        f"[{spec.name}] CmdGUESS output {out_path} never appeared"
    )
    sz = out_path.stat().st_size
    assert sz > 0, f"[{spec.name}] CmdGUESS output is zero bytes"

    # GDF format starts with `nodedef>...` on the first non-empty line.
    # Encoding varies per form: Kinship/Networks default UTF-8; Office
    # writes UTF-16LE.  Detect via BOM.
    raw = out_path.read_bytes()
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        text = raw.decode("utf-16", errors="replace")
    else:
        text = raw.decode("utf-8", errors="replace")
    text = text.lstrip("﻿")
    lines = [ln for ln in text.replace("\r\n", "\n").split("\n") if ln.strip()]
    assert lines, f"[{spec.name}] decoded to no lines: {out_path.read_bytes()[:80]!r}"
    assert lines[0].lower().startswith("nodedef"), (
        f"[{spec.name}] CmdGUESS output doesn't start with 'nodedef>': "
        f"first line = {lines[0]!r}"
    )
    print(f"[{spec.name}] CmdGUESS OK ({sz} bytes, {len(lines)} lines)",
          flush=True)

    # ---- Depth checks (PR R) ----------------------------------
    # CmdGUESS shares the .gdf format with CmdGephi; reuse the same
    # depth helper from test_vba_pajek_gephi_cross_form.py.
    from test_vba_pajek_gephi_cross_form import _assert_gephi_depth
    _assert_gephi_depth(spec.name, lines)
