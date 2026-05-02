"""
`CmdGISPeople_Click` test for LookAtOffice (roadmap item 8, fourth slice).

CmdGISPeople is hosted only on LookAtOffice — it dumps the
`ZZ_SCRATCH_P_OFFICE` recordset to a `.tab` file (people-level
GIS export, distinct from CmdGIS which dumps `ZZ_SCRATCH_OFFICE`).

Same structural-assertion pattern as test_vba_cmdgis_other_forms.py:
- chain CmdQuery → CmdGISPeople via Form.Tag
- assert file exists, non-empty, header has the expected columns

The header begins with `Name\tNameChn\t...` in UTF-8 mode (FrameGISPeople
= 2).  Force that branch via set_control so output bytes are stable.
"""
from __future__ import annotations

from pathlib import Path

import pytest

from cbdb_driver.vba_session import VbaSession, make_fixture
from test_vba_matrix_all_forms import _all_fixtures, CrossFixture, SRC

WORK = Path(__file__).resolve().parent.parent / "analysis" / "_cmdgispeople_test_copy.mdb"


@pytest.fixture(scope="function")
def vba():
    yield from make_fixture(SRC, WORK)


def _office_fixture() -> CrossFixture:
    for fx in _all_fixtures():
        if fx.spec.name == "LookAtOffice":
            return fx
    pytest.skip("no LookAtOffice fixture available")


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
    # Force UTF-8 branch so header bytes are stable.  ChkPeopleKML
    # must be False (otherwise CmdGISPeople trampolines to writePersonKML).
    try:
        vba.set_control(spec.name, "FrameGISPeople", 2)
    except Exception:
        pass
    try:
        vba.set_control(spec.name, "ChkPeopleKML", False)
    except Exception:
        pass


def test_cmd_gis_people_produces_file(vba: VbaSession, tmp_path):
    fx = _office_fixture()
    spec = fx.spec
    vba.patch_filedialog(spec.name)
    _seed_query_inputs(vba, fx)

    out_path = tmp_path / "office_people.tab"
    vba.set_form_tag(spec.name, f"{spec.cmd_name},CmdGISPeople", str(out_path))

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
        f"[{spec.name}] CmdGISPeople output {out_path} never appeared"
    )
    sz = out_path.stat().st_size
    assert sz > 0, f"[{spec.name}] CmdGISPeople output is zero bytes"

    raw = out_path.read_bytes()
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        text = raw.decode("utf-16", errors="replace")
    else:
        text = raw.decode("utf-8", errors="replace")
    text = text.lstrip("﻿")
    lines = [ln for ln in text.replace("\r\n", "\n").split("\n") if ln.strip()]
    assert lines, (
        f"[{spec.name}] CmdGISPeople decoded to no lines: {raw[:80]!r}"
    )
    header = lines[0]
    sep = "\t" if "\t" in header else ","
    cols = header.split(sep)
    assert "NameChn" in cols, (
        f"[{spec.name}] CmdGISPeople header missing NameChn (sep={sep!r}): "
        f"{cols!r}"
    )
    assert len(lines) >= 2, (
        f"[{spec.name}] CmdGISPeople produced only header — no data"
    )
    print(f"[{spec.name}] CmdGISPeople OK ({sz} bytes, {len(cols)} cols, "
          f"{len(lines) - 1} data rows)", flush=True)
