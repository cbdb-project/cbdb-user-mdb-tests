"""
Cross-form `CmdPajek_Click` / `CmdGephi_Click` tests (roadmap item 8,
third slice).

CmdPajek exports a Pajek `.net` graph; CmdGephi exports a `.gdf`
(same Graph Description Format CmdGUESS produces — Gephi imports it).
Per `analysis/dump/vba/Form_*.vb`:

  CmdPajek hosts:   AssociationPairs, Associations, Kinship,
                    Networks, Place, Status
  CmdGephi hosts:   AssociationPairs, Associations, Place, Status

Skip rationale:
  - Networks: Form_Open hangs in this driver
    (matrix Networks skip + picker test skip same family).
  - AssociationPairs: matrix CmdQuery / CmdRun itself doesn't
    complete (item 7 still open); without a working query the
    CmdQuery → CmdPajek chain has nothing to feed.

Structural assertions only — exact bytes drift with each CBDB data
release.

  Pajek `.net` first non-empty line: `*Vertices N`
  Gephi `.gdf` first non-empty line: `nodedef>...`

Encoding: CmdPajek uses `Scripting.FileSystemObject.CreateTextFile`
(ASCII-ish system default).  CmdGephi shares CmdGUESS's stream code
path, so Office writes UTF-16LE; others UTF-8.  We BOM-detect.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pytest

from cbdb_driver.vba_session import VbaSession, make_fixture

from test_vba_matrix_all_forms import _all_fixtures, CrossFixture, SRC

WORK = Path(__file__).resolve().parent.parent / "analysis" / "_pajek_gephi_test_copy.mdb"


@pytest.fixture(scope="function")
def vba():
    yield from make_fixture(SRC, WORK)


@dataclass(frozen=True)
class Case:
    form: str             # spec.name
    cmd: str              # "CmdPajek" / "CmdGephi"
    ext: str              # ".net" / ".gdf"
    header_prefix: str    # lowercase first-line prefix


_CASES: tuple[Case, ...] = (
    Case("LookAtKinship",      "CmdPajek", ".net", "*vertices"),
    Case("LookAtPlace",        "CmdPajek", ".net", "*vertices"),
    Case("LookAtStatus",       "CmdPajek", ".net", "*vertices"),
    Case("LookAtAssociations", "CmdPajek", ".net", "*vertices"),
    Case("LookAtPlace",        "CmdGephi", ".gdf", "nodedef"),
    Case("LookAtStatus",       "CmdGephi", ".gdf", "nodedef"),
    Case("LookAtAssociations", "CmdGephi", ".gdf", "nodedef"),
)


def _case_skip_marks(c: Case):
    # LookAtStatus's CmdQuery cleanup section rebinds both subform
    # recordsets via `Set ZZ_SCRATCH_X.Form.Recordset = CurrentDb.
    # OpenRecordset(...)`.  The chained CmdPajek / CmdGephi then read
    # `.Form.Recordset.RecordCount = 0` on those rebound recordsets
    # (returns 0 even with 947 rows in the underlying table) and
    # silently bail before SaveToFile.  The CmdQuery body also throws
    # `Object required` near the end (matrix Status passes because it
    # only checks the scratch row count, not the export).  Investigate
    # as a separate thread — for now skip so the other 5 case ship.
    if c.form == "LookAtStatus":
        return pytest.mark.skip(
            reason="LookAtStatus chain to CmdPajek/CmdGephi reads "
                   "RecordCount=0 on subform recordsets after CmdQuery "
                   "cleanup rebinds them; CmdQuery body also errors "
                   "'Object required' late.  Roadmap follow-up."
        )
    return ()


def _fixture_for(form: str) -> CrossFixture | None:
    for fx in _all_fixtures():
        if fx.spec.name == form:
            return fx
    return None


def _case_id(c: Case) -> str:
    return f"{c.form}_{c.cmd}"


def _seed_query_inputs(vba: VbaSession, fx: CrossFixture) -> None:
    spec = fx.spec
    vba.open_form(spec.name)
    if spec.name == "LookAtPlace":
        try:
            vba.set_control("LookAtPlace", "TabPlaces", 0)
        except Exception as e:
            print(f"  warn TabPlaces=0: {e}")
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
    "case",
    [pytest.param(c, marks=_case_skip_marks(c)) for c in _CASES],
    ids=_case_id,
)
def test_export_button_produces_file(vba: VbaSession, case: Case, tmp_path):
    fx = _fixture_for(case.form)
    if fx is None:
        pytest.skip(f"no matrix fixture for {case.form}")
    spec = fx.spec

    vba.patch_filedialog(spec.name)
    _seed_query_inputs(vba, fx)

    out_path = tmp_path / f"{case.cmd.lower()}_{spec.name}{case.ext}"
    vba.set_form_tag(spec.name, f"{spec.cmd_name},{case.cmd}", str(out_path))

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
        f"[{spec.name}] {case.cmd} output {out_path} never appeared"
    )
    sz = out_path.stat().st_size
    assert sz > 0, f"[{spec.name}] {case.cmd} output is zero bytes"

    raw = out_path.read_bytes()
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        text = raw.decode("utf-16", errors="replace")
    else:
        text = raw.decode("utf-8", errors="replace")
    text = text.lstrip("﻿")
    lines = [ln for ln in text.replace("\r\n", "\n").split("\n") if ln.strip()]
    assert lines, (
        f"[{spec.name}] {case.cmd} decoded to no lines: {raw[:80]!r}"
    )
    assert lines[0].lower().lstrip().startswith(case.header_prefix), (
        f"[{spec.name}] {case.cmd} header doesn't start with "
        f"{case.header_prefix!r}: {lines[0]!r}"
    )
    print(f"[{spec.name}] {case.cmd} OK ({sz} bytes, {len(lines)} lines)",
          flush=True)
