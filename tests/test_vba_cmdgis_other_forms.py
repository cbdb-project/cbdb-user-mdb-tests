"""
Cross-form `CmdGIS_Click` tests (roadmap item 8, first slice).

`tests/test_vba_export.py` already covers `LookAtEntry.CmdGIS` with a
byte-level diff against a frozen golden.  This file generalises to
the other 5 forms whose CmdQuery/CmdRun + CmdGIS chain is known to
work in the matrix (Status, Texts, Associations, Office, Place,
Kinship), using **structural** assertions instead of byte-diff:

  - The output file exists and is non-empty.
  - The header line (first newline-terminated record) starts with the
    form's expected column anchor (`Name\\tNameChn\\t` for the UTF-8
    branch — set via `GISFrame=1`).
  - At least one data row follows the header.

Byte-diff goldens are intentionally left out of this slice — adding
them would require blessing 6 goldens whose exact bytes change every
time CBDB releases new data.  The structural check still catches the
high-value regressions (file not produced, header column dropped,
header column renamed).

Skipped: AssociationPairs, Networks, GroupData (matrix CmdQuery /
CmdRun already skipped — the chain can't fire CmdGIS without a
working query).  LookAtNetworks Form_Open also hangs in this driver.
"""
from __future__ import annotations

from pathlib import Path

import pytest

from cbdb_driver.vba_session import VbaSession, make_fixture
from cbdb_driver.form_specs import FormSpec

from test_vba_matrix_all_forms import (
    SRC, CrossFixture, _all_fixtures,
)

WORK = Path(__file__).resolve().parent.parent / "analysis" / "_cmdgis_test_copy.mdb"


@pytest.fixture(scope="function")
def vba():
    yield from make_fixture(SRC, WORK)


# Forms covered here.  LookAtEntry is omitted — already covered by the
# byte-diff test in tests/test_vba_export.py.  AssociationPairs /
# Networks / GroupData are omitted because their CmdQuery / CmdRun
# itself doesn't complete in the matrix; without that, the
# CmdQuery → CmdGIS chain has nothing to feed CmdGIS.
_FORMS_WITH_CMDGIS_TESTABLE_HERE = {
    "LookAtStatus", "LookAtTexts", "LookAtAssociations",
    "LookAtOffice", "LookAtPlace", "LookAtKinship",
}

# LookAtKinship's CmdGIS reads a saved-query-bound subform
# (frmZZ_SCRATCH_KIN); the chain block now requeries it via
# cbdb_driver.vba_session._SUBFORMS_TO_REQUERY → it passes.
#
# LookAtPlace's CmdGIS reads frmZZZ_PLACE which sits INSIDE
# `TabPlaces` / `PlacePage` (a tab control).  Subforms inside tab
# pages don't materialise their `.Form` object until the page is
# activated AND the subform actually renders — programmatically
# setting TabPlaces.Value = 0 isn't sufficient on its own.  CmdGIS
# then errors with "Object required" reading
# `frmZZZ_PLACE.Form.Recordset.RecordCount`.  Logged thanks to the
# fixed Err-insert SQL in this PR; the actual fix (force the
# subform to load) is left as follow-up.
def _skip_marks(fx: CrossFixture):
    if fx.spec.name == "LookAtPlace":
        return pytest.mark.skip(
            reason="LookAtPlace's frmZZZ_PLACE subform is nested in "
                   "TabPlaces / PlacePage and doesn't load until the "
                   "tab is rendered; CmdGIS errors with `Object "
                   "required` reading its recordset.  Setting "
                   "TabPlaces.Value=0 isn't enough."
        )
    return ()


def _gis_fixtures() -> list[CrossFixture]:
    """One fixture per form from the matrix's high-density inputs.
    Pick the first matrix fixture per form so we don't blow the
    runtime up by parametrising over every fixture × every form."""
    by_form: dict[str, CrossFixture] = {}
    for fx in _all_fixtures():
        if fx.spec.name not in _FORMS_WITH_CMDGIS_TESTABLE_HERE:
            continue
        by_form.setdefault(fx.spec.name, fx)
    return list(by_form.values())


def _seed_query_inputs(vba: VbaSession, fx: CrossFixture) -> None:
    """Same setup the matrix test does — open form first, then
    populate pickers (Form_Open of LookAtOffice wipes
    ZZ_OFFICE_CODE).  Then set GISFrame=1 to force UTF-8 output so
    the header bytes are predictable."""
    spec = fx.spec
    vba.open_form(spec.name)
    # LookAtPlace's frmZZZ_PLACE subform sits inside TabPlaces /
    # PlacePage; subforms inside Tab controls don't materialise their
    # `.Form` object until the page is activated.  Force the active
    # page to PlacePage so the chain block's `frmZZZ_PLACE.Form.Requery`
    # has something to operate on (otherwise CmdGIS later errors with
    # "Object required" reading the recordset).
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
    # Force UTF-8 GIS output so header bytes are predictable.
    # (default unset = 0 → falls into else: gb18030; both encodings
    # have identical ASCII-only header bytes, but UTF-8 is what we
    # exercise in test_vba_export.py too.)
    try:
        vba.set_control(spec.name, "GISFrame", 1)
    except Exception:
        pass  # form may not have GISFrame (Kinship uses different name)


@pytest.mark.parametrize(
    "fx",
    [pytest.param(f, marks=_skip_marks(f)) for f in _gis_fixtures()],
    ids=lambda f: f.name,
)
def test_cmd_gis_produces_file(vba: VbaSession, fx: CrossFixture, tmp_path):
    """Generic CmdGIS test for any LookAt form.  Steps mirror the
    matrix's test_cross_form_matrix but chain CmdGIS after CmdQuery /
    CmdRun via Form.Tag, and assert on the resulting file rather than
    the scratch table."""
    spec = fx.spec

    # 1. patch FileDialog so .Show short-circuits to our path
    vba.patch_filedialog(spec.name)

    # 2. populate inputs (open form first; OFFICE Form_Open wipes pickers)
    _seed_query_inputs(vba, fx)

    # 3. wire chain via Form.Tag — `<query_or_run>,CmdGIS|<path>`
    out_path = tmp_path / f"gis_{spec.name}.tab"
    vba.set_form_tag(spec.name, f"{spec.cmd_name},CmdGIS", str(out_path))

    # 4. fire CmdQuery / CmdRun via timer — the autodetect-injected
    # chain block at end of the body dispatches to CmdGIS
    n = vba.click_via_timer(
        spec.name, ctl=spec.cmd_name,
        result_table=spec.result_table, timeout=120,
    )
    print(f"\n[{spec.name}] {spec.cmd_name} produced {n} scratch rows",
          flush=True)
    assert n >= fx.expected_min_rows, (
        f"[{spec.name}] {spec.cmd_name} only produced {n} rows "
        f"(expected ≥ {fx.expected_min_rows}) — fixture stale or "
        f"VBA filter changed"
    )

    # 5. file should exist + non-empty
    assert out_path.exists(), (
        f"[{spec.name}] CmdGIS output {out_path} never appeared"
    )
    sz = out_path.stat().st_size
    assert sz > 0, f"[{spec.name}] CmdGIS output is zero bytes"
    print(f"[{spec.name}] CmdGIS wrote {sz} bytes", flush=True)

    # 6. structural header check.  CmdGIS prepends a UTF-8 BOM
    # (matches the existing Entry golden in tests/golden/exports/);
    # strip it before checking columns.
    raw = out_path.read_bytes()
    text = raw.decode("utf-8", errors="replace").lstrip("﻿")
    lines = [ln for ln in text.replace("\r\n", "\n").split("\n") if ln.strip()]
    assert lines, f"[{spec.name}] file decoded to no lines: {raw[:80]!r}"
    header = lines[0]
    cols = header.split("\t")
    # `NameChn` is the strongest cross-form anchor — Status / Texts /
    # Associations / Office / Place / Kinship / Entry all carry it in
    # the non-Pinyin branch.  Catch a column drop or rename.
    assert "NameChn" in cols, (
        f"[{spec.name}] CmdGIS header has no `NameChn` column: "
        f"{cols!r}"
    )
    # Sanity: at least 4 columns (every CmdGIS variant emits
    # ≥ Name / NameChn / IndexYear / something-else).
    assert len(cols) >= 4, (
        f"[{spec.name}] CmdGIS header has only {len(cols)} columns: "
        f"{cols!r}"
    )

    # 7. at least one data row beyond the header.
    assert len(lines) >= 2, (
        f"[{spec.name}] CmdGIS produced only the header — "
        f"no data rows.  Lines: {len(lines)}"
    )
    print(f"[{spec.name}] CmdGIS OK — {len(cols)} cols, "
          f"{len(lines)-1} data rows", flush=True)
