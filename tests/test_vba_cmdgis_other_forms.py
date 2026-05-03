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
# LookAtPlace's CmdGIS used to fail with "Object required" — root
# cause was Bug #4 in reports/CBDB_Issues_Report_EN.md (CBDB code references a non-
# existent `GISFrame` control).  The driver now applies a per-form
# rewrite (`_PER_FORM_CMDGIS_PATCHES["Form_LookAtPlace"]`) that
# substitutes the correct control name `CodeFrame` so the test
# passes.  The underlying CBDB bug remains — production users
# clicking GIS on LookAtPlace still see "Object required".
def _skip_marks(fx: CrossFixture):
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
    # Per-form separator: most forms use tab (Chr(9)); LookAtPlace
    # uses comma (Chr(44) — see Form_LookAtPlace.vb:1582).  Detect
    # rather than hard-code so the test stays portable if CBDB ever
    # standardises.
    sep = "," if ("\t" not in header and "," in header) else "\t"
    cols = header.split(sep)
    # `NameChn` is the strongest cross-form anchor — Status / Texts /
    # Associations / Office / Place / Kinship / Entry all carry it in
    # the non-Pinyin branch.  Catch a column drop or rename.
    assert "NameChn" in cols, (
        f"[{spec.name}] CmdGIS header has no `NameChn` column "
        f"(sep={sep!r}): {cols!r}"
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

    # ---- 8. depth checks (PR P) ---------------------------------
    # Catch silent column drops, off-by-one row shifts, and
    # large-scale data emptiness that the surface checks (file
    # exists, NameChn in header) miss.
    _assert_gis_export_depth(spec.name, header, lines, sep,
                              scratch_rows=n)


# ----------------------------------------------------------------------
# PR P: per-form GIS export schema manifest + depth checks
#
# Required-column lists are derived from the existing goldens at
# tests/golden/exports/real_lookat<form>_gis_*.tab — anything that's
# in the golden header is asserted as a required column here.  When
# CBDB intentionally adds / drops a column, update both the golden
# (if any) AND this manifest in the same commit.
# ----------------------------------------------------------------------

# Required columns per form's CmdGIS UTF-8 (NameChn-branch) header.
# Forms not listed here fall back to the loose anchor check
# (`NameChn` present + ≥4 cols), which the existing test already does.
_GIS_REQUIRED_COLUMNS: dict[str, list[str]] = {
    "LookAtStatus": [
        "Name", "NameChn", "Sex", "IndexYear",
        "AddrName", "AddrChn", "X", "Y", "xy_count",
    ],
    "LookAtAssociations": [
        "Name", "NameChn", "Female", "IndexYear",
        "AddrName", "AddrChn", "X", "Y", "xy_count",
    ],
    "LookAtOffice": [
        "Name", "NameChn", "IndexYear", "Sex",
        "AddrName", "AddrChn", "PersonX", "PersonY",
        "Office", "OfficeChn", "FirstYear", "LastYear",
        "Dynasty",
        "OfficeAddr", "OfficeAddrChn", "X", "Y", "xy_count",
    ],
    # PR T: LookAtTexts/LookAtPlace/LookAtKinship manifests, derived
    # from the VBA source (analysis/dump/vba/Form_LookAt<F>.vb,
    # GISFrame=1 / NameChn branch).  LookAtTexts also has a committed
    # byte-level golden (tests/golden/exports/real_lookattexts_gis_
    # biblcat_1.tab) whose header matches this manifest exactly.
    # LookAtPlace and LookAtKinship have no committed byte-level
    # golden — blessing one is deferred (would require a stable
    # picker fixture + commit against current CBDB_20260430_DATA.mdb).
    "LookAtTexts": [
        # Form_LookAtTexts.vb:1386-1387 (UTF-8 / NameChn branch).
        "Name", "NameChn", "Sex", "IndexYear",
        "AddrName", "AddrChn", "X", "Y", "xy_count",
    ],
    "LookAtPlace": [
        # Form_LookAtPlace.vb:1588-1590.  Separator is comma
        # (Chr(44), Form_LookAtPlace.vb:1582), not tab — the test's
        # `sep` autodetect handles that.  Note: no `Sex` column.
        "Name", "NameChn", "IndexYear",
        "AddrName", "AddrChn", "X", "Y", "xy_count",
    ],
    "LookAtKinship": [
        # Form_LookAtKinship.vb:220-221.  Note the column-name
        # difference vs the other forms: `XY_count` is capitalised
        # (the others emit `xy_count`).  Don't normalise — the
        # exported file uses exactly this casing.
        "Name", "NameChn", "IndexYear", "Sex",
        "AddrName", "AddrChn", "X", "Y", "XY_count",
    ],
}

# Columns that should be non-empty for the vast majority of rows
# (>= 80 %).  Catches a silent column-bind regression that leaves
# everything blank (the user-reported "lost columns" bug pattern).
_GIS_EXPECTED_NON_EMPTY = {
    "Name", "NameChn", "IndexYear",
}

# Strings the form uses as "no data" placeholders.  We treat these
# as empty when computing the non-empty rate.
_GIS_EMPTY_PLACEHOLDERS = {"", "[ ]", "[Addr Name Missing]",
                            "[Addr Chn Missing]"}


def _assert_gis_export_depth(form_name: str,
                              header: str,
                              lines: list[str],
                              sep: str,
                              scratch_rows: int) -> None:
    """Run the depth assertions on a parsed GIS export.

    `lines` includes the header at index 0; data rows are
    `lines[1:]`.  `scratch_rows` is the row count that came out of
    CmdQuery (or CmdRun) so we can sanity-check that the export
    didn't silently drop most of them.
    """
    cols = header.split(sep)
    n_cols = len(cols)
    data_rows = lines[1:]

    # 8a. Required-column manifest (per form).
    required = _GIS_REQUIRED_COLUMNS.get(form_name)
    if required is not None:
        missing = [c for c in required if c not in cols]
        assert not missing, (
            f"[{form_name}] CmdGIS header is missing required "
            f"columns {missing} (header was {cols!r}).  If this is "
            f"intentional, update _GIS_REQUIRED_COLUMNS in "
            f"tests/test_vba_cmdgis_other_forms.py + the golden."
        )

    # 8b. Per-row field count must match header — catches off-by-one
    # column shifts that would silently slide every value one column
    # to the left/right.
    bad_width = []
    for i, line in enumerate(data_rows, start=1):
        cells = line.split(sep)
        if len(cells) != n_cols:
            bad_width.append((i, len(cells), line[:120]))
        if len(bad_width) >= 5:
            break
    assert not bad_width, (
        f"[{form_name}] CmdGIS rows have wrong field count "
        f"(header has {n_cols} cols).  First mismatches "
        f"(row_index, field_count, snippet):\n"
        + "\n".join(f"  row {i}: {n} cells — {snip!r}"
                     for i, n, snip in bad_width)
    )

    # 8c. Key columns non-empty rate.  CBDB's #1 user-reported bug
    # class is "this column is empty for everyone" (silent backfill
    # failures, wrong control source, etc.).  Demand ≥ 80% non-empty
    # for Name / NameChn / IndexYear.
    col_index = {c: i for i, c in enumerate(cols)}
    for key_col in _GIS_EXPECTED_NON_EMPTY:
        if key_col not in col_index:
            continue
        idx = col_index[key_col]
        non_empty = 0
        for line in data_rows:
            cells = line.split(sep)
            if idx >= len(cells):
                continue
            v = cells[idx].strip()
            if v not in _GIS_EMPTY_PLACEHOLDERS:
                non_empty += 1
        if not data_rows:
            continue
        rate = non_empty / len(data_rows)
        assert rate >= 0.80, (
            f"[{form_name}] CmdGIS column {key_col!r} is non-empty "
            f"in only {non_empty}/{len(data_rows)} rows "
            f"({100*rate:.1f}%).  Below 80% threshold — likely a "
            f"silent column-bind regression of the kind that gave "
            f"us Bugs #10/#11/#12."
        )

    # 8d. Row-count sanity: GIS export should produce roughly one
    # row per scratch_table row.  CmdGIS sometimes deduplicates or
    # one-row-per-(person, addr, …) so this isn't a strict
    # equality, but a > 5x mismatch should never happen.
    if scratch_rows > 0 and data_rows:
        ratio = len(data_rows) / scratch_rows
        assert 0.20 <= ratio <= 5.0, (
            f"[{form_name}] CmdGIS row count {len(data_rows)} is "
            f"{ratio:.2f}× the scratch row count {scratch_rows}.  "
            f"Outside the [0.2, 5.0] sanity band — either CmdGIS "
            f"silently dropped most rows or it's joining and "
            f"exploding."
        )

    print(f"[{form_name}] CmdGIS depth checks OK — "
          f"{n_cols} cols, {len(data_rows)} data rows, "
          f"required-col coverage "
          f"{'manifested' if required is not None else 'loose-check'}",
          flush=True)
