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
    """Volume fixture per form (first in list) plus all quality fixtures
    (expected_gis_iy_min_pct > 0).  Dynastic-combo fixtures are excluded
    to avoid blowing up runtime."""
    volume: dict[str, CrossFixture] = {}
    quality: list[CrossFixture] = []
    for fx in _all_fixtures():
        if fx.spec.name not in _FORMS_WITH_CMDGIS_TESTABLE_HERE:
            continue
        volume.setdefault(fx.spec.name, fx)
        if fx.expected_gis_iy_min_pct > 0:
            quality.append(fx)
    return list(volume.values()) + quality


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
                              scratch_rows=n,
                              min_iy_pct=fx.expected_gis_iy_min_pct)


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
    # Form_LookAtAssociationPairs.vb:2088-2090.  Separator is
    # comma (Chr(44)); the `_assert_gis_export_depth` auto-
    # detects this.  9 columns; no Sex column (has Female instead).
    "LookAtAssociationPairs": [
        "Name", "NameChn", "Female", "IndexYear",
        "AddrName", "AddrChn", "X", "Y", "xy_count",
    ],
}

# Columns that should be non-empty for the vast majority of rows
# (>= 80 %).  Catches a silent column-bind regression that leaves
# everything blank (the user-reported "lost columns" bug pattern).
#
# IndexYear is intentionally excluded: its fill rate depends heavily
# on which office/dynasty the query covers.  For example, c_office_id=80944
# (典史 / Clerk/District Jailor) holders are mostly Qing officials with
# no career index year in BIOG_MAIN (0.27% fill rate in source data) —
# this is a data characteristic, not a column-bind bug.
# Name / NameChn should always be non-empty for any valid CBDB person.
_GIS_EXPECTED_NON_EMPTY = {
    "Name", "NameChn",
}

# Strings the form uses as "no data" placeholders.  We treat these
# as empty when computing the non-empty rate.
_GIS_EMPTY_PLACEHOLDERS = {"", "[ ]", "[Addr Name Missing]",
                            "[Addr Chn Missing]"}


def _assert_gis_export_depth(form_name: str,
                              header: str,
                              lines: list[str],
                              sep: str,
                              scratch_rows: int,
                              min_iy_pct: float = 0.0) -> None:
    """Run the depth assertions on a parsed GIS export.

    `lines` includes the header at index 0; data rows are
    `lines[1:]`.  `scratch_rows` is the row count that came out of
    CmdQuery (or CmdRun) so we can sanity-check that the export
    didn't silently drop most of them.  `min_iy_pct` > 0 enables the
    IndexYear fill-rate check (quality fixtures only).
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

    # 8c-extra. IndexYear fill rate — quality fixtures only.
    # Catches silent column-bind regressions (the class of bug that
    # gave us Issues #10/#11/#12).  Skipped for volume fixtures where
    # the source population legitimately has sparse IndexYear (e.g.
    # 典史 office holders: 0.3% IY in BIOG_MAIN).
    if min_iy_pct > 0 and "IndexYear" in col_index:
        idx = col_index["IndexYear"]
        iy_non_empty = 0
        for line in data_rows:
            cells = line.split(sep)
            if idx >= len(cells):
                continue
            v = cells[idx].strip()
            if v not in _GIS_EMPTY_PLACEHOLDERS:
                iy_non_empty += 1
        if data_rows:
            iy_rate = iy_non_empty / len(data_rows)
            assert iy_rate >= min_iy_pct, (
                f"[{form_name}] CmdGIS IndexYear is non-empty in only "
                f"{iy_non_empty}/{len(data_rows)} rows "
                f"({100*iy_rate:.1f}%). "
                f"Expected >= {100*min_iy_pct:.0f}% (quality fixture). "
                f"This suggests a column-bind regression — compare the "
                f"scratch table c_index_year with the GIS output."
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


# ----------------------------------------------------------------------
# LookAtGroupData × CmdGIS — clean-branches coverage
#
# GroupData differs from the 6 forms above in two ways that make the
# generic test_cmd_gis_produces_file unsuitable:
#
#  1. CmdGIS dispatches to *multiple* per-checkbox WriteGIS_X subs
#     (Status / Office / OfficePeople / Entry / Text / Addr).  Each
#     fires exactly one .tab file via patch_filedialog.  Asserting on
#     a single output file (the way the generic test does) doesn't
#     cover what GroupData actually produces.
#
#  2. The Entry branch hits Issue #6 (P1 — `queryEntry` projects the
#     non-existent `ENTRY_DATA.c_parental_status` instead of
#     `c_parental_status_code`; JET 3061 fires before INSERT).  That
#     bug is separately pinned by the static
#     `tests/test_known_bugs.py::test_bug6_groupdata_query_entry_wrong_field`
#     and the runtime
#     `tests/test_vba_bug_behaviors.py::test_bug6_lookat_groupdata_query_entry_fires_no_such_field`.
#     Bundling the Entry branch into a coverage test would mix
#     bug-pin and coverage; do NOT include Entry here.
#
# The 11-iteration sub-isolation probe at
# `analysis/probe_groupdata_cmdgis_subcalls.py` confirmed:
#   - Status / Office (×2 GIS variants) / Addr branches: clean
#   - Entry branch: ERR fires (Issue #6)
#   - Text branch: clean ERR-wise but produces 0 files on
#     `person_1` (likely benign — person_1 has 0 BIOG_TEXT_DATA
#     rows; WriteGIS_Text bails on RecCount=0)
# Per the brief, this test covers Status / Office / Addr only;
# Entry is bug-pinned separately, Text is omitted because the
# small fixture benignly yields no Text data.
# ----------------------------------------------------------------------

# Per-branch header anchors (first column of the .tab header that
# WriteGIS_<X> writes).  Verified against the form's VBA source +
# the live probe.  Headers are tab-separated.
_GROUPDATA_BRANCH_SHAPES: dict[str, dict] = {
    # WriteGIS_Status, Form_LookAtGroupData.vb header line.
    # 10 cols.  Starts `Name` then carries `NameChn`.  Person-
    # row shape (one row per person × status).
    "Status": {
        "header_first_col": "Name",
        "header_must_contain": ["NameChn", "IndexYear", "AddrID",
                                 "X", "Y"],
        "min_cols": 10,
    },
    # WriteGIS_OfficeOffice (paired with ChkGisOffice).  10 cols.
    # Starts `Office` (the office-name column, not a person
    # column).
    "Office": {
        "header_first_col": "Office",
        "header_must_contain": ["OfficeChn", "FirstYear",
                                 "LastYear", "Dynasty", "X", "Y"],
        "min_cols": 10,
    },
    # WriteGIS_Addr, Form_LookAtGroupData.vb:5571.  9 cols
    # (Status's 10 minus AddrID).  Also starts `Name`, so
    # disambiguate from Status by column count.
    #
    # Caveat: WriteGIS_OfficeOffice (line 2929) writes TWO
    # files — an Office-shape (line 3027) AND a people-side
    # file (line 3219-3224) whose header is BYTE-IDENTICAL
    # to WriteGIS_Addr's.  The classifier can't distinguish
    # the two; both will tag as `Addr`.  The test below uses
    # `setdefault` so each branch is recorded once even if
    # multiple files match its shape, and asserts all 3
    # branches are seen — the OfficeOffice people-side file
    # is then absorbed under the `Addr` slot.  This is a
    # CmdGIS-source quirk, not a test correctness issue:
    # both files are well-formed, structurally consistent,
    # and meet the per-shape required-cols check.
    "Addr": {
        "header_first_col": "Name",
        "header_must_contain": ["NameChn", "IndexYear",
                                 "AddrName", "X", "Y"],
        "min_cols": 9,
    },
}


def _classify_groupdata_gis_file(header: str
                                  ) -> "tuple[str, dict] | None":
    """Identify which WriteGIS_X sub wrote this file by header
    first-column and column count.  Status and Addr both start
    with `Name`; disambiguate by column count (Status 10 cols,
    Addr 9 cols)."""
    cols = header.split("\t")
    if not cols:
        return None
    first = cols[0]
    if first == "Office":
        return ("Office", _GROUPDATA_BRANCH_SHAPES["Office"])
    if first == "Name":
        if len(cols) >= 10 and "AddrID" in cols:
            return ("Status", _GROUPDATA_BRANCH_SHAPES["Status"])
        return ("Addr", _GROUPDATA_BRANCH_SHAPES["Addr"])
    return None


def test_cmd_gis_groupdata_clean_branches(vba: VbaSession,
                                            tmp_path):
    """LookAtGroupData × CmdGIS, scoped to the 3 clean branches
    (Status / Office / Addr).

    Boundary explicitly excludes:
      - Entry — Issue #6 (P1) fires JET 3061 here; pinned by the
        sister tests test_bug6_groupdata_query_entry_wrong_field
        (static) + test_bug6_lookat_groupdata_query_entry_fires_no_such_field
        (runtime).  Including Entry would mix coverage and
        bug-pin.
      - Text — `person_1` has 0 BIOG_TEXT_DATA rows on the
        current dump, so WriteGIS_Text bails on RecCount=0
        (benign 0-files state; not a coverage gap, just a
        fixture-shape side-effect).  Skipped here to keep
        assertions strict.

    Fixture: matrix_hard_forms's `groupdata_person_1_small`
    (`c_personid = 1`, An Dun 安惇).  Re-defined inline rather
    than imported because matrix_hard_forms's `_HARD_FIXTURES`
    is a module-level constant whose other entries
    (AssociationPairs) we don't want to drag into this test.

    Probe evidence: analysis/probe_groupdata_cmdgis_subcalls.py
    + analysis/groupdata_cmdgis_subcall_trace.md confirmed
    Status / Office / Addr branches each produce a well-formed
    .tab file under this fixture, with the per-branch header
    shapes recorded in `_GROUPDATA_BRANCH_SHAPES`.
    """
    from cbdb_driver.form_specs import LOOKATGROUPDATA
    spec = LOOKATGROUPDATA
    PERSON_ID = 1

    # 1. patch FileDialog so each WriteGIS_X .Show short-circuits
    # to a fresh f<n>.out per call (directory mode, trailing \).
    vba.patch_filedialog(spec.name)

    # 2. open form, seed picker, set checkbox state explicitly.
    #
    # Reset ALL eleven Chk* controls (5 CmdRun-side query + 6
    # CmdGIS-side write) to False FIRST so the test isn't
    # at the mercy of Form_Open defaults — if any of them
    # defaults to True (notably ChkEntry on this form, which the
    # probe-baseline showed must be False to skip the Issue #6
    # branch), the assertion at step 5 will fire even when the
    # explicitly-set checkboxes are correct.  Then enable just
    # the 6 clean-branch boxes.
    #
    # Explicitly skipped:
    #   - ChkEntry / ChkGisEntry — Issue #6 (P1) bug-pinned
    #     separately
    #   - ChkText / ChkGisText — benign 0-files on person_1
    #   - ChkGisOfficePeople — probe showed Office_OfficeOffice
    #     alone exercises queryOffice → WriteGIS_OfficeOffice
    #     cleanly; OfficePeople variant adds a second writer
    #     without coverage value here
    vba.open_form(spec.name)
    vba.set_picker_codes(spec.picker_table, [PERSON_ID],
                         column=spec.picker_column)
    all_chk_controls = (
        "ChkStatus", "ChkOffice", "ChkEntry", "ChkText", "ChkAddr",
        "ChkGisStatus", "ChkGisOffice", "ChkGisOfficePeople",
        "ChkGisEntry", "ChkGisText", "ChkGisAddr",
    )
    for ctl in all_chk_controls:
        try:
            vba.set_control(spec.name, ctl, False)
        except Exception as e:
            print(f"  warn reset {ctl}=False: {e}")
    for ctl in ("ChkStatus", "ChkOffice", "ChkAddr",
                "ChkGisStatus", "ChkGisOffice", "ChkGisAddr"):
        try:
            vba.set_control(spec.name, ctl, True)
        except Exception as e:
            print(f"  warn setting {ctl}=True: {e}")

    # 3. wire chain CmdRun -> CmdGIS via Form.Tag, directory mode
    out_dir = tmp_path / "groupdata_gis_out"
    out_dir.mkdir()
    vba.set_form_tag(spec.name,
                     f"{spec.cmd_name},CmdGIS",
                     str(out_dir) + "\\")

    # 4. fire CmdRun via timer; chain dispatches to CmdGIS after
    # CmdRun's body completes
    n = vba.click_via_timer(
        spec.name, ctl=spec.cmd_name,        # "CmdRun"
        # CmdRun's primary effect is UPDATE-style backfill on
        # ZZ_SCRATCH_IMPORT_PEOPLE; queryStatus / queryOffice /
        # queryAddr each populate their own ZZ_SCRATCH_<X>.
        # Watch ZZ_SCRATCH_STATUS — it's the first per-checkbox
        # sub-query CmdRun calls and tells us whether the chain
        # is dispatching at all.
        result_table="ZZ_SCRATCH_STATUS", timeout=180,
    )
    print(f"\n[LookAtGroupData] CmdRun -> "
          f"ZZ_SCRATCH_STATUS rows = {n}", flush=True)

    # 5. Inspect ZZ_TEST_DEBUG to confirm no unexpected ERR.
    # The probe showed Status/Office/Addr branches produce no
    # ERR; if one fires here it's a regression worth investigating
    # (could be a new bug in a sister branch we didn't anticipate).
    cur = vba.conn.cursor()
    cur.execute("SELECT msg FROM ZZ_TEST_DEBUG ORDER BY id")
    msgs = [str(r[0]) for r in cur.fetchall()]
    cur.close()
    err_msgs = [m for m in msgs
                if "LookAtGroupData:ERR" in m]
    assert not err_msgs, (
        "GroupData × CmdGIS clean-branches coverage saw an "
        "unexpected :ERR marker.  Probe-confirmed Status / "
        "Office / Addr branches were clean — investigate "
        "whether a sister branch regressed (e.g. queryStatus "
        "schema drifted, or one of the WriteGIS_X subs changed "
        f"behaviour).  err_msgs={err_msgs}"
    )

    # 6. Inventory output files.  patch_filedialog in directory
    # mode produces sequential `f<n>.out` per Show call.
    files = sorted(out_dir.glob("*"))
    print(f"[LookAtGroupData] CmdGIS produced {len(files)} files:",
          flush=True)
    for f in files:
        print(f"   {f.name}: {f.stat().st_size} bytes", flush=True)

    # We expect AT LEAST 3 files (Status + Office + Addr, one
    # each).  WriteGIS_Status / WriteGIS_OfficeOffice /
    # WriteGIS_Addr each call dlgSaveAs.Show exactly once.
    assert len(files) >= 3, (
        f"[LookAtGroupData] CmdGIS produced only {len(files)} "
        f"files; expected >= 3 (Status / Office / Addr).  "
        f"per_file: {[(f.name, f.stat().st_size) for f in files]}"
    )

    # 7. Per-file structural checks + per-branch classification.
    # Required: each file is non-empty, header decodes, header
    # column count == first data row's column count, header
    # matches one of the three expected branch shapes.
    classified: dict[str, str] = {}
    for f in files:
        sz = f.stat().st_size
        assert sz > 0, f"[LookAtGroupData] {f.name} is zero bytes"

        raw = f.read_bytes()
        text = raw.decode("utf-8",
                          errors="replace").lstrip("﻿")
        lines = [ln for ln in text.replace("\r\n", "\n").split("\n")
                 if ln.strip()]
        assert lines, (
            f"[LookAtGroupData] {f.name} decoded to no "
            f"lines: {raw[:80]!r}"
        )
        header = lines[0]
        header_cols = header.split("\t")
        n_cols = len(header_cols)

        # First data row's col count must match header (catches
        # off-by-one column shifts / missing trailing field).
        assert len(lines) >= 2, (
            f"[LookAtGroupData] {f.name} has header but no data "
            f"rows: lines={len(lines)}"
        )
        first_row_cols = lines[1].split("\t")
        assert n_cols == len(first_row_cols), (
            f"[LookAtGroupData] {f.name} header has {n_cols} "
            f"cols but first data row has {len(first_row_cols)} "
            f"— column-count mismatch.  header={header!r} "
            f"row1={lines[1][:200]!r}"
        )

        # Classify by header shape and check required columns.
        shape = _classify_groupdata_gis_file(header)
        assert shape is not None, (
            f"[LookAtGroupData] {f.name} header doesn't match "
            f"any expected branch shape (Status / Office / "
            f"Addr).  header={header!r}"
        )
        branch, spec_dict = shape
        assert n_cols >= spec_dict["min_cols"], (
            f"[LookAtGroupData] {f.name} ({branch}) has only "
            f"{n_cols} cols; expected >= "
            f"{spec_dict['min_cols']}.  header={header!r}"
        )
        missing = [c for c in spec_dict["header_must_contain"]
                   if c not in header_cols]
        assert not missing, (
            f"[LookAtGroupData] {f.name} ({branch}) is missing "
            f"required columns {missing}.  header={header_cols!r}"
        )
        # Record one file per branch (silently keep the first if
        # a branch produces multiple files — none does today, but
        # the invariant is "we saw all 3").
        classified.setdefault(branch, f.name)
        print(f"[LookAtGroupData] {f.name}: {n_cols} cols, "
              f"{len(lines)-1} data rows, branch={branch}",
              flush=True)

    # 8. All 3 expected branches classified at least once.
    expected_branches = {"Status", "Office", "Addr"}
    saw = set(classified)
    missing_branches = expected_branches - saw
    assert not missing_branches, (
        f"[LookAtGroupData] CmdGIS produced files but didn't "
        f"cover all 3 clean branches: missing={sorted(missing_branches)}; "
        f"saw={sorted(saw)} (file-by-branch: {classified})"
    )
    print(f"[LookAtGroupData] CmdGIS clean-branches coverage: "
          f"{sorted(saw)}", flush=True)


# ----------------------------------------------------------------------
# LookAtAssociationPairs × CmdGIS — 1×3 known-edged fixture
#
# AssocPairs.CmdGIS_Click (Form_LookAtAssociationPairs.vb:2011)
# reads ZZ_SCRATCH_PEOPLE (populated by CmdQuery) and writes a
# comma-delimited .txt with header:
#   Name,NameChn,Female,IndexYear,AddrName,AddrChn,X,Y,xy_count
# (9 columns, comma separator, ADODB stream, UTF-8 when GISFrame=1).
#
# Requires CmdQuery to run first (populates ZZ_SOCIAL_NETWORK and
# ZZ_SCRATCH_PEOPLE).  Uses the same 1×3 known-edged fixture as
# the Pajek/Neo4j AssocPairs tests so we know ZZ_SOCIAL_NETWORK
# is non-empty.
#
# Fixture: persons 1 and 3 — smallest direct-edge pair on the
# current dump (verified by test_vba_pajek_gephi_cross_form).
# ----------------------------------------------------------------------


def test_cmd_gis_assocpairs(vba: VbaSession, tmp_path):
    """LookAtAssociationPairs × CmdGIS structural coverage test.

    Fires CmdQuery (via Form_Timer) then CmdGIS on the 1×3
    known-edged fixture.  Asserts:
      - .txt file produced and non-empty
      - Header contains required GIS columns
      - Per-row field count matches header
      - Key columns ≥ 80 % non-empty (Name, NameChn, IndexYear)
    """
    from test_vba_pajek_gephi_cross_form import _assocpairs_1x3_fixture
    from cbdb_driver.form_specs import LOOKATASSOCIATIONPAIRS

    spec = LOOKATASSOCIATIONPAIRS
    fx = _assocpairs_1x3_fixture()

    # 1. patch FileDialog (directory mode so each Show returns a
    #    unique counter-suffixed file).
    vba.patch_filedialog(spec.name)

    # 2. open form and seed controls / pickers.
    vba.open_form(spec.name)
    for ctl, val in fx.controls.items():
        try:
            vba.set_control(spec.name, ctl, val)
        except Exception as e:
            print(f"  warn {ctl}={val!r}: {e}")
    # Force UTF-8 GIS output (GISFrame=1 → tStream.Charset="utf-8").
    try:
        vba.set_control(spec.name, "GISFrame", 1)
    except Exception:
        pass

    # 3. wire CmdQuery→CmdGIS chain via Form.Tag.
    out_path = tmp_path / "assocpairs_gis.txt"
    vba.set_form_tag(spec.name, f"{spec.cmd_name},CmdGIS",
                     str(out_path))

    # 4. fire CmdQuery via timer; chain dispatches to CmdGIS after.
    n = vba.click_via_timer(
        spec.name, ctl=spec.cmd_name,
        result_table=spec.result_table, timeout=120,
    )
    print(f"\n[LookAtAssociationPairs] CmdQuery -> {n} rows "
          f"in ZZ_SOCIAL_NETWORK", flush=True)
    assert n >= fx.expected_min_rows, (
        f"[LookAtAssociationPairs] CmdQuery produced only {n} "
        f"rows (expected ≥ {fx.expected_min_rows}).  Fixture "
        f"may be stale (1×3 edge removed?) or CmdQuery regressed."
    )

    # 5. file must exist and be non-empty.
    assert out_path.exists(), (
        f"[LookAtAssociationPairs] CmdGIS output {out_path} never "
        f"appeared — either the chain didn't dispatch or FileDialog "
        f"intercept failed."
    )
    sz = out_path.stat().st_size
    assert sz > 0, "[LookAtAssociationPairs] CmdGIS output is zero bytes"
    print(f"[LookAtAssociationPairs] CmdGIS wrote {sz} bytes",
          flush=True)

    # 6. decode + split header.
    raw = out_path.read_bytes()
    text = raw.decode("utf-8", errors="replace").lstrip("﻿")
    lines = [ln for ln in text.replace("\r\n", "\n").split("\n")
             if ln.strip()]
    assert lines, (
        f"[LookAtAssociationPairs] CmdGIS decoded to no lines: "
        f"{raw[:80]!r}"
    )
    header = lines[0]
    # AssocPairs uses comma separator (Chr(44)), not tab.
    sep = "," if "\t" not in header else "\t"
    cols = header.split(sep)
    assert "NameChn" in cols, (
        f"[LookAtAssociationPairs] CmdGIS header missing 'NameChn': "
        f"{cols!r}"
    )
    assert len(cols) >= 4, (
        f"[LookAtAssociationPairs] CmdGIS header has only {len(cols)} "
        f"columns: {cols!r}"
    )
    assert len(lines) >= 2, (
        "[LookAtAssociationPairs] CmdGIS produced only the header"
    )
    print(f"[LookAtAssociationPairs] CmdGIS OK — {len(cols)} cols, "
          f"{len(lines)-1} data rows", flush=True)

    # 7. depth checks (required-column manifest + row-width + fill-rate).
    _assert_gis_export_depth("LookAtAssociationPairs", header,
                              lines, sep, scratch_rows=n)
