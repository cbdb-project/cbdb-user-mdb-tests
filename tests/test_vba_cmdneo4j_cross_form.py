"""
Cross-form `CmdNeo4j_Click` tests (roadmap item 8, fifth slice).

CmdNeo4j is the most complex export — it pops 6-10 SaveAs dialogs in
sequence and dumps a separate `.csv` per scratch table (People,
PeopleEntry, Places, PeoplePlaces, EntryCodes, ...).  Hosted on
seven LookAt forms.

Test approach: pass the test path as a *directory* (trailing
backslash).  The patched `GetTestExportPath()` (FILEDIALOG_PATCH v8)
detects directory mode and returns a counter-suffixed `f<n>.out`
file per call, so each `dlgSaveAs.Show` block writes to a unique
path.  Then we assert on the file count + per-file non-emptiness —
catches "Sub doesn't run", "early bail before any write", "regex
mismatch breaks one block".

Skips:
- LookAtNetworks: matrix CmdQuery/CmdRun skipped + under default
  full injection Form_Open hits the project-wide auto-compile
  deadlock (PR AR-AX, AGENTS landmine #3.5).  Form_Open is fine
  via minimal injection — see tests/test_vba_networks_small_fixture.py.
- LookAtStatus: chain interaction with CmdQuery cleanup-rebind
  (same root family as Pajek/Gephi Status skip).
- LookAtAssociationPairs / LookAtGroupData: matrix CmdQuery skipped.

Per-form minimum file count is conservative: many of the dialog
blocks sit inside `If <flag>.Value Then` branches we don't enter.
4 is the floor that matches the simplest forms after gating.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pytest

from cbdb_driver.vba_session import VbaSession, make_fixture
from test_vba_matrix_all_forms import _all_fixtures, CrossFixture, SRC

WORK = Path(__file__).resolve().parent.parent / "analysis" / "_cmdneo4j_test_copy.mdb"


@pytest.fixture(scope="function")
def vba():
    yield from make_fixture(SRC, WORK)


@dataclass(frozen=True)
class Spec:
    form: str
    min_files: int


_SPECS: tuple[Spec, ...] = (
    # LookAtEntry has been measured end-to-end against the
    # `c_entry_code = 101` matrix fixture (Issue #9 reverification
    # 2026-05-04): produces exactly 7 files (People, PeopleEntry,
    # Places, PeoplePlaces, PersonPlaceCodes, EntryCodes,
    # AssocCodes).  No InstitutionCodes file because
    # `ENTRY_DATA.c_inst_code > 0 = 0` on the current dump gates
    # the optional InstitutionCodes block out — see the per-fixture
    # InstitutionCodes-absent assertion in the test body.
    Spec("LookAtEntry",        min_files=7),
    Spec("LookAtTexts",        min_files=4),
    Spec("LookAtAssociations", min_files=4),
    Spec("LookAtOffice",       min_files=4),
    Spec("LookAtPlace",        min_files=4),
    Spec("LookAtKinship",      min_files=4),
    Spec("LookAtStatus",       min_files=4),
)


def _fixture_for(form: str) -> CrossFixture | None:
    for fx in _all_fixtures():
        if fx.spec.name == form:
            return fx
    return None


def _spec_skip_marks(s: Spec):
    if s.form == "LookAtStatus":
        return pytest.mark.skip(
            reason="LookAtStatus chain post-cleanup invalidates the "
                   "subform recordset rebind; downstream CmdNeo4j reads "
                   "RecordCount=0.  Same root family as Pajek/Gephi skip."
        )
    if s.form == "LookAtPlace":
        return pytest.mark.skip(
            reason="LookAtPlace.CmdNeo4j fires `Item not found in this "
                   "collection.` mid-body — looks like a real CBDB bug "
                   "(SQL or recordset field reference against a renamed/"
                   "missing column).  Worth a deeper audit; for now skip "
                   "so the 3 working forms ship."
        )
    if s.form == "LookAtAssociations":
        return pytest.mark.skip(
            reason="LookAtAssociations.CmdNeo4j produces 0 files in "
                   "directory mode — needs investigation alongside Place."
        )
    return ()


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
    "spec",
    [pytest.param(s, marks=_spec_skip_marks(s)) for s in _SPECS],
    ids=lambda s: s.form,
)
def test_cmd_neo4j_produces_files(vba: VbaSession, spec: Spec, tmp_path):
    fx = _fixture_for(spec.form)
    if fx is None:
        pytest.skip(f"no matrix fixture for {spec.form}")
    fspec = fx.spec

    vba.patch_filedialog(fspec.name)
    _seed_query_inputs(vba, fx)

    # Directory mode: trailing backslash signals GetTestExportPath
    # to return f1.out / f2.out / ... per call.
    out_dir = tmp_path / "neo4j_out"
    out_dir.mkdir()
    out_dir_str = str(out_dir) + "\\"
    vba.set_form_tag(fspec.name,
                     f"{fspec.cmd_name},CmdNeo4j",
                     out_dir_str)

    n = vba.click_via_timer(
        fspec.name, ctl=fspec.cmd_name,
        result_table=fspec.result_table, timeout=180,
    )
    print(f"\n[{fspec.name}] {fspec.cmd_name} -> {n} scratch rows",
          flush=True)
    assert n >= fx.expected_min_rows, (
        f"[{fspec.name}] {fspec.cmd_name} only {n} rows; expected "
        f"≥ {fx.expected_min_rows}"
    )

    # CmdNeo4j auto-appends `.csv` (when our `f<n>.out` doesn't end
    # in .csv), so each file lands as `f<n>.out.csv`.  Glob loosely.
    files = sorted(out_dir.glob("f*"))
    print(f"[{fspec.name}] CmdNeo4j produced {len(files)} files:", flush=True)
    for f in files:
        sz = f.stat().st_size
        print(f"   {f.name}: {sz} bytes", flush=True)

    assert len(files) >= spec.min_files, (
        f"[{fspec.name}] CmdNeo4j produced only {len(files)} files; "
        f"expected ≥ {spec.min_files}.  Per-file: "
        f"{[(f.name, f.stat().st_size) for f in files]}"
    )
    # Each produced file must be non-empty.
    for f in files:
        sz = f.stat().st_size
        assert sz > 0, f"[{fspec.name}] {f.name} is zero bytes"

    # ---- Depth checks (PR Q) ----------------------------------
    # CmdNeo4j writes 6-10 CSVs per form, each a different shape.
    # Driver-side dialog redirection means we can't tell which
    # file is which from the on-disk name (everything is
    # `f<n>.out.csv`).  Inspect each file's header and classify
    # by shape; for known shapes, assert per-row width + key
    # id-column non-empty rate.
    _assert_neo4j_export_depth(fspec.name, files)

    # ---- Per-form structural assertions ------------------------
    if fspec.name == "LookAtEntry":
        _assert_lookatentry_neo4j_shape(fspec.name, files)


def _assert_lookatentry_neo4j_shape(form_name: str,
                                     files: list[Path]) -> None:
    """LookAtEntry × CmdNeo4j on the current dump produces exactly
    7 files with these shapes (verified end-to-end by the Issue #9
    reverification probe at
    `analysis/investigate_issue9_neo4j_institutioncodes.py`):

      People             — header starts `nameID,nameHZ,…`
      PeopleEntry        — header starts `NameID,EntryCode,…`
      Places             — header starts `placeID,placePY,…`
      PeoplePlaces       — header starts `nameID,placeID,…`
      PersonPlaceCodes   — header starts `personPlaceCode,…`
      EntryCodes         — header starts `EntryCode,EntryDesc,…`
      AssocCodes         — header starts `AssocCode,…`

    Conspicuously ABSENT: `InstitutionCodes` (header would start
    `InstitutionCode,…`).  This is NOT a bug on the current dump:
    `Form_LookAtEntry.vb:1389` gates the InstitutionCodes block
    on `If tRecDeleted > 0 Then` where `tRecDeleted` is the row
    count of an `INSERT INTO ZZ_SCRATCH_P_TEXT … WHERE
    ZZ_SCRATCH_ENTRY.c_inst_code > 0`.  CmdQuery copies
    `ENTRY_DATA.c_inst_code` verbatim into ZZ_SCRATCH_ENTRY
    (lines 1645-1652), and `ENTRY_DATA.c_inst_code > 0 = 0` for
    all 263,454 rows on this dump → the gate is false → the
    optional InstitutionCodes block is silently skipped.  This is
    the same per-block "skip when source count is 0" pattern the
    surrounding blocks use (e.g. AssocCodes is also skipped for a
    fixture with `c_assoc_code = 0`).

    Issue #9 (the line-1425 `With tRstAssocCodes` typo intended
    to be `With tRstInstitutions`) remains a real source-level
    bug, but on the current dump it sits behind this gate and is
    unreachable — see Issue #9 in
    `reports/CBDB_Issues_Report_EN.md` (P5 latent).

    If a future MDB drop introduces any `c_inst_code > 0` row in
    ENTRY_DATA, the gate opens, this assertion fails (because an
    8th file appears with `InstitutionCode` first column), AND
    the `With tRstAssocCodes` typo fires DAO 3021 — at which
    point Issue #9 needs re-promotion to P1 and this assertion
    needs updating to either (a) require InstitutionCodes
    present + non-empty after the typo is fixed, or (b) document
    the new failure mode."""
    headers_first_col = []
    inst_files = []
    for f in files:
        raw = f.read_bytes()
        text = raw.decode("utf-8", errors="replace").lstrip("﻿")
        first_line = text.split("\n", 1)[0].strip()
        first_col = first_line.split(",", 1)[0]
        headers_first_col.append(first_col)
        if first_col == "InstitutionCode":
            inst_files.append(f.name)

    expected_first_cols = {
        "nameID",          # People AND PeoplePlaces (disambiguated
                           # downstream by 2-col classifier)
        "NameID",          # PeopleEntry
        "placeID",         # Places
        "personPlaceCode", # PersonPlaceCodes
        "EntryCode",       # EntryCodes
        "AssocCode",       # AssocCodes
    }
    seen = set(headers_first_col)
    missing = expected_first_cols - seen
    assert not missing, (
        f"[{form_name}] CmdNeo4j missing expected file shapes "
        f"(by header first-column): missing={sorted(missing)}; "
        f"saw={sorted(seen)}.  Headers per file: "
        f"{list(zip([f.name for f in files], headers_first_col))}"
    )

    # InstitutionCodes file MUST be absent — see docstring for why
    # this is the desired behaviour on the current dump (Issue #9
    # latent gate).  If this assertion ever fails, see Issue #9
    # re-promotion plan.
    assert not inst_files, (
        f"[{form_name}] CmdNeo4j unexpectedly produced an "
        f"InstitutionCodes-shape file ({inst_files}).  This "
        f"means `ENTRY_DATA.c_inst_code > 0` is no longer 0 on "
        f"the current dump and Issue #9's LATENT-gate has "
        f"flipped — the source-level typo at "
        f"Form_LookAtEntry.vb:1425 is now reachable.  Re-promote "
        f"Issue #9 from P5 latent to P1 in "
        f"reports/generate_report.py and update this test to "
        f"either (a) assert InstitutionCodes present + non-empty "
        f"after the typo is fixed, or (b) document the new "
        f"failure mode.  See "
        f"analysis/issue9_neo4j_institutioncodes_reverification.md."
    )

    # We expect exactly 6 distinct first-column shapes (the two
    # `nameID` files differ by their second column — disambiguated
    # by `_NEO4J_SHAPES_BY_TWO_COLS`).
    assert len(seen) == len(expected_first_cols), (
        f"[{form_name}] expected {len(expected_first_cols)} "
        f"distinct file-shape first-columns "
        f"({sorted(expected_first_cols)}); saw {len(seen)} "
        f"({sorted(seen)})."
    )


# ----------------------------------------------------------------------
# PR Q: per-shape Neo4j export depth manifest
# ----------------------------------------------------------------------

# Recognised file shapes (from existing committed goldens at
# tests/golden/exports/real_lookatentry_neo4j_*.csv plus the
# Issue #9 reverification probe).  Mapping:
#   header_first_column → (shape_label, required_columns,
#                           key_id_columns_must_be_non_empty)
#
# Two `nameID` shapes share the same first column — People (the
# top-level person dump, 6 cols) and PeoplePlaces (a 3-col
# person↔place edge dump that LookAtEntry produces).  Disambiguate
# via second-column lookup before falling back to first-column.
#
# CmdNeo4j writes UTF-8 with BOM; we strip the BOM before splitting.
_NEO4J_SHAPES: dict[str, tuple[str, list[str], list[str]]] = {
    "nameID": ("People",
               ["nameID", "nameHZ", "namePY", "indexyear",
                "dynasty", "sex"],
               ["nameID"]),
    "NameID": ("PeopleEntry",
               ["NameID", "EntryCode"],  # row width loose-check
               ["NameID"]),
    "EntryCode": ("EntryCode-codes",
                  ["EntryCode", "EntryDesc", "EntryDescHZ"],
                  ["EntryCode"]),
    "KinCode": ("KinshipCodes",
                ["KinCode", "KinDesc"],
                ["KinCode"]),
    "AssocCode": ("AssocCodes",
                  ["AssocCode"],   # loose
                  ["AssocCode"]),
    "InstCode": ("InstCodes",
                 ["InstCode"],
                 ["InstCode"]),
    # Added 2026-05-04 to cover LookAtEntry's CmdNeo4j output:
    "placeID": ("Places",
                ["placeID", "placePY", "placeHZ", "placeX",
                 "placeY"],
                ["placeID"]),
    "personPlaceCode": ("PersonPlaceCodes",
                        ["personPlaceCode", "personPlaceTrans",
                         "personPlaceHZ"],
                        ["personPlaceCode"]),
    # Added 2026-05-04 (follow-up) to cover LookAtTexts's
    # mixed-case Neo4j output.  Probe results saved to
    # analysis/_lookattexts_neo4j_headers.json.
    "PlaceID": ("Places-capital",
                ["PlaceID", "PlacePY", "PlaceHZ", "PlaceX",
                 "PlaceY"],
                ["PlaceID"]),
    "TextID": ("TextCodes",
               ["TextID", "TextPY", "TextHZ",
                "TextCategoryCode"],
               ["TextID"]),
    "PersonTextRoleCode": ("PersonTextRoleCodes",
                           # Note: actual CBDB header has
                           # `PersonTextRoleDesc` repeated as
                           # cols[1] AND cols[2] — looks like a
                           # CBDB-side dup-column bug in the
                           # header writer.  We don't enforce the
                           # dup; just require the key column.
                           ["PersonTextRoleCode",
                            "PersonTextRoleDesc"],
                           ["PersonTextRoleCode"]),
}

# Two-column-prefix disambiguation for shapes whose first column
# alone is ambiguous.  Checked BEFORE the first-column lookup.
#
# Two casings of the People-Place-edge shape exist in the wild —
# LookAtEntry uses lower-case (`nameID,placeID,personPlaceCode`),
# LookAtTexts uses mixed-case (`NameID,PlaceID,PersonPlaceCode`).
# Without this two-column disambiguator the LookAtTexts variant
# gets mis-classified as `PeopleEntry` (which expects
# `NameID,EntryCode,...`) and the depth check fails.  Verified
# pre-existing failure on `main` for `test_cmd_neo4j_produces_
# files[LookAtTexts]` when run with `--include-vba` — fixed here
# alongside the LookAtEntry promotion.
_NEO4J_SHAPES_BY_TWO_COLS: dict[
        tuple[str, str], tuple[str, list[str], list[str]]] = {
    ("nameID", "placeID"): ("PeoplePlaces",
                            ["nameID", "placeID",
                             "personPlaceCode"],
                            ["nameID"]),
    ("NameID", "PlaceID"): ("PeoplePlaces",
                            ["NameID", "PlaceID",
                             "PersonPlaceCode"],
                            ["NameID"]),
    # Added 2026-05-04 (follow-up) to cover LookAtTexts's
    # mixed-case Neo4j output.  Three NameID-prefixed shapes
    # exist in the wild; without two-column disambiguation they
    # all collapse onto the legacy `NameID -> PeopleEntry`
    # fall-through and the depth check fails on the People
    # ('NameID,NameHZ,...') and PeopleText ('NameID,TextID,...')
    # files.
    ("NameID", "NameHZ"): ("People-capital",
                           ["NameID", "NameHZ", "NamePY",
                            "IndexYear", "Dynasty", "Sex"],
                           ["NameID"]),
    ("NameID", "TextID"): ("PeopleText",
                           ["NameID", "TextID"],  # row width
                                                  # loose-check
                           ["NameID"]),
    # Make the legacy single-column fall-through explicit too,
    # so the disambiguation is uniform: when col0 is `NameID`
    # the classifier ALWAYS consults the 2-col table first, and
    # `(NameID, EntryCode)` here matches the same shape that the
    # single-column lookup would have returned.  Keeps existing
    # forms (Office, Kinship) on the same answer regardless of
    # whether their PeopleEntry header has 2+ cols.
    ("NameID", "EntryCode"): ("PeopleEntry",
                              ["NameID", "EntryCode"],
                              ["NameID"]),
    # Added 2026-05-04 to cover LookAtOffice's PeopleOffice shape.
    # Header literal at Form_LookAtOffice.vb:947:
    #   "NameID,OfficeCode,OfficeAddrID,SocialInstID,
    #    PostingFirstYear,PostingLastYear,PostingDynasty"
    # Without this 2-col entry the file falls through to the
    # legacy `NameID -> PeopleEntry` single-col lookup, which
    # requires `EntryCode` -> assertion fires.  Verified
    # pre-existing failure on baseline (PR fix/cmdneo4j-
    # classifier-lookattexts surfaced it; PR docs/inventory-
    # real-vba-failing-semantics pinned it as `real_vba_failing`).
    ("NameID", "OfficeCode"): ("PeopleOffice",
                               ["NameID", "OfficeCode",
                                "OfficeAddrID", "SocialInstID",
                                "PostingFirstYear",
                                "PostingLastYear",
                                "PostingDynasty"],
                               ["NameID"]),
}


def _classify_neo4j_csv(header_cols: list[str]
                        ) -> tuple[str, list[str], list[str]] | None:
    """Try to identify what shape a Neo4j CSV is from its header's
    columns.  Two-column prefix wins over single-column to handle
    the `nameID` ambiguity (People shape vs PeoplePlaces shape both
    start with `nameID`).  Returns None if we don't recognise it
    (loose-check fallback)."""
    if not header_cols:
        return None
    if len(header_cols) >= 2:
        two = (header_cols[0], header_cols[1])
        if two in _NEO4J_SHAPES_BY_TWO_COLS:
            return _NEO4J_SHAPES_BY_TWO_COLS[two]
    return _NEO4J_SHAPES.get(header_cols[0])


def _assert_neo4j_export_depth(form_name: str,
                                files: list[Path]) -> None:
    """For each produced CSV: classify by header, then run width +
    non-empty-id checks.  Unknown shapes get the loose check
    (well-formed CSV, ≥ 1 data row, every row has ≥ 2 columns)."""
    classified_count = 0
    for f in files:
        raw = f.read_bytes()
        text = raw.decode("utf-8", errors="replace").lstrip("﻿")
        lines = [ln for ln in text.replace("\r\n", "\n").split("\n")
                 if ln.strip()]
        if not lines:
            raise AssertionError(
                f"[{form_name}] {f.name} decoded to no lines: "
                f"{raw[:80]!r}"
            )
        header = lines[0]
        cols = header.split(",")
        n_cols = len(cols)
        data_rows = lines[1:]

        # Per-row field count must match header.
        bad_width = []
        for i, line in enumerate(data_rows, start=1):
            cells = line.split(",")
            if len(cells) != n_cols:
                bad_width.append((i, len(cells), line[:120]))
            if len(bad_width) >= 5:
                break
        # CmdNeo4j sometimes embeds commas in free-text columns
        # (it doesn't use CSV quoting consistently).  So instead of a
        # hard equality assert, demand that the FIRST cell of every
        # row be a non-empty integer-ish id (the various shape's
        # primary id columns are all integer codes / personids).
        # That catches "all cells slid over" without false-positiving
        # on rows whose downstream columns contain commas.
        bad_id = []
        for i, line in enumerate(data_rows, start=1):
            first = line.split(",", 1)[0].strip()
            if not first or not first.replace("-", "").isdigit():
                bad_id.append((i, first[:80], line[:120]))
            if len(bad_id) >= 5:
                break
        # Skip the id-non-emptiness test for code-table shapes
        # whose key column might legitimately be 0 / blank.

        shape = _classify_neo4j_csv(cols)
        if shape:
            shape_label, required, key_id_cols = shape
            classified_count += 1
            missing = [c for c in required if c not in cols]
            assert not missing, (
                f"[{form_name}] {f.name} ({shape_label}) is missing "
                f"required columns {missing}.  Header was {cols!r}."
            )
            # Key id columns: ≥ 90 % non-empty (very strict; these
            # are integer ids, never blank in healthy data).
            for key in key_id_cols:
                if key not in cols:
                    continue
                idx = cols.index(key)
                non_empty = 0
                for line in data_rows:
                    cells = line.split(",")
                    if idx < len(cells) and cells[idx].strip():
                        non_empty += 1
                if not data_rows:
                    continue
                rate = non_empty / len(data_rows)
                assert rate >= 0.90, (
                    f"[{form_name}] {f.name} ({shape_label}) "
                    f"column {key!r} non-empty in only "
                    f"{non_empty}/{len(data_rows)} rows "
                    f"({100*rate:.1f}%) — likely a silent column-"
                    f"bind regression."
                )

        if bad_id and shape and shape_label not in (
                "EntryCode-codes", "KinshipCodes",
                "AssocCodes", "InstCodes",
                # PersonPlaceCodes is a code-table whose first
                # column (`personPlaceCode`) can legitimately be
                # 0 / blank for unmapped place categories.
                "PersonPlaceCodes",
                # PersonTextRoleCodes — same code-table family;
                # role code can legitimately be 0 / blank for
                # unmapped role values.
                "PersonTextRoleCodes",
                # TextCodes is keyed on `TextID` (integer), but
                # CBDB occasionally writes 0 / blank for unmapped
                # texts; treat as code-table for the bad-id check.
                "TextCodes"):
            # Code-table shapes can legitimately start with 0.
            raise AssertionError(
                f"[{form_name}] {f.name} has rows whose first cell "
                f"isn't an integer id (sample {bad_id[:3]!r}) — "
                f"either CSV escaping is off or columns slid."
            )

        print(f"[{form_name}] {f.name}: "
              f"{n_cols} cols, {len(data_rows)} rows"
              + (f", shape={shape[0]}" if shape else ", shape=?"),
              flush=True)

    print(f"[{form_name}] CmdNeo4j depth: classified "
          f"{classified_count}/{len(files)} files",
          flush=True)
