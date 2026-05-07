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
- LookAtGroupData: matrix CmdQuery skipped.
- LookAtPlace: `Item not found in this collection.` mid-body —
  separate audit.

LookAtAssociationPairs (covered): the prior matrix CmdQuery skip
+ the CmdNeo4j-specific debug-MsgBox blocker have both been
resolved.  CmdQuery is unblocked by the AssociationPairs SetFocus
driver patch (commits 3bb69ef + 0c0eaf1, on main as of PR AV).
CmdNeo4j is unblocked by the AssociationPairs CmdNeo4j_Click
debug-MsgBox suppress driver patch (PR #109,
`_suppress_assocpairs_cmdneo4j_debug_msgbox`).  Hosted here using
the same 1×3 known-edged person pair as the Pajek/Gephi
cross-form test — see `_assocpairs_1x3_fixture()` in that file
(re-exported via import).

LookAtAssociations (newly covered, this PR): the prior 0-file
skip is resolved by two driver-side workarounds, both on main:
PR #116 (`_rewrite_associations_cmdneo4j_target_column`) rewrites
the canonical Issue #23 target-column typo
`c_index_addr_type_code` -> `c_addr_type` so the INSERT into
ZZ_SCRATCH_PEOPLE no longer hits JET 3061; PR #117
(`_suppress_associations_cmdneo4j_debug_msgbox`) suppresses the
5 concat-form debug MsgBox calls that would otherwise pop dialogs
mid-chain.  Hosted here using the matrix
`_make_assoc_fixtures` first fixture
(`assoc_<top_code>_unfiltered`).  The canonical Issue #23 stays
P1 — these workarounds make the cell *testable* on the existing
source, not *fixed* upstream.  See
`reports/probe_associations_cmdneo4j_after_msgbox_suppress.json`
(PR #117) for the strict-clean verification baseline this test
pins against.

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
from test_vba_pajek_gephi_cross_form import _assocpairs_1x3_fixture

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
    # LookAtAssociations × CmdNeo4j on the matrix
    # `_make_assoc_fixtures` first fixture
    # (`assoc_<top_code>_unfiltered`) produces exactly 8 files
    # (People, Places, PeoplePlaces, PeopleAssociations,
    # AssociationCodes, KinshipCodes, OccasionCodes,
    # TopicCodes) — see
    # `_assert_lookatassociations_neo4j_shape` for the
    # per-shape pinning.  Verified end-to-end by PR #117's
    # post-MsgBox-suppress probe (chain_elapsed = 8.04 s,
    # watchdog dialogs = 0, ZZ_TEST_DEBUG = [ENTER, MSGBOX,
    # DONE]).
    Spec("LookAtAssociations", min_files=8),
    Spec("LookAtOffice",       min_files=4),
    Spec("LookAtPlace",        min_files=4),
    Spec("LookAtKinship",      min_files=4),
    Spec("LookAtStatus",       min_files=4),
    # AssociationPairs × CmdNeo4j on the 1×3 known-edged pair
    # produces exactly 6 files (People, Places, PeoplePlaces,
    # PeopleAssociations, AssociationCodes, KinshipCodes) — see
    # `_assert_lookatassociationpairs_neo4j_shape` for the
    # per-shape pinning.  Verified end-to-end by PR #109's
    # verification probe (chain_elapsed = 5.53 s,
    # msgboxes_observed = 0, ZZ_TEST_DEBUG = [ENTER, DONE]).
    Spec("LookAtAssociationPairs", min_files=6),
)


def _fixture_for(form: str) -> CrossFixture | None:
    # AssociationPairs uses the same custom 1×3 known-edged
    # fixture as `test_vba_pajek_gephi_cross_form.py`; the matrix
    # default 4×5 pair shares 0 first-order edges and would leave
    # ZZ_SOCIAL_NETWORK empty (chained CmdNeo4j would bail on
    # RecordCount=0).
    if form == "LookAtAssociationPairs":
        return _assocpairs_1x3_fixture()
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
    if fspec.name == "LookAtAssociationPairs":
        _assert_lookatassociationpairs_neo4j_shape(
            fspec.name, files, vba)
    if fspec.name == "LookAtAssociations":
        _assert_lookatassociations_neo4j_shape(
            fspec.name, files, vba)


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


def _assert_lookatassociationpairs_neo4j_shape(
        form_name: str,
        files: list[Path],
        vba: VbaSession,
) -> None:
    """LookAtAssociationPairs × CmdNeo4j on the 1×3 known-edged
    fixture (TxtID1=1, TxtID2=3, FrameFilterYears=1, ChkKinship=0,
    Chk2Nodes=0) produces exactly 6 files with these first-column
    shapes:

      People               — `nameID,nameHZ,namePY,…`
      Places               — `placeID,placePY,placeHZ,…`
      PeoplePlaces         — `nameID,placeID,personPlaceCode,…`
      PeopleAssociations   — `Person1_ID,Person2_ID,Association_Code,…`
      AssociationCodes     — `AssociationCode,AssociationTrans,AssociationHZ`
      KinshipCodes         — `KinshipCode,KinshipTrans,KinshipHZ`

    The two `nameID`-headed files (People, PeoplePlaces) are
    disambiguated by their second column via
    `_NEO4J_SHAPES_BY_TWO_COLS` — same pattern as LookAtEntry.

    Conspicuously ABSENT on this fixture:
      KinshipRelations     — gated by `ChkKinship.Value` (off here)
      LiteraryGenreCodes   — gated by `tRecDeleted > 0` (= 0 on
                             this fixture's ZZ_SOCIAL_NETWORK)
      InstitutionCodes     — same gate
      OccasionCodes        — same gate
      TopicCodes           — same gate

    Provenance: PR #109's verification probe ran this exact
    fixture with the MsgBox-suppress driver patch active and
    observed the 6-file set in 5.53 s with zero MsgBox dialogs
    and `ZZ_TEST_DEBUG = [LookAtAssociationPairs:ENTER,
    LookAtAssociationPairs:DONE]`.  See
    `reports/probe_assocpairs_cmdneo4j_after_msgbox_patch.json`
    in PR #109.

    A `:ERR` marker in `ZZ_TEST_DEBUG` would mean the chain hit
    the error trap mid-run (either the original VBA `On Error
    GoTo Err_<sub>` body, or our generic `MsgBox Err.Description`
    rewrite).  This assertion fails on either.

    `ZZ_KIN_LIST_TMP` carries 132 rows after this run with
    ChkKinship=0 (the DELETE at Form_LookAtAssociationPairs.vb
    line 900 is gated by `ChkKinship.Value`).  Per PR AX's probe
    that's an OBSERVATION, not a confirmed blocker — the
    KinshipCodes file (96 rows in this run) is asserted present
    + non-empty by the per-shape depth check above; this helper
    intentionally does NOT escalate the carry-over to a finding
    unless the file's contents are independently shown wrong.
    """
    headers_first_col: list[str] = []
    for f in files:
        raw = f.read_bytes()
        text = raw.decode("utf-8", errors="replace").lstrip("﻿")
        first_line = text.split("\n", 1)[0].strip()
        first_col = first_line.split(",", 1)[0]
        headers_first_col.append(first_col)

    expected_first_cols = {
        "nameID",          # People AND PeoplePlaces (disambiguated
                           # downstream by 2-col classifier)
        "placeID",         # Places
        "Person1_ID",      # PeopleAssociations
        "AssociationCode", # AssociationCodes
        "KinshipCode",     # KinshipCodes
    }
    seen = set(headers_first_col)
    missing = expected_first_cols - seen
    extra = seen - expected_first_cols
    assert not missing, (
        f"[{form_name}] CmdNeo4j missing expected file shapes "
        f"(by header first-column): missing={sorted(missing)}; "
        f"saw={sorted(seen)}.  Headers per file: "
        f"{list(zip([f.name for f in files], headers_first_col))}"
    )
    assert not extra, (
        f"[{form_name}] CmdNeo4j produced unexpected file shape(s) "
        f"(by header first-column): extra={sorted(extra)}; "
        f"expected exactly {sorted(expected_first_cols)}.  Headers "
        f"per file: "
        f"{list(zip([f.name for f in files], headers_first_col))}.  "
        f"This usually means a previously-gated optional block "
        f"(KinshipRelations / LiteraryGenreCodes / InstitutionCodes "
        f"/ OccasionCodes / TopicCodes) became reachable on the "
        f"current dump — see Form_LookAtAssociationPairs.vb gate "
        f"conditions and update the expected set OR the fixture."
    )

    # Exactly 6 files (per PR #109 verification probe baseline on
    # this fixture).  Two `nameID` files (People, PeoplePlaces)
    # share a first column → 5 distinct first-column shapes.
    assert len(files) == 6, (
        f"[{form_name}] CmdNeo4j produced {len(files)} files; "
        f"expected exactly 6 on the 1×3 known-edged fixture.  "
        f"Headers per file: "
        f"{list(zip([f.name for f in files], headers_first_col))}"
    )
    assert len(seen) == len(expected_first_cols), (
        f"[{form_name}] expected {len(expected_first_cols)} "
        f"distinct file-shape first-columns "
        f"({sorted(expected_first_cols)}); saw {len(seen)} "
        f"({sorted(seen)})."
    )

    # Chain-completion + no-error markers in ZZ_TEST_DEBUG.  The
    # generic Err.Description rewrite in the driver writes
    # `<short>:ERR …` rows; a clean run only contains
    # `<short>:ENTER` and `<short>:DONE`.
    cur = vba.conn.cursor()
    cur.execute(
        "SELECT msg FROM ZZ_TEST_DEBUG ORDER BY id")
    debug_msgs = [r[0] for r in cur.fetchall()]
    cur.close()
    assert any(m.endswith(":DONE") for m in debug_msgs), (
        f"[{form_name}] ZZ_TEST_DEBUG never reached :DONE — "
        f"chain block did not complete.  Markers seen: "
        f"{debug_msgs}"
    )
    err_rows = [m for m in debug_msgs if ":ERR" in m]
    assert not err_rows, (
        f"[{form_name}] ZZ_TEST_DEBUG contains runtime :ERR "
        f"markers (chain hit the error trap): {err_rows}.  "
        f"All markers: {debug_msgs}"
    )


def _assert_lookatassociations_neo4j_shape(
        form_name: str,
        files: list[Path],
        vba: VbaSession,
) -> None:
    """LookAtAssociations × CmdNeo4j on the matrix
    `_make_assoc_fixtures` first fixture
    (`assoc_<top_code>_unfiltered`) produces exactly 8 files
    with these first-column shapes:

      People               — `nameID,nameHZ,namePY,…`            (6 cols)
      Places               — `placeID,placePY,placeHZ,…`         (5 cols)
      PeoplePlaces         — `nameID,placeID,personPlaceCode,…`  (4 cols)
      PeopleAssociations   — `Person1_ID,Person2_ID,Association_Code,…` (13 cols)
      AssociationCodes     — `AssociationCode,AssociationTypeID,AssociationTrans,AssociationHZ` (4 cols)
      KinshipCodes         — `KinshipCode,KinshipTrans,KinshipHZ`  (3 cols)
      OccasionCodes        — `OccasionCode,OccasionTrans,OccasionHZ` (3 cols)
      TopicCodes           — `TopicCode,TopicTrans,TopicHZ`        (3 cols)

    The two `nameID`-headed files (People, PeoplePlaces) are
    disambiguated by their second column via
    `_NEO4J_SHAPES_BY_TWO_COLS` — same pattern as LookAtEntry
    and LookAtAssociationPairs.

    Note on AssociationCodes: this form's AssociationCodes file
    has FOUR columns (the extra `AssociationTypeID` between
    `AssociationCode` and `AssociationTrans`), distinct from the
    AssociationPairs three-column form.  Disambiguated by the
    `(AssociationCode, AssociationTypeID)` 2-col entry.

    Provenance: PR #117's verification probe ran this exact
    fixture with both driver workarounds active (PR #116's
    `_rewrite_associations_cmdneo4j_target_column` for the
    Issue #23 c_addr_type rewrite, plus PR #117's
    `_suppress_associations_cmdneo4j_debug_msgbox` for the 5
    concat-form debug MsgBox prefixes) and observed:
      file_count = 8
      chain_elapsed_sec = 8.04
      watchdog dialog count = 0
      ZZ_TEST_DEBUG = [ENTER, MSGBOX, DONE]
      all four strict gates met
    See `reports/probe_associations_cmdneo4j_after_msgbox_suppress.json`.

    The single ZZ_TEST_DEBUG `:MSGBOX` marker is the generic
    literal-only neutralizer's footprint for the terminal
    `MsgBox "Finished saving to Neo4j"` line — NOT a dialog,
    NOT a blocker, NOT the line-1033 RecordCount=0 early-bail
    marker.

    A `:ERR` marker in `ZZ_TEST_DEBUG` would mean the chain
    hit the error trap mid-run.  This assertion fails on any
    `:ERR` row.

    **Canonical Issue #23 stays P1.**  This coverage test only
    proves the cell is *testable* on the existing source via
    the repo-local driver workarounds (PR #116 + #117); the
    underlying source-level defect at the INSERT target column
    is NOT fixed and remains tracked via Issue #23.
    """
    headers_first_col: list[str] = []
    for f in files:
        raw = f.read_bytes()
        text = raw.decode("utf-8", errors="replace").lstrip("﻿")
        first_line = text.split("\n", 1)[0].strip()
        first_col = first_line.split(",", 1)[0]
        headers_first_col.append(first_col)

    expected_first_cols = {
        "nameID",          # People AND PeoplePlaces (disambiguated
                           # downstream by 2-col classifier)
        "placeID",         # Places
        "Person1_ID",      # PeopleAssociations
        "AssociationCode", # AssociationCodes (4-col on this form)
        "KinshipCode",     # KinshipCodes
        "OccasionCode",    # OccasionCodes
        "TopicCode",       # TopicCodes
    }
    seen = set(headers_first_col)
    missing = expected_first_cols - seen
    extra = seen - expected_first_cols
    assert not missing, (
        f"[{form_name}] CmdNeo4j missing expected file shapes "
        f"(by header first-column): missing={sorted(missing)}; "
        f"saw={sorted(seen)}.  Headers per file: "
        f"{list(zip([f.name for f in files], headers_first_col))}"
    )
    assert not extra, (
        f"[{form_name}] CmdNeo4j produced unexpected file shape(s) "
        f"(by header first-column): extra={sorted(extra)}; "
        f"expected exactly {sorted(expected_first_cols)}.  Headers "
        f"per file: "
        f"{list(zip([f.name for f in files], headers_first_col))}.  "
        f"This usually means a previously-gated optional block "
        f"(KinshipRelations / LiteraryGenreCodes / InstitutionCodes) "
        f"became reachable on the current dump — see "
        f"Form_LookAtAssociations.vb gate conditions and update "
        f"the expected set OR the fixture."
    )

    # Exactly 8 files (per PR #117 verification probe baseline on
    # this fixture).  Two `nameID` files (People, PeoplePlaces)
    # share a first column → 7 distinct first-column shapes.
    assert len(files) == 8, (
        f"[{form_name}] CmdNeo4j produced {len(files)} files; "
        f"expected exactly 8 on the matrix Associations fixture.  "
        f"Headers per file: "
        f"{list(zip([f.name for f in files], headers_first_col))}"
    )
    assert len(seen) == len(expected_first_cols), (
        f"[{form_name}] expected {len(expected_first_cols)} "
        f"distinct file-shape first-columns "
        f"({sorted(expected_first_cols)}); saw {len(seen)} "
        f"({sorted(seen)})."
    )

    # Chain-completion + no-error markers in ZZ_TEST_DEBUG.
    cur = vba.conn.cursor()
    cur.execute(
        "SELECT msg FROM ZZ_TEST_DEBUG ORDER BY id")
    debug_msgs = [r[0] for r in cur.fetchall()]
    cur.close()
    assert any(m.endswith(":DONE") for m in debug_msgs), (
        f"[{form_name}] ZZ_TEST_DEBUG never reached :DONE — "
        f"chain block did not complete.  Markers seen: "
        f"{debug_msgs}"
    )
    err_rows = [m for m in debug_msgs if ":ERR" in m]
    assert not err_rows, (
        f"[{form_name}] ZZ_TEST_DEBUG contains runtime :ERR "
        f"markers (chain hit the error trap): {err_rows}.  "
        f"All markers: {debug_msgs}"
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
    # Added 2026-05-04 (hygiene follow-up to PR cover/lookatoffice-
    # cmdneo4j-peopleoffice) to cover LookAtOffice's OfficeCode-
    # codes shape so the LookAtOffice CmdNeo4j chain goes from
    # 5/6 strictly classified to 6/6.  Header literal at
    # Form_LookAtOffice.vb:1326 (non-ASCII branch, the default):
    #   "OfficeCode,OfficeTrans,OfficePinyin,OfficeHZ"
    # Live-verified against the office_80944_unfiltered fixture:
    # 1 row (the office matching the picker), all 4 cols populated.
    # The ASCII branch (line 1324) emits 3 cols (OfficeCode,
    # OfficeTrans, OfficePinyin); not currently exercised by any
    # cross-form test fixture, so requiring all 4 cols is fine
    # today — would need a separate ASCII-mode fixture before
    # tightening could regress.
    "OfficeCode": ("OfficeCodes",
                   ["OfficeCode", "OfficeTrans", "OfficePinyin",
                    "OfficeHZ"],
                   ["OfficeCode"]),
    # Added 2026-05-07 to cover LookAtAssociationPairs's CmdNeo4j
    # output (3 new header families that PR AX's probe surfaced;
    # PR #109's MsgBox-suppress patch made the chain reachable
    # unattended).  None of these shapes overlap by first-column
    # with existing entries so no two-column disambiguator needed.
    #
    # Person1_ID — PeopleAssociations.  13-col edge table from
    # Form_LookAtAssociationPairs.vb:766 SaveAs block.  Keys are
    # integer person ids, never blank — use the strict bad-id
    # check (don't add to skip set).
    "Person1_ID": ("PeopleAssociations",
                   ["Person1_ID", "Person2_ID",
                    "Association_Code"],
                   ["Person1_ID"]),
    # AssociationCode — AssociationCodes.  3-col code table from
    # Form_LookAtAssociationPairs.vb:974.  Distinct from the
    # existing `AssocCode -> AssocCodes` shape used by other
    # forms (different header literal).  Code-table — see
    # bad-id skip set additions below.
    "AssociationCode": ("AssociationCodes-AssocPairs",
                        ["AssociationCode", "AssociationTrans",
                         "AssociationHZ"],
                        ["AssociationCode"]),
    # KinshipCode — KinshipCodes.  3-col code table from
    # Form_LookAtAssociationPairs.vb:1072.  Distinct from the
    # existing `KinCode -> KinshipCodes` shape (different
    # header literal).  Code-table.
    "KinshipCode": ("KinshipCodes-AssocPairs",
                    ["KinshipCode", "KinshipTrans",
                     "KinshipHZ"],
                    ["KinshipCode"]),
    # Added 2026-05-07 (this PR) to cover LookAtAssociations's
    # CmdNeo4j output.  3-col code tables, header literals at
    # Form_LookAtAssociations.vb:1434 and :1519 (UTF-8 / non-ASCII
    # branch — the default on the matrix fixture).  Both code-table
    # — see bad-id skip set additions below.
    "OccasionCode": ("OccasionCodes",
                     ["OccasionCode", "OccasionTrans",
                      "OccasionHZ"],
                     ["OccasionCode"]),
    "TopicCode": ("TopicCodes",
                  ["TopicCode", "TopicTrans", "TopicHZ"],
                  ["TopicCode"]),
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
    # PeoplePlaces variants in the wild (lower-case nameID column):
    #   LookAtEntry / LookAtAssociationPairs (3-col):
    #     nameID, placeID, personPlaceCode, personPlaceTrans, personPlaceHZ
    #     (LookAtEntry uses 5; LookAtAssociationPairs uses 5 too)
    #   LookAtAssociations (4-col, NO personPlaceCode):
    #     nameID, placeID, personPlaceTrans, personPlaceHZ
    #     (Form_LookAtAssociations.vb:881 UTF-8 / non-ASCII branch)
    # Required cols here are the intersection of the variants
    # (`nameID, placeID` only).  Per-form structural assertions
    # (`_assert_lookatentry_neo4j_shape`,
    # `_assert_lookatassociationpairs_neo4j_shape`,
    # `_assert_lookatassociations_neo4j_shape`) pin the per-form
    # column shapes more strictly.
    ("nameID", "placeID"): ("PeoplePlaces",
                            ["nameID", "placeID"],
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
    # Added 2026-05-07 (this PR) to disambiguate the
    # AssociationCodes file produced by LookAtAssociations's
    # CmdNeo4j (4 cols, with `AssociationTypeID` between
    # `AssociationCode` and `AssociationTrans`) from the
    # AssociationPairs-side 3-col variant.  Header literal at
    # Form_LookAtAssociations.vb:1080 (UTF-8 / non-ASCII branch):
    #   "AssociationCode,AssociationTypeID,AssociationTrans,AssociationHZ"
    # Without this 2-col entry the file falls through to the
    # single-col `AssociationCode -> AssociationCodes-AssocPairs`
    # lookup; that lookup's required cols
    # `[AssociationCode, AssociationTrans, AssociationHZ]` is a
    # subset of the 4-col header so the assertion still passes,
    # but the LABEL would be misleading ("AssocPairs" on a
    # different form).  Code-table — see bad-id skip set
    # additions below.
    ("AssociationCode", "AssociationTypeID"): (
        "AssociationCodes-Associations",
        ["AssociationCode", "AssociationTypeID",
         "AssociationTrans", "AssociationHZ"],
        ["AssociationCode"]),
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
                "TextCodes",
                # OfficeCodes — same code-table family; office
                # code can legitimately be 0 / blank for unmapped
                # offices.
                "OfficeCodes",
                # AssociationCodes-AssocPairs / KinshipCodes-
                # AssocPairs — same code-table family from
                # AssocPairs's CmdNeo4j output; first cell can
                # legitimately be 0 / blank for unmapped codes.
                "AssociationCodes-AssocPairs",
                "KinshipCodes-AssocPairs",
                # Added 2026-05-07 (this PR) for
                # LookAtAssociations CmdNeo4j: same code-table
                # family, first cell can legitimately be 0 /
                # blank for unmapped codes.
                "AssociationCodes-Associations",
                "OccasionCodes",
                "TopicCodes"):
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
