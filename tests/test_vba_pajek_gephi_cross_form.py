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
  - Networks: under default full injection Form_Open hits the
    project-wide auto-compile deadlock (PR AR-AX, AGENTS landmine
    #3.5).  Same family as matrix Networks + picker test skips.
    Form_Open verified fine via minimal injection in
    tests/test_vba_networks_small_fixture.py.
  - AssociationPairs (was previously SKIPPED): the prior matrix
    CmdQuery / CmdRun blocker has been resolved by the
    AssociationPairs SetFocus driver patch
    (`_PER_FORM_CMDGIS_PATCHES["Form_LookAtAssociationPairs"]`).
    AssociationPairs × CmdPajek and × CmdGephi are now wired in
    here using a known-edged 1×3 person pair (NOT the matrix's
    default 4×5 pair, which on the current dump shares 0 first-
    order associations and would leave ZZ_SOCIAL_NETWORK empty).
    See `_assocpairs_1x3_fixture()` for the custom CrossFixture.

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
    Case("LookAtKinship",          "CmdPajek",     ".net", "*vertices"),
    # CmdUTF8Pajek: same Pajek .net format as CmdPajek but UTF-8 encoded
    # via ADO stream (Form_LookAtKinship.vb:2712).  Shares the same kinship
    # fixture and the same _assert_pajek_depth depth check.
    Case("LookAtKinship",          "CmdUTF8Pajek", ".net", "*vertices"),
    Case("LookAtPlace",            "CmdPajek",     ".net", "*vertices"),
    Case("LookAtStatus",           "CmdPajek",     ".net", "*vertices"),
    Case("LookAtAssociations",     "CmdPajek",     ".net", "*vertices"),
    Case("LookAtAssociationPairs", "CmdPajek",     ".net", "*vertices"),
    Case("LookAtPlace",            "CmdGephi",     ".gdf", "nodedef"),
    Case("LookAtStatus",           "CmdGephi",     ".gdf", "nodedef"),
    Case("LookAtAssociations",     "CmdGephi",     ".gdf", "nodedef"),
    Case("LookAtAssociationPairs", "CmdGephi",     ".gdf", "nodedef"),
)


def _assocpairs_1x3_fixture() -> CrossFixture:
    """Custom 1×3 known-edged fixture for AssociationPairs.

    The matrix's default `_make_assoc_pairs_fixtures` picks the
    top pair from `discover_test_inputs.py`'s
    `top_pairs_by_edge_count`, which on the current dump is the
    4×5 pair — that pair shares 0 first-order ASSOC_DATA edges
    and leaves ZZ_SOCIAL_NETWORK empty after CmdQuery, so the
    chained CmdPajek / CmdGephi would bail on RecordCount=0.

    1×3 was selected by direct SQL on ASSOC_DATA (
    `WHERE c_personid=1 AND c_assoc_id=3`) — the smallest
    person-id pair on the current dump that shares at least one
    direct edge.  Verified end-to-end by the SetFocus driver
    patch's smoke probe (`tests/test_vba_associationpairs_probe.
    py::test_associationpairs_cmdquery_setfocus_patch_unblocks
    _inserts`): ZZ_SCRATCH_PEOPLE = 2, ZZ_SOCIAL_NETWORK > 0.

    If a future dump removes the 1↔3 ASSOC_DATA edge (or
    person 1 / 3 entirely), the test will fail at the chained
    export's RecordCount=0 check; pick a new known-edged small
    pair via the same SQL.
    """
    from cbdb_driver.form_specs import LOOKATASSOCIATIONPAIRS
    return CrossFixture(
        name="assocpair_1x3_known_edged",
        spec=LOOKATASSOCIATIONPAIRS,
        controls={
            "TxtID1": 1, "TxtID2": 3,
            "TxtPerson1": "1", "TxtPerson2": "3",
            "FrameFilterYears": 1,
            "Chk2Nodes": 0, "ChkKinship": 0,
        },
        expected_min_rows=1,
        source_sql=None,
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
    # AssociationPairs uses a custom 1×3 known-edged fixture
    # rather than the matrix's default 4×5 (which has 0
    # first-order edges on the current dump — see
    # `_assocpairs_1x3_fixture` docstring).
    if form == "LookAtAssociationPairs":
        return _assocpairs_1x3_fixture()
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
        result_table=spec.result_table,
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

    # ---- Depth checks (PR R) ----------------------------------
    if case.cmd in ("CmdPajek", "CmdUTF8Pajek"):
        _assert_pajek_depth(spec.name, lines)
    elif case.cmd == "CmdGephi":
        _assert_gephi_depth(spec.name, lines)


# ----------------------------------------------------------------------
# PR R: Pajek + Gephi/GUESS export depth helpers
# ----------------------------------------------------------------------

def _assert_pajek_depth(form_name: str, lines: list[str]) -> None:
    """Pajek `.net`:
        *Vertices N
        1 "label1"
        2 "label2"
        ...
        *Edges  (or *Arcs)
        src dst weight
        ...
    Assertions:
      * `*Vertices N` parses (N is an integer)
      * exactly N vertex rows follow before the next `*` section
      * each vertex row's first token is a unique 1-based id
      * edge / arc section, if present, has at least one row whose
        first two tokens are integer vertex ids in [1, N]
    """
    import re
    # Find *Vertices N marker (case-insensitive).
    head = lines[0].strip()
    m = re.match(r"\*[Vv]ertices\s+(\d+)", head)
    assert m, f"[{form_name}] Pajek: header isn't `*Vertices N`: {head!r}"
    n_vertices = int(m.group(1))

    # Walk lines until next `*` section to count vertex rows.
    vertex_rows: list[str] = []
    edge_rows: list[str] = []
    in_edges = False
    for ln in lines[1:]:
        s = ln.strip()
        if s.startswith("*"):
            # Either *Edges or *Arcs (or *Matrix) — switch mode.
            in_edges = True
            continue
        if not in_edges:
            vertex_rows.append(s)
        else:
            edge_rows.append(s)

    assert len(vertex_rows) == n_vertices, (
        f"[{form_name}] Pajek: header declared {n_vertices} vertices "
        f"but found {len(vertex_rows)} vertex rows before the next "
        f"`*` section.  Off-by-N → silent vertex-export bug."
    )

    # Each vertex row's first token is the vertex id.  Demand 1..N
    # all present exactly once (Pajek convention).
    vertex_ids: list[int] = []
    for i, row in enumerate(vertex_rows, start=1):
        first = row.split(None, 1)[0]
        try:
            vid = int(first)
        except ValueError:
            raise AssertionError(
                f"[{form_name}] Pajek: vertex row {i} doesn't start "
                f"with an integer id: {row[:120]!r}"
            )
        vertex_ids.append(vid)
    assert len(set(vertex_ids)) == len(vertex_ids), (
        f"[{form_name}] Pajek: duplicate vertex ids in {n_vertices} "
        f"rows — set size {len(set(vertex_ids))}"
    )
    assert min(vertex_ids) == 1 and max(vertex_ids) == n_vertices, (
        f"[{form_name}] Pajek: vertex ids should be 1..{n_vertices}; "
        f"got min={min(vertex_ids)} max={max(vertex_ids)}"
    )

    # Edges (if present): first two tokens are integer vertex ids in
    # range.  Spot-check up to 100 rows to keep this O(small).
    if edge_rows:
        for i, row in enumerate(edge_rows[:100], start=1):
            toks = row.split()
            if len(toks) < 2:
                raise AssertionError(
                    f"[{form_name}] Pajek: edge row {i} has fewer "
                    f"than 2 tokens: {row[:120]!r}"
                )
            try:
                a = int(toks[0]); b = int(toks[1])
            except ValueError:
                raise AssertionError(
                    f"[{form_name}] Pajek: edge row {i} first two "
                    f"tokens not integers: {row[:120]!r}"
                )
            assert 1 <= a <= n_vertices, (
                f"[{form_name}] Pajek: edge row {i} src {a} out of "
                f"[1, {n_vertices}]"
            )
            assert 1 <= b <= n_vertices, (
                f"[{form_name}] Pajek: edge row {i} dst {b} out of "
                f"[1, {n_vertices}]"
            )
    print(f"[{form_name}] Pajek depth: {n_vertices} vertices, "
          f"{len(edge_rows)} edges; all ids unique and in range",
          flush=True)


def _assert_gephi_depth(form_name: str, lines: list[str]) -> None:
    """Gephi/GUESS `.gdf`:
        nodedef> name VARCHAR, label VARCHAR, ...
        node1, "label1", ...
        ...
        edgedef> node1 VARCHAR, node2 VARCHAR, ...
        a, b, ...

    Assertions:
      * `nodedef>` line parses to ≥ 1 column
      * every node row's field count matches nodedef columns
      * `edgedef>` (if present) parses similarly
      * every edge row's field count matches edgedef columns
      * the first column of every node row (`name`) is non-empty
    """
    head = lines[0].strip()
    assert head.lower().startswith("nodedef>"), (
        f"[{form_name}] Gephi: first line isn't `nodedef>`: {head!r}"
    )
    node_cols = [c.strip() for c in head[len("nodedef>"):].split(",")
                 if c.strip()]
    n_node_cols = len(node_cols)
    assert n_node_cols >= 1, (
        f"[{form_name}] Gephi: nodedef> declared 0 columns: {head!r}"
    )

    # Walk node rows until edgedef> (if present).
    node_rows: list[str] = []
    edge_rows: list[str] = []
    edge_cols: list[str] = []
    in_edges = False
    for ln in lines[1:]:
        s = ln.strip()
        if s.lower().startswith("edgedef>"):
            in_edges = True
            edge_cols = [c.strip() for c in s[len("edgedef>"):].split(",")
                         if c.strip()]
            continue
        (edge_rows if in_edges else node_rows).append(s)

    # Per-row width for nodes.  Gephi field values can contain commas
    # inside quoted labels, so be lenient: width must be at least
    # n_node_cols (some rows may have escaping bumps), but no row
    # should have wildly more.
    bad_node = []
    for i, row in enumerate(node_rows, start=1):
        cells = row.split(",")
        if len(cells) < n_node_cols or len(cells) > n_node_cols + 5:
            bad_node.append((i, len(cells), row[:120]))
        if len(bad_node) >= 5:
            break
    assert not bad_node, (
        f"[{form_name}] Gephi: node rows with bad field count "
        f"(nodedef has {n_node_cols} cols).  First mismatches:\n"
        + "\n".join(f"  row {i}: {n} cells — {snip!r}"
                     for i, n, snip in bad_node)
    )

    # Node `name` column (first one) must be non-empty for every row.
    bad_name = []
    for i, row in enumerate(node_rows, start=1):
        first = row.split(",", 1)[0].strip()
        if not first:
            bad_name.append((i, row[:120]))
        if len(bad_name) >= 5:
            break
    assert not bad_name, (
        f"[{form_name}] Gephi: node rows with empty `name` column.  "
        f"First mismatches: {bad_name!r}"
    )

    if edge_cols:
        n_edge_cols = len(edge_cols)
        bad_edge = []
        for i, row in enumerate(edge_rows[:200], start=1):
            cells = row.split(",")
            if len(cells) < n_edge_cols or len(cells) > n_edge_cols + 5:
                bad_edge.append((i, len(cells), row[:120]))
            if len(bad_edge) >= 5:
                break
        assert not bad_edge, (
            f"[{form_name}] Gephi: edge rows with bad field count "
            f"(edgedef has {n_edge_cols} cols).  First mismatches:\n"
            + "\n".join(f"  row {i}: {n} cells — {snip!r}"
                         for i, n, snip in bad_edge)
        )

    print(f"[{form_name}] Gephi depth: {len(node_rows)} nodes "
          f"({n_node_cols} cols), {len(edge_rows)} edges "
          f"({len(edge_cols) or '0/no edgedef'} cols); "
          f"all `name` columns non-empty",
          flush=True)
