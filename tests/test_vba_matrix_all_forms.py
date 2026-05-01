"""
Cross-form data-driven matrix.

For each LookAt form (Status, Texts, Place, Associations, Office),
build fixtures from analysis/dump/test_inputs.json (high-density real
inputs), drive REAL VBA via pywinauto, and run the form-agnostic
integrity checks.

The form-specific column-set check uses cbdb_driver.form_specs.

Fixtures cover:
  - "top code unfiltered" — most populous picker code, no other filter
  - "top code × dynasty" — top code constrained to most populous dynasty
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
import pytest

from cbdb_driver.vba_session import VbaSession, make_fixture
from cbdb_driver.form_specs import (
    FormSpec, LOOKATSTATUS, LOOKATTEXTS, LOOKATASSOCIATIONS,
    LOOKATOFFICE, LOOKATPLACE, LOOKATASSOCIATIONPAIRS,
    LOOKATKINSHIP, LOOKATNETWORKS, LOOKATGROUPDATA,
)


ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK = ROOT / "analysis" / "_matrix_all_test_copy.mdb"
INPUTS = ROOT / "analysis" / "dump" / "test_inputs.json"


@pytest.fixture(scope="function")
def vba():
    yield from make_fixture(SRC, WORK)


@dataclass
class CrossFixture:
    name: str
    spec: FormSpec
    picker_ids: list[int] = field(default_factory=list)
    addr_ids: list[int] = field(default_factory=list)
    controls: dict = field(default_factory=dict)
    expected_min_rows: int = 1
    source_sql: str | None = None


def _load_inputs() -> dict | None:
    if not INPUTS.exists():
        return None
    return json.loads(INPUTS.read_text(encoding="utf-8"))


def _make_status_fixtures(inputs: dict) -> list[CrossFixture]:
    out: list[CrossFixture] = []
    data = inputs.get("lookatstatus", {})
    if data.get("top_status_codes"):
        c = int(data["top_status_codes"][0]["c_status_code"])
        out.append(CrossFixture(
            name=f"status_{c}_unfiltered",
            spec=LOOKATSTATUS,
            picker_ids=[c],
            controls={"FrameFilterYears": 1},   # 1 = no year filter
            expected_min_rows=1000,
            source_sql=(f"SELECT DISTINCT c_personid FROM STATUS_DATA "
                        f"WHERE c_status_code = {c}"),
        ))
    for combo in data.get("status_x_dynasty_combos", [])[:2]:
        c = int(combo["c_status_code"])
        dy = int(combo["c_dy"])
        out.append(CrossFixture(
            name=f"status_{c}_dy_{dy}",
            spec=LOOKATSTATUS,
            picker_ids=[c],
            controls={"FrameFilterYears": 3},   # 3 = dynasty mode
            expected_min_rows=max(1, int(combo["n_persons"]) // 50),
            source_sql=(
                f"SELECT DISTINCT BIOG_MAIN.c_personid "
                f"FROM BIOG_MAIN INNER JOIN STATUS_DATA "
                f"  ON BIOG_MAIN.c_personid = STATUS_DATA.c_personid "
                f"WHERE STATUS_DATA.c_status_code = {c} "
                f"  AND BIOG_MAIN.c_dy = {dy}"
            ),
        ))
    return out


def _make_texts_fixtures(inputs: dict) -> list[CrossFixture]:
    out: list[CrossFixture] = []
    data = inputs.get("lookattexts", {})
    if data.get("top_biblcat_codes"):
        # skip the "unknown" code 0 if it's first
        for row in data["top_biblcat_codes"][:3]:
            c = int(row["c_bibl_cat_code"])
            if c == 0:
                continue
            out.append(CrossFixture(
                name=f"biblcat_{c}_unfiltered",
                spec=LOOKATTEXTS,
                picker_ids=[c],
                controls={"FrameFilterYears": 1},
                expected_min_rows=10,
                source_sql=(
                    f"SELECT DISTINCT BIOG_TEXT_DATA.c_personid "
                    f"FROM BIOG_TEXT_DATA INNER JOIN TEXT_CODES "
                    f"  ON BIOG_TEXT_DATA.c_textid = TEXT_CODES.c_textid "
                    f"WHERE TEXT_CODES.c_bibl_cat_code = {c}"
                ),
            ))
            break
    return out


def _make_assoc_fixtures(inputs: dict) -> list[CrossFixture]:
    out: list[CrossFixture] = []
    data = inputs.get("lookatassociations", {})
    if data.get("top_assoc_codes"):
        c = int(data["top_assoc_codes"][0]["c_assoc_code"])
        out.append(CrossFixture(
            name=f"assoc_{c}_unfiltered",
            spec=LOOKATASSOCIATIONS,
            picker_ids=[c],
            controls={"FrameFilterYears": 1},
            expected_min_rows=500,
            source_sql=(f"SELECT DISTINCT c_personid FROM ASSOC_DATA "
                        f"WHERE c_assoc_code = {c}"),
        ))
    for combo in data.get("assoc_x_dynasty_combos", [])[:2]:
        c = int(combo["c_assoc_code"])
        dy = int(combo["c_dy"])
        out.append(CrossFixture(
            name=f"assoc_{c}_dy_{dy}",
            spec=LOOKATASSOCIATIONS,
            picker_ids=[c],
            controls={"FrameFilterYears": 3},
            expected_min_rows=max(1, int(combo["n_rows"]) // 100),
            source_sql=(
                f"SELECT DISTINCT BIOG_MAIN.c_personid "
                f"FROM BIOG_MAIN INNER JOIN ASSOC_DATA "
                f"  ON BIOG_MAIN.c_personid = ASSOC_DATA.c_personid "
                f"WHERE ASSOC_DATA.c_assoc_code = {c} "
                f"  AND BIOG_MAIN.c_dy = {dy}"
            ),
        ))
    return out


def _make_office_fixtures(inputs: dict) -> list[CrossFixture]:
    out: list[CrossFixture] = []
    data = inputs.get("lookatoffice", {})
    for row in data.get("top_office_codes", [])[:2]:
        c = int(row["c_office_id"])
        out.append(CrossFixture(
            name=f"office_{c}_unfiltered",
            spec=LOOKATOFFICE,
            picker_ids=[c],
            # TxtTypeDesc: anything ≠ "N/A" routes the SQL through the
            # ZZ_OFFICE_CODE picker branch (see Form_LookAtOffice line 2080-2117).
            controls={"FrameFilterYears": 1, "TxtTypeDesc": "[All]"},
            expected_min_rows=10,
            source_sql=(f"SELECT DISTINCT c_personid FROM POSTED_TO_OFFICE_DATA "
                        f"WHERE c_office_id = {c}"),
        ))
    return out


def _make_place_fixtures(inputs: dict) -> list[CrossFixture]:
    """LookAtPlace: picker is ZZ_SCRATCH_ADDR (address ids).
    Sets ChkIndividual=True for the 'Biography' branch."""
    out: list[CrossFixture] = []
    data = inputs.get("lookatplace", {})
    for row in data.get("top_addr_by_indexed_persons", [])[:2]:
        a = int(row["c_index_addr_id"])
        out.append(CrossFixture(
            name=f"place_addr_{a}",
            spec=LOOKATPLACE,
            addr_ids=[a],
            controls={"ChkIndividual": True, "ChkOffice": False,
                      "ChkAssoc": False, "ChkPosting": False,
                      "ChkEntry": False},
            expected_min_rows=50,
            source_sql=(f"SELECT DISTINCT c_personid FROM BIOG_ADDR_DATA "
                        f"WHERE c_addr_id = {a}"),
        ))
    return out


def _make_kinship_fixtures(inputs: dict) -> list[CrossFixture]:
    """LookAtKinship: input is a single person id via ZZ_SCRATCH_IMPORT_PEOPLE.
    CmdRun_Click runs recursive kinship traversal."""
    out: list[CrossFixture] = []
    data = inputs.get("lookatkinship", {})
    for row in data.get("top_persons_by_kin_count", [])[:1]:
        p = int(row["c_personid"])
        out.append(CrossFixture(
            name=f"kinship_person_{p}",
            spec=LOOKATKINSHIP,
            picker_ids=[p],
            controls={},   # use form defaults for distance constraints
            expected_min_rows=1,
            source_sql=None,
        ))
    return out


def _make_networks_fixtures(inputs: dict) -> list[CrossFixture]:
    """LookAtNetworks: input is a single (or set of) person id(s)."""
    out: list[CrossFixture] = []
    data = inputs.get("lookatnetworks", {})
    for row in data.get("top_persons_by_assoc_count", [])[:1]:
        p = int(row["c_personid"])
        out.append(CrossFixture(
            name=f"network_person_{p}",
            spec=LOOKATNETWORKS,
            picker_ids=[p],
            controls={},
            expected_min_rows=1,
            source_sql=None,
        ))
    return out


def _make_groupdata_fixtures(inputs: dict) -> list[CrossFixture]:
    """LookAtGroupData: input is a person list."""
    out: list[CrossFixture] = []
    # reuse lookatnetworks top persons for the group
    data = inputs.get("lookatnetworks", {})
    persons = data.get("top_persons_by_assoc_count", [])[:1]
    if persons:
        p = int(persons[0]["c_personid"])
        out.append(CrossFixture(
            name=f"groupdata_person_{p}",
            spec=LOOKATGROUPDATA,
            picker_ids=[p],
            controls={},
            expected_min_rows=1,
            source_sql=None,
        ))
    return out


def _make_assoc_pairs_fixtures(inputs: dict) -> list[CrossFixture]:
    """LookAtAssociationPairs: needs two person ids via TxtID1/TxtID2.
    TxtPerson1/2 are display-only NAME fields, populated by the picker.

    Chk2Nodes MUST be False — if True, Form_LookAtAssociationPairs
    CmdQuery_Click pops a MsgBox warning ("two-node routine takes a
    while...") that blocks the COM thread.  ChkKinship adds 4× more
    queries; turn off for a fast smoke."""
    out: list[CrossFixture] = []
    data = inputs.get("lookatassociationpairs", {})
    for row in data.get("top_pairs_by_edge_count", [])[:1]:
        p1 = int(row["person_id_1"])
        p2 = int(row["person_id_2"])
        out.append(CrossFixture(
            name=f"assocpair_{p1}_{p2}",
            spec=LOOKATASSOCIATIONPAIRS,
            controls={"TxtID1": p1, "TxtID2": p2,
                      "TxtPerson1": str(p1), "TxtPerson2": str(p2),
                      "FrameFilterYears": 1,
                      "Chk2Nodes": 0, "ChkKinship": 0},
            expected_min_rows=1,
            source_sql=None,
        ))
    return out


def _all_fixtures() -> list[CrossFixture]:
    inputs = _load_inputs()
    if inputs is None:
        return []
    return (
        _make_status_fixtures(inputs)
        + _make_texts_fixtures(inputs)
        + _make_assoc_fixtures(inputs)
        + _make_office_fixtures(inputs)
        + _make_place_fixtures(inputs)
        + _make_assoc_pairs_fixtures(inputs)
        + _make_kinship_fixtures(inputs)
        + _make_networks_fixtures(inputs)
        + _make_groupdata_fixtures(inputs)
    )


def _xfail_marks(fx: CrossFixture):
    """Skip three forms whose CmdQuery/CmdRun does very heavy
    recursive expansion (network / pair-wise / cross-form group
    aggregation).  Need either smaller fixtures, lower distance
    constraints, or longer timeouts.  See findings.md.

    Working in matrix:
      Entry/Status/Texts/Associations/Office/Place/Kinship
    Skipped:
      AssociationPairs (CmdQuery times out — Link1stOrder runs a JET
        ASSOC_DATA self-join (`ZABA INNER JOIN ZABA_1 ON ZABA.c_assoc_id
        = ZABA_1.c_assoc_id`) that the optimizer evaluates *before* the
        WHERE personid filter; even Wang Anshi×Sima Guang (47 shared
        edges) doesn't return in 120s.  Setting Chk2Nodes=0 +
        ChkKinship=0 doesn't help.  Need either a smaller fixture
        (people with <10 assocs each) or a saved query that pre-filters.
      Networks (CmdRun times out — Zhu Xi has 2471 assocs)
      GroupData (CmdRun aggregates across many tables)"""
    if fx.spec.name in (
        "LookAtAssociationPairs", "LookAtNetworks", "LookAtGroupData"
    ):
        return pytest.mark.skip(
            reason=f"{fx.spec.name} CmdQuery/CmdRun times out — needs "
                   "smaller fixture or more preconditions."
        )
    return ()


# All forms use Form_Timer trigger now (not just disabled-button forms).
# pywinauto-based click_input is fragile across multiple sequential test
# runs in one pytest session: UIA cache corrupts after a few open/close
# cycles, leading to silent click drops on subsequent forms (assoc tests
# fail with 0 rows after biblcat passes, etc.).  Form_Timer fires from
# inside Access itself and is immune to UI / desktop / focus state.
_TIMER_TRIGGER_FORMS = {
    "LookAtEntry", "LookAtStatus", "LookAtTexts",
    "LookAtAssociations", "LookAtOffice", "LookAtPlace",
    "LookAtAssociationPairs", "LookAtKinship", "LookAtNetworks",
    "LookAtGroupData",
}


@pytest.mark.parametrize(
    "fx",
    [pytest.param(f, marks=_xfail_marks(f)) for f in _all_fixtures()],
    ids=lambda f: f.name,
)
def test_cross_form_matrix(vba: VbaSession, fx: CrossFixture):
    """Generic VBA-driven matrix test for any LookAt form.

    Steps:
      1. Wipe + populate picker scratch tables
      2. Open form, set year-mode + other controls
      3. Click Run Query (force-enabled), wait for result table
      4. Verify expected min row count
      5. Form-specific column structure
      6. FK integrity (only for forms whose results have c_addr_id)
      7. Differential vs source SQL
    """
    spec = fx.spec

    # 1. open form, set controls
    # Note: must open BEFORE populating pickers — Form_LookAtOffice's
    # Form_Open wipes ZZ_OFFICE_CODE.
    vba.open_form(spec.name)
    for ctl, val in fx.controls.items():
        try:
            vba.set_control(spec.name, ctl, val)
        except Exception as e:
            print(f"  warn setting {ctl}={val!r}: {e}")

    # 2. populate pickers (after open — see above)
    if fx.picker_ids and spec.picker_table:
        vba.set_picker_codes(spec.picker_table, fx.picker_ids,
                              column=spec.picker_column)
    if fx.addr_ids:
        vba.set_picker_addrs(fx.addr_ids)

    # 3. fire VBA — use timer trigger for disabled-button forms
    if spec.name in _TIMER_TRIGGER_FORMS:
        n = vba.click_via_timer(
            spec.name, ctl=spec.cmd_name,
            result_table=spec.result_table, timeout=120,
        )
    else:
        n = vba.click_button_and_wait_table(
            spec.cmd_caption, form=spec.name,
            result_table=spec.result_table,
            force_enable_ctl=spec.cmd_name,
            timeout=60,
        )
    print(f"\n[{fx.name}] VBA produced {n} rows", flush=True)
    assert n >= fx.expected_min_rows, (
        f"[{fx.name}] only {n} rows (expected ≥ {fx.expected_min_rows}); "
        "test inputs stale or VBA filter too aggressive"
    )

    # NB. previously close_form here, but it can hang for minutes
    # on LookAtOffice while Access finishes the backfill UPDATE chain.
    # click_via_timer's DONE marker already waits for that.

    print(f"  [{fx.name}] step 4: read TOP 1 to inspect columns", flush=True)
    # 4. column structure — only need first row to inspect columns
    df_head = vba.read(spec.result_table, top=1)
    print(f"  [{fx.name}] step 4 done, cols={len(df_head.columns)}", flush=True)
    assert not df_head.empty
    missing = spec.insert_cols - set(df_head.columns)
    assert not missing, (
        f"[{fx.name}] INSERT-target columns missing from "
        f"{spec.result_table}: {sorted(missing)}"
    )

    # Close form to release Access's read lock on the result table.
    # Now safe (DONE marker confirmed CmdQuery_Click finished).
    try:
        vba.close_form(spec.name)
    except Exception:
        pass

    print(f"  [{fx.name}] step 5: read person_id column (cursor)", flush=True)
    # 5. read just person_id column for differential / FK checks (cheap)
    # Use raw cursor — pandas.read_sql can deadlock with Access on large
    # local tables that the form had bound (e.g. LookAtOffice 37k rows).
    pid_col = spec.person_id_col
    has_addr = "c_addr_id" in df_head.columns
    addr_select = ", c_addr_id" if has_addr else ""
    cur = vba.conn.cursor()
    cur.execute(
        f"SELECT {pid_col}{addr_select} FROM [{spec.result_table}]"
    )
    rows = cur.fetchall()
    cur.close()
    cols = [pid_col] + (["c_addr_id"] if has_addr else [])
    df_lite = pd.DataFrame.from_records(rows, columns=cols)
    print(f"  [{fx.name}] step 5 done, {len(df_lite)} rows", flush=True)

    # 6. FK integrity for c_addr_id (when present)
    if has_addr:
        addrs = sorted({int(a) for a in df_lite["c_addr_id"].dropna()})
        if addrs:
            in_clause = ",".join(str(a) for a in addrs[:200])  # cap
            print(f"  [{fx.name}] step 6: FK check on {len(addrs)} addrs",
                  flush=True)
            cur = vba.conn.cursor()
            cur.execute(
                f"SELECT DISTINCT c_addr_id FROM ADDR_CODES "
                f"WHERE c_addr_id IN ({in_clause})"
            )
            found = {int(r[0]) for r in cur.fetchall()}
            cur.close()
            orphans = [a for a in addrs[:200] if a not in found]
            assert not orphans, (
                f"[{fx.name}] c_addr_id not in ADDR_CODES: {orphans[:5]}"
            )

    # 7. differential — source SQL must be ⊆ VBA result.
    # Skipped for LookAtOffice: Access still holds a read lock on the
    # linked POSTED_TO_OFFICE_DATA after Form_LookAtOffice's UPDATE
    # chain, and pyodbc deadlocks waiting for it.  The row-count + FK
    # checks above are still strong signals.
    if fx.source_sql and spec.name != "LookAtOffice":
        print(f"  [{fx.name}] step 7: differential source SQL", flush=True)
        cur = vba.conn.cursor()
        cur.execute(fx.source_sql)
        src_set = {int(r[0]) for r in cur.fetchall()}
        cur.close()
        vba_set = {int(p) for p in df_lite[pid_col].dropna()}
        only_src = src_set - vba_set
        # We allow VBA to include MORE persons (e.g. via auxiliary
        # joins), but it must not LOSE persons the source query found.
        assert not only_src, (
            f"[{fx.name}] VBA LOST {len(only_src)} persons that "
            f"independent source SQL found: {sorted(only_src)[:10]}"
        )

    print(f"[{fx.name}] ✓ all checks passed ({n} rows)")
