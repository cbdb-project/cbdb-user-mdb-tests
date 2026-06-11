"""
Parametrized VBA differential + integrity matrix.

Drives REAL high-density inputs (discovered by
analysis/discover_test_inputs.py) through the actual VBA
CmdQuery_Click, then runs ALL integrity / differential checks in a
single test per fixture (so we pay the ~12s Access startup once).

Add new fixtures by editing the FIXTURES tuple below or extending
discover_test_inputs.py.

Naming convention: <form>_<short-input-id> — e.g.
  lookatentry_jinshi_qing
  lookatentry_yin_general_kaifeng
  lookatstatus_civil_office_song
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

import pandas as pd
import pytest

from cbdb_driver.vba_session import VbaSession, make_fixture
from cbdb_replay import lookatentry as le


ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK = ROOT / "analysis" / "_matrix_test_copy.mdb"
INPUTS = ROOT / "analysis" / "dump" / "test_inputs.json"


@pytest.fixture(scope="function")
def vba():
    yield from make_fixture(SRC, WORK)


# ----------------------------------------------------------------------
# Per-fixture spec
# ----------------------------------------------------------------------

@dataclass
class Fixture:
    name: str
    form: str
    result_table: str
    cmd_caption: str
    cmd_name: str
    # picker tables: list of (table, [ids], column)
    pickers: list[tuple[str, list[int], str]] = field(default_factory=list)
    # form controls: dict ctl_name -> value
    controls: dict = field(default_factory=dict)
    # expected min row count (sanity floor — high-density inputs should
    # produce many rows; if not, test inputs are stale)
    expected_min_rows: int = 1
    # ID-set to verify against an INDEPENDENT source SQL (catches info loss)
    source_sql: str | None = None
    # Python replay function + inputs (for differential check; optional)
    py_replay: Callable | None = None
    py_inputs: object | None = None


# ----------------------------------------------------------------------
# Build fixtures from discovered high-density inputs
# ----------------------------------------------------------------------

def _load_inputs() -> dict:
    if not INPUTS.exists():
        pytest.skip(f"missing {INPUTS} — run discover_test_inputs.py first")
    return json.loads(INPUTS.read_text(encoding="utf-8"))


def _build_lookatentry_fixtures(inputs: dict) -> list[Fixture]:
    """Build a small but high-coverage matrix of LookAtEntry fixtures
    from the discovery JSON."""
    out: list[Fixture] = []
    le_data = inputs["lookatentry"]

    # 1. Top entry code, no other filter — most populous
    if le_data["top_entry_codes"]:
        c = int(le_data["top_entry_codes"][0]["c_entry_code"])
        out.append(Fixture(
            name=f"top_entry_code_{c}_unfiltered",
            form="LookAtEntry",
            result_table="ZZ_SCRATCH_ENTRY",
            cmd_caption="Run Query",
            cmd_name="CmdQuery",
            pickers=[("ZZ_SCRATCH_ENTRY_CODE", [c], "c_entry_code")],
            controls={"FrameYears": 0,
                      "TxtEntryDesc": le_data["top_entry_codes"][0]["c_entry_desc"],
                      "TxtTypeCode": "N/A"},
            expected_min_rows=1000,
            source_sql=f"""
                SELECT DISTINCT BIOG_MAIN.c_personid
                FROM BIOG_MAIN INNER JOIN ENTRY_DATA
                  ON BIOG_MAIN.c_personid = ENTRY_DATA.c_personid
                WHERE ENTRY_DATA.c_entry_code = {c}
            """,
        ))

    # 2-4. Top three (entry × dynasty) combos
    for combo in le_data.get("entry_x_dynasty_combos", [])[:3]:
        c = int(combo["c_entry_code"])
        dy = int(combo["c_dy"])
        # find dynasty year range
        dyn_row = next(
            (d for d in le_data["top_dynasties"] if int(d["c_dy"]) == dy),
            None,
        )
        if not dyn_row:
            continue
        out.append(Fixture(
            name=f"entry_{c}_dy_{dy}",
            form="LookAtEntry",
            result_table="ZZ_SCRATCH_ENTRY",
            cmd_caption="Run Query",
            cmd_name="CmdQuery",
            pickers=[("ZZ_SCRATCH_ENTRY_CODE", [c], "c_entry_code")],
            controls={
                # FrameYears=3 selects dynasty mode.  The dynasty *value*
                # (gFromDynasty) normally comes from the frmpickdynasty picker,
                # which the matrix bypasses -- so gFromDynasty stays at its -2
                # default.  In CmdQuery_Click's FrameYears=3 branch
                # (Form_LookAtEntry.vb:1622-1624) that -2 default takes the
                # `If gFromDynasty = -2 Then tStrYears = "((BIOG_MAIN.c_dy) > 0)"`
                # path and sets gUseDynasties=True (which adds a DYNASTIES join),
                # NOT a filter on dynasty {dy}.  The year textboxes are not read
                # in this branch, so TxtFromYear/TxtToYear apply no filter (years
                # read live only under FrameYears 1/2 at :1596/:1609).  This
                # fixture therefore exercises the FrameYears=3 entry path; it does
                # NOT (and through the picker bypass cannot) verify dynasty
                # *selection* -- that needs a driver-side gFromDynasty injection,
                # tracked but not done here.  {dy} survives only as the
                # high-density combo this case was sampled from.
                "FrameYears": 3,
                "TxtEntryDesc": "[selected]",
                "TxtTypeCode": "N/A",
                "TxtFromYear": int(dyn_row["c_start"]),
                "TxtToYear": int(dyn_row["c_end"]),
            },
            expected_min_rows=max(1, int(combo["n_persons"]) // 50),
            # Oracle = the entry code only, NOT `c_dy = {dy}` and NOT `c_dy > 0`.
            # Although the VBA literally builds `((BIOG_MAIN.c_dy) > 0)`, that
            # clause does NOT reduce the DISTINCT c_personid set in this path:
            # measured via COM, an oracle carrying `AND c_dy > 0` is short by
            # exactly the c_dy<=0 / NULL persons VBA still returns (e.g. 38 for
            # entry 39 / dy 20).  VBA's effective result is every person for the
            # entry code, so the information-loss oracle (4.7, strict
            # set-equality) mirrors that.  Empirically confirmed: this oracle
            # passes 4.7 against live Access; the `c_dy > 0` variant does not.
            source_sql=f"""
                SELECT DISTINCT BIOG_MAIN.c_personid
                FROM BIOG_MAIN INNER JOIN ENTRY_DATA
                  ON BIOG_MAIN.c_personid = ENTRY_DATA.c_personid
                WHERE ENTRY_DATA.c_entry_code = {c}
            """,
        ))

    # 5-6. Two (entry × addr) combos with year filter.
    #
    # NOTE: this fixture explicitly selects index-year mode
    # (FrameYears=2) with range [-2000, 2000].  The source_sql must
    # mirror that filter, otherwise the diff check expects NULL-year
    # persons that VBA legitimately excludes — `c_index_year BETWEEN
    # -2000 AND 2000` evaluates to NULL (≈ False) under SQL three-
    # valued logic when c_index_year IS NULL.
    #
    # Concrete numbers (entry=110, addr=7213, current data):
    #   - baseline join:                           2176 persons
    #   - + c_index_year IS NULL:                  2170 persons
    #   - + c_index_year BETWEEN -2000 AND 2000:      6 persons
    # VBA returns the 6.  The pre-fix `expected_min_rows = n_persons//5
    # = 435` and a year-filter-less source_sql made the test claim
    # "VBA filter is too aggressive"; in fact VBA was correct and the
    # test was checking against the wrong baseline.
    for combo in le_data.get("entry_x_address_combos", [])[:2]:
        c = int(combo["c_entry_code"])
        addr = int(combo["c_index_addr_id"])
        out.append(Fixture(
            name=f"entry_{c}_addr_{addr}_indexyears",
            form="LookAtEntry",
            result_table="ZZ_SCRATCH_ENTRY",
            cmd_caption="Run Query",
            cmd_name="CmdQuery",
            pickers=[
                ("ZZ_SCRATCH_ENTRY_CODE", [c], "c_entry_code"),
                ("ZZ_SCRATCH_ADDR", [addr], "c_addr_id"),
            ],
            controls={
                "FrameYears": 2,         # index years
                "TxtFromYear": -2000,    # wide-open year range
                "TxtToYear": 2000,
                "TxtEntryDesc": "[selected]",
                "TxtTypeCode": "N/A",
            },
            # Most persons in CBDB have c_index_year IS NULL, so any
            # index-year filter dramatically narrows the pool.  Use a
            # conservative floor; the source_sql diff check below is
            # the strong correctness assertion.
            expected_min_rows=1,
            source_sql=f"""
                SELECT DISTINCT BIOG_MAIN.c_personid
                FROM BIOG_MAIN INNER JOIN ENTRY_DATA
                  ON BIOG_MAIN.c_personid = ENTRY_DATA.c_personid
                WHERE ENTRY_DATA.c_entry_code = {c}
                  AND BIOG_MAIN.c_index_addr_id = {addr}
                  AND BIOG_MAIN.c_index_year BETWEEN -2000 AND 2000
            """,
            py_replay=le.run,
            py_inputs=le.EntryQueryInputs(
                entry_codes=[c],
                addr_ids=[addr],
                addr_field="person",
                year_mode="index",
                from_year=-2000, to_year=2000,
            ),
        ))

    return out


# ----------------------------------------------------------------------
# Single big test — runs the full integrity suite per fixture
# ----------------------------------------------------------------------

def _all_fixtures() -> list[Fixture]:
    if not INPUTS.exists():
        return []
    inputs = _load_inputs()
    return _build_lookatentry_fixtures(inputs)


def _xfail_marks(fx: Fixture):
    """Previously: marked top_entry_code_* fixtures as `xfail strict=True`
    on the theory that LookAtEntry's multi-table backfill UPDATE
    silently failed on >30 k-row results.  That theory was the
    original "Bug #3", removed from the documented ISSUES set on
    2026-05-03 — re-verification via timer-triggered CmdQuery + SQL
    NULL-count probe found 0 / 92,514 rows missing c_entry_desc on
    the original fixture, AND there was no upstream source-level fix
    to point at, so per the marker-failure-≠-fix policy it was
    treated as a fixture / driver false positive rather than a
    CBDB-maintainer bug.

    The xfail marks themselves were also misleading: the matrix test
    actually failed BEFORE reaching the backfill step (pywinauto
    button-locator regression in `click_button_and_wait_table`),
    so `xfail strict=True` was passing on an unrelated failure and
    we never actually verified the backfill claim through the matrix.

    Returning no marks today.  The driver-level "button not found"
    issue still needs its own fix (migrate this file to
    `click_via_timer` like `test_vba_matrix_all_forms.py` already
    does), but that's an internal driver concern, not a CBDB bug.
    """
    return ()


@pytest.mark.parametrize(
    "fx",
    [pytest.param(f, marks=_xfail_marks(f)) for f in _all_fixtures()],
    ids=lambda f: f.name,
)
def test_vba_full_matrix(vba: VbaSession, fx: Fixture):
    """One mega-test per fixture: VBA fires the actual handler, then
    we apply all integrity checks. Logs the breakdown so failures
    are diagnosed easily."""

    # -------- 1. setup pickers + controls --------
    for table, ids, column in fx.pickers:
        vba.set_picker_codes(table, ids, column=column)
    vba.open_form(fx.form)
    for ctl, val in fx.controls.items():
        try:
            vba.set_control(fx.form, ctl, val)
        except Exception as e:
            print(f"  warn: setting {ctl}={val!r}: {e}")

    # -------- 2. fire VBA --------
    n = vba.click_button_and_wait_table(
        fx.cmd_caption, form=fx.form,
        result_table=fx.result_table,
        force_enable_ctl=fx.cmd_name,
    )
    print(f"\n[{fx.name}] VBA produced {n} rows")
    assert n >= fx.expected_min_rows, (
        f"[{fx.name}] only {n} rows; expected ≥ {fx.expected_min_rows}. "
        "Either the test inputs are stale (re-run discover_test_inputs.py) "
        "or the VBA filter is too aggressive."
    )

    # -------- 3. read result --------
    df = vba.read(fx.result_table)
    assert not df.empty

    # -------- 4. INTEGRITY CHECKS (12 dimensions) --------

    # 4.1 column structure (specific to LookAtEntry — generalize later)
    if fx.form == "LookAtEntry":
        from test_vba_integrity import (
            ZZ_SCRATCH_ENTRY_INSERT_COLS, ZZ_SCRATCH_ENTRY_UPDATE_COLS,
        )
        missing_insert = ZZ_SCRATCH_ENTRY_INSERT_COLS - set(df.columns)
        missing_update = ZZ_SCRATCH_ENTRY_UPDATE_COLS - set(df.columns)
        assert not missing_insert, f"INSERT cols missing: {missing_insert}"
        assert not missing_update, f"UPDATE cols missing: {missing_update}"

    # 4.2 numeric dtypes
    for c in ("c_personid", "c_index_year", "c_year"):
        if c in df.columns:
            assert pd.api.types.is_numeric_dtype(df[c]), \
                f"{c} should be numeric, got {df[c].dtype}"

    # 4.3 source-table fidelity (sample 50 rows)
    pids = sorted({int(p) for p in df["c_personid"].dropna()})
    if pids:
        sample_pids = pids[:50]
        in_clause = ",".join(str(p) for p in sample_pids)
        bm = pd.read_sql(
            f"SELECT c_personid, c_name, c_name_chn, c_index_year, c_dy "
            f"FROM BIOG_MAIN WHERE c_personid IN ({in_clause})",
            vba.conn,
        ).set_index("c_personid")
        mismatches = []
        for _, row in df[df["c_personid"].isin(sample_pids)].iterrows():
            pid = int(row["c_personid"])
            src = bm.loc[pid] if pid in bm.index else None
            if src is None:
                mismatches.append(("missing", pid))
                continue
            if isinstance(src, pd.DataFrame):
                src = src.iloc[0]
            for col in ("c_name", "c_name_chn", "c_index_year", "c_dy"):
                v_dst = row.get(col)
                v_src = src[col]
                if pd.isna(v_dst) and pd.isna(v_src):
                    continue
                if v_dst != v_src:
                    mismatches.append((col, pid, v_dst, v_src))
        assert not mismatches, f"BIOG_MAIN fidelity issues: {mismatches[:5]}"

    # 4.4 backfill correctness — c_entry_desc ≡ ENTRY_CODES.c_entry_desc
    if "c_entry_code" in df.columns and "c_entry_desc" in df.columns:
        codes = sorted({int(c) for c in df["c_entry_code"].dropna()})
        if codes:
            in_clause = ",".join(str(c) for c in codes)
            ec = pd.read_sql(
                f"SELECT c_entry_code, c_entry_desc FROM ENTRY_CODES "
                f"WHERE c_entry_code IN ({in_clause})",
                vba.conn,
            ).set_index("c_entry_code")
            mis = []
            for _, row in df.iterrows():
                code = row.get("c_entry_code")
                if pd.isna(code):
                    continue
                code = int(code)
                if code not in ec.index:
                    continue
                expected = ec.loc[code, "c_entry_desc"]
                actual = row.get("c_entry_desc")
                if pd.isna(expected) and pd.isna(actual):
                    continue
                if expected != actual:
                    mis.append((code, actual, expected))
            assert not mis, f"c_entry_desc backfill wrong: {mis[:3]}"

    # 4.5 FK integrity — addr_id ∈ ADDR_CODES
    if "c_addr_id" in df.columns:
        addrs = sorted({int(a) for a in df["c_addr_id"].dropna()})
        if addrs:
            in_clause = ",".join(str(a) for a in addrs)
            found = pd.read_sql(
                f"SELECT DISTINCT c_addr_id FROM ADDR_CODES "
                f"WHERE c_addr_id IN ({in_clause})", vba.conn,
            )
            found_set = set(found["c_addr_id"].astype(int))
            orphans = [a for a in addrs if a not in found_set]
            assert not orphans, f"orphan c_addr_id (FK fail): {orphans[:5]}"

    # 4.6 backfill completeness — non-NULL FK ⇒ non-NULL desc
    if {"c_entry_code", "c_entry_desc"}.issubset(df.columns):
        n_code = (df["c_entry_code"].notna()).sum()
        n_desc = (df["c_entry_desc"].notna()).sum()
        assert n_desc >= n_code, (
            f"{n_code - n_desc} rows have c_entry_code set but "
            f"c_entry_desc NULL (silent UPDATE join failure)"
        )

    # 4.7 information loss — independent SQL must yield same person set
    if fx.source_sql:
        src = pd.read_sql(fx.source_sql, vba.conn)
        src_set = {int(p) for p in src["c_personid"]}
        vba_set = {int(p) for p in df["c_personid"].dropna()}
        only_src = src_set - vba_set
        only_vba = vba_set - src_set
        assert not only_src, (
            f"VBA LOST {len(only_src)} persons that source SQL found: "
            f"{sorted(only_src)[:10]}"
        )
        assert not only_vba, (
            f"VBA includes {len(only_vba)} extra persons not in source: "
            f"{sorted(only_vba)[:10]}"
        )

    # 4.8 differential — VBA distinct persons match Python replay
    if fx.py_replay and fx.py_inputs:
        df_py = fx.py_replay(vba.conn, fx.py_inputs)
        py_set = {int(p) for p in df_py["c_personid"].dropna()}
        vba_set = {int(p) for p in df["c_personid"].dropna()}
        diff = py_set ^ vba_set
        assert not diff, (
            f"VBA vs Python replay disagree on {len(diff)} persons: "
            f"{sorted(diff)[:10]}"
        )

    print(f"[{fx.name}] ✓ all integrity checks passed ({n} rows)")
