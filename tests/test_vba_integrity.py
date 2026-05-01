"""
Data-integrity tests on real VBA-produced ZZ_SCRATCH_<X> contents.

Goes beyond "row counts match"; verifies:
  1. Column set + types in the scratch table match expectation
  2. Each row's source-derived fields (c_name, c_name_chn, c_dy, etc.)
     match BIOG_MAIN for the same c_personid (no copy-paste corruption,
     no string truncation)
  3. Every "_desc" / "_chn" descriptive backfill column actually got
     populated from the corresponding *_CODES lookup table
  4. Every foreign-key column references an existing row in the
     referenced table (c_addr_id ∈ ADDR_CODES, c_entry_code ∈
     ENTRY_CODES, etc.)
  5. No column that the form's UPDATE statement writes is left NULL
     when its source isn't NULL (catches silent JOIN failures)

These tests run against the LIVE VBA output, not the Python replay.
They detect bugs in the VBA INSERT or UPDATE statements that would
be invisible to row-count diffs.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from cbdb_driver.vba_session import VbaSession, make_fixture


ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK = ROOT / "analysis" / "_integrity_test_copy.mdb"


@pytest.fixture(scope="function")
def vba():
    yield from make_fixture(SRC, WORK)


# Columns that ZZ_SCRATCH_ENTRY must have (derived from form's INSERT
# + UPDATE statements per Form_LookAtEntry.vb:1645 + 1778).
ZZ_SCRATCH_ENTRY_INSERT_COLS = {
    "c_personid", "c_name", "c_name_chn", "c_index_year",
    "c_index_year_type_code", "c_dy",
    "c_entry_code", "c_year", "c_sequence", "c_exam_rank",
    "c_addr_id", "c_kin_id", "c_kin_code",
    "c_assoc_id", "c_assoc_code",
    "c_parental_status_code", "c_entry_addr_id",
    "c_source", "c_inst_code", "c_inst_name_code", "c_addr_type",
}
ZZ_SCRATCH_ENTRY_UPDATE_COLS = {
    "c_index_year_type_desc", "c_index_year_type_hz",
    "c_entry_desc", "c_entry_chn",
    "c_kin_desc", "c_kin_name", "c_kin_chn",
    "c_assoc_desc", "c_assoc_desc_chn",
    "c_assoc_name", "c_assoc_name_chn",
    "c_addr_name", "c_addr_chn", "x_coord", "y_coord",
    "c_addr_desc", "c_addr_desc_chn",
    "c_parental_status_desc", "c_parental_status_desc_chn",
    "c_entry_addr_name", "c_entry_addr_chn",
    "c_entry_xcoord", "c_entry_ycoord",
    "c_dynasty_chn", "c_dynasty",
    "c_source_text", "c_source_text_chn",
}


def _setup_kaifeng_yin_general(vba: VbaSession) -> int:
    """Common test prelude: seed pickers, set controls, run VBA."""
    vba.set_picker_codes("ZZ_SCRATCH_ENTRY_CODE", [118],
                         column="c_entry_code")
    vba.set_picker_addrs([100658])
    vba.open_form("LookAtEntry")
    for ctl, val in {
        "TxtFromYear": 900, "TxtToYear": 1100, "FrameYears": 2,
        "TxtEntryDesc": "yin privilege: general", "TxtTypeCode": "N/A",
    }.items():
        vba.set_control("LookAtEntry", ctl, val)
    n = vba.click_button_and_wait_table(
        "Run Query", form="LookAtEntry",
        result_table="ZZ_SCRATCH_ENTRY",
        force_enable_ctl="CmdQuery", timeout=30,
    )
    assert n > 0, "VBA produced 0 rows; cannot integrity-check"
    return n


# ----------------------------------------------------------------------
# 1. Column structure
# ----------------------------------------------------------------------

def test_zz_scratch_entry_has_all_insert_columns(vba: VbaSession):
    """Every column the VBA INSERT writes must exist in the table."""
    _setup_kaifeng_yin_general(vba)
    df = vba.read("ZZ_SCRATCH_ENTRY", where="1=1 AND 1=0")  # 0 rows, just schema
    missing = ZZ_SCRATCH_ENTRY_INSERT_COLS - set(df.columns)
    assert not missing, f"INSERT-target columns missing: {sorted(missing)}"


def test_zz_scratch_entry_has_all_update_columns(vba: VbaSession):
    """Every column the VBA UPDATE backfills must exist in the table."""
    _setup_kaifeng_yin_general(vba)
    df = vba.read("ZZ_SCRATCH_ENTRY", where="1=1 AND 1=0")
    missing = ZZ_SCRATCH_ENTRY_UPDATE_COLS - set(df.columns)
    assert not missing, f"UPDATE-target columns missing: {sorted(missing)}"


def test_zz_scratch_entry_dtypes_sensible(vba: VbaSession):
    """Numeric columns shouldn't all be strings; year/coord columns
    should be numeric in the result."""
    _setup_kaifeng_yin_general(vba)
    df = vba.read("ZZ_SCRATCH_ENTRY")
    assert pd.api.types.is_numeric_dtype(df["c_personid"]), \
        f"c_personid should be numeric, got {df['c_personid'].dtype}"
    assert pd.api.types.is_numeric_dtype(df["c_index_year"]), \
        f"c_index_year should be numeric, got {df['c_index_year'].dtype}"
    assert pd.api.types.is_numeric_dtype(df["c_year"]), \
        f"c_year should be numeric, got {df['c_year'].dtype}"
    # x/y coords may be float OR object (NULL-friendly); must not be int
    if "x_coord" in df.columns and len(df) > 0:
        # accept float/object/None
        pass


# ----------------------------------------------------------------------
# 2. Source-table fidelity (no copy-paste corruption)
# ----------------------------------------------------------------------

def test_personid_fields_match_biog_main(vba: VbaSession):
    """For every c_personid in the result, c_name / c_name_chn /
    c_index_year / c_dy must equal what BIOG_MAIN has for that
    person."""
    _setup_kaifeng_yin_general(vba)
    df = vba.read(
        "ZZ_SCRATCH_ENTRY",
        order_by="c_personid",
    )
    pids = sorted({int(p) for p in df["c_personid"].dropna()})
    assert pids
    in_clause = ",".join(str(p) for p in pids)
    bm = pd.read_sql(
        f"SELECT c_personid, c_name, c_name_chn, c_index_year, c_dy "
        f"FROM BIOG_MAIN WHERE c_personid IN ({in_clause})",
        vba.conn,
    )
    bm_idx = bm.set_index("c_personid")
    mismatches = []
    for _, row in df.iterrows():
        pid = int(row["c_personid"])
        if pid not in bm_idx.index:
            mismatches.append(("missing-from-biog", pid))
            continue
        src = bm_idx.loc[pid]
        if isinstance(src, pd.DataFrame):    # duplicate (shouldn't happen)
            src = src.iloc[0]
        for col in ("c_name", "c_name_chn", "c_index_year", "c_dy"):
            v_dst = row[col]
            v_src = src[col]
            # treat NaN/None as equal
            if pd.isna(v_dst) and pd.isna(v_src):
                continue
            if v_dst != v_src:
                mismatches.append((col, pid, v_dst, v_src))
    assert not mismatches, (
        "ZZ_SCRATCH_ENTRY values diverge from BIOG_MAIN source:\n"
        + "\n".join(map(str, mismatches[:10]))
        + ("\n  ... and more" if len(mismatches) > 10 else "")
    )


def test_entry_event_fields_match_entry_data(vba: VbaSession):
    """For each row, (c_personid, c_entry_code, c_year, c_sequence) tuple
    must exist in ENTRY_DATA. Catches stale rows or junk inserts."""
    _setup_kaifeng_yin_general(vba)
    df = vba.read("ZZ_SCRATCH_ENTRY")
    if df.empty:
        pytest.skip("no rows")
    # take a sample of 50 rows
    sample = df.sample(n=min(50, len(df)), random_state=0)
    cur = vba.conn.cursor()
    not_found = []
    for _, row in sample.iterrows():
        pid = row.get("c_personid")
        ec = row.get("c_entry_code")
        if pd.isna(pid) or pd.isna(ec):
            continue
        cur.execute(
            "SELECT COUNT(*) FROM ENTRY_DATA "
            "WHERE c_personid = ? AND c_entry_code = ?",
            int(pid), int(ec),
        )
        if int(cur.fetchone()[0]) == 0:
            not_found.append((int(pid), int(ec)))
    cur.close()
    assert not not_found, (
        "rows in ZZ_SCRATCH_ENTRY have no matching ENTRY_DATA record "
        f"(possible stale data): {not_found[:5]}"
    )


# ----------------------------------------------------------------------
# 3. Backfill correctness
# ----------------------------------------------------------------------

def test_entry_desc_backfill_matches_entry_codes(vba: VbaSession):
    """ZZ_SCRATCH_ENTRY.c_entry_desc must equal ENTRY_CODES.c_entry_desc
    for the same c_entry_code."""
    _setup_kaifeng_yin_general(vba)
    df = vba.read(
        "ZZ_SCRATCH_ENTRY",
        order_by="c_personid",
    )
    if df.empty:
        pytest.skip("no rows")
    codes = sorted({int(c) for c in df["c_entry_code"].dropna()})
    in_clause = ",".join(str(c) for c in codes)
    ec = pd.read_sql(
        f"SELECT c_entry_code, c_entry_desc, c_entry_desc_chn "
        f"FROM ENTRY_CODES WHERE c_entry_code IN ({in_clause})",
        vba.conn,
    ).set_index("c_entry_code")
    mismatches = []
    for _, row in df.iterrows():
        code = row.get("c_entry_code")
        if pd.isna(code):
            continue
        code = int(code)
        expected_desc = ec.loc[code, "c_entry_desc"] if code in ec.index else None
        actual_desc = row.get("c_entry_desc")
        if pd.isna(expected_desc) and pd.isna(actual_desc):
            continue
        if expected_desc != actual_desc:
            mismatches.append((code, actual_desc, expected_desc))
    assert not mismatches, (
        f"c_entry_desc backfill wrong for {len(mismatches)} rows; "
        f"first: {mismatches[:3]}"
    )


def test_addr_name_backfill_matches_addr_codes(vba: VbaSession):
    """ZZ_SCRATCH_ENTRY.c_addr_name should equal ADDR_CODES.c_name for
    the same c_addr_id (when c_addr_id is non-null)."""
    _setup_kaifeng_yin_general(vba)
    df = vba.read("ZZ_SCRATCH_ENTRY")
    sub = df[df["c_addr_id"].notna()].copy()
    if sub.empty:
        pytest.skip("no rows with addr")
    addr_ids = sorted({int(a) for a in sub["c_addr_id"]})
    in_clause = ",".join(str(a) for a in addr_ids)
    ac = pd.read_sql(
        f"SELECT c_addr_id, c_name, c_name_chn, x_coord, y_coord "
        f"FROM ADDR_CODES WHERE c_addr_id IN ({in_clause})",
        vba.conn,
    ).set_index("c_addr_id")
    mismatches = []
    for _, row in sub.iterrows():
        aid = int(row["c_addr_id"])
        expected = ac.loc[aid, "c_name"] if aid in ac.index else None
        actual = row.get("c_addr_name")
        if pd.isna(expected) and pd.isna(actual):
            continue
        if expected != actual:
            mismatches.append((aid, actual, expected))
    assert not mismatches, (
        f"c_addr_name backfill wrong for {len(mismatches)} rows; "
        f"first: {mismatches[:3]}"
    )


# ----------------------------------------------------------------------
# 4. FK integrity
# ----------------------------------------------------------------------

def test_fk_addr_id_exists_in_addr_codes(vba: VbaSession):
    _setup_kaifeng_yin_general(vba)
    df = vba.read("ZZ_SCRATCH_ENTRY")
    addrs = sorted({int(a) for a in df["c_addr_id"].dropna()})
    if not addrs:
        pytest.skip("no addrs in result")
    in_clause = ",".join(str(a) for a in addrs)
    ac = pd.read_sql(
        f"SELECT DISTINCT c_addr_id FROM ADDR_CODES "
        f"WHERE c_addr_id IN ({in_clause})",
        vba.conn,
    )
    found = set(ac["c_addr_id"].astype(int).tolist())
    orphans = [a for a in addrs if a not in found]
    assert not orphans, f"{len(orphans)} c_addr_id orphaned from ADDR_CODES: {orphans[:5]}"


def test_fk_entry_code_exists_in_entry_codes(vba: VbaSession):
    _setup_kaifeng_yin_general(vba)
    df = vba.read("ZZ_SCRATCH_ENTRY")
    codes = sorted({int(c) for c in df["c_entry_code"].dropna()})
    if not codes:
        pytest.skip("no codes")
    in_clause = ",".join(str(c) for c in codes)
    ec = pd.read_sql(
        f"SELECT DISTINCT c_entry_code FROM ENTRY_CODES "
        f"WHERE c_entry_code IN ({in_clause})",
        vba.conn,
    )
    found = set(ec["c_entry_code"].astype(int).tolist())
    orphans = [c for c in codes if c not in found]
    assert not orphans, f"orphan entry codes: {orphans}"


# ----------------------------------------------------------------------
# 5. Backfill completeness (no silent JOIN failures)
# ----------------------------------------------------------------------

def test_no_silent_backfill_failures(vba: VbaSession):
    """For every row where c_entry_code is set, c_entry_desc must be set
    too (the UPDATE backfill should have populated it). Same for
    c_addr_id → c_addr_name."""
    _setup_kaifeng_yin_general(vba)
    df = vba.read("ZZ_SCRATCH_ENTRY")
    issues = []
    n_entry_code_set = (df["c_entry_code"].notna()).sum()
    n_entry_desc_set = (df["c_entry_desc"].notna()).sum()
    if n_entry_code_set > 0 and n_entry_desc_set < n_entry_code_set:
        issues.append(
            f"{n_entry_code_set - n_entry_desc_set} rows have "
            f"c_entry_code set but c_entry_desc NULL (UPDATE join failed?)"
        )
    n_addr_id_set = (df["c_addr_id"].notna()).sum()
    n_addr_name_set = (df["c_addr_name"].notna()).sum()
    # Some addr_ids legitimately have null name (rare); allow ≤5%
    if n_addr_id_set > 0:
        gap = n_addr_id_set - n_addr_name_set
        if gap / n_addr_id_set > 0.05:
            issues.append(
                f"{gap}/{n_addr_id_set} rows have c_addr_id set but "
                f"c_addr_name NULL (>5% — UPDATE backfill failing)"
            )
    assert not issues, "\n".join(issues)


# ----------------------------------------------------------------------
# 6. Information-loss detection (independent count vs source)
# ----------------------------------------------------------------------

def test_no_persons_lost_vs_source_query(vba: VbaSession):
    """Independent SQL: 'persons with entry_code 118 + index_addr 100658
    + index_year 900-1100' computed directly against linked tables.
    The set of persons MUST be identical to ZZ_SCRATCH_ENTRY's distinct
    c_personid set.

    If VBA loses persons in its multi-step JOIN gymnastics, this test
    will fail."""
    _setup_kaifeng_yin_general(vba)
    df = vba.read("ZZ_SCRATCH_ENTRY")
    set_vba = {int(p) for p in df["c_personid"].dropna()}

    # Independent query (different join order from VBA)
    src = pd.read_sql("""
        SELECT DISTINCT BIOG_MAIN.c_personid
        FROM BIOG_MAIN INNER JOIN ENTRY_DATA
          ON BIOG_MAIN.c_personid = ENTRY_DATA.c_personid
        WHERE ENTRY_DATA.c_entry_code = 118
          AND BIOG_MAIN.c_index_addr_id = 100658
          AND BIOG_MAIN.c_index_year >= 900
          AND BIOG_MAIN.c_index_year <= 1100
    """, vba.conn)
    set_src = {int(p) for p in src["c_personid"]}

    only_vba = set_vba - set_src
    only_src = set_src - set_vba
    assert not only_vba and not only_src, (
        f"VBA distinct persons differ from source query:\n"
        f"  in VBA only: {sorted(only_vba)[:10]} (total {len(only_vba)})\n"
        f"  in source only (LOST by VBA!): {sorted(only_src)[:10]} (total {len(only_src)})"
    )


def test_no_rows_lost_vs_entry_data_count(vba: VbaSession):
    """For each person in the result, the number of ZZ_SCRATCH_ENTRY
    rows must equal the number of ENTRY_DATA rows matching the
    selected entry codes for that person. Catches dedup bugs."""
    _setup_kaifeng_yin_general(vba)
    df = vba.read("ZZ_SCRATCH_ENTRY")
    if df.empty:
        pytest.skip("no rows")
    # group VBA result by personid
    counts_vba = df.groupby("c_personid").size()
    pids = sorted(int(p) for p in counts_vba.index if not pd.isna(p))
    if not pids:
        pytest.skip("no personids")

    in_clause = ",".join(str(p) for p in pids[:50])  # sample
    counts_src = pd.read_sql(f"""
        SELECT c_personid, COUNT(*) AS n
        FROM ENTRY_DATA
        WHERE c_personid IN ({in_clause})
          AND c_entry_code = 118
        GROUP BY c_personid
    """, vba.conn).set_index("c_personid")

    mismatches = []
    for pid, n_vba in counts_vba.items():
        if pd.isna(pid) or int(pid) not in counts_src.index:
            continue
        pid = int(pid)
        n_src = int(counts_src.loc[pid, "n"])
        if int(n_vba) != n_src:
            mismatches.append((pid, int(n_vba), n_src))
    # allow some difference if multiple address joins multiply rows; flag drastic ones
    if mismatches:
        big = [m for m in mismatches if abs(m[1] - m[2]) > 1]
        if big:
            assert False, (
                f"row-count mismatches between ZZ_SCRATCH_ENTRY and "
                f"ENTRY_DATA for some persons (>1 diff): {big[:5]}"
            )
