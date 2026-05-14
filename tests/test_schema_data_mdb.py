"""Fast-suite tests verifying that TablesFields and ForeignKeys
documentation tables inside CBDB_20260430_DATA.mdb accurately reflect
the actual database structure.

No VBA or COM required.  Runs with the standard fast suite:
    python -m pytest tests/test_schema_data_mdb.py -v

Requires data/CBDB_20260430_DATA.mdb to be present (gitignored).
If the file is absent every test is skipped automatically.
"""
from __future__ import annotations

import json
from pathlib import Path

import pytest
import pyodbc

REPO = Path(__file__).resolve().parent.parent
DATA_MDB = REPO / "data" / "CBDB_20260430_DATA.mdb"
SCHEMA_DIFF_JSON = REPO / "reports" / "schema_diff.json"


# ---------------------------------------------------------------------------
# Module-scoped fixture
# ---------------------------------------------------------------------------

@pytest.fixture(scope="module")
def data_mdb_conn():
    if not DATA_MDB.exists():
        pytest.skip("DATA mdb not present (gitignored)")
    conn = pyodbc.connect(
        f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={DATA_MDB};"
    )
    yield conn
    conn.close()


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_fresh_conn() -> pyodbc.Connection:
    """Open a fresh pyodbc connection to DATA_MDB.

    Used wherever catalog calls (tables/columns/primaryKeys) might leave
    the ODBC driver in a corrupted state if reused on the same connection
    that also serves live SQL queries.
    """
    cs = (
        "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={DATA_MDB};"
    )
    return pyodbc.connect(cs, autocommit=True)


def _existing_tables() -> set[str]:
    """Return the set of real (non-system) table names via ODBC catalog.
    Names are stored in upper-case for case-insensitive comparison
    (Access table names are case-insensitive).
    Uses a fresh connection to avoid polluting the shared fixture connection.
    """
    fresh = _make_fresh_conn()
    try:
        cur = fresh.cursor()
        cur.tables(tableType="TABLE")
        return {
            r.table_name.upper()
            for r in cur.fetchall()
            if not r.table_name.startswith("MSys")
            and not r.table_name.startswith("~")
        }
    finally:
        fresh.close()


def _catalog_table_names() -> dict[str, str]:
    """Return mapping of UPPER(table_name) â†’ actual_table_name from the
    ODBC catalog.  Used to resolve the correct case before calling
    cursor.columns(), which is case-sensitive in the Access ODBC driver
    and can return corrupt metadata for some capitalisation variants.
    Uses a fresh connection.
    """
    fresh = _make_fresh_conn()
    try:
        cur = fresh.cursor()
        cur.tables(tableType="TABLE")
        return {
            r.table_name.upper(): r.table_name
            for r in cur.fetchall()
            if not r.table_name.startswith("MSys")
            and not r.table_name.startswith("~")
        }
    finally:
        fresh.close()


def _columns_for_table(table_name: str) -> set[str]:
    """Return column names (upper-cased) for a single table.

    Always opens a FRESH connection to avoid ODBC internal-state
    corruption that occurs when cursor.columns() is called many times on
    the same shared connection (the Access ODBC driver can produce a
    UTF-16-LE decode error on subsequent catalog queries if any earlier
    call partially failed or left cursor state dirty).
    """
    fresh = _make_fresh_conn()
    try:
        cur = fresh.cursor()
        cur.columns(table=table_name)
        names: set[str] = set()
        while True:
            try:
                row = cur.fetchone()
            except UnicodeDecodeError:
                break
            if row is None:
                break
            names.add(row.column_name.upper())
        return names
    finally:
        fresh.close()


# ---------------------------------------------------------------------------
# Test 1 â€” all tables documented in TablesFields actually exist
# ---------------------------------------------------------------------------

@pytest.mark.xfail(
    reason=(
        "TablesFields documents 'PersonIDSource' which does not exist in the "
        "2026-04-30 dump. This xfail documents the stale documentation; "
        "promote to strict once the doc table is updated."
    ),
    strict=False,
)
def test_tables_fields_all_tables_exist(data_mdb_conn):
    """Every distinct AccessTblNm in TablesFields must exist as a real table.
    Comparison is case-insensitive (Access table names are case-insensitive).
    """
    cur = data_mdb_conn.cursor()
    cur.execute("SELECT DISTINCT AccessTblNm FROM TablesFields ORDER BY AccessTblNm")
    documented_tables = [r[0] for r in cur.fetchall()]

    existing = _existing_tables()  # upper-cased

    missing = [t for t in documented_tables if t.upper() not in existing]
    assert not missing, (
        f"{len(missing)} tables documented in TablesFields but not found "
        f"in the database: {missing}"
    )


# ---------------------------------------------------------------------------
# Test 2 â€” all columns documented in TablesFields actually exist
# ---------------------------------------------------------------------------

@pytest.mark.xfail(
    reason=(
        "TablesFields documents columns that do not exist in the 2026-04-30 dump "
        "(ADMIN_CAT_CODE_TYPE_REL.c_admin_type_code, ADMIN_CAT_TYPES 3 cols, "
        "ENTRY_DATA.c_addr_id, ENTRY_DATA.c_posting_id, "
        "MERGED_PERSON_DATA.c_merged_to_personid, TMP_ADDR_C.Max_c_belongs_first_year "
        "and others). These are stale documentation entries; see schema_diff.json "
        "only_in_current for the full list. Promote to strict once docs are updated."
    ),
    strict=False,
)
def test_tables_fields_all_columns_exist(data_mdb_conn):
    """Every (AccessTblNm, AccessFldNm) in TablesFields must exist as
    a real column in the corresponding table.
    Comparison is case-insensitive (Access is case-insensitive).
    """
    cur = data_mdb_conn.cursor()
    cur.execute(
        "SELECT AccessTblNm, AccessFldNm FROM TablesFields ORDER BY AccessTblNm, AccessFldNm"
    )
    rows = cur.fetchall()

    existing_tables = _existing_tables()  # upper-cased
    catalog_names = _catalog_table_names()  # upperâ†’actual_case
    # col_cache key is upper-cased table name
    col_cache: dict[str, set[str]] = {}
    missing: list[tuple[str, str]] = []

    for tbl, col in rows:
        tbl_up = tbl.upper()
        if tbl_up not in existing_tables:
            # Already caught by test 1; skip to avoid noise
            continue
        if tbl_up not in col_cache:
            # Use actual catalog case to avoid ODBC UTF-16 decode errors
            actual_name = catalog_names.get(tbl_up, tbl)
            col_cache[tbl_up] = _columns_for_table(actual_name)
        if col.upper() not in col_cache[tbl_up]:
            missing.append((tbl, col))

    assert not missing, (
        f"{len(missing)} columns documented in TablesFields but not found "
        f"in the actual table: {missing[:20]}"
        + (" ... (truncated)" if len(missing) > 20 else "")
    )


# ---------------------------------------------------------------------------
# Test 3 â€” FK referenced tables exist
# ---------------------------------------------------------------------------

def test_foreign_keys_referenced_tables_exist(data_mdb_conn):
    """For every row in ForeignKeys where ForeignKey IS NOT NULL, the
    referenced table must exist in the database.
    Comparison is case-insensitive.
    """
    cur = data_mdb_conn.cursor()
    cur.execute(
        "SELECT AccessTblNm, AccessFldNm, ForeignKey "
        "FROM ForeignKeys "
        "WHERE ForeignKey IS NOT NULL "
        "ORDER BY ForeignKey"
    )
    rows = cur.fetchall()

    existing = _existing_tables()  # upper-cased
    missing: list[dict] = []
    for src_tbl, src_col, ref_tbl in rows:
        if ref_tbl.upper() not in existing:
            missing.append({
                "AccessTblNm": src_tbl,
                "AccessFldNm": src_col,
                "ForeignKey": ref_tbl,
            })

    assert not missing, (
        f"{len(missing)} FK-referenced tables in ForeignKeys not found "
        f"in the database: {missing[:20]}"
        + (" ... (truncated)" if len(missing) > 20 else "")
    )


# ---------------------------------------------------------------------------
# Test 4 â€” FK referenced columns exist
# ---------------------------------------------------------------------------

def test_foreign_keys_referenced_columns_exist(data_mdb_conn):
    """For every row in ForeignKeys where both ForeignKey and
    ForeignKeyBaseField are not NULL, the referenced column must exist
    in the referenced table.
    Comparison is case-insensitive.
    """
    cur = data_mdb_conn.cursor()
    cur.execute(
        "SELECT AccessTblNm, AccessFldNm, ForeignKey, ForeignKeyBaseField "
        "FROM ForeignKeys "
        "WHERE ForeignKey IS NOT NULL AND ForeignKeyBaseField IS NOT NULL "
        "ORDER BY ForeignKey, ForeignKeyBaseField"
    )
    rows = cur.fetchall()

    existing_tables = _existing_tables()  # upper-cased
    catalog_names = _catalog_table_names()  # upperâ†’actual_case
    col_cache: dict[str, set[str]] = {}
    missing: list[dict] = []

    for src_tbl, src_col, ref_tbl, ref_col in rows:
        ref_tbl_up = ref_tbl.upper()
        if ref_tbl_up not in existing_tables:
            continue  # Caught by test 3
        if ref_tbl_up not in col_cache:
            # Use actual catalog case to avoid ODBC UTF-16 decode errors
            actual_name = catalog_names.get(ref_tbl_up, ref_tbl)
            col_cache[ref_tbl_up] = _columns_for_table(actual_name)
        if ref_col.upper() not in col_cache[ref_tbl_up]:
            missing.append({
                "AccessTblNm": src_tbl,
                "AccessFldNm": src_col,
                "ForeignKey": ref_tbl,
                "ForeignKeyBaseField": ref_col,
            })

    assert not missing, (
        f"{len(missing)} FK-referenced columns in ForeignKeys not found "
        f"in the referenced table: {missing[:20]}"
        + (" ... (truncated)" if len(missing) > 20 else "")
    )


# ---------------------------------------------------------------------------
# Test 5 â€” schema_diff.json is clean (no stale / missing entries)
# ---------------------------------------------------------------------------

@pytest.mark.xfail(
    reason=(
        "TablesFields is known to be out of sync with the actual DB schema "
        "in the 2026-04-30 dump (22 stale rows, 131 undocumented columns, "
        "337 DataFormat/NULL_allowed mismatches â€” see schema_diff.json). "
        "This xfail documents the finding; promote to a strict assertion "
        "once the docs are updated."
    ),
    strict=False,
)
def test_schema_diff_is_clean(data_mdb_conn):
    """If reports/schema_diff.json exists, assert that:
    - tables_fields.only_in_current is empty (no stale doc rows)
    - tables_fields.only_in_regen is empty (no undocumented columns)
    - foreign_keys.only_in_current and only_in_regen are empty
      (only if FK introspection was available)
    """
    if not SCHEMA_DIFF_JSON.exists():
        pytest.skip("run reports/collect_schema_diffs.py first")

    diff = json.loads(SCHEMA_DIFF_JSON.read_text(encoding="utf-8"))

    tf = diff["tables_fields"]
    fk = diff["foreign_keys"]

    problems: list[str] = []
    if tf["only_in_current"]:
        problems.append(
            f"tables_fields.only_in_current has {len(tf['only_in_current'])} "
            f"rows (TablesFields documents columns that no longer exist): "
            f"{tf['only_in_current'][:5]}"
        )
    if tf["only_in_regen"]:
        problems.append(
            f"tables_fields.only_in_regen has {len(tf['only_in_regen'])} "
            f"rows (actual DB columns not documented in TablesFields): "
            f"{tf['only_in_regen'][:5]}"
        )
    if fk.get("fk_introspection_available"):
        if fk["only_in_current"]:
            problems.append(
                f"foreign_keys.only_in_current has {len(fk['only_in_current'])} "
                f"rows: {fk['only_in_current'][:5]}"
            )
        if fk["only_in_regen"]:
            problems.append(
                f"foreign_keys.only_in_regen has {len(fk['only_in_regen'])} "
                f"rows: {fk['only_in_regen'][:5]}"
            )

    assert not problems, "\n".join(problems)
