"""Fast-suite tests verifying that TablesFields and ForeignKeys
documentation tables inside the DATA mdb accurately reflect the actual
database structure.

No VBA or COM required.  Runs with the standard fast suite:
    python -m pytest tests/test_schema_data_mdb.py -v

The DATA mdb is resolved dynamically (newest CBDB_*_DATA.mdb in data/,
gitignored) — never a hardcoded build date.  On headless/non-Windows
(no pyodbc) the whole module is skipped; on a box that CAN run it but has
no DATA mdb, the tests FAIL (not skip) so the schema/FK coverage cannot
silently vanish.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

# pyodbc is Windows/Access-only; skip the whole module on headless/Linux
# so the fast suite still collects cleanly there (mirrors conftest's rule).
pyodbc = pytest.importorskip("pyodbc")

REPO = Path(__file__).resolve().parent.parent
SCHEMA_DIFF_JSON = REPO / "reports" / "schema_diff.json"

# Resolve the DATA mdb dynamically (newest CBDB_*_DATA.mdb in data/) via the
# shared finder — NEVER hardcode a build date.  A pinned CBDB_20260430_DATA.mdb
# previously made every test pytest.skip() on any other build, silently dropping
# the whole schema/FK module (coverage-floor regression).  None => no DATA mdb
# found; the fixture turns that into a FAIL, not a skip.
_ANALYSIS_DIR = str(REPO / "analysis")
if _ANALYSIS_DIR not in sys.path:
    sys.path.insert(0, _ANALYSIS_DIR)
# Let an ImportError propagate (a corrupt checkout should fail loudly, not
# masquerade as "no DATA mdb").  Only "no file found" maps to None.
from _data_mdb_finder import find_data_mdb  # noqa: E402
try:
    DATA_MDB = find_data_mdb(REPO)
except FileNotFoundError:
    DATA_MDB = None


# ---------------------------------------------------------------------------
# Module-scoped fixture
# ---------------------------------------------------------------------------

@pytest.fixture(scope="module")
def data_mdb_conn():
    if DATA_MDB is None:
        pytest.fail(
            "No CBDB_*_DATA.mdb found in data/ — the TablesFields/ForeignKeys "
            "schema validation cannot run.  Place the DATA mdb in data/ (it is "
            "gitignored).  This is a FAILURE, not a skip: silently skipping "
            "would drop the whole schema/FK module from the standardized run "
            "(the coverage-floor regression this guards against)."
        )
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

    If cursor.columns() itself raises UnicodeDecodeError (e.g. the table's
    own ODBC column metadata is corrupt), falls back to SELECT TOP 1 to
    enumerate column names from the live data.  DataFormat / nullability
    are not needed here -- we only need the name set for existence checks.
    """
    fresh = _make_fresh_conn()
    try:
        cur = fresh.cursor()
        names: set[str] = set()
        partial_failure = False
        try:
            cur.columns(table=table_name)
            while True:
                try:
                    row = cur.fetchone()
                except UnicodeDecodeError:
                    # Partial read -- fall through to SELECT fallback
                    partial_failure = True
                    break
                if row is None:
                    break
                names.add(row.column_name.upper())
        except UnicodeDecodeError:
            partial_failure = True

        if partial_failure or not names:
            # SELECT TOP 1 fallback: enumerates columns without catalog
            try:
                fb = fresh.cursor()
                fb.execute(f"SELECT TOP 1 * FROM [{table_name}]")
                fb.fetchall()
                names = {d[0].upper() for d in fb.description}
            except Exception:
                pass  # return whatever partial names we already have

        return names
    finally:
        fresh.close()


# ---------------------------------------------------------------------------
# Test 1 â€" all tables documented in TablesFields actually exist
# ---------------------------------------------------------------------------

@pytest.mark.xfail(
    reason=(
        "TablesFields is known to document table(s) that don't exist in the "
        "DATA mdb (build-dependent; see reports/schema_diff.json "
        "tables_fields.only_in_current for the authoritative per-build list). "
        "xfail documents the doc drift; promote to strict once docs are updated."
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
# Test 2 â€" all columns documented in TablesFields actually exist
# ---------------------------------------------------------------------------

@pytest.mark.xfail(
    reason=(
        "TablesFields is known to document columns that don't exist in the "
        "DATA mdb (build-dependent count + list; see reports/schema_diff.json "
        "tables_fields.only_in_current for the authoritative per-build list). "
        "xfail documents the doc drift; promote to strict once docs are updated."
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
# Test 3 â€" FK referenced tables exist
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
# Test 4 â€" FK referenced columns exist
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
# Test 5 â€" schema_diff.json is clean (no stale / missing entries)
# ---------------------------------------------------------------------------

@pytest.mark.xfail(
    reason=(
        "TablesFields is known to be out of sync with the actual DB schema "
        "(build-dependent counts: stale rows / undocumented columns / "
        "DataFormat-NULL mismatches — see reports/schema_diff.json for the "
        "authoritative per-build numbers).  This xfail documents the finding; "
        "promote to a strict assertion once the docs are updated."
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
