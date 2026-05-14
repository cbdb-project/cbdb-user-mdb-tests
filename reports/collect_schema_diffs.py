"""Collect diffs between the documentation tables (TablesFields,
ForeignKeys) inside CBDB_20260430_DATA.mdb and the actual database
structure reconstructed from ODBC catalog calls.

Outputs (all in reports/):
  tables_fields_current.csv   — live dump of TablesFields
  foreign_keys_current.csv    — live dump of ForeignKeys
  tables_fields_regen.csv     — reconstructed from ODBC catalog
  foreign_keys_regen.csv      — reconstructed from ODBC foreignKeys()
  schema_diff.json            — diff summary consumed by generate_report.py

Run:
    python reports/collect_schema_diffs.py
"""
from __future__ import annotations

import csv
import json
from datetime import datetime, timezone
from pathlib import Path

import pyodbc

ROOT = Path(__file__).resolve().parent.parent
DATA_MDB = ROOT / "data" / "CBDB_20260430_DATA.mdb"
REPORTS = ROOT / "reports"

OUT_TF_CURRENT = REPORTS / "tables_fields_current.csv"
OUT_FK_CURRENT = REPORTS / "foreign_keys_current.csv"
OUT_TF_REGEN   = REPORTS / "tables_fields_regen.csv"
OUT_FK_REGEN   = REPORTS / "foreign_keys_regen.csv"
OUT_JSON       = REPORTS / "schema_diff.json"


def _open_data_mdb() -> pyodbc.Connection:
    cs = (
        "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={DATA_MDB};"
    )
    return pyodbc.connect(cs, autocommit=True)


# ---------------------------------------------------------------------------
# ODBC TYPE_NAME → Access DataFormat
# ---------------------------------------------------------------------------
_TYPE_MAP: dict[str, str] = {
    "LONGCHAR": "Memo",
    "MEMO":     "Memo",
    "VARCHAR":  "Text",
    "CHAR":     "Text",
    "INTEGER":  "Long",
    "LONG":     "Long",
    "COUNTER":  "Long",
    "SMALLINT": "Integer",
    "SHORT":    "Integer",
    "DOUBLE":   "Double",
    "FLOAT":    "Double",
    "SINGLE":   "Double",
    "DATETIME": "Date/Time",
    "DATE":     "Date/Time",
    "LOGICAL":  "Yes/No",
    "BIT":      "Yes/No",
    "BOOLEAN":  "Yes/No",
    "CURRENCY": "Currency",
    "DECIMAL":  "Decimal",
    "NUMERIC":  "Decimal",
    "BYTE":     "Byte",
}


def _map_type(odbc_type_name: str) -> str:
    return _TYPE_MAP.get(odbc_type_name.upper(), odbc_type_name)


# ---------------------------------------------------------------------------
# Step 1 — dump current table contents to CSV
# ---------------------------------------------------------------------------

def _dump_current(conn: pyodbc.Connection) -> tuple[list[dict], list[dict]]:
    cur = conn.cursor()

    # TablesFields
    cur.execute(
        "SELECT RowNum, DumpTblNm, DumpFldNm, AccessTblNm, AccessFldNm, "
        "IndexOnField, DataFormat, NULL_allowed, ForeignKey, ForeignKeyBaseField "
        "FROM TablesFields "
        "ORDER BY AccessTblNm, AccessFldNm"
    )
    tf_cols = [d[0] for d in cur.description]
    tf_rows = [dict(zip(tf_cols, row)) for row in cur.fetchall()]

    with OUT_TF_CURRENT.open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=tf_cols)
        writer.writeheader()
        writer.writerows(tf_rows)
    print(f"wrote {OUT_TF_CURRENT}  ({len(tf_rows)} rows)")

    # ForeignKeys
    cur.execute(
        "SELECT AccessTblNm, AccessFldNm, ForeignKey, ForeignKeyBaseField, "
        "FKString, FKName, skip, IndexOnField, DataFormat, NULL_allowed "
        "FROM ForeignKeys "
        "ORDER BY AccessTblNm, AccessFldNm"
    )
    fk_cols = [d[0] for d in cur.description]
    fk_rows = [dict(zip(fk_cols, row)) for row in cur.fetchall()]

    with OUT_FK_CURRENT.open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=fk_cols)
        writer.writeheader()
        writer.writerows(fk_rows)
    print(f"wrote {OUT_FK_CURRENT}  ({len(fk_rows)} rows)")

    return tf_rows, fk_rows


# ---------------------------------------------------------------------------
# Step 2 — reconstruct from ODBC catalog
# ---------------------------------------------------------------------------

def _regen_from_catalog(conn: pyodbc.Connection) -> tuple[list[dict], list[dict], bool]:
    # Use a fresh dedicated cursor just for the initial table enumeration so
    # we don't pollute any cursor state before the per-table column loops.
    enum_cur = conn.cursor()
    enum_cur.tables(tableType="TABLE")
    all_tables = [
        row.table_name
        for row in enum_cur.fetchall()
        if not row.table_name.startswith("MSys")
        and not row.table_name.startswith("~")
    ]
    all_tables.sort()
    print(f"  found {len(all_tables)} user tables via ODBC catalog")

    # Build upper-case → actual-case mapping so that TablesFields entries
    # (which may use a different capitalisation, e.g. KIN_MOURNING vs
    # KIN_Mourning) can be matched case-insensitively when we later diff.
    # We ALWAYS call cursor.columns() with the exact case from the catalog
    # because the Access ODBC driver is case-sensitive for catalog queries.
    upper_to_actual: dict[str, str] = {t.upper(): t for t in all_tables}

    tf_regen: list[dict] = []
    fk_regen: list[dict] = []
    fk_introspection_available = False

    for table_name in all_tables:
        # ---- Primary keys (fresh cursor per table) ----
        try:
            pk_cur = conn.cursor()
            pk_cur.primaryKeys(table=table_name)
            pk_rows = pk_cur.fetchall()
            pk_cols = {row[3] for row in pk_rows}  # COLUMN_NAME is index 3
        except Exception:
            pk_cols = set()

        # ---- Columns (fresh cursor per table + UnicodeDecodeError fallback)
        # The Access ODBC driver corrupts shared cursor state when
        # cursor.columns() is called sequentially on many tables.  Using a
        # fresh cursor per call eliminates that state pollution.  If
        # UnicodeDecodeError still fires (e.g. corrupted column metadata in
        # the .mdb itself), fall back to SELECT TOP 1 to enumerate names.
        col_rows: list | None = None
        try:
            col_cur = conn.cursor()
            col_cur.columns(table=table_name)
            col_rows = []
            while True:
                try:
                    row = col_cur.fetchone()
                except UnicodeDecodeError:
                    # Partial read — fall through to SELECT fallback
                    col_rows = None
                    break
                if row is None:
                    break
                col_rows.append(row)
        except UnicodeDecodeError:
            col_rows = None

        if col_rows is not None:
            # Happy path — full ODBC metadata available
            for col in col_rows:
                col_name  = col.column_name
                type_name = col.type_name if col.type_name else ""
                nullable  = col.nullable   # 1 = nullable, 0 = not nullable

                index_on = "Primary Key" if col_name in pk_cols else None
                data_fmt = _map_type(type_name)
                null_allowed = bool(nullable)

                tf_regen.append({
                    "AccessTblNm":       table_name,
                    "AccessFldNm":       col_name,
                    "IndexOnField":      index_on,
                    "DataFormat":        data_fmt,
                    "NULL_allowed":      null_allowed,
                    "ForeignKey":        None,
                    "ForeignKeyBaseField": None,
                })
        else:
            # Fallback — cursor.columns() failed; enumerate via SELECT TOP 1
            print(f"  [WARN] UnicodeDecodeError on cursor.columns({table_name!r})"
                  f" — falling back to SELECT TOP 1")
            try:
                fb_cur = conn.cursor()
                fb_cur.execute(f"SELECT TOP 1 * FROM [{table_name}]")
                fb_cur.fetchall()
                col_names = [d[0] for d in fb_cur.description]
                for col_name in col_names:
                    index_on = "Primary Key" if col_name in pk_cols else None
                    tf_regen.append({
                        "AccessTblNm":       table_name,
                        "AccessFldNm":       col_name,
                        "IndexOnField":      index_on,
                        # DataFormat and NULL_allowed unknown via this path
                        "DataFormat":        None,
                        "NULL_allowed":      None,
                        "ForeignKey":        None,
                        "ForeignKeyBaseField": None,
                    })
            except Exception as exc:
                print(f"  [ERROR] SELECT TOP 1 fallback also failed for "
                      f"{table_name!r}: {exc}")

        # ---- Foreign keys (fresh cursor per table) ----
        try:
            fk_cur = conn.cursor()
            fk_cur.foreignKeys(table=table_name)
            fk_rows_odbc = fk_cur.fetchall()
        except Exception:
            fk_rows_odbc = []

        for fk_row in fk_rows_odbc:
            # Standard ODBC columns: 0=PKTABLE_CAT, 1=PKTABLE_SCHEM,
            # 2=PKTABLE_NAME, 3=PKCOLUMN_NAME, 4=FKTABLE_CAT,
            # 5=FKTABLE_SCHEM, 6=FKTABLE_NAME, 7=FKCOLUMN_NAME, ...
            try:
                pktable  = fk_row[2]
                pkcol    = fk_row[3]
                fktable  = fk_row[6]
                fkcol    = fk_row[7]
            except (IndexError, TypeError):
                continue
            if pktable is None:
                continue
            fk_introspection_available = True
            fk_regen.append({
                "AccessTblNm":       fktable,
                "AccessFldNm":       fkcol,
                "ForeignKey":        pktable,
                "ForeignKeyBaseField": pkcol,
                "FKString":          None,
                "FKName":            None,
                "skip":              None,
            })

    # Update FK entries in tf_regen with FK information
    fk_lookup: dict[tuple[str, str], dict] = {}
    for fk in fk_regen:
        key = (fk["AccessTblNm"], fk["AccessFldNm"])
        fk_lookup[key] = fk

    for row in tf_regen:
        key = (row["AccessTblNm"], row["AccessFldNm"])
        if key in fk_lookup:
            row["ForeignKey"] = fk_lookup[key]["ForeignKey"]
            row["ForeignKeyBaseField"] = fk_lookup[key]["ForeignKeyBaseField"]

    # Write tables_fields_regen.csv
    tf_regen_cols = [
        "AccessTblNm", "AccessFldNm", "IndexOnField",
        "DataFormat", "NULL_allowed", "ForeignKey", "ForeignKeyBaseField",
    ]
    with OUT_TF_REGEN.open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=tf_regen_cols)
        writer.writeheader()
        writer.writerows(tf_regen)
    print(f"wrote {OUT_TF_REGEN}  ({len(tf_regen)} rows)")

    # Write foreign_keys_regen.csv
    fk_regen_cols = [
        "AccessTblNm", "AccessFldNm", "ForeignKey", "ForeignKeyBaseField",
        "FKString", "FKName", "skip",
    ]
    with OUT_FK_REGEN.open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=fk_regen_cols)
        writer.writeheader()
        if fk_introspection_available:
            writer.writerows(fk_regen)
    print(f"wrote {OUT_FK_REGEN}  "
          f"({'%d rows' % len(fk_regen) if fk_introspection_available else 'empty — FK introspection unavailable'})")

    return tf_regen, fk_regen, fk_introspection_available


# ---------------------------------------------------------------------------
# Step 3 — compute diffs and write schema_diff.json
# ---------------------------------------------------------------------------

def _compute_diffs(
    tf_current: list[dict],
    tf_regen: list[dict],
    fk_current: list[dict],
    fk_regen: list[dict],
    fk_introspection_available: bool,
) -> dict:
    # ---- TablesFields diff ----
    # Key: (AccessTblNm.upper(), AccessFldNm.upper()) — case-insensitive, since
    # Access table/column names are case-insensitive and the documentation table
    # may use a different capitalisation than the ODBC catalog returns.
    # We store the ORIGINAL names from each source for display.
    cur_map: dict[tuple, dict] = {}
    for row in tf_current:
        key = (row["AccessTblNm"].upper(), row["AccessFldNm"].upper())
        cur_map[key] = row

    reg_map: dict[tuple, dict] = {}
    for row in tf_regen:
        key = (row["AccessTblNm"].upper(), row["AccessFldNm"].upper())
        reg_map[key] = row

    cur_keys = set(cur_map)
    reg_keys = set(reg_map)

    tf_only_in_current = [
        {
            "AccessTblNm": cur_map[k]["AccessTblNm"],
            "AccessFldNm": cur_map[k]["AccessFldNm"],
        }
        for k in sorted(cur_keys - reg_keys)
    ]
    tf_only_in_regen = [
        {
            "AccessTblNm": reg_map[k]["AccessTblNm"],
            "AccessFldNm": reg_map[k]["AccessFldNm"],
            "DataFormat":  reg_map[k]["DataFormat"],
            "NULL_allowed": reg_map[k]["NULL_allowed"],
        }
        for k in sorted(reg_keys - cur_keys)
    ]
    tf_mismatches: list[dict] = []
    for k in sorted(cur_keys & reg_keys):
        c = cur_map[k]
        r = reg_map[k]
        for field in ("DataFormat", "NULL_allowed"):
            cv = c.get(field)
            rv = r.get(field)
            # Normalise NULL_allowed: the DB may store True/False or 1/0
            if field == "NULL_allowed":
                cv = bool(cv) if cv is not None else None
                rv = bool(rv) if rv is not None else None
            if cv != rv:
                tf_mismatches.append({
                    "AccessTblNm": c["AccessTblNm"],
                    "AccessFldNm": c["AccessFldNm"],
                    "field": field,
                    "current": cv,
                    "regen": rv,
                })

    # ---- ForeignKeys diff ----
    fk_only_in_current: list[dict] = []
    fk_only_in_regen: list[dict] = []
    fk_mismatches: list[dict] = []

    if fk_introspection_available:
        def _fk_key(row: dict) -> tuple:
            return (
                row["AccessTblNm"],
                row["AccessFldNm"],
                row["ForeignKey"],
                row["ForeignKeyBaseField"],
            )

        fk_cur_keys = set(_fk_key(r) for r in fk_current if r.get("ForeignKey"))
        fk_reg_keys = set(_fk_key(r) for r in fk_regen)

        fk_only_in_current = [
            {"AccessTblNm": k[0], "AccessFldNm": k[1],
             "ForeignKey": k[2], "ForeignKeyBaseField": k[3]}
            for k in sorted(fk_cur_keys - fk_reg_keys)
        ]
        fk_only_in_regen = [
            {"AccessTblNm": k[0], "AccessFldNm": k[1],
             "ForeignKey": k[2], "ForeignKeyBaseField": k[3]}
            for k in sorted(fk_reg_keys - fk_cur_keys)
        ]

    result = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "data_mdb": DATA_MDB.name,
        "tables_fields": {
            "total_current": len(tf_current),
            "total_regen": len(tf_regen),
            "only_in_current": tf_only_in_current,
            "only_in_regen": tf_only_in_regen,
            "mismatches": tf_mismatches,
        },
        "foreign_keys": {
            "fk_introspection_available": fk_introspection_available,
            "total_current": len(fk_current),
            "total_regen": len(fk_regen),
            "only_in_current": fk_only_in_current,
            "only_in_regen": fk_only_in_regen,
            "mismatches": fk_mismatches,
        },
    }
    OUT_JSON.write_text(
        json.dumps(result, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"wrote {OUT_JSON}")
    return result


def _print_summary(diff: dict) -> None:
    tf = diff["tables_fields"]
    fk = diff["foreign_keys"]
    print("\n=== Schema Diff Summary ===")
    print(f"TablesFields:")
    print(f"  total_current : {tf['total_current']}")
    print(f"  total_regen   : {tf['total_regen']}")
    print(f"  only_in_current : {len(tf['only_in_current'])}")
    print(f"  only_in_regen   : {len(tf['only_in_regen'])}")
    print(f"  mismatches      : {len(tf['mismatches'])}")
    print(f"ForeignKeys:")
    print(f"  fk_introspection_available : {fk['fk_introspection_available']}")
    print(f"  total_current : {fk['total_current']}")
    print(f"  total_regen   : {fk['total_regen']}")
    print(f"  only_in_current : {len(fk['only_in_current'])}")
    print(f"  only_in_regen   : {len(fk['only_in_regen'])}")
    print(f"  mismatches      : {len(fk['mismatches'])}")


def main() -> int:
    if not DATA_MDB.exists():
        raise SystemExit(f"DATA mdb not found: {DATA_MDB}")

    print(f"connecting to {DATA_MDB.name} ...")
    conn = _open_data_mdb()

    print("\nStep 1 — dumping current table contents ...")
    tf_current, fk_current = _dump_current(conn)

    print("\nStep 2 — reconstructing from ODBC catalog ...")
    tf_regen, fk_regen, fk_available = _regen_from_catalog(conn)

    conn.close()

    print("\nStep 3 — computing diffs ...")
    diff = _compute_diffs(tf_current, tf_regen, fk_current, fk_regen, fk_available)

    _print_summary(diff)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
