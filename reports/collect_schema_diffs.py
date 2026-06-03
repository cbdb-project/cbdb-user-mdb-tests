"""Collect diffs between the documentation tables (TablesFields,
ForeignKeys) inside CBDB_20260430_DATA.mdb and the actual database
structure reconstructed from Access DAO (TableDefs) via COM.

Outputs (all in reports/):
  tables_fields_current.csv   — live dump of TablesFields
  foreign_keys_current.csv    — live dump of ForeignKeys
  tables_fields_regen.csv     — reconstructed from DAO TableDefs
  foreign_keys_regen.csv      — reconstructed from DAO Relations
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
import glob as _glob
_data_candidates = sorted(_glob.glob(str(ROOT / "data" / "CBDB_*_DATA.mdb")))
if not _data_candidates:
    raise SystemExit("No CBDB_*_DATA.mdb found in data/")
DATA_MDB = Path(_data_candidates[-1])
REPORTS = ROOT / "reports"

OUT_TF_CURRENT = REPORTS / "tables_fields_current.csv"
OUT_FK_CURRENT = REPORTS / "foreign_keys_current.csv"
OUT_TF_REGEN   = REPORTS / "tables_fields_regen.csv"
OUT_FK_REGEN   = REPORTS / "foreign_keys_regen.csv"
OUT_JSON       = REPORTS / "schema_diff.json"

# Diff-result CSV files (written by _write_diff_csvs)
OUT_TF_ONLY_CURRENT  = REPORTS / "schema_diff_tables_fields_only_in_current.csv"
OUT_TF_ONLY_REGEN    = REPORTS / "schema_diff_tables_fields_only_in_regen.csv"
OUT_TF_MISMATCHES    = REPORTS / "schema_diff_tables_fields_mismatches.csv"
OUT_FK_ONLY_CURRENT  = REPORTS / "schema_diff_foreign_keys_only_in_current.csv"
OUT_FK_ONLY_REGEN    = REPORTS / "schema_diff_foreign_keys_only_in_regen.csv"
OUT_FK_MISMATCHES    = REPORTS / "schema_diff_foreign_keys_mismatches.csv"


def _open_data_mdb() -> pyodbc.Connection:
    cs = (
        "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={DATA_MDB};"
    )
    return pyodbc.connect(cs, autocommit=True)


# ---------------------------------------------------------------------------
# DAO integer type code → Access DataFormat
# ---------------------------------------------------------------------------
DAO_TYPE_TO_FORMAT: dict[int, str] = {
    1:   "Yes/No",          # dbBoolean
    2:   "Byte",            # dbByte
    3:   "Integer",         # dbInteger
    4:   "Long",            # dbLong
    5:   "Currency",        # dbCurrency
    6:   "Single",          # dbSingle
    7:   "Double",          # dbDouble
    8:   "Date/Time",       # dbDate
    10:  "Text",            # dbText
    11:  "OLE Object",      # dbLongBinary
    12:  "Memo",            # dbMemo
    15:  "Replication ID",  # dbGUID
    16:  "Long",            # dbBigInt (Access 2016+ Large Number → map to Long for compat)
    17:  "Double",          # dbFloat
    20:  "Decimal",         # dbDecimal
    101: "Attachment",      # dbAttachment
}
# Fallback for unknown types: f"DAO_type_{type_int}"


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
# Step 2 — reconstruct from Access DAO (TableDefs + Relations)
# ---------------------------------------------------------------------------

def _get_dao_data(data_mdb_path: Path) -> tuple[list[dict], list[dict]]:
    """Open Access once via COM, collect TablesFields regen rows (from
    db.TableDefs) and FK regen rows (from db.Relations), then quit.

    Returns (tf_rows, fk_rows).  Raises on COM failure so the caller can
    surface the error clearly.
    """
    try:
        import win32com.client  # type: ignore
    except ImportError:
        raise RuntimeError("win32com not available — install pywin32")

    app = win32com.client.DispatchEx("Access.Application")
    app.Visible = False
    try:
        app.OpenCurrentDatabase(str(data_mdb_path), False)
        db = app.CurrentDb()

        # --- TablesFields regen ---
        tf_rows: list[dict] = []
        tdefs = db.TableDefs
        print(f"  DAO db.TableDefs count (raw): {tdefs.Count}")
        for i in range(tdefs.Count):
            tdef = tdefs[i]
            tname = tdef.Name
            if tname.startswith("MSys") or tname.startswith("~"):
                continue
            # Build PK field set for this table
            pk_fields: set[str] = set()
            indexes = tdef.Indexes
            for k in range(indexes.Count):
                idx = indexes[k]
                if idx.Primary:
                    idx_fields = idx.Fields
                    for m in range(idx_fields.Count):
                        pk_fields.add(idx_fields[m].Name.upper())
            # Enumerate fields
            fields = tdef.Fields
            for j in range(fields.Count):
                fld = fields[j]
                data_format = DAO_TYPE_TO_FORMAT.get(fld.Type, f"DAO_type_{fld.Type}")
                null_allowed = not bool(fld.Required)
                index_on_field = "Primary Key" if fld.Name.upper() in pk_fields else None
                tf_rows.append({
                    "AccessTblNm": tname,
                    "AccessFldNm": fld.Name,
                    "IndexOnField": index_on_field,
                    "DataFormat": data_format,
                    "NULL_allowed": null_allowed,
                    "ForeignKey": None,
                    "ForeignKeyBaseField": None,
                })

        # --- FK regen (same session) ---
        fk_rows: list[dict] = []
        rels = db.Relations
        print(f"  DAO db.Relations count: {rels.Count}")
        for i in range(rels.Count):
            rel = rels[i]
            rel_fields = rel.Fields
            for j in range(rel_fields.Count):
                fld = rel_fields[j]
                fk_rows.append({
                    "AccessTblNm": rel.ForeignTable,
                    "AccessFldNm": fld.ForeignName,
                    "ForeignKey": rel.Table,
                    "ForeignKeyBaseField": fld.Name,
                    "RelationName": rel.Name,
                })
    finally:
        try:
            app.CloseCurrentDatabase()
            app.Quit()
        except Exception:
            pass

    return tf_rows, fk_rows


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
        # Normalise to uppercase for comparison — DAO returns mixed-case names
        # (e.g. 'Assoc_data') while the ForeignKeys documentation table uses
        # uppercase.  We normalise the KEY but preserve the original case of
        # each source when building the display rows.
        def _fk_upper_key(row: dict) -> tuple:
            return (
                (row["AccessTblNm"] or "").upper(),
                (row["AccessFldNm"] or "").upper(),
                (row["ForeignKey"] or "").upper(),
                (row["ForeignKeyBaseField"] or "").upper(),
            )

        # Build lookup: normalised key → original-case row (for display)
        fk_cur_map: dict[tuple, dict] = {}
        for r in fk_current:
            if r.get("ForeignKey"):
                fk_cur_map[_fk_upper_key(r)] = r

        fk_reg_map: dict[tuple, dict] = {}
        for r in fk_regen:
            fk_reg_map[_fk_upper_key(r)] = r

        fk_cur_set = set(fk_cur_map)
        fk_reg_set = set(fk_reg_map)

        fk_only_in_current = [
            {
                "AccessTblNm": fk_cur_map[k]["AccessTblNm"],
                "AccessFldNm": fk_cur_map[k]["AccessFldNm"],
                "ForeignKey": fk_cur_map[k]["ForeignKey"],
                "ForeignKeyBaseField": fk_cur_map[k]["ForeignKeyBaseField"],
            }
            for k in sorted(fk_cur_set - fk_reg_set)
        ]
        fk_only_in_regen = [
            {
                "AccessTblNm": fk_reg_map[k]["AccessTblNm"],
                "AccessFldNm": fk_reg_map[k]["AccessFldNm"],
                "ForeignKey": fk_reg_map[k]["ForeignKey"],
                "ForeignKeyBaseField": fk_reg_map[k]["ForeignKeyBaseField"],
            }
            for k in sorted(fk_reg_set - fk_cur_set)
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


def _write_diff_csvs(diff: dict) -> None:
    """Write the six diff-result CSV files from the in-memory diff data."""
    tf = diff["tables_fields"]
    fk = diff["foreign_keys"]

    # --- TablesFields: only_in_current ---
    rows = tf["only_in_current"]
    with OUT_TF_ONLY_CURRENT.open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=["AccessTblNm", "AccessFldNm"])
        writer.writeheader()
        writer.writerows(rows)
    print(f"wrote {OUT_TF_ONLY_CURRENT}  ({len(rows)} rows)")

    # --- TablesFields: only_in_regen ---
    rows = tf["only_in_regen"]
    with OUT_TF_ONLY_REGEN.open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=["AccessTblNm", "AccessFldNm", "DataFormat", "NULL_allowed"])
        writer.writeheader()
        writer.writerows(rows)
    print(f"wrote {OUT_TF_ONLY_REGEN}  ({len(rows)} rows)")

    # --- TablesFields: mismatches ---
    rows = tf["mismatches"]
    with OUT_TF_MISMATCHES.open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=["AccessTblNm", "AccessFldNm", "field", "current", "regen"])
        writer.writeheader()
        writer.writerows(rows)
    print(f"wrote {OUT_TF_MISMATCHES}  ({len(rows)} rows)")

    # --- ForeignKeys: only_in_current ---
    rows = fk["only_in_current"]
    fk_cols = ["AccessTblNm", "AccessFldNm", "ForeignKey", "ForeignKeyBaseField"]
    with OUT_FK_ONLY_CURRENT.open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=fk_cols)
        writer.writeheader()
        writer.writerows(rows)
    print(f"wrote {OUT_FK_ONLY_CURRENT}  ({len(rows)} rows)")

    # --- ForeignKeys: only_in_regen ---
    rows = fk["only_in_regen"]
    with OUT_FK_ONLY_REGEN.open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=fk_cols)
        writer.writeheader()
        writer.writerows(rows)
    print(f"wrote {OUT_FK_ONLY_REGEN}  ({len(rows)} rows)")

    # --- ForeignKeys: mismatches ---
    rows = fk["mismatches"]
    with OUT_FK_MISMATCHES.open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=fk_cols)
        writer.writeheader()
        writer.writerows(rows)
    print(f"wrote {OUT_FK_MISMATCHES}  ({len(rows)} rows)")



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

    conn.close()

    print("\nStep 2 — reconstructing schema via Access DAO (TableDefs + Relations) ...")
    tf_regen, fk_regen = _get_dao_data(DATA_MDB)
    fk_available = bool(fk_regen)

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
        "AccessTblNm", "AccessFldNm", "ForeignKey",
        "ForeignKeyBaseField", "RelationName",
    ]
    fk_regen_sorted = sorted(
        fk_regen,
        key=lambda r: (r["AccessTblNm"].upper(), r["AccessFldNm"].upper()),
    )
    with OUT_FK_REGEN.open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=fk_regen_cols)
        writer.writeheader()
        writer.writerows(fk_regen_sorted)
    print(f"wrote {OUT_FK_REGEN}  ({len(fk_regen_sorted)} rows, via DAO)")

    print("\nStep 3 — computing diffs ...")
    diff = _compute_diffs(tf_current, tf_regen, fk_current, fk_regen, fk_available)

    print("\nStep 4 — writing diff CSV files ...")
    _write_diff_csvs(diff)

    _print_summary(diff)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
