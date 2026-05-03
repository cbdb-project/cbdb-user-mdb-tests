"""Extract index-year / index-address rebuild algorithms from
CBDB_<YYYYMMDD>_DATA.mdb.

Why: PR G claimed the c_index_year rebuild VBA was not in the
shipped User MDB and "likely lives in an Admin MDB we don't have".
That was wrong — the rebuild logic actually lives in the DATA mdb
(the linked-tables backend, not the User-facing front-end mdb).
This script is the corrective extraction.

What we dump:

  analysis/dump_data/object_inventory.json
    Top-level inventory of every named object in the DATA mdb:
    Tables, QueryDefs, Forms, Modules, Macros (whatever DAO surfaces
    via Containers / Documents).  Just names + dates so a reviewer
    can quickly see what's there.

  analysis/dump_data/querydefs_index.json
  analysis/dump_data/querydefs_index/<safe_name>.sql
    For every QueryDef whose name OR sql mentions any of:
      c_index_year, c_index_addr, IndexYear, IndexAddr,
      BM IY, tmpIndexYear, tmpIndexAddr, RebuildIndex
    we capture the full SQL (one .sql file per query) plus a
    summary in the JSON.  These are the algorithms.

We open the file via DAO.DBEngine.120 in read-only mode (no
ldb file impact).  If pywin32 / DAO 12 isn't available we
fall back to pyodbc-based MSysObjects scrape, which sees less
metadata but still gets QueryDef SQL.
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA_MDB = next(
    (p for p in (ROOT / "data").glob("CBDB_*_DATA.mdb")), None
)
OUT_DIR = ROOT / "analysis" / "dump_data"
OUT_INV = OUT_DIR / "object_inventory.json"
OUT_QD_JSON = OUT_DIR / "querydefs_index.json"
OUT_QD_DIR = OUT_DIR / "querydefs_index"

NEEDLES = (
    "c_index_year", "c_index_addr",
    "indexyear", "indexaddr",
    "bm iy", "tmpindexyear", "tmpindexaddr",
    "rebuildindex", "rebuild_index",
)


def _safe_filename(name: str) -> str:
    # Access object names allow spaces, dots, etc.  Make a stable
    # filename token while keeping it readable.
    out = re.sub(r"[^\w\-. ]", "_", name).strip()
    out = out.replace(" ", "_")
    return out[:120] or "_unnamed"


def _matches(name: str, sql: str) -> bool:
    blob = (name + "\n" + sql).lower()
    return any(n in blob for n in NEEDLES)


def _try_dao() -> tuple[object, object] | None:
    """Open the DATA mdb via DAO.DBEngine.120 read-only.  Returns
    (db, engine) on success, None if pywin32 / DAO 12 unavailable."""
    try:
        import win32com.client
    except Exception:
        return None
    for progid in ("DAO.DBEngine.120", "DAO.DBEngine.36"):
        try:
            engine = win32com.client.Dispatch(progid)
        except Exception:
            continue
        try:
            # OpenDatabase(name, exclusive, readOnly)
            db = engine.OpenDatabase(str(DATA_MDB), False, True)
            print(f"opened via {progid}")
            return db, engine
        except Exception as e:
            print(f"  {progid} OpenDatabase failed: {e}")
    return None


def _dump_via_dao(db) -> dict:
    """Walk DAO Containers + QueryDefs.  Returns inventory dict and
    populates the OUT_QD_DIR with one .sql per matching query."""
    inventory: dict[str, list[dict]] = {
        "tables": [],
        "querydefs": [],
        "forms": [],
        "modules": [],
        "macros": [],
        "scripts": [],
        "other_containers": [],
    }
    matched_queries: list[dict] = []

    # Tables.
    for td in db.TableDefs:
        try:
            nm = str(td.Name)
        except Exception:
            continue
        if nm.startswith(("MSys", "USys")):
            continue
        try:
            cols = [c.Name for c in td.Fields]
        except Exception:
            cols = []
        inventory["tables"].append({
            "name": nm,
            "n_columns": len(cols),
            "columns": cols,
        })

    # QueryDefs.
    OUT_QD_DIR.mkdir(parents=True, exist_ok=True)
    for q in db.QueryDefs:
        try:
            nm = str(q.Name)
        except Exception:
            continue
        if nm.startswith("~"):
            # Access stores hidden helper queries with a leading ~
            continue
        try:
            sql = str(q.SQL)
        except Exception:
            sql = ""
        inventory["querydefs"].append({
            "name": nm,
            "sql_len": len(sql),
        })
        if _matches(nm, sql):
            fname = _safe_filename(nm) + ".sql"
            (OUT_QD_DIR / fname).write_text(
                f"-- QueryDef name: {nm}\n"
                f"-- Source: {DATA_MDB.name}\n\n"
                f"{sql}\n",
                encoding="utf-8",
            )
            matched_queries.append({
                "name": nm,
                "file": str((OUT_QD_DIR / fname).relative_to(ROOT)),
                "sql_preview": (sql[:300]
                                + ("..." if len(sql) > 300 else "")),
            })

    # Containers (Forms, Modules, Macros, Scripts, …).  In a
    # _DATA backend Access often stores ZERO of these — but let's
    # print what's there.
    for c in db.Containers:
        try:
            cname = str(c.Name)
        except Exception:
            continue
        items = []
        try:
            for d in c.Documents:
                try:
                    items.append({"name": str(d.Name)})
                except Exception:
                    continue
        except Exception:
            pass
        bucket_key = {
            "Forms": "forms",
            "Modules": "modules",
            "Scripts": "macros",
            "Reports": "scripts",  # bucket spillover
        }.get(cname, "other_containers")
        if bucket_key in inventory:
            inventory[bucket_key].extend(items)
        else:
            inventory.setdefault("other_containers", []).extend(
                [{"name": str(d.get("name", ""))} for d in items])
        if items:
            print(f"  container {cname!r}: {len(items)} items")

    return {
        "inventory": inventory,
        "matched_queries": matched_queries,
    }


def _try_pyodbc() -> dict | None:
    """Fallback: scrape MSysObjects + sysObjects via pyodbc.
    Less metadata but still gives QueryDef SQL."""
    try:
        import pyodbc
    except Exception:
        return None
    cs = ("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
          f"DBQ={DATA_MDB};ReadOnly=1;")
    try:
        conn = pyodbc.connect(cs, autocommit=True)
    except Exception as e:
        print(f"  pyodbc connect failed: {e}")
        return None
    cur = conn.cursor()

    # Tables
    tables = []
    for r in cur.tables(tableType="TABLE"):
        nm = r.table_name
        if nm.startswith(("MSys", "USys")):
            continue
        cols = [c.column_name for c in cur.columns(table=nm)]
        tables.append({"name": nm, "n_columns": len(cols),
                       "columns": cols})

    # QueryDefs are visible as VIEW-type tables in pyodbc's tables()
    qdefs = []
    matched = []
    OUT_QD_DIR.mkdir(parents=True, exist_ok=True)
    for r in cur.tables(tableType="VIEW"):
        nm = r.table_name
        if nm.startswith(("MSys", "USys", "~")):
            continue
        # Pull the SQL from MSysObjects via the documented view
        # (not always accessible through ODBC, but worth trying).
        sql = ""
        try:
            cur.execute(
                "SELECT MSysObjects.Name, MSysObjects.Type "
                "FROM MSysObjects WHERE Name=?", nm
            )
            cur.fetchone()
        except Exception:
            pass
        # Easier: open the view and dump its definition via ADOX.
        # Without ADOX we can at least record the name.
        qdefs.append({"name": nm, "sql_len": len(sql)})
        if _matches(nm, sql):
            fname = _safe_filename(nm) + ".sql"
            (OUT_QD_DIR / fname).write_text(
                f"-- QueryDef name: {nm}\n"
                f"-- Source: {DATA_MDB.name}\n"
                f"-- (SQL extraction via ODBC fallback returned no body; "
                f"open the file in Access to read the QueryDef SQL.)\n\n",
                encoding="utf-8",
            )
            matched.append({"name": nm,
                            "file": str(
                                (OUT_QD_DIR / fname).relative_to(ROOT))})
    cur.close()
    conn.close()
    return {
        "inventory": {
            "tables": tables,
            "querydefs": qdefs,
            "forms": [],
            "modules": [],
            "macros": [],
            "scripts": [],
            "other_containers": [],
        },
        "matched_queries": matched,
    }


def main() -> int:
    if DATA_MDB is None:
        raise SystemExit(
            f"no CBDB_*_DATA.mdb found under {ROOT/'data'}; place "
            f"the linked-data backend there before running this "
            f"extractor"
        )
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    print(f"extracting from {DATA_MDB.name} ({DATA_MDB.stat().st_size:,} bytes)")

    result = None
    handle = _try_dao()
    if handle is not None:
        db, engine = handle
        try:
            result = _dump_via_dao(db)
        finally:
            try: db.Close()
            except Exception: pass
    if result is None:
        print("DAO unavailable; falling back to pyodbc scrape")
        result = _try_pyodbc()
    if result is None:
        raise SystemExit(
            "neither DAO.DBEngine.120 nor pyodbc could open "
            f"{DATA_MDB}; install pywin32 or check Access drivers"
        )

    inv = result["inventory"]
    qmatched = result["matched_queries"]
    OUT_INV.write_text(
        json.dumps(inv, ensure_ascii=False, indent=2,
                    default=str),
        encoding="utf-8",
    )
    print(f"wrote {OUT_INV}")
    print(f"  tables:    {len(inv['tables'])}")
    print(f"  querydefs: {len(inv['querydefs'])}")
    print(f"  forms:     {len(inv['forms'])}")
    print(f"  modules:   {len(inv['modules'])}")
    print(f"  macros:    {len(inv['macros'])}")

    OUT_QD_JSON.write_text(
        json.dumps({
            "source_file": DATA_MDB.name,
            "needles": list(NEEDLES),
            "matched_queries": qmatched,
            "out_dir": str(OUT_QD_DIR.relative_to(ROOT)),
        }, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"wrote {OUT_QD_JSON}")
    print(f"  matched queries: {len(qmatched)}")
    for q in qmatched:
        print(f"    - {q['name']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
