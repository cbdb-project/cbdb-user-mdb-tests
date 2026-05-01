"""
Extract complete metadata + VBA source from CBDB_BJ_User.mdb.

Outputs (all in ./dump/):
  - tables.json         : table list + column schema (from DAO)
  - queries.json        : every query with its SQL
  - forms.json          : every form's properties + control tree + module code
  - modules.json        : every standalone module's full VBA source
  - macros.json         : list of macros (name only -- bodies require XML export)
  - relationships.json  : foreign-key style relations between tables
  - summary.txt         : human-readable inventory

Run from the project root:
    python analysis/dump_metadata.py
"""

from __future__ import annotations

import json
import os
import sys
import time
from pathlib import Path

import win32com.client  # pywin32

ROOT = Path(__file__).resolve().parent.parent
USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
DUMP_DIR = Path(__file__).resolve().parent / "dump"
DUMP_DIR.mkdir(exist_ok=True)


def _json_default(o):
    # Best-effort: convert COM objects + everything weird to a string repr.
    try:
        return str(o)
    except Exception:
        return repr(o)


def jdump(obj, name: str) -> None:
    path = DUMP_DIR / name
    path.write_text(
        json.dumps(obj, ensure_ascii=False, indent=2, default=_json_default),
        encoding="utf-8",
    )
    print(f"  wrote {path.name}  ({path.stat().st_size:,} bytes)")


def safe(getter, default=None):
    try:
        return getter()
    except Exception:
        return default


# ---------- DAO: tables, queries, relationships ----------

def dump_via_dao(mdb_path: Path):
    print("[1/4] DAO: tables / queries / relationships")
    dao = win32com.client.Dispatch("DAO.DBEngine.120")
    db = dao.OpenDatabase(str(mdb_path))

    # ----- tables -----
    tables = []
    for td in db.TableDefs:
        if td.Name.startswith(("MSys", "~")):
            continue
        cols = []
        for f in td.Fields:
            cols.append({
                "name": f.Name,
                "type": int(f.Type),
                "size": int(f.Size),
                "required": bool(f.Required),
                "allow_zero_length": bool(safe(lambda: f.AllowZeroLength, False)),
            })
        idxs = []
        for ix in td.Indexes:
            ix_fields = []
            try:
                for ixf in ix.Fields:
                    ix_fields.append(ixf.Name)
            except Exception:
                # Some Index.Fields are exposed as a parseable string instead of a collection
                try:
                    ix_fields = [str(ix.Fields)]
                except Exception:
                    ix_fields = []
            idxs.append({
                "name": ix.Name,
                "primary": bool(ix.Primary),
                "unique": bool(ix.Unique),
                "fields": ix_fields,
            })
        tables.append({
            "name": td.Name,
            "record_count": int(safe(lambda: td.RecordCount, -1)),
            "columns": cols,
            "indexes": idxs,
        })
    jdump(tables, "tables.json")

    # ----- queries -----
    queries = []
    for qd in db.QueryDefs:
        if qd.Name.startswith("~"):
            continue
        params = []
        for p in qd.Parameters:
            params.append({"name": p.Name, "type": int(p.Type)})
        queries.append({
            "name": qd.Name,
            "type": int(qd.Type),
            "sql": qd.SQL,
            "parameters": params,
        })
    jdump(queries, "queries.json")

    # ----- relationships -----
    rels = []
    for r in db.Relations:
        fields = [{"name": f.Name, "foreign": f.ForeignName} for f in r.Fields]
        rels.append({
            "name": r.Name,
            "table": r.Table,
            "foreign_table": r.ForeignTable,
            "attributes": int(r.Attributes),
            "fields": fields,
        })
    jdump(rels, "relationships.json")

    db.Close()
    print(f"  tables={len(tables)} queries={len(queries)} relations={len(rels)}")


# ---------- Access.Application: forms, modules, macros, VBA ----------

def dump_via_access(mdb_path: Path):
    print("[2/4] Access.Application: opening database (this may take a moment)...")
    app = win32com.client.Dispatch("Access.Application")
    app.Visible = False
    try:
        app.OpenCurrentDatabase(str(mdb_path))
    except Exception as e:
        print("  ERROR opening DB:", e)
        app.Quit()
        raise

    try:
        # AccessObjectType: Forms=2, Reports=3, Macros=4, Modules=5, Queries=1
        proj = app.CurrentProject

        # ----- standalone modules -----
        print("[3/4] dumping standalone modules")
        modules = []
        for ao in proj.AllModules:
            name = ao.Name
            try:
                # open as Access module, read all lines from CodeModule
                app.DoCmd.OpenModule(name)
                mod = app.Modules(name)
                cm = mod.CodeModule
                line_count = cm.CountOfLines
                code = cm.Lines(1, line_count) if line_count > 0 else ""
                modules.append({"name": name, "lines": line_count, "code": code})
                # close it to keep things light
                app.DoCmd.Close(5, name)  # acModule = 5
            except Exception as e:
                modules.append({"name": name, "error": str(e)})
        jdump(modules, "modules.json")

        # ----- forms (open in design view to read controls + module code) -----
        print("[4/4] dumping forms (open each in design view)")
        forms = []
        form_names = [ao.Name for ao in proj.AllForms]
        print(f"  total forms: {len(form_names)}")
        for i, fname in enumerate(form_names, 1):
            entry = {"name": fname}
            try:
                # acDesign = 1
                app.DoCmd.OpenForm(fname, 1, "", "", 0, 1)
                f = app.Forms(fname)
                # form-level properties
                props = {}
                for pn in ("RecordSource", "Filter", "OrderBy", "Caption",
                           "DefaultView", "AllowEdits", "AllowAdditions",
                           "AllowDeletions", "DataEntry"):
                    props[pn] = safe(lambda pn=pn: getattr(f, pn))
                entry["properties"] = props

                # controls
                controls = []
                for c in f.Controls:
                    cdesc = {
                        "name": safe(lambda c=c: c.Name),
                        "control_type": safe(lambda c=c: int(c.ControlType)),
                        "control_source": safe(lambda c=c: getattr(c, "ControlSource", None)),
                        "row_source": safe(lambda c=c: getattr(c, "RowSource", None)),
                        "row_source_type": safe(lambda c=c: getattr(c, "RowSourceType", None)),
                        "tag": safe(lambda c=c: getattr(c, "Tag", None)),
                        "caption": safe(lambda c=c: getattr(c, "Caption", None)),
                    }
                    # collect any *On* event property bound to a procedure
                    events = {}
                    for ev in ("OnClick", "OnDblClick", "OnChange", "OnGotFocus",
                               "OnLostFocus", "OnEnter", "OnExit", "AfterUpdate",
                               "BeforeUpdate", "OnCurrent", "OnLoad", "OnOpen",
                               "OnClose", "OnUnload"):
                        v = safe(lambda ev=ev, c=c: getattr(c, ev, None))
                        if v:
                            events[ev] = v
                    if events:
                        cdesc["events"] = events
                    controls.append(cdesc)
                entry["controls"] = controls

                # form module code
                try:
                    if f.HasModule:
                        cm = f.Module.CodeModule
                        n = cm.CountOfLines
                        entry["code_lines"] = n
                        entry["code"] = cm.Lines(1, n) if n > 0 else ""
                except Exception as e:
                    entry["code_error"] = str(e)

                # acForm = 2, acSaveNo = 2
                app.DoCmd.Close(2, fname, 2)
            except Exception as e:
                entry["error"] = str(e)
            forms.append(entry)
            if i % 10 == 0 or i == len(form_names):
                print(f"    [{i}/{len(form_names)}] {fname}")
        jdump(forms, "forms.json")

        # ----- macros (just names; XML export is heavier) -----
        macros = [{"name": ao.Name} for ao in proj.AllMacros]
        jdump(macros, "macros.json")

    finally:
        try:
            app.CloseCurrentDatabase()
        except Exception:
            pass
        app.Quit()


def write_summary():
    lines = []
    for fname in ("tables.json", "queries.json", "forms.json", "modules.json",
                  "macros.json", "relationships.json"):
        p = DUMP_DIR / fname
        if not p.exists():
            continue
        data = json.loads(p.read_text(encoding="utf-8"))
        lines.append(f"{fname}: {len(data)} entries")
    (DUMP_DIR / "summary.txt").write_text("\n".join(lines), encoding="utf-8")
    print("\n=== summary ===")
    print("\n".join(lines))


def main():
    if not USER_MDB.exists():
        print("missing:", USER_MDB)
        sys.exit(1)
    t0 = time.time()
    dump_via_dao(USER_MDB)
    dump_via_access(USER_MDB)
    write_summary()
    print(f"\ndone in {time.time() - t0:.1f}s")


if __name__ == "__main__":
    main()
