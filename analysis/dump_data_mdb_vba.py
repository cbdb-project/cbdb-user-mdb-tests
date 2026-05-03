"""SaveAsText dumper for forms + modules in
data/CBDB_<YYYYMMDD>_DATA.mdb (the linked-tables backend).

PR H got the QueryDef SQLs out via DAO read-only.  But form / module
*VBA source* needs `Access.Application.SaveAsText`, which has to
open the database via the Access COM application (not just DAO).

What this dumps:
  forms:    frmBaseMaintenance, FrmCopyTables
  modules:  Class1, FixCBDB_extra_programs, Module1, Module2

Output paths (matching the front-end VBA dump style at
analysis/dump/vba/Form_*.vb):
  analysis/dump_data/vba/Form_<name>.vb
  analysis/dump_data/vba/Module_<name>.bas
  analysis/dump_data/vba/Class_<name>.cls
  (Access produces .cls / .bas-ish text from SaveAsText; we keep the
  .vb suffix for forms to match the existing convention.)

Safety:

  - Copies the source DATA mdb to a working location FIRST and
    operates on the copy.  Never opens the original 1 GB file
    via COM.  This avoids any chance of marking it dirty,
    auto-compacting, or accidentally executing AutoExec.
  - Does NOT execute any maintenance button or sub.  Only
    SaveAsText.
  - Closes the working copy and quits Access cleanly.

Run via:
  python analysis/dump_data_mdb_vba.py
"""
from __future__ import annotations

import os
import shutil
import sys
import time
from pathlib import Path

import win32com.client

ROOT = Path(__file__).resolve().parent.parent
DATA_MDB = next(
    (p for p in (ROOT / "data").glob("CBDB_*_DATA.mdb")), None
)
WORK_MDB = ROOT / "analysis" / "_data_mdb_vba_work.mdb"
OUT_DIR = ROOT / "analysis" / "dump_data" / "vba"

# Access object-type constants (acObjectType enum).
ACCESS_OBJECT_TYPES = {
    "Form": 2,
    "Macro": 4,
    "Module": 5,
    "Report": 3,
}

# Targets — what to dump.  (object_type, name, output_filename).
TARGETS: list[tuple[str, str, str]] = [
    ("Form",   "frmBaseMaintenance", "Form_frmBaseMaintenance.vb"),
    ("Form",   "FrmCopyTables",      "Form_FrmCopyTables.vb"),
    # The 4 modules.  Class1 is technically a class module but
    # SaveAsText with acModule handles both; output suffix follows
    # the convention.
    ("Module", "Class1",                  "Class_Class1.cls"),
    ("Module", "FixCBDB_extra_programs",  "Module_FixCBDB_extra_programs.bas"),
    ("Module", "Module1",                 "Module_Module1.bas"),
    ("Module", "Module2",                 "Module_Module2.bas"),
]


def _open_access_for_saveastext():
    """Open Access against a writable WORKING COPY of the DATA mdb.
    Returns the Access.Application COM object."""
    if DATA_MDB is None:
        raise SystemExit(
            f"no CBDB_*_DATA.mdb under {ROOT/'data'}")
    # Refresh the working copy.  The DATA mdb is ~1GB so this takes
    # ~10s on a fast SSD.
    if WORK_MDB.exists():
        try:
            WORK_MDB.unlink()
        except PermissionError:
            time.sleep(1); WORK_MDB.unlink()
    print(f"copying {DATA_MDB.name} → {WORK_MDB.name} "
          f"({DATA_MDB.stat().st_size:,} bytes) ...")
    shutil.copy2(DATA_MDB, WORK_MDB)

    app = win32com.client.DispatchEx("Access.Application")
    try:
        app.AutomationSecurity = 1  # msoAutomationSecurityLow
    except Exception:
        pass
    # Open VISIBLE False — we don't need the UI.  But Access still
    # has to run its AutoExec on open, which on this DATA mdb
    # appears to be a no-op (it's a backend).
    app.Visible = False
    print(f"opening {WORK_MDB.name} via Access COM ...")
    app.OpenCurrentDatabase(str(WORK_MDB))
    return app


def _saveastext(app, obj_type_name: str, name: str, out_path: Path
                 ) -> bool:
    """Call Application.SaveAsText to dump one object's design text.
    Returns True on success."""
    out_path.parent.mkdir(parents=True, exist_ok=True)
    if out_path.exists():
        try:
            out_path.unlink()
        except Exception:
            pass
    obj_type_const = ACCESS_OBJECT_TYPES[obj_type_name]
    try:
        app.SaveAsText(obj_type_const, name, str(out_path))
    except Exception as e:
        print(f"  ✗ {obj_type_name} {name!r}: {e}")
        return False
    if not out_path.exists():
        print(f"  ✗ {obj_type_name} {name!r}: SaveAsText reported "
              f"no error but file not written")
        return False
    print(f"  ✓ {obj_type_name} {name!r} → {out_path.relative_to(ROOT)} "
          f"({out_path.stat().st_size:,} bytes)")
    return True


def main() -> int:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    app = _open_access_for_saveastext()
    ok = 0
    fail = 0
    try:
        for obj_type, name, fname in TARGETS:
            out = OUT_DIR / fname
            if _saveastext(app, obj_type, name, out):
                ok += 1
            else:
                fail += 1
    finally:
        try:
            app.CloseCurrentDatabase()
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass
    print(f"\n{ok} dumped, {fail} failed")
    print(f"\nworking copy left at {WORK_MDB} — safe to delete; "
          f"original DATA mdb at {DATA_MDB} not touched")
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
