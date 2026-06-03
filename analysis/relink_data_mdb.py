"""
Relink the linked tables in CBDB_BJ_User.mdb to the current DATA mdb.

Run after every DATA mdb swap (i.e. whenever CBDB_*_DATA.mdb is replaced
with a newer build).  The User mdb stores linked-table connection strings
that embed the DATA mdb path.  If the DATA mdb is replaced without running
this script, pyodbc and Access COM will both fail to access the linked
tables.

How linking works (VBA Form_NAVIGATION_PANE.Form_Open):
  1. Read LinkListInit.c_path.  If it equals CurrentProject.FullName → skip
     (fast-path bypass used by conftest to prevent the relink-hang).
  2. Otherwise: derive DATA mdb path as
       Left(CurrentProject.FullName, Len - 12) + "_" + c_dataset + "_DATA.mdb"
     and relink all tables listed in LinkedTables via tdf.RefreshLink().
  3. Write c_path = CurrentProject.FullName back to LinkListInit.

This script replicates step 2 & 3 via DAO so it can run without Access
being open.  It also updates c_dataset so the fast-path bypass stays
consistent with the new DATA mdb.

Usage:
    python analysis/relink_data_mdb.py

Runs automatically from conftest.py when the DATA mdb date in
LinkListInit.c_dataset doesn't match the DATA mdb found in data/.
"""
from __future__ import annotations

import sys
from pathlib import Path

import win32com.client  # pywin32

ROOT = Path(__file__).resolve().parent.parent
USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"


def find_data_mdb(root: Path) -> Path:
    matches = list((root / "data").glob("CBDB_*_DATA.mdb"))
    if not matches:
        raise FileNotFoundError("No CBDB_*_DATA.mdb found in data/")
    return sorted(matches, key=lambda p: p.stem.split("_")[1])[-1]


def relink(user_mdb: Path, data_mdb: Path) -> int:
    """Relink all linked tables and update LinkListInit.  Returns count."""
    dao = win32com.client.Dispatch("DAO.DBEngine.120")
    db = dao.OpenDatabase(str(user_mdb), False, False)  # writable

    connect_str = f"MS Access;DATABASE={data_mdb}"
    n = 0
    errors: list[str] = []
    for td in db.TableDefs:
        try:
            if td.Connect and ";DATABASE=" in td.Connect:
                td.Connect = connect_str
                td.RefreshLink()
                n += 1
        except Exception as e:
            errors.append(f"  {td.Name}: {e}")

    if errors:
        # Do NOT update LinkListInit — partial relink must fail hard so
        # future runs don't trust a broken state as "already fixed".
        db.Close()
        msg = "\n".join(errors)
        raise RuntimeError(
            f"{len(errors)} table(s) failed to relink:\n{msg}"
        )

    # All tables relinked successfully — update LinkListInit so the
    # fast-path bypass stays consistent with the new DATA mdb.
    rst = db.OpenRecordset("LinkListInit", 2)  # 2 = dbOpenDynaset
    rst.MoveFirst()
    rst.Edit()
    rst.Fields("c_path").Value = str(user_mdb)
    rst.Fields("c_dataset").Value = data_mdb.stem.split("_")[1]  # YYYYMMDD
    rst.Update()
    rst.Close()

    db.Close()

    return n


def main(argv: list[str] | None = None) -> None:
    import argparse
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[1].strip())
    parser.add_argument(
        "--user-mdb", default=str(USER_MDB),
        help="Path to CBDB_BJ_User.mdb to relink (default: data/CBDB_BJ_User.mdb)",
    )
    args = parser.parse_args(argv)
    user_mdb = Path(args.user_mdb).resolve()
    data_mdb = find_data_mdb(ROOT)
    print(f"[relink] linking {user_mdb.name} -> {data_mdb.name} ...")
    try:
        n = relink(user_mdb, data_mdb)
    except RuntimeError as exc:
        print(f"[relink] FAILED: {exc}", file=sys.stderr)
        sys.exit(1)
    print(f"[relink] done: {n} tables relinked")


if __name__ == "__main__":
    main()
