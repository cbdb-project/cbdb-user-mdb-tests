"""Diagnose why LookAtOffice CmdQuery_Click returns 0 rows.

1. Verify ZZ_OFFICE_CODE schema + that we can write to it.
2. Insert top office id (80944).
3. Run the exact SQL CmdQuery_Click would produce in our mode
   (TxtTypeDesc != "N/A", no addresses, no year filter):
      INSERT INTO ZZ_SCRATCH_OFFICE (...) SELECT ...
      FROM BIOG_MAIN INNER JOIN ((POD INNER JOIN PAD ON ...)
        INNER JOIN ZZ_OFFICE_CODE ON POD.c_office_id = ZZ_OFFICE_CODE.c_office_id)
        ON BIOG_MAIN.c_personid = POD.c_personid
4. Count rows in ZZ_SCRATCH_OFFICE.

If this works -> the issue is in CmdQuery_Click's runtime gating
(probably gUseOfficeID still False, so it took the TxtOfficeID path
and TxtOfficeID = -1 -> 0 rows).

If this fails -> the issue is the SQL itself or picker schema.
"""
from __future__ import annotations
import shutil
from pathlib import Path
import pyodbc

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK = ROOT / "analysis" / "_diag_office.mdb"
OFFICE_ID = 80944


def conn(p: Path):
    s = (
        "Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={p};"
    )
    c = pyodbc.connect(s, autocommit=True)
    return c


def main():
    if WORK.exists():
        WORK.unlink()
    shutil.copy(SRC, WORK)
    c = conn(WORK)
    cur = c.cursor()

    # 1. schema
    cur.execute("SELECT TOP 1 * FROM ZZ_OFFICE_CODE")
    cols = [d[0] for d in cur.description]
    print(f"ZZ_OFFICE_CODE columns: {cols}")

    # 2. populate
    cur.execute("DELETE FROM ZZ_OFFICE_CODE")
    cur.execute(
        f"INSERT INTO ZZ_OFFICE_CODE (c_office_id) VALUES ({OFFICE_ID})"
    )
    cur.execute("SELECT COUNT(*) FROM ZZ_OFFICE_CODE")
    print(f"ZZ_OFFICE_CODE rows: {cur.fetchone()[0]}")

    # 3. clear ZZ_SCRATCH_OFFICE
    cur.execute("DELETE FROM ZZ_SCRATCH_OFFICE")

    # 3a. count POD rows for this office_id
    cur.execute(
        f"SELECT COUNT(*) FROM POSTED_TO_OFFICE_DATA "
        f"WHERE c_office_id = {OFFICE_ID}"
    )
    n_pod = cur.fetchone()[0]
    print(f"POD rows for office_id={OFFICE_ID}: {n_pod}")

    # 3b. count PAD rows
    cur.execute(
        f"SELECT COUNT(*) FROM POSTED_TO_ADDR_DATA "
        f"WHERE c_office_id = {OFFICE_ID}"
    )
    n_pad = cur.fetchone()[0]
    print(f"PAD rows for office_id={OFFICE_ID}: {n_pad}")

    # 4. Run the exact SQL the VBA would build for our test mode
    # (no addresses, no year filter, TxtTypeDesc != "N/A" -> picker branch)
    insert_sql = (
        "INSERT INTO ZZ_SCRATCH_OFFICE ("
        "c_posting_id, c_personid, c_index_year, c_female, c_person_dy, "
        "c_office_id, c_sequence, c_firstyear, c_fy_nh_code, c_fy_nh_year, "
        "c_fy_range, c_lastyear, c_ly_nh_code, c_ly_nh_year, c_ly_range, "
        "c_appt_code, c_assume_office_code, c_inst_code, c_inst_name_code, "
        "c_source, c_pages, c_notes, c_fy_intercalary, c_fy_month, "
        "c_ly_intercalary, c_ly_month, c_fy_day, c_ly_day, c_fy_day_gz, "
        "c_ly_day_gz, c_dy, c_office_category_id, c_addr_id, c_addr_type, "
        "c_office_addr_id, c_index_year_type_code) "
        "SELECT POD.c_posting_id, POD.c_personid, BIOG_MAIN.c_index_year, "
        "BIOG_MAIN.c_female, BIOG_MAIN.c_dy, POD.c_office_id, POD.c_sequence, "
        "POD.c_firstyear, POD.c_fy_nh_code, POD.c_fy_nh_year, POD.c_fy_range, "
        "POD.c_lastyear, POD.c_ly_nh_code, POD.c_ly_nh_year, POD.c_ly_range, "
        "POD.c_appt_code, POD.c_assume_office_code, POD.c_inst_code, "
        "POD.c_inst_name_code, POD.c_source, POD.c_pages, POD.c_notes, "
        "POD.c_fy_intercalary, POD.c_fy_month, POD.c_ly_intercalary, "
        "POD.c_ly_month, POD.c_fy_day, POD.c_ly_day, POD.c_fy_day_gz, "
        "POD.c_ly_day_gz, BIOG_MAIN.c_dy, POD.c_office_category_id, "
        "BIOG_MAIN.c_index_addr_id, BIOG_MAIN.c_index_addr_type_code, "
        "PAD.c_addr_id, BIOG_MAIN.c_index_year_type_code "
        "FROM BIOG_MAIN INNER JOIN ((POSTED_TO_OFFICE_DATA AS POD "
        "INNER JOIN POSTED_TO_ADDR_DATA AS PAD "
        "ON (POD.c_posting_id = PAD.c_posting_id) "
        "AND (POD.c_office_id = PAD.c_office_id)) "
        "INNER JOIN ZZ_OFFICE_CODE ON POD.c_office_id = ZZ_OFFICE_CODE.c_office_id) "
        "ON BIOG_MAIN.c_personid = POD.c_personid"
    )
    try:
        cur.execute(insert_sql)
        print(f"INSERT executed.  rowcount={cur.rowcount}")
    except Exception as e:
        print(f"INSERT FAILED: {e}")
        # try single-stage
        try:
            cur.execute(
                "SELECT COUNT(*) FROM POSTED_TO_OFFICE_DATA AS POD "
                "INNER JOIN POSTED_TO_ADDR_DATA AS PAD "
                "ON (POD.c_posting_id = PAD.c_posting_id) "
                "AND (POD.c_office_id = PAD.c_office_id) "
                f"WHERE POD.c_office_id = {OFFICE_ID}"
            )
            print(f"  POD ⨝ PAD rows: {cur.fetchone()[0]}")
        except Exception as e2:
            print(f"  even POD⨝PAD failed: {e2}")

    cur.execute("SELECT COUNT(*) FROM ZZ_SCRATCH_OFFICE")
    print(f"ZZ_SCRATCH_OFFICE final rows: {cur.fetchone()[0]}")

    cur.close()
    c.close()


if __name__ == "__main__":
    main()
