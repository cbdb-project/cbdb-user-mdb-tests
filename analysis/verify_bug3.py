"""Direct verification of Bug #3 — does LookAtEntry.CmdQuery's
backfill UPDATE actually leave c_entry_desc NULL on a large result
set in the current dump?

Method:
  1. Open LookAtEntry via the driver (timer-trigger path, NOT
     pywinauto button click — the latter has a separate failure mode
     that masked Bug #3 verification all along).
  2. Set entry code 36 (jinshi general), no year filter — the original
     fixture that supposedly triggered Bug #3.
  3. Wait for CmdQuery_Click to complete (DONE marker).
  4. Read ZZ_SCRATCH_ENTRY directly via pyodbc.  Check:
       - row count
       - how many rows have c_entry_desc = NULL despite c_entry_code IS NOT NULL
       - sample 5 rows that should have a desc per ENTRY_CODES
  5. Report verdict.
"""
import sys, shutil
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "tests"))

from cbdb_driver.vba_session import VbaSession
from cbdb_driver.form_specs import LOOKATENTRY

ROOT = Path(__file__).parent.parent
SRC = ROOT / "data" / "CBDB_BJ_User.mdb"
WORK = ROOT / "analysis" / "_bug3_verify.mdb"


def main() -> int:
    if WORK.exists():
        try:
            WORK.unlink()
        except PermissionError:
            import time; time.sleep(1); WORK.unlink()
    shutil.copy2(SRC, WORK)
    vba = VbaSession(SRC, WORK)
    vba.open()
    try:
        spec = LOOKATENTRY
        vba.open_form(spec.name)
        # Top entry code 36 = "examination: jinshi (general)" — the
        # original Bug #3 fixture.
        vba.set_picker_codes(spec.picker_table, [36],
                              column=spec.picker_column)
        # FrameYears = 1 means "no year filter".
        vba.set_control(spec.name, "FrameYears", 1)
        print(f"firing {spec.name}.{spec.cmd_name} via timer ...")
        n = vba.click_via_timer(
            spec.name, ctl=spec.cmd_name,
            result_table=spec.result_table,
            timeout=180,  # large fixture — backfill chain takes ~60s+
        )
        print(f"  ZZ_SCRATCH_ENTRY row count after CmdQuery: {n:,}")

        cur = vba.conn.cursor()

        cur.execute("SELECT COUNT(*) FROM ZZ_SCRATCH_ENTRY")
        total = int(cur.fetchone()[0])
        print(f"\n  total rows: {total:,}")

        cur.execute(
            "SELECT COUNT(*) FROM ZZ_SCRATCH_ENTRY "
            "WHERE c_entry_code IS NOT NULL AND c_entry_desc IS NULL"
        )
        nullish = int(cur.fetchone()[0])
        print(f"  rows with c_entry_code NOT NULL but c_entry_desc IS NULL: {nullish:,}")
        if nullish > 0:
            pct = 100.0 * nullish / max(total, 1)
            print(f"    → {pct:.1f}% of rows have c_entry_desc NULL "
                  f"(Bug #3 STILL PRESENT)")
        else:
            print(f"    → 0 rows are missing c_entry_desc → "
                  f"Bug #3 IS NOT REPRODUCIBLE on this dump")

        # Sample: pick 5 c_entry_code values, see how many rows per
        # code are missing the desc that ENTRY_CODES would supply.
        cur.execute(
            "SELECT TOP 5 c_entry_code, COUNT(*) "
            "FROM ZZ_SCRATCH_ENTRY WHERE c_entry_code IS NOT NULL "
            "GROUP BY c_entry_code ORDER BY COUNT(*) DESC"
        )
        top_codes = cur.fetchall()
        print(f"\n  top entry codes in scratch:")
        for code, cnt in top_codes:
            cur2 = vba.conn.cursor()
            cur2.execute(
                f"SELECT c_entry_desc FROM ENTRY_CODES "
                f"WHERE c_entry_code = {int(code)}"
            )
            row = cur2.fetchone()
            expected_desc = row[0] if row else "?"
            cur2.execute(
                f"SELECT COUNT(*) FROM ZZ_SCRATCH_ENTRY "
                f"WHERE c_entry_code = {int(code)} "
                f"  AND c_entry_desc IS NULL"
            )
            missing = int(cur2.fetchone()[0])
            cur2.close()
            print(f"    code {code} (×{cnt:,} rows, "
                  f"expected desc='{expected_desc}'): "
                  f"{missing:,} NULL desc")

        # Also check c_addr_name (another column the multi-table UPDATE
        # supposedly fails on).
        cur.execute(
            "SELECT COUNT(*) FROM ZZ_SCRATCH_ENTRY "
            "WHERE c_addr_id IS NOT NULL AND c_addr_id > 0 "
            "  AND c_addr_name IS NULL"
        )
        addr_null = int(cur.fetchone()[0])
        print(f"\n  rows with c_addr_id > 0 but c_addr_name IS NULL: "
              f"{addr_null:,}")

        cur.close()
    finally:
        vba.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
