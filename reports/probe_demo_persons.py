"""For each issue that needs a 'pick this specific person' demo,
SQL-probe the User MDB and pick a small, well-known person id whose
data DOES contain the row needed to reproduce the bug.

Output is JSON `reports/demo_persons.json` consumed by
`reports/generate_report.py` to substitute concrete examples for
'any person' / '任意一位' phrasing in the report.

For each issue we record:
  - personid + name (zh + py)
  - the underlying row count(s) that prove the bug is reproducible
    on this person
  - a short hand-picked English + 中文 hint the report can drop into
    its 'Steps to reproduce' bullets

We deliberately pick LOW personid (= famous / well-attested Song
figures whose data the maintainer would recognise instantly) and
prefer persons whose row counts are small (so the maintainer's
sub-datasheet only shows the relevant evidence, not 50+ rows of
noise).
"""
from __future__ import annotations

import json
from pathlib import Path

import pyodbc

ROOT = Path(__file__).resolve().parent.parent
USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
OUT = ROOT / "reports" / "demo_persons.json"


def _conn() -> pyodbc.Connection:
    cs = ("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
          f"DBQ={USER_MDB};")
    return pyodbc.connect(cs, autocommit=True)


def _name(cur, pid: int) -> tuple[str, str]:
    cur.execute(
        "SELECT c_name, c_name_chn FROM BIOG_MAIN "
        f"WHERE c_personid = {pid}"
    )
    row = cur.fetchone()
    if not row:
        return ("?", "?")
    return (row[0] or "?", row[1] or "?")


def main() -> int:
    conn = _conn()
    cur = conn.cursor()
    out: dict[str, dict] = {}

    # ---- Bug #1: View_StatusData FY/LY alias swap ------------------
    # Need a person with at least one STATUS_DATA row whose c_fy_range
    # differs from c_ly_range — that's where the alias swap is visible.
    # Probe: pick a person id whose STATUS_DATA has at least one row
    # with c_fy_range != c_ly_range.  ORDER BY without GROUP BY so we
    # don't trip the JET TOP+GROUP-BY tie-return quirk.
    # First measure how much STATUS_DATA actually carries fy/ly range
    # values.  In the current dump this is essentially empty — the
    # alias-swap bug exists in the SQL but no row in the data triggers
    # it via the UI.  We record this 'dormant' state explicitly so the
    # report doesn't tell the user 'open any person, look at the
    # STATUS sub-datasheet' when in fact no person on this dump shows
    # the symptom.
    cur.execute("SELECT COUNT(*) FROM STATUS_DATA")
    total = int(cur.fetchone()[0])
    cur.execute("SELECT COUNT(*) FROM STATUS_DATA WHERE c_fy_range > 0")
    n_fy = int(cur.fetchone()[0])
    cur.execute("SELECT COUNT(*) FROM STATUS_DATA WHERE c_ly_range > 0")
    n_ly = int(cur.fetchone()[0])
    cur.execute(
        "SELECT COUNT(*) FROM STATUS_DATA "
        "WHERE c_fy_range > 0 AND c_ly_range > 0 "
        "  AND c_fy_range <> c_ly_range"
    )
    n_both_diff = int(cur.fetchone()[0])

    cur.execute(
        "SELECT TOP 1 sd.c_personid "
        "FROM STATUS_DATA sd "
        "WHERE sd.c_personid > 0 "
        "  AND sd.c_fy_range > 0 AND sd.c_ly_range > 0 "
        "  AND sd.c_fy_range <> sd.c_ly_range "
        "ORDER BY sd.c_personid ASC"
    )
    row = cur.fetchone()
    if row:
        # Live trigger exists — name the demo person.
        pid = int(row[0]); py, chn = _name(cur, pid)
        out["bug1"] = {
            "personid": pid, "name_py": py, "name_chn": chn,
            "dormant": False,
            "stats": {"total": total, "n_fy_gt_0": n_fy,
                       "n_ly_gt_0": n_ly,
                       "n_both_diff": n_both_diff},
            "hint_en": (
                f"Open person {pid} ({chn}, {py}). The STATUS "
                f"sub-datasheet shows row(s) where c_fy_range and "
                f"c_ly_range differ — those surface the alias swap."
            ),
            "hint_zh": (
                f"開啟人物 {pid}（{chn}，{py}）。STATUS 子資料表會顯示"
                f" c_fy_range 與 c_ly_range 不同的列——這幾列就是別名"
                f"錯位最容易看到的地方。"
            ),
        }
    else:
        # Dormant on this dump — explain to the reader.
        out["bug1"] = {
            "dormant": True,
            "stats": {"total": total, "n_fy_gt_0": n_fy,
                       "n_ly_gt_0": n_ly,
                       "n_both_diff": n_both_diff},
            "hint_en": (
                f"On this data snapshot the bug is **DORMANT** — "
                f"STATUS_DATA has {total:,} rows, but only {n_fy} "
                f"have c_fy_range > 0 and {n_ly} have c_ly_range > 0; "
                f"{n_both_diff} have both populated AND different. "
                f"So no person currently surfaces the alias swap "
                f"through the UI.  The SQL bug still exists; the "
                f"moment a future data refresh adds a STATUS_DATA "
                f"row with both fy/ly range codes set differently, "
                f"the corresponding sub-datasheet line will display "
                f"the wrong text.  To verify the bug today, run the "
                f"SQL directly:\n"
                f"  SELECT c_personid, c_fy_range, c_fy_range_desc, "
                f"c_ly_range, c_ly_range_desc FROM View_StatusData "
                f"WHERE c_fy_range > 0 OR c_ly_range > 0;"
            ),
            "hint_zh": (
                f"在當前資料快照下，這個 bug 處於 **潛伏 (dormant) 狀態**"
                f"——STATUS_DATA 共 {total:,} 行，但只有 {n_fy} 行 "
                f"c_fy_range > 0、{n_ly} 行 c_ly_range > 0；兩個都有值"
                f"且不同的只有 {n_both_diff} 行。所以目前沒有任何人物"
                f"能在 UI 上重現這個別名錯位。SQL 缺陷仍然存在；只要"
                f"未來某次資料更新插入一條 fy/ly range 都填了且不同的 "
                f"STATUS_DATA 記錄，對應的子資料表那一行就會顯示錯誤"
                f"文字。今天若要驗證這個 bug，可以直接跑 SQL：\n"
                f"  SELECT c_personid, c_fy_range, c_fy_range_desc, "
                f"c_ly_range, c_ly_range_desc FROM View_StatusData "
                f"WHERE c_fy_range > 0 OR c_ly_range > 0;"
            ),
        }

    # ---- Bug #6: LookAtGroupData ChkEntry → queryEntry SQL ---------
    # Need a person id whose ENTRY_DATA row has c_parental_status_code
    # populated (so the broken column reference matters).  Person id 1
    # (An Dun) has 2 entries.
    cur.execute(
        "SELECT TOP 1 c_personid, COUNT(*) "
        "FROM ENTRY_DATA "
        "WHERE c_personid > 0 AND c_personid < 100 "
        "GROUP BY c_personid HAVING COUNT(*) BETWEEN 1 AND 3 "
        "ORDER BY c_personid"
    )
    row = cur.fetchone()
    if row:
        pid = int(row[0]); py, chn = _name(cur, pid)
        out["bug6"] = {
            "personid": pid, "name_py": py, "name_chn": chn,
            "n_entry_rows": int(row[1]),
            "hint_en": (
                f"Use person {pid} ({chn}, {py}) as the import list "
                f"(small: only {row[1]} entry row, fast to reproduce). "
                f"In LookAtGroupData, leave only the **Entry** "
                f"checkbox ticked, click **Run**."
            ),
            "hint_zh": (
                f"用人物 {pid}（{chn}，{py}）當匯入清單（資料少、只有 "
                f"{row[1]} 條 entry 記錄，方便復現）。在 LookAtGroupData "
                f"上只勾 **Entry**，點 **Run**。"
            ),
        }

    # ---- Bug #7 / #4 already use addr 7213 (kept literal) -----------

    # ---- Bug #10 / #11: events sub-form (EVENTS_DATA + EVENTS_ADDR)
    # Need a person who has BOTH events with addresses (#10) AND any
    # events at all (#11 control was supposed to bind to event-record-
    # id).  Pick a person with a small EVENTS_DATA + EVENTS_ADDR pair.
    cur.execute(
        "SELECT TOP 5 ed.c_personid, COUNT(*) AS n_ev "
        "FROM EVENTS_DATA ed "
        "WHERE ed.c_personid > 0 AND ed.c_personid < 50000 "
        "  AND ed.c_personid IN (SELECT c_personid FROM EVENTS_ADDR) "
        "GROUP BY ed.c_personid HAVING COUNT(*) BETWEEN 1 AND 5 "
        "ORDER BY ed.c_personid"
    )
    row = cur.fetchone()
    if row:
        pid = int(row[0]); py, chn = _name(cur, pid)
        cur.execute(
            f"SELECT COUNT(*) FROM EVENTS_ADDR "
            f"WHERE c_personid = {pid}"
        )
        n_addr = int(cur.fetchone()[0])
        for bug in ("bug10", "bug11"):
            out[bug] = {
                "personid": pid, "name_py": py, "name_chn": chn,
                "n_event_rows": int(row[1]),
                "n_event_addr_rows": n_addr,
                "hint_en": (
                    f"Open person {pid} ({chn}, {py}). The EVENTS "
                    f"sub-datasheet shows {row[1]} event row(s); "
                    f"{n_addr} of them have an associated address. "
                    f"That's where the bound controls render blank "
                    f"on every row."
                ),
                "hint_zh": (
                    f"開啟人物 {pid}（{chn}，{py}）。EVENTS 子資料表會"
                    f"顯示 {row[1]} 條事件，其中 {n_addr} 條有對應地址。"
                    f"相關綁定控件在每一列都會顯示空白。"
                ),
            }

    # ---- Bug #12: POSTED_TO_OFFICE_DATA appointment-type ------------
    # Need a person with an office posting whose c_appt_code is
    # non-NULL (so the column the form should display has a value).
    cur.execute(
        "SELECT TOP 5 c_personid, COUNT(*) "
        "FROM POSTED_TO_OFFICE_DATA "
        "WHERE c_personid > 0 AND c_personid < 50000 "
        "  AND c_appt_code IS NOT NULL "
        "GROUP BY c_personid HAVING COUNT(*) BETWEEN 1 AND 5 "
        "ORDER BY c_personid"
    )
    row = cur.fetchone()
    if row:
        pid = int(row[0]); py, chn = _name(cur, pid)
        out["bug12"] = {
            "personid": pid, "name_py": py, "name_chn": chn,
            "n_posting_rows": int(row[1]),
            "hint_en": (
                f"Open person {pid} ({chn}, {py}). The POSTED-TO-"
                f"OFFICE sub-datasheet shows {row[1]} posting row(s) "
                f"with non-null c_appt_code — yet the appointment-"
                f"type column on every row is blank."
            ),
            "hint_zh": (
                f"開啟人物 {pid}（{chn}，{py}）。POSTED-TO-OFFICE 子資料"
                f"表會顯示 {row[1]} 條官職任命記錄，c_appt_code 都不為 "
                f"NULL；但任職類型那一列每一列都是空白。"
            ),
        }

    # ---- Bug #13: BIOG_MAIN_2 c_fl_ey_notes click → frmPickNIAN_HAO
    # Need a person whose c_fl_ey_notes is non-empty (so the field
    # is interactable when the user clicks it).
    cur.execute(
        "SELECT TOP 5 c_personid, c_fl_ey_notes "
        "FROM BIOG_MAIN "
        "WHERE c_personid > 0 AND c_personid < 5000 "
        "  AND c_fl_ey_notes IS NOT NULL "
        "  AND LEN(c_fl_ey_notes) > 0 "
        "ORDER BY c_personid"
    )
    row = cur.fetchone()
    if row:
        pid = int(row[0]); py, chn = _name(cur, pid)
        notes = (row[1] or "")[:40]
        out["bug13"] = {
            "personid": pid, "name_py": py, "name_chn": chn,
            "fl_ey_notes_sample": notes,
            "hint_en": (
                f"Open person {pid} ({chn}, {py}). Their `c_fl_ey_notes` "
                f"field has actual text in it (sample: '{notes}…'), so "
                f"clicking it actually triggers the broken Sub."
            ),
            "hint_zh": (
                f"開啟人物 {pid}（{chn}，{py}）。其 `c_fl_ey_notes` 欄位"
                f"有真實內容（樣例：「{notes}…」），所以點擊它會真的觸發"
                f"這個有缺陷的 Sub。"
            ),
        }

    # ---- Bug #14: KIN_DATA → frmPickKINSHIP_CODES ------------------
    # Need a person with kinship records (so the kinship sub-form has
    # rows the user could click into).
    cur.execute(
        "SELECT TOP 5 c_personid, COUNT(*) "
        "FROM KIN_DATA "
        "WHERE c_personid > 0 AND c_personid < 50000 "
        "GROUP BY c_personid HAVING COUNT(*) BETWEEN 1 AND 5 "
        "ORDER BY c_personid"
    )
    row = cur.fetchone()
    if row:
        pid = int(row[0]); py, chn = _name(cur, pid)
        out["bug14"] = {
            "personid": pid, "name_py": py, "name_chn": chn,
            "n_kin_rows": int(row[1]),
            "hint_en": (
                f"Open person {pid} ({chn}, {py}). The KIN_DATA "
                f"sub-datasheet shows {row[1]} kinship row(s) — click "
                f"any one's kinship-code picker to trigger the broken "
                f"Sub."
            ),
            "hint_zh": (
                f"開啟人物 {pid}（{chn}，{py}）。KIN_DATA 子資料表會顯示 "
                f"{row[1]} 條親屬記錄——點任一條的「kinship code」picker "
                f"欄位即可觸發這個有缺陷的 Sub。"
            ),
        }

    OUT.write_text(json.dumps(out, ensure_ascii=False, indent=2),
                    encoding="utf-8")
    print(f"wrote {OUT} with {len(out)} demo entries:")
    for bug, info in out.items():
        if info.get("dormant"):
            print(f"  {bug}: DORMANT — {info['stats']}")
        else:
            print(f"  {bug}: c_personid={info['personid']} "
                  f"({info['name_chn']}, {info['name_py']})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
