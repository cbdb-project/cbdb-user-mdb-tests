"""Generate two Word documents (English + Chinese) summarising every
documented bug in CBDB_BJ_User.mdb, ordered by severity, with
screenshots from `reports/screenshots/` embedded under each issue.

Run:
    python reports/generate_report.py
Outputs:
    reports/CBDB_Issues_Report_EN.docx
    reports/CBDB_Issues_Report_ZH.docx

Tone: deferential / polite throughout — the maintainer is a
respected senior researcher, and these reports are a courtesy hand-
off, not a pull request.
"""
from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Inches, Pt, RGBColor

import opencc
_S2T = opencc.OpenCC("s2twp")  # Simplified -> Traditional (Taiwan idiom)


def t(s: str) -> str:
    """Convert Simplified Chinese strings to Traditional (Taiwan).
    Safe to call on ASCII / English (the converter is a no-op there)."""
    if not s:
        return s
    return _S2T.convert(s)

REPO = Path(__file__).resolve().parent.parent
SHOT_DIR = REPO / "reports" / "screenshots"
OUT_EN = REPO / "reports" / "CBDB_Issues_Report_EN.docx"
OUT_ZH = REPO / "reports" / "CBDB_Issues_Report_ZH-Hant.docx"
OUT_EN_MD = REPO / "reports" / "CBDB_Issues_Report_EN.md"
OUT_ZH_MD = REPO / "reports" / "CBDB_Issues_Report_ZH-Hant.md"


# ---------------------------------------------------------------------
# Issue catalogue — single source of truth for both language reports.
# Ordered by severity (highest first within each tier).
# ---------------------------------------------------------------------

ISSUES = [
    # ========== Tier 1: silent data corruption ==========
    {
        "id": 1,
        "tier": "P5_dormant_or_latent",
        "form": "View_StatusData",
        "title_en": "View_StatusData would display last-year range in the first-year column — DORMANT (no source rows trigger it on this dump)",
        "title_zh": "View_StatusData 會把首年份範圍顯示成末年份範圍 — DORMANT（當前 dump 沒有源資料能觸發）",
        "summary_en": (
            "The saved query `View_StatusData` joins `YEAR_RANGE_CODES` "
            "twice (once aliased as `YEAR_RANGE_CODES_1` for the last-year "
            "range), but the SELECT list pulls every range field from the "
            "_1 alias. As a result, every status row displayed in the "
            "Status sub-datasheet shows the last-year range value in the "
            "first-year range column."
        ),
        "summary_zh": (
            "存档查询 `View_StatusData` 把 `YEAR_RANGE_CODES` 表 JOIN 了两次"
            "（其中一次别名是 `YEAR_RANGE_CODES_1`，用于末年份范围），但"
            " SELECT 列表里所有范围字段都从 _1 别名取值。结果是 Status 子数据"
            "表里每一行显示的「首年份范围」其实是末年份范围。"
        ),
        "steps_en": [
            "Because no STATUS_DATA row in the current dump has both "
            "c_fy_range AND c_ly_range populated, this bug cannot be "
            "demonstrated through the UI today.  Verify it directly "
            "in SQL instead:",
            "Open the .mdb in Access.  Press F11 to show the "
            "navigation pane, then double-click query "
            "**View_StatusData**.",
            "Inspect the SELECT clause: every `c_fy_range_*` alias is "
            "pulled from `YEAR_RANGE_CODES_1`, but the FROM clause "
            "joins that alias on the LAST-year range.  That's the "
            "swap.",
            "(Optional) Run `SELECT TOP 100 c_personid, c_fy_range, "
            "c_fy_range_desc, c_ly_range, c_ly_range_desc FROM "
            "View_StatusData WHERE c_fy_range > 0 OR c_ly_range > 0` "
            "in the Access query window — once a future data refresh "
            "populates both fields differently, every such row will "
            "display the wrong first-year text.",
        ],
        "steps_zh": [
            "由於本 .mdb 當前快照下，沒有任何 STATUS_DATA 列同時填了 "
            "c_fy_range 和 c_ly_range，這個 bug 暫時無法在 UI 上復現。"
            "請直接用 SQL 驗證：",
            "在 Access 裡打開 .mdb，按 F11 顯示導航窗格，雙擊查詢 "
            "**View_StatusData**。",
            "查看 SELECT 子句：所有 `c_fy_range_*` 別名都從 "
            "`YEAR_RANGE_CODES_1` 取值，但 FROM 子句把這個別名 JOIN "
            "在末年份範圍上——這就是錯位。",
            "（可選）在 Access 查詢視窗執行 `SELECT TOP 100 c_personid, "
            "c_fy_range, c_fy_range_desc, c_ly_range, c_ly_range_desc "
            "FROM View_StatusData WHERE c_fy_range > 0 OR c_ly_range > 0` "
            "——未來某次資料更新如果同時填了這兩個欄位且取值不同，每一條"
            "結果都會顯示錯誤的首年份文字。",
        ],
        "fix_en": (
            "In `View_StatusData` change `YEAR_RANGE_CODES_1.c_range AS "
            "c_fy_range_desc` and `YEAR_RANGE_CODES_1.c_range_chn AS "
            "c_fy_range_chn` to use the un-aliased `YEAR_RANGE_CODES.*` "
            "fields (which the FROM clause already joins on `c_fy_range`)."
        ),
        "fix_zh": (
            "在 `View_StatusData` 中，把 `YEAR_RANGE_CODES_1.c_range AS "
            "c_fy_range_desc` 和 `YEAR_RANGE_CODES_1.c_range_chn AS "
            "c_fy_range_chn` 改成不带别名的 `YEAR_RANGE_CODES.*`（FROM "
            "子句已经按 `c_fy_range` JOIN 了它）。"
        ),
        "screenshots": [],
        "severity_en": "P5 — Dormant on this dump (would be P0 if any STATUS_DATA row had both fy/ly range codes set differently)",
        "severity_zh": "P5 — 在當前 dump 上潛伏（若任何 STATUS_DATA 列同時填了 fy/ly range 且不同，會升為 P0）",
    },
    {
        "id": 7,
        "tier": "P0_silent_data",
        "form": "Form_LookAtPlace.CmdNeo4j_Click",
        "title_en": "LookAtPlace.CmdNeo4j people-CSV silently fails on the first record",
        "title_zh": "LookAtPlace.CmdNeo4j 在写入第一条 people-CSV 时静默失败",
        "summary_en": (
            "The People-CSV section of `LookAtPlace.CmdNeo4j_Click` "
            "(line ~322 onward) builds a recordset from a SELECT that "
            "projects only four `ZZ_SCRATCH_P_TEXT` columns, but the "
            "row-write loop reads `!c_dynasty`, `!c_dynasty_chn`, and "
            "`!c_female` from that recordset. As soon as the loop hits "
            "the first row, JET raises 'Item not found in this collection'. "
            "The error handler silences it with `MsgBox`, so the user sees "
            "a single popup, then NO files are produced for any of the "
            "downstream Neo4j export steps."
        ),
        "summary_zh": (
            "`LookAtPlace.CmdNeo4j_Click` 中负责生成 People-CSV 的部分（约"
            "第 322 行起）用 SELECT 打开记录集，但 SELECT 里只投影了四个 "
            "`ZZ_SCRATCH_P_TEXT` 字段；接下来的写入循环却试着读 `!c_dynasty`、"
            "`!c_dynasty_chn`、`!c_female`。循环一碰到第一行，JET 立即报"
            "「集合中找不到项目」（Item not found in this collection）。"
            "错误处理把它弹了一个 MsgBox 就结束了，所以用户只看到一个对话框，"
            "之后整个 Neo4j 导出链下游的任何文件都不会产生。"
        ),
        "steps_en": [
            "Open **LookAtPlace**.  Pick the address picker, choose a "
            "well-attested address — for example **c_addr_id = 100658** "
            "(Kaifeng 開封; this is also the addr_id used by "
            "`tests/test_vba_inline.py`'s kaifeng-yin fixture) — so the "
            "resulting query has plenty of people to feed the People-CSV "
            "loop.  Click **Run Query**.",
            "Once the query finishes, click the **Neo4j** export button.",
            "Pick a save location at the first SaveAs prompt (the "
            "'People file' prompt).",
            "A `Run-time error 3265 — Item not found in this "
            "collection` popup appears almost immediately.",
            "After clicking OK, the chosen folder is empty — no Neo4j "
            "export file was written.",
        ],
        "steps_zh": [
            "打開 **LookAtPlace**。透過地址 picker 選一個資料量足夠的"
            "地址——例如 **c_addr_id = 100658（開封）**（這也是 "
            "`tests/test_vba_inline.py` 的 kaifeng-yin fixture 用的 "
            "addr_id）——這樣查詢結果有足夠人物餵給 People-CSV 迴圈。"
            "點 **Run Query**。",
            "等查詢跑完，點 **Neo4j** 匯出按鈕。",
            "在第一個另存對話框（「People 檔」對話框）裡選好儲存路徑。",
            "幾乎立刻彈出 `執行時錯誤 3265 ——集合中找不到項目` 對話框。",
            "點確定後，剛才選的資料夾裡一個檔案也沒有——整個 Neo4j 匯出"
            "什麼都沒寫出。",
        ],
        "fix_en": (
            "Extend the SELECT in the People-CSV branch to project the "
            "fields the loop reads, e.g. `DYNASTIES.c_dynasty`, "
            "`DYNASTIES.c_dynasty_chn`, `BIOG_MAIN.c_female` (the JOINs "
            "already expose them)."
        ),
        "fix_zh": (
            "把 People-CSV 部分的 SELECT 扩展，加入循环里读到的字段，例如 "
            "`DYNASTIES.c_dynasty`、`DYNASTIES.c_dynasty_chn`、"
            "`BIOG_MAIN.c_female`（FROM 子句里 JOIN 已经把它们暴露出来了）。"
        ),
        "screenshots": [
            ("bug7_step1_annotated.png",
             "Step 1 — open LookAtPlace, run any query, click "
             "**Neo4j**.  The CmdNeo4j button is in the bottom "
             "export-button row; reachability verified against "
             "`analysis/dump/control_inventory.json` (LookAtPlace "
             "has a `CmdNeo4j` control with the `CmdNeo4j_Click` "
             "event bound — re-checked 2026-05-03)."),
            ("bug7_step2_faux_popup.png",
             "Step 2 — the popup users see.  Reconstructed in PIL "
             "because the real popup would block the COM test "
             "driver; the error code (DAO 3265 'Item not found in "
             "this collection') and message text come from JET's "
             "documented behaviour for a recordset field that "
             "isn't in the underlying SELECT."),
        ],
        "severity_en": "P0 — Silent data corruption (export silently produces nothing)",
        "severity_zh": "P0 — 静默数据缺失（导出无声地什么都没生成）",
    },
    {
        "id": 8,
        "tier": "P0_silent_data",
        "form": "Form_LookAtNetworks.CmdNeo4j_Click",
        "title_en": "LookAtNetworks.CmdNeo4j people/place CSVs silently fail on the first record",
        "title_zh": "LookAtNetworks.CmdNeo4j 的 people/place CSV 在第一条上静默失败",
        "summary_en": (
            "Same shape as Issue #7 but on a different form. Two SELECTs "
            "in `LookAtNetworks.CmdNeo4j_Click` are missing fields that "
            "the row-write loop reads:\n\n"
            "  • `tRstPlace` SELECT (line 2458) projects 3 columns; the "
            "loop reads `!x_coord` / `!y_coord` (not projected).\n"
            "  • `tRstPeoplePlace` SELECT similarly omits `c_person_id` / "
            "`c_index_addr_id` that the loop reads.\n\n"
            "Same silent-fail symptom as Issue #7."
        ),
        "summary_zh": (
            "症状与 Issue #7 相同，只是在另一个表单上。"
            "`LookAtNetworks.CmdNeo4j_Click` 中两条 SELECT 都漏写了循环里要"
            "读的字段：\n\n"
            "  • `tRstPlace` 的 SELECT（第 2458 行）只投影 3 个字段，循环"
            "却读 `!x_coord` / `!y_coord`（没在 SELECT 里）。\n"
            "  • `tRstPeoplePlace` 的 SELECT 也漏了 `c_person_id` / "
            "`c_index_addr_id`，循环要读它们。\n\n"
            "症状与 Issue #7 完全相同：静默失败。"
        ),
        "steps_en": [
            "Open **LookAtNetworks** (note: this form has a known opening-"
            "delay issue; please allow several seconds).",
            "Run a query, then click **Neo4j**.",
            "When the export reaches the People-with-Place file, the same "
            "`Item not found` popup appears, and no further files are written.",
        ],
        "steps_zh": [
            "打开 **LookAtNetworks**（注意：这个表单已知打开会延迟，请给它几秒钟）。",
            "跑一次查询，然后点 **Neo4j**。",
            "导出走到 People-with-Place 文件那一步时，同样的「Item not found」"
            "对话框弹出来，之后的文件都不会再写了。",
        ],
        "fix_en": (
            "Extend each SELECT to project every field the loop reads. "
            "For tRstPlace: add `ADDR_CODES.x_coord`, `ADDR_CODES.y_coord`. "
            "For tRstPeoplePlace: add the missing `c_person_id` / "
            "`c_index_addr_id` columns from the joined tables."
        ),
        "fix_zh": (
            "把两条 SELECT 都扩展，加入循环里读到的字段。"
            "对 tRstPlace 加上 `ADDR_CODES.x_coord`、`ADDR_CODES.y_coord`。"
            "对 tRstPeoplePlace 加上 `c_person_id` 和 `c_index_addr_id`。"
        ),
        "screenshots": [
            ("bug8_faux_popup.png",
             "Reconstructed-in-PIL popup showing the JET 'Item not "
             "found in this collection' error users would see.  "
             "**Important caveat:** the backdrop in this image is "
             "LookAtPlace, NOT LookAtNetworks — LookAtNetworks's "
             "`Form_Open` currently hangs the COM test driver, so a "
             "real runtime view of the host form couldn't be "
             "captured.  The popup text is reconstructed from VBA "
             "static inspection of `Form_LookAtNetworks.vb:2458` / "
             "`:2475`."),
        ],
        "severity_en": "P0 — Silent data corruption",
        "severity_zh": "P0 — 静默数据缺失",
    },
    {
        "id": 9,
        "tier": "P5_dormant_or_latent",
        "form": "Form_LookAtEntry.CmdNeo4j_Click",
        "title_en": "LookAtEntry.CmdNeo4j Institutions block uses the wrong recordset variable — LATENT (gated unreachable on this dump; no ENTRY_DATA row has c_inst_code > 0)",
        "title_zh": "LookAtEntry.CmdNeo4j 的機構 (Institutions) 部分用錯了記錄集變數 — LATENT（被資料 gate 跳過、不可達；當前 dump 中沒有 c_inst_code > 0 的 ENTRY_DATA 行）",
        "summary_en": (
            "**Source-level typo, currently unreachable on this dump.**\n\n"
            "Line 1415 of `Form_LookAtEntry.vb` opens the institutions "
            "recordset as `Set tRstInstitutions = CurrentDb."
            "OpenRecordset(tQueryStr)`.  Ten lines later, line 1425 "
            "says `With tRstAssocCodes` and the loop reads "
            "`!c_inst_code`, `!c_inst_name_code`, etc. against THAT "
            "recordset — which was bound much earlier to the "
            "AssocCodes SELECT and was already `Close`d at line 1373 "
            "of the AssocCodes block.  If executed, this would raise "
            "DAO 3021 (`No current record`) on the `.MoveFirst` line; "
            "the misnamed reference is a genuine source-level bug.\n\n"
            "However, on the current dump the entire SaveAs prompt "
            "and buggy `With` block sit inside the gate "
            "`If tRecDeleted > 0 Then` at line 1389, where "
            "`tRecDeleted` is the row count of an "
            "`INSERT INTO ZZ_SCRATCH_P_TEXT … WHERE "
            "ZZ_SCRATCH_ENTRY.c_inst_code > 0`.  "
            "`ZZ_SCRATCH_ENTRY.c_inst_code` is copied verbatim from "
            "`ENTRY_DATA.c_inst_code` by CmdQuery (lines 1645-1652), "
            "and on this MDB **0 of 263,454 ENTRY_DATA rows have "
            "`c_inst_code > 0`** (also 0 with `c_inst_name_code > "
            "0`).  So `tRecDeleted = 0` for every possible "
            "LookAtEntry fixture, the gate evaluates false, the "
            "SaveAs prompt is never shown, the `With tRstAssocCodes` "
            "line is never executed, and CmdNeo4j proceeds cleanly "
            "to \"Finished saving to Neo4j\" — silently omitting "
            "`InstitutionCodes_*.csv` because there are no "
            "institution rows to write.\n\n"
            "**The missing `InstitutionCodes_*.csv` is not itself a "
            "user-visible bug** on the current dump.  Skipping an "
            "optional per-block file when its source-table count is "
            "0 is the same gating pattern the surrounding blocks use "
            "(when `c_assoc_code = 0` for a fixture, the AssocCodes "
            "block is also silently skipped — see the matching "
            "behaviour in the re-verification artifacts).  The "
            "latent typo only becomes user-visible if a future MDB "
            "drop introduces any `ENTRY_DATA` row with "
            "`c_inst_code > 0`.\n\n"
            "**Re-verification evidence:** SQL pre-image plus real "
            "Access COM probes for `c_entry_code = 36` (jinshi) and "
            "`c_entry_code = 101` (recommendation / 薦舉) confirmed "
            "no popup, chain finishes cleanly, no `InstitutionCode`-"
            "shape file in the produced set.  Details in "
            "`analysis/issue9_neo4j_institutioncodes_reverification."
            "md` and `reports/"
            "issue9_neo4j_institutioncodes_reverification.json`."
        ),
        "summary_zh": (
            "**Source-level typo，目前在此 dump 上不可達。**\n\n"
            "`Form_LookAtEntry.vb` 第 1415 行用 "
            "`Set tRstInstitutions = CurrentDb.OpenRecordset"
            "(tQueryStr)` 打開 institutions 記錄集。十行之後，"
            "第 1425 行卻寫成 `With tRstAssocCodes`，接下來迴圈讀 "
            "`!c_inst_code`、`!c_inst_name_code` 等欄位 ── 這個 "
            "recordset 早在 AssocCodes 區塊的第 1373 行就已經 "
            "`Close` 掉了。一旦真的執行到，`.MoveFirst` 那行就會"
            "丟出 DAO 3021（「No current record」）。這個錯名 "
            "reference 確實是 source-level bug。\n\n"
            "但在當前 dump，整個 SaveAs 對話框與有 bug 的 `With` "
            "區塊都包在第 1389 行的 gate `If tRecDeleted > 0 Then` "
            "裡；`tRecDeleted` 是緊鄰的 `INSERT INTO "
            "ZZ_SCRATCH_P_TEXT … WHERE "
            "ZZ_SCRATCH_ENTRY.c_inst_code > 0` 寫入列數。CmdQuery "
            "把 `ENTRY_DATA.c_inst_code` 原樣複製到 "
            "`ZZ_SCRATCH_ENTRY.c_inst_code`（第 1645-1652 行），而 "
            "**當前 MDB 263,454 筆 ENTRY_DATA 中，`c_inst_code > 0` "
            "的列數為 0**（`c_inst_name_code > 0` 也為 0）。因此任何 "
            "LookAtEntry 條件下 `tRecDeleted` 都會是 0、gate 永遠"
            "為假、SaveAs 對話框不會出現、`With tRstAssocCodes` "
            "永遠不會被執行，CmdNeo4j 會順利走完並顯示"
            "「Finished saving to Neo4j」── 只是因為沒有 institution "
            "rows，所以靜默地省略 `InstitutionCodes_*.csv`。\n\n"
            "**缺 `InstitutionCodes_*.csv` 本身在當前 dump 不是"
            "用戶可見錯誤**。當該分區資料表計數為 0 時跳過該 "
            "optional 檔，與鄰近區塊的 gating 行為完全一致（fixture "
            "若 `c_assoc_code = 0`，AssocCodes 區塊同樣會被靜默跳"
            "過 ── 詳見 re-verification artifacts 的對照行為）。"
            "只有當未來某次 MDB 更新引入任一 `c_inst_code > 0` 的 "
            "ENTRY_DATA 列，這個 latent typo 才會變成使用者可見"
            "錯誤。\n\n"
            "**Re-verification 證據：** SQL pre-image 與對 "
            "`c_entry_code = 36`（科舉：進士）、`c_entry_code = 101`"
            "（薦舉/保任）的真實 Access COM 探查均確認：無 popup、"
            "chain 順利完成、產出檔案中沒有 `InstitutionCode` 形狀"
            "的檔。細節請見 `analysis/"
            "issue9_neo4j_institutioncodes_reverification.md` 與 "
            "`reports/"
            "issue9_neo4j_institutioncodes_reverification.json`。"
        ),
        "steps_en": [
            "**On the current dump this bug cannot be triggered "
            "through the UI** — the `If tRecDeleted > 0 Then` gate "
            "at Form_LookAtEntry.vb:1389 is false for every "
            "possible LookAtEntry fixture (0 of 263,454 ENTRY_DATA "
            "rows have `c_inst_code > 0`).  Verify the source-level "
            "typo statically instead:",
            "Open `analysis/dump/vba/Form_LookAtEntry.vb` and read "
            "lines 1415-1425.  Line 1415: `Set tRstInstitutions = "
            "CurrentDb.OpenRecordset(tQueryStr)`.  Line 1425: "
            "`With tRstAssocCodes` (intended: `With "
            "tRstInstitutions`).  `tRstAssocCodes` was already "
            "`Close`d at line 1373 of the AssocCodes block, so "
            "`.MoveFirst` would raise DAO 3021.",
            "(Optional, SQL) Confirm the gate condition on the "
            "current dump: `SELECT COUNT(*) FROM ENTRY_DATA WHERE "
            "c_inst_code > 0` returns 0 (same with `c_inst_name_code "
            "> 0`).  This is what makes the buggy block unreachable.",
            "(Optional, runtime evidence) Pick `c_entry_code = 36` "
            "(jinshi) or `c_entry_code = 101` (recommendation / "
            "薦舉) on LookAtEntry → Run Query → Neo4j.  The chain "
            "finishes cleanly with `Finished saving to Neo4j`; no "
            "popup, no `InstitutionCodes_*.csv` in the output "
            "folder.  These are not user-visible errors today — they "
            "are evidence that the gate works.",
            "The source-level typo will become user-visible the "
            "first time a future MDB drop introduces any "
            "`ENTRY_DATA` row with `c_inst_code > 0`.  At that "
            "point the gate opens, the `With tRstAssocCodes` line "
            "executes against a `Close`d recordset, and DAO 3021 "
            "(`No current record`) fires on `.MoveFirst`.",
        ],
        "steps_zh": [
            "**當前 dump 此 bug 無法從 UI 觸發** ── "
            "Form_LookAtEntry.vb:1389 的 `If tRecDeleted > 0 Then` "
            "gate 對任何 LookAtEntry 條件都為假（263,454 筆 "
            "ENTRY_DATA 中 `c_inst_code > 0` 為 0）。請改用 "
            "source-level 靜態驗證：",
            "打開 `analysis/dump/vba/Form_LookAtEntry.vb`，讀第 "
            "1415-1425 行。第 1415 行寫 `Set tRstInstitutions = "
            "CurrentDb.OpenRecordset(tQueryStr)`，第 1425 行寫 "
            "`With tRstAssocCodes`（原意應為 `With "
            "tRstInstitutions`）。`tRstAssocCodes` 早在 AssocCodes "
            "區塊的第 1373 行就已經 `Close` 掉，因此 "
            "`.MoveFirst` 一旦執行即丟 DAO 3021。",
            "（可選，SQL）在當前 dump 確認 gate 條件：`SELECT "
            "COUNT(*) FROM ENTRY_DATA WHERE c_inst_code > 0` 結果"
            "為 0（`c_inst_name_code > 0` 同樣為 0）。正是這點讓有 "
            "bug 的區塊變成不可達。",
            "（可選，runtime 證據）在 LookAtEntry 選 "
            "`c_entry_code = 36`（科舉：進士）或 `c_entry_code = "
            "101`（薦舉/保任）→ Run Query → Neo4j。chain 會順利"
            "完成並顯示「Finished saving to Neo4j」；無 popup，"
            "輸出資料夾中沒有 `InstitutionCodes_*.csv`。這在今天"
            "不是使用者可見錯誤 ── 它是 gate 確實起作用的證據。",
            "Source-level typo 會在未來某次 MDB 更新引入任一 "
            "`c_inst_code > 0` 的 ENTRY_DATA 列時變成使用者可見"
            "錯誤。屆時 gate 打開，`With tRstAssocCodes` 那行對"
            "已 `Close` 的 recordset 執行，`.MoveFirst` 立即丟出 "
            "DAO 3021（「No current record」）。",
        ],
        "concrete_reproduction_en": (
            "These two `c_entry_code` values are used as "
            "investigation evidence in the re-verification — they "
            "exercise CmdQuery + CmdNeo4j end-to-end and demonstrate "
            "that the InstitutionCodes branch is gated out today.  "
            "**They are NOT a popup reproduction** — both produce "
            "`Finished saving to Neo4j` with no error and no "
            "`InstitutionCodes_*.csv`:\n\n"
            "  - `c_entry_code = 36` (jinshi / 科舉: 進士) — "
            "92,514 ENTRY_DATA rows, 0 with `c_inst_code > 0`\n"
            "  - `c_entry_code = 101` (recommendation / 薦舉) — "
            "878 ENTRY_DATA rows, 0 with `c_inst_code > 0`\n\n"
            "Re-run the evidence with `python "
            "analysis/investigate_issue9_neo4j_institutioncodes.py` "
            "(SQL only) or `… --com` (also runs real Access COM)."
        ),
        "concrete_reproduction_zh": (
            "下列兩個 `c_entry_code` 值是 re-verification 用的"
            "investigation 證據 ── 它們會走完 CmdQuery + CmdNeo4j，"
            "用來示範 InstitutionCodes branch 在當前 dump 被 gate "
            "跳過。**這兩個 fixture 不是 popup 復現**，兩者都會顯示"
            "「Finished saving to Neo4j」、無錯誤、無 "
            "`InstitutionCodes_*.csv`：\n\n"
            "  - `c_entry_code = 36`（科舉：進士）── "
            "ENTRY_DATA 92,514 筆，`c_inst_code > 0` 為 0\n"
            "  - `c_entry_code = 101`（薦舉/保任）── "
            "ENTRY_DATA 878 筆，`c_inst_code > 0` 為 0\n\n"
            "用 `python analysis/"
            "investigate_issue9_neo4j_institutioncodes.py` 重跑 "
            "SQL evidence；加 `--com` 會額外跑真實 Access COM。"
        ),
        "fix_en": (
            "Change `With tRstAssocCodes` on line 1425 to `With "
            "tRstInstitutions`.  Single-character class of fix; the "
            "underlying recordset variable was simply mis-named.  "
            "Although currently unreachable on this dump, fixing it "
            "costs nothing and prevents a future-data regression."
        ),
        "fix_zh": (
            "把第 1425 行的 `With tRstAssocCodes` 改成 "
            "`With tRstInstitutions`。屬於一字之差的筆誤，底層"
            "記錄集變數只是寫錯了。雖然目前不可達，順手修掉成本"
            "極低，也能避免未來資料一旦變動就回歸成 user-visible "
            "bug。"
        ),
        "screenshots": [],
        "severity_en": "P5 — Latent source-level typo (gated unreachable on this dump; would re-promote to P1 if any future ENTRY_DATA row has c_inst_code > 0)",
        "severity_zh": "P5 — Source-level latent typo（在當前 dump 被 gate 跳過、不可達；若未來任一 ENTRY_DATA 列出現 c_inst_code > 0，會回歸為 P1）",
    },
    {
        "id": 20,
        "tier": "P0_silent_data",
        "form": "ADDR_CODES + Form_LookAt*.CmdGIS_Click",
        "title_en": "BOM-prefixed address names can become embedded tabs and misalign GIS exports",
        "title_zh": "地址名中的 BOM 会在 GIS 导出中变成 tab，造成栏位错位",
        "summary_en": (
            "315 rows of `ADDR_CODES` carry a stray `U+FEFF` (BOM) "
            "prefix in **both** `c_name` and `c_name_chn`, almost "
            "certainly the residue of a UTF-8-with-BOM paste at "
            "data-import time. When `LookAtStatus.CmdQuery` (and the "
            "equivalent CmdQuery / CmdRun on the other LookAt forms) "
            "copies one of these rows into its scratch staging table "
            "via SQL UPDATE/INSERT, JET strips the BOM and "
            "re-interprets the remaining UTF-16 LE bytes as "
            "single-byte chars — promoting them back to Unicode but "
            "with mangled values. For `c_addr_id = 702559` (Wei Shi / "
            "尉氏) the source string `﻿尉氏` (UTF-16 bytes "
            "`ff fe 09 5c 0f 6c`) becomes the staged string "
            "`\\t\\\\\\x0fl` (UTF-16 bytes `09 00 5c 00 0f 00 6c 00`) — "
            "with a literal **TAB character at position 0**.\n\n"
            "`Form_LookAtStatus.CmdGIS_Click` (lines 1554–1636) then "
            "writes each cell as `tStr + value + tC` with `tC = "
            "Chr(9)` (line 1552) — no escaping is performed. The "
            "embedded TAB becomes a delimiter, splits AddrChn into "
            "two cells, and silently shifts every column to its "
            "right. A user opening the resulting `.tab` file in "
            "Excel sees coordinates land in the wrong column and "
            "an extra trailing column. The same `tStr + value + tC` "
            "pattern is present in the CmdGIS body of "
            "LookAtTexts / LookAtPlace / LookAtAssociations / "
            "LookAtOffice / LookAtKinship, so any LookAt form whose "
            "query happens to include one of the 315 dirty addresses "
            "reproduces the same misalignment.\n\n"
            "Evidence — full byte-level trace in "
            "`analysis/gis_status_embedded_delim_root_cause.md`; "
            "source-side scan in "
            "`reports/gis_embedded_delimiter_findings.json`; "
            "exported-file dump in "
            "`reports/gis_status_export_bytes_dump.json`. The "
            "regression test "
            "`tests/test_addr_codes_embedded_delim.py` will fail "
            "(intentionally) if the upstream data is cleaned, "
            "prompting a re-evaluation.\n\n"
            "**Known reach (PR W).** Of the 315 dirty `ADDR_CODES` "
            "rows, **only 1** (`c_addr_id = 702559` / Wei Shi 尉氏) "
            "is referenced by any person record in `BIOG_MAIN` or "
            "`BIOG_ADDR_DATA`; the other 314 are orphan rows with "
            "no person attached.  So today's user-facing surface is "
            "small: byte-confirmed in **LookAtStatus** "
            "(`c_status_code=40` fixture, the row this report was "
            "filed on); **likely reachable** in **LookAtKinship** "
            "(any of the 3 persons whose kin includes Ruan Fu) and "
            "in **LookAtPlace** (if a user picks `c_addr_id = "
            "702559`); **not currently reachable** in **LookAtTexts "
            "/ LookAtAssociations / LookAtOffice** under existing "
            "source data.  Full per-form reach analysis in "
            "`analysis/gis_embedded_delimiter_reach.md` and "
            "`reports/gis_embedded_delimiter_reach.json`.  The 314 "
            "orphan rows are a latent data-quality issue — they "
            "would reproduce the same misalignment the moment any "
            "of them gains its first person link.  Both candidate "
            "fixes above remain warranted."
        ),
        "summary_zh": (
            "`ADDR_CODES` 中有 315 行在 `c_name` **和** `c_name_chn` "
            "里都带着 `U+FEFF`（BOM）前缀，几乎可以确定是数据导入"
            "时从 UTF-8-with-BOM 文档复制粘贴留下的痕迹。"
            "当 `LookAtStatus.CmdQuery`（以及其他 LookAt 表单的对应 "
            "CmdQuery / CmdRun）把这些行通过 SQL UPDATE/INSERT 复制"
            "到自己的 scratch 暂存表时，JET 会先把 BOM 去掉，再把"
            "剩下的 UTF-16 LE 字节重新当成单字节字符——升回 Unicode "
            "之后值就被破坏了。以 `c_addr_id = 702559`（尉氏）为例，"
            "源字符串 `﻿尉氏`（UTF-16 字节 "
            "`ff fe 09 5c 0f 6c`）变成了暂存字符串 "
            "`\\t\\\\\\x0fl`（UTF-16 字节 "
            "`09 00 5c 00 0f 00 6c 00`），第 0 位上多了一个**真正的 "
            "TAB 字符**。\n\n"
            "随后 `Form_LookAtStatus.CmdGIS_Click`（第 1554–1636 行）"
            "把每个字段写成 `tStr + value + tC`，其中 `tC = Chr(9)` "
            "（第 1552 行）——完全没有做任何转义。这个嵌入的 TAB 就"
            "被当作分隔符，把 AddrChn 拆成两栏，往后所有的栏位都"
            "悄无声息地往右挪一格。用户在 Excel 里打开这份 `.tab` "
            "档，会看到坐标落在错误的栏位、还多出一个尾栏。"
            "LookAtTexts / LookAtPlace / LookAtAssociations / "
            "LookAtOffice / LookAtKinship 的 CmdGIS 都用同样的 "
            "`tStr + value + tC` 模式，所以任何 LookAt 表单只要查询"
            "结果里碰到这 315 个脏地址里的任何一个，都会重现同样的"
            "栏位错位。\n\n"
            "证据——完整的字节级追踪在 "
            "`analysis/gis_status_embedded_delim_root_cause.md`；"
            "源端扫描在 "
            "`reports/gis_embedded_delimiter_findings.json`；"
            "实际导出档的字节级 dump 在 "
            "`reports/gis_status_export_bytes_dump.json`。回归测试 "
            "`tests/test_addr_codes_embedded_delim.py` 会在上游数据"
            "被清理后**主动失败**，提醒重新评估。\n\n"
            "**已知影响面（PR W）。** 在这 315 行脏 `ADDR_CODES` "
            "里，**只有 1 行**（`c_addr_id = 702559` / 尉氏）真的"
            "被任何人物记录引用——透过 `BIOG_MAIN.c_index_addr_id` "
            "或 `BIOG_ADDR_DATA`；其余 314 行在 ADDR_CODES 表里是"
            "孤立的，没有任何人物挂上去。所以今天的用户实际"
            "影响面其实很小：在 **LookAtStatus**（`c_status_code=40` "
            "fixture，正是本 issue 立案的那一行）已有字节级实证；"
            "在 **LookAtKinship**（如果选到那 3 位以阮孚为亲属的"
            "人）和 **LookAtPlace**（如果用户选 `c_addr_id = "
            "702559`）属于「按源资料看应该会触达」；在 "
            "**LookAtTexts / LookAtAssociations / LookAtOffice** "
            "在当前源资料下根本触达不到。完整的逐表分析在 "
            "`analysis/gis_embedded_delimiter_reach.md` 与 "
            "`reports/gis_embedded_delimiter_reach.json`。其余 314 "
            "行是一个**潜伏的资料品质问题**——它们一旦有第一个人"
            "物挂上去，就会重现同样的栏位错位。前面建议的两条"
            "修法依然都值得做。"
        ),
        "steps_en": [
            "Open **LookAtStatus**. Pick the status picker and "
            "choose status code **40** (civil office / [為官者：文]) "
            "without setting any year filter — `FrameFilterYears = 1` "
            "in the test fixture.",
            "Click **Run Query**. ~17 000 rows populate the result "
            "grid.",
            "Click **GIS** with the encoding selector set to UTF-8 "
            "(`GISFrame = 1`). Save the resulting `.tab` file.",
            "Open the file in any tab-aware tool (Excel / a text "
            "editor with column rulers). Around row **11476** "
            "(corresponding to person Ruan Fu / 阮孚, "
            "`c_addr_id = 702559` / Wei Shi 尉氏) one row has 10 "
            "tab cells against the 9-column header. AddrChn is "
            "blank, X column contains text, the real X / Y values "
            "have all shifted one column to the right.",
        ],
        "steps_zh": [
            "打开 **LookAtStatus**。在 status picker 里挑 status "
            "code **40**（[為官者：文] / civil office），不要设年份"
            "过滤——测试 fixture 里 `FrameFilterYears = 1`。",
            "点 **Run Query**。结果网格里大约填进 17 000 行。",
            "点 **GIS**，把编码选成 UTF-8（`GISFrame = 1`）。把"
            "导出的 `.tab` 档存下来。",
            "在任意支援 tab 的工具（Excel / 带栏位标尺的文本编辑器）"
            "里打开这个档。第 **11476** 行附近（对应人物阮孚，"
            "`c_addr_id = 702559` / 尉氏）有一行包含 10 个 tab 栏位、"
            "却对着 9 栏的表头。AddrChn 是空的、X 栏里塞了文字，"
            "真正的 X / Y 值都往右挪了一栏。",
        ],
        "fix_en": (
            "Two complementary fixes, both worth doing:\n\n"
            "  1. **One-shot data cleanup.** Strip the leading "
            "`U+FEFF` from the 315 affected `ADDR_CODES.c_name` / "
            "`c_name_chn` rows (e.g. `UPDATE ADDR_CODES SET c_name "
            "= Mid(c_name, 2) WHERE Left(c_name, 1) = ChrW(65279)` "
            "and the parallel statement for `c_name_chn`). This "
            "removes the immediate user-visible misalignment.\n\n"
            "  2. **Defensive sanitisation in the export writers.** "
            "Before each `tStr = tStr + value + tC` append in the "
            "CmdGIS bodies of LookAtStatus / Texts / Place / "
            "Associations / Office / Kinship, replace any embedded "
            "Chr(9), Chr(10), Chr(13), Chr(11), Chr(12), or `U+FEFF` "
            "in `value` with a space. This protects the same export "
            "writers against any future similar dirty data — "
            "without it, the next tab character that creeps into "
            "`ADDR_CODES.c_name` (or `BIOG_MAIN.c_name`, or any "
            "other text field these exports touch) will reproduce "
            "the same silent misalignment.\n\n"
            "  3. **Optional pre-release audit.** A short script "
            "scanning every export-bound text column for delimiter "
            "or control characters before each release would "
            "catch this class of dirty-data issue before it ships. "
            "`analysis/probe_status_gis_embedded_delim.py` is a "
            "concrete starting point."
        ),
        "fix_zh": (
            "两条互补的修法，建议都做：\n\n"
            "  1. **一次性数据清理。** 把这 315 行 `ADDR_CODES.c_name` / "
            "`c_name_chn` 开头的 `U+FEFF` 去掉（例如 "
            "`UPDATE ADDR_CODES SET c_name = Mid(c_name, 2) "
            "WHERE Left(c_name, 1) = ChrW(65279)`，再对 "
            "`c_name_chn` 重复一遍）。这一步可以立即解决用户能看到的"
            "栏位错位。\n\n"
            "  2. **导出端做防御性 sanitisation。** 在 LookAtStatus / "
            "Texts / Place / Associations / Office / Kinship 各自"
            "的 CmdGIS 里，每一个 `tStr = tStr + value + tC` 之前，"
            "先把 `value` 里的 Chr(9)、Chr(10)、Chr(13)、Chr(11)、"
            "Chr(12)、`U+FEFF` 全部替换成空格。这样以后任何 text "
            "字段如果再混进类似的脏字符，导出依然能保持栏位对齐"
            "——少了这一层，下一次只要 `ADDR_CODES.c_name`（或 "
            "`BIOG_MAIN.c_name`、或其他这些导出会碰到的 text 字段）"
            "里悄悄塞进一个 tab 字符，又会重现完全一样的静默错位。\n\n"
            "  3. **建议增加一个发布前的检查脚本。** 写一个简短的"
            "脚本，发布前扫描所有会被导出的 text 栏位，看里面有没有"
            "分隔符或控制字符，可以在每次发布前提前抓到这一类脏"
            "数据问题。`analysis/probe_status_gis_embedded_delim.py` "
            "是一个现成的起点。"
        ),
        "screenshots": [],
        "severity_en": "P0 — Silent export column misalignment (numeric fields land in text columns; values shifted by one and one extra trailing column appears)",
        "severity_zh": "P0 — 静默导出栏位错位（数字字段落到文本栏，所有栏位向右挪一格，结尾多出一栏）",
    },
    # ========== Tier 2: visible runtime crash (popup blocks user) ==========
    {
        "id": 4,
        "tier": "P5_dormant_or_latent",
        "form": "Form_LookAtPlace.CmdGIS_Click",
        "title_en": "LookAtPlace.CmdGIS would abort with 'Object required' — LATENT, masked by Issue #15 (no CmdGIS button on the form)",
        "title_zh": "LookAtPlace.CmdGIS 會報「Object required」 — LATENT，被 Issue #15（表單上沒有 CmdGIS 按鈕）所遮蔽",
        "summary_en": (
            "Note: this issue is moot in the current dump because there "
            "is no CmdGIS button on LookAtPlace's design (Issue #15) — "
            "users physically cannot click it. But the underlying VBA "
            "problem remains: line 1539 of `Form_LookAtPlace.vb` reads "
            "`If GISFrame.Value = 1 Then`, and there is no control named "
            "`GISFrame` on this form (the actual encoding control is "
            "`CodeFrame`). If the missing button is ever re-added (Issue "
            "#15) without first fixing this line, every click will throw."
        ),
        "summary_zh": (
            "说明：在当前 .mdb 上这个问题暂时不会被用户触发，因为 LookAtPlace"
            "的设计视图里根本没有 CmdGIS 按钮（即 Issue #15）——用户无法点击。"
            "但底层 VBA 问题依然存在：`Form_LookAtPlace.vb` 第 1539 行写的是 "
            "`If GISFrame.Value = 1 Then`，而该表单上根本没有 `GISFrame` "
            "控件（真正的编码选择控件叫 `CodeFrame`）。一旦 Issue #15 "
            "里把缺失的按钮加回去而没先修这一行，每一次点击都会抛错。"
        ),
        "steps_en": [
            "(Hypothetical, after Issue #15 is fixed.) Open **LookAtPlace**.",
            "Run any query.",
            "Click the GIS button.",
            "A `Run-time error 424 — Object required` popup appears, "
            "the export does nothing.",
        ],
        "steps_zh": [
            "（在 Issue #15 修好之后才能复现）打开 **LookAtPlace**。",
            "跑任意一次查询。",
            "点 GIS 按钮。",
            "弹出 `运行时错误 424 ——必要的对象（Object required）` 对话框，"
            "导出什么都没做。",
        ],
        "fix_en": (
            "Change `GISFrame.Value` to `CodeFrame.Value` on line 1539 "
            "of `Form_LookAtPlace.vb`. Same change `CmdNeo4j_Click`, "
            "`CmdGephi_Click`, and `CmdPajek_Click` on the same form "
            "already use correctly."
        ),
        "fix_zh": (
            "把 `Form_LookAtPlace.vb` 第 1539 行的 `GISFrame.Value` 改成"
            " `CodeFrame.Value`。同表单的 `CmdNeo4j_Click`、`CmdGephi_Click`、"
            "`CmdPajek_Click` 已经写对了，可以参考。"
        ),
        "screenshots": [
            ("bug4_step3_faux_popup.png",
             "**Hypothetical** popup, reconstructed in PIL.  Users "
             "currently CAN'T trigger this — Bug #15 means the "
             "CmdGIS button does not exist on LookAtPlace, so the "
             "click that would fire `CmdGIS_Click` (and produce "
             "this 'Object required' error) has nowhere to come "
             "from.  This image shows what the user would see if a "
             "future change restored the CmdGIS button without "
             "first fixing the GISFrame → CodeFrame typo on line "
             "1539.  The earlier bug4 step1 / step2 runtime "
             "screenshots were misleading (their annotations "
             "implied a clickable GIS button) and were removed in "
             "PR C — only this faux popup is kept as latent-state "
             "evidence."),
        ],
        "severity_en": "P5 — Latent (would be P1 if Issue #15 fixed without first fixing this)",
        "severity_zh": "P5 — 潛伏（若先修了 Issue #15 而沒同時修本條，會變成 P1）",
    },
    {
        "id": 5,
        "tier": "P5_dormant_or_latent",
        "form": "Form_LookAtStatus.CmdPajek_Click",
        "title_en": "LookAtStatus.CmdPajek references a missing control AND uses three columns that don't exist",
        "title_zh": "LookAtStatus.CmdPajek 引用了不存在的控件，且 SQL 用了三个不存在的列",
        "summary_en": (
            "Two related defects in the same handler:\n\n"
            "  (a) Line 2308 reads `If ChkIDs.Value Then`, but Status has "
            "no control named `ChkIDs` — only `ChkXYRef`, `ChkKML`, and "
            "`ChkSubUnits`.\n\n"
            "  (b) Lines 2335–2338 build a SELECT that references "
            "`ZZ_SCRATCH_STATUS.c_person_id`, `c_status_id`, and "
            "`c_status_count` — none of which exist in the schema (the "
            "real columns are `c_personid`, `c_status_code`, no count "
            "column at all).\n\n"
            "The whole sub looks copy-pasted from "
            "`LookAtAssociations.CmdPajek_Click` where these names ARE "
            "valid; the rename pass missed both spots. Like Issue #4 this "
            "is also somewhat moot because LookAtStatus has no Pajek "
            "button (Issue #16); the SQL still fails the moment the sub "
            "is invoked though, so adding the button without fixing the "
            "SQL would just expose the failure to users."
        ),
        "summary_zh": (
            "同一个 handler 里有两个相关缺陷：\n\n"
            "  (a) 第 2308 行写 `If ChkIDs.Value Then`，但 Status 上没有名为"
            " `ChkIDs` 的控件——只有 `ChkXYRef`、`ChkKML`、`ChkSubUnits`。\n\n"
            "  (b) 第 2335–2338 行构造的 SELECT 引用 "
            "`ZZ_SCRATCH_STATUS.c_person_id`、`c_status_id`、`c_status_count`"
            "——这三列都不在 schema 里（真实列名是 `c_personid`、"
            "`c_status_code`，count 列根本没有）。\n\n"
            "整段 sub 看起来是从 `LookAtAssociations.CmdPajek_Click` 整段"
            "拷过来的，那边列名都对得上；改名时这两处都漏了。和 Issue #4 一样，"
            "因为 LookAtStatus 当前也没有 Pajek 按钮（Issue #16），用户暂时碰"
            "不到；但只要按钮加回去而没先修这两处，用户就会立刻看到错误。"
        ),
        "steps_en": [
            "(Hypothetical, after Issue #16 is fixed.) Open **LookAtStatus**.",
            "Run a query, then click the Pajek button.",
            "First: an `Object required` popup appears (the ChkIDs reference).",
            "If that's worked around, the next click hits the SQL: "
            "a `No such field` error from the SELECT that references three "
            "non-existent columns.",
        ],
        "steps_zh": [
            "（在 Issue #16 修好之后才能复现）打开 **LookAtStatus**。",
            "跑一次查询，然后点 Pajek 按钮。",
            "第一次会弹 `Object required`（ChkIDs 引用所致）。",
            "如果绕过它，下一次点就会触发 SQL：因为 SELECT 引用了三个不存在的"
            "列，会报 `No such field` 之类的错误。",
        ],
        "fix_en": (
            "Two fixes:\n"
            "  (a) Replace `ChkIDs.Value` with either a constant `False` "
            "(if the optional behaviour isn't needed) or add a real "
            "ChkIDs control to LookAtStatus's design.\n"
            "  (b) Rewrite the SELECT to use `ZZ_SCRATCH_STATUS.c_personid` "
            "and `ZZ_SCRATCH_STATUS.c_status_code`, and either drop the "
            "count aggregate or compute it some other way (the source "
            "table doesn't have `c_status_count`).\n\n"
            "Realistically the whole sub probably needs a thoughtful "
            "rewrite rather than spot fixes — it was clearly inherited "
            "from another form without verification."
        ),
        "fix_zh": (
            "两处都要改：\n"
            "  (a) 把 `ChkIDs.Value` 替换成常量 `False`（如果这个可选行为可以"
            "去掉），或者在 LookAtStatus 的设计视图里真的加一个 ChkIDs 控件。\n"
            "  (b) 把 SELECT 改成 `ZZ_SCRATCH_STATUS.c_personid` 和 "
            "`ZZ_SCRATCH_STATUS.c_status_code`，并去掉对 `c_status_count` 的"
            "聚合，或用别的方式计算（源表里就没有 c_status_count 列）。\n\n"
            "建议整段 sub 通盘重写而不是单点修补——它显然是从另一个表单整段"
            "拷贝过来的，列名没校对过。"
        ),
        "screenshots": [],
        "severity_en": "P5 — Latent (would be P1 if Issue #16 fixed without first fixing this)",
        "severity_zh": "P5 — 潛伏（若先修了 Issue #16 而沒同時修本條，會變成 P1）",
    },
    {
        "id": 6,
        "tier": "P1_visible_crash",
        "form": "Form_LookAtGroupData.queryEntry",
        "title_en": "LookAtGroupData ChkEntry path references a non-existent column ENTRY_DATA.c_parental_status",
        "title_zh": "LookAtGroupData 的 ChkEntry 路径引用了不存在的列 ENTRY_DATA.c_parental_status",
        "summary_en": (
            "`Form_LookAtGroupData.vb` line 2621 has an INSERT INTO whose "
            "target column list ends with `c_parental_status_code` but "
            "whose SELECT projection ends with `ENTRY_DATA.c_parental_status` "
            "(no `_code` suffix). The actual column on `ENTRY_DATA` is "
            "`c_parental_status_code`; the typo means the SQL crashes with "
            "'No such field' the moment the user checks **Entry** and "
            "clicks **Run**.\n\n"
            "`Form_LookAtEntry.vb:1650` does the same logical query and "
            "uses the correct name, so this is a single-line drift."
        ),
        "summary_zh": (
            "`Form_LookAtGroupData.vb` 第 2621 行的 INSERT INTO 目标列里写的"
            "是 `c_parental_status_code`，但 SELECT 投影写的是 "
            "`ENTRY_DATA.c_parental_status`（少了 `_code` 后缀）。`ENTRY_DATA` "
            "上真实列名是 `c_parental_status_code`；这个笔误让用户一旦勾上"
            " **Entry** 子类型再点 **Run**，SQL 就会报「无此字段」。\n\n"
            "`Form_LookAtEntry.vb:1650` 写的是同一段逻辑查询而且名字是对的，"
            "可以参考。"
        ),
        "steps_en": [
            "In **LookAtGroupData**, populate the import list with one "
            "entry — c_personid = 1 (An Dun 安惇) is enough; he has "
            "exactly 2 ENTRY_DATA rows so the broken queryEntry SQL "
            "will run on a tiny well-known sample.",
            "Tick **only** the **Entry** checkbox (leave Status / "
            "Office / Text / Addr unchecked so the unrelated query "
            "branches don't fire).",
            "Click **Run**.",
            "A popup appears reporting that a field doesn't exist "
            "(JET reports this as 'No value given for one or more "
            "required parameters' / 'No such field' depending on the "
            "Office build — both mean the SQL referenced "
            "`ENTRY_DATA.c_parental_status` which doesn't exist).",
        ],
        "steps_zh": [
            "在 **LookAtGroupData** 上把匯入清單設為一個人——例如 "
            "c_personid = 1（安惇 An Dun），他只有 2 條 ENTRY_DATA "
            "記錄，足以讓有缺陷的 queryEntry SQL 在一個小而熟知的"
            "樣本上跑起來。",
            "**只**勾 **Entry** 複選框（Status / Office / Text / "
            "Addr 都不勾，避免無關的查詢分支干擾）。",
            "點 **Run**。",
            "彈出「欄位不存在」之類的對話框（JET 在不同 Office 版本"
            "下給出的措辭是「沒有為一個或多個必要參數提供值」或「沒有"
            "此欄位」——都是因為 SQL 引用了根本不存在的 "
            "`ENTRY_DATA.c_parental_status`）。",
        ],
        "fix_en": (
            "Change `ENTRY_DATA.c_parental_status` to "
            "`ENTRY_DATA.c_parental_status_code` on line 2621. One-line fix."
        ),
        "fix_zh": (
            "把第 2621 行的 `ENTRY_DATA.c_parental_status` 改成 "
            "`ENTRY_DATA.c_parental_status_code`。一行修复。"
        ),
        "screenshots": [
            ("bug6_form_annotated.png",
             "Step 1 — open LookAtGroupData, leave only the Entry "
             "checkbox ticked, click Run.  (Demo input from "
             "`reports/demo_persons.json`: import list = "
             "c_personid 1, 安惇.)"),
            ("bug6_faux_popup.png",
             "Step 2 — the JET error popup users see.  The popup "
             "graphic is reconstructed in PIL because the real popup "
             "would block the COM test driver; the error code (3061) "
             "and message text come from JET's documented behaviour "
             "for unknown-identifier-as-parameter on the line cited "
             "in the summary."),
        ],
        "severity_en": "P1 — Visible crash on a common path (Entry sub-query)",
        "severity_zh": "P1 — 常用路径上的可见报错（Entry 子查询）",
    },
    {
        "id": 13,
        "tier": "P1_visible_crash",
        "form": "Form_BIOG_MAIN_2_Subform.c_fl_ey_notes_Click",
        "title_en": "BIOG_MAIN_2 Subform tries to open a picker form (frmPickNIAN_HAO) that doesn't exist",
        "title_zh": "BIOG_MAIN_2 子表单试图打开一个不存在的 picker 表单 (frmPickNIAN_HAO)",
        "summary_en": (
            "When the user clicks the `c_fl_ey_notes` field on a person's "
            "biographical detail subform, `Sub c_fl_ey_notes_Click` runs "
            "`DoCmd.OpenForm \"frmPickNIAN_HAO\"`. There is no form named "
            "`frmPickNIAN_HAO` in the .mdb's CurrentProject.AllForms "
            "collection. Access raises 'Item not found …' and the field "
            "click does nothing useful for the user.\n\n"
            "Likely cause: a picker form was renamed or consolidated in an "
            "earlier refactor, and this caller wasn't updated."
        ),
        "summary_zh": (
            "用户在某位人物生平详情子表单上点击 `c_fl_ey_notes` 字段时，"
            "`Sub c_fl_ey_notes_Click` 会调用 `DoCmd.OpenForm "
            "\"frmPickNIAN_HAO\"`。.mdb 的 CurrentProject.AllForms 集合中"
            "并没有名为 `frmPickNIAN_HAO` 的表单。Access 报「集合中找不到"
            "项目」，用户的这一次点击就此无效。\n\n"
            "可能原因：picker 表单在某次重构中被重命名或合并了，而这个调用"
            "处没有跟着更新。"
        ),
        "steps_en": [
            "Open the biographical detail form for **c_personid = 5 "
            "(Zha Yue 查籥)** — picked because his BIOG_MAIN row has a "
            "non-empty `c_fl_ey_notes` value, so the field is "
            "interactable (clicking an empty field doesn't fire the "
            "Sub).",
            "On the BIOG_MAIN_2 subform, click the `c_fl_ey_notes` "
            "field — that fires the `c_fl_ey_notes_Click` Sub.",
            "An `Item not found in this collection.` popup appears "
            "(because the Sub tries `DoCmd.OpenForm \"frmPickNIAN_HAO\"` "
            "and that form doesn't exist).",
        ],
        "steps_zh": [
            "打開人物 **c_personid = 5（查籥 Zha Yue）** 的生平詳情"
            "——之所以選他，是因為其 BIOG_MAIN 上 `c_fl_ey_notes` 欄位"
            "有實際內容，欄位可點（點一個空欄位不會觸發這個 Sub）。",
            "在 BIOG_MAIN_2 子表單上點 `c_fl_ey_notes` 欄位——這會觸發"
            " `c_fl_ey_notes_Click` Sub。",
            "彈出「集合中找不到項目」對話框（因為 Sub 試圖 "
            "`DoCmd.OpenForm \"frmPickNIAN_HAO\"`，而該表單根本不存在）。",
        ],
        "fix_en": (
            "Either restore the picker form `frmPickNIAN_HAO`, or update "
            "the caller in `Form_BIOG_MAIN_2_Subform.c_fl_ey_notes_Click` "
            "to use whichever picker form replaced it."
        ),
        "fix_zh": (
            "要么把 `frmPickNIAN_HAO` 表单恢复回来，要么在 "
            "`Form_BIOG_MAIN_2_Subform.c_fl_ey_notes_Click` 里把调用改成"
            "替代的那个 picker 表单。"
        ),
        "screenshots": [
            ("bug13_browser_annotated.png",
             "Step 1 — open CBDB_Browser_2, navigate to c_personid=5 "
             "(查籥 Zha Yue), click the `c_fl_ey_notes` field on the "
             "Birth/Death sub-tab (this fires `c_fl_ey_notes_Click`)."),
            ("bug13_faux_popup.png",
             "Step 2 — the popup users see.  Reconstructed in PIL "
             "because the real popup would block the COM test "
             "driver; error 2102 + 'misspelled or refers to a form "
             "that doesn't exist' is Access's standard message when "
             "DoCmd.OpenForm targets a form not in "
             "CurrentProject.AllForms."),
        ],
        "severity_en": "P1 — Visible crash on a user click",
        "severity_zh": "P1 — 用户点击时可见的报错",
    },
    {
        "id": 21,
        "tier": "P1_visible_crash",
        "form": "Form_LookAtGroupData.CmdNeo4j_Click",
        "title_en": "LookAtGroupData.CmdNeo4j crashes with 'No current record' on empty sections",
        "title_zh": "LookAtGroupData.CmdNeo4j 在汇出空分部时崩溃报「No current record」",
        "summary_en": (
            "All 11 CSV export blocks in `Form_LookAtGroupData."
            "CmdNeo4j_Click` open a scratch recordset and "
            "immediately call `.MoveFirst` without first checking "
            "if the recordset is empty (no `.EOF` or "
            "`.RecordCount > 0` guard).  When a user queries a "
            "group whose data has no rows in a given category — "
            "for example, no Entry data so `ZZ_SCRATCH_ENTRY` is "
            "empty — the `.MoveFirst` call instantly raises DAO "
            "3021 'No current record', a popup blocks the user, "
            "and the entire Neo4j export chain aborts.\n\n"
            "On the current dump and a typical small fixture the "
            "**first user-reachable failure is block #9 "
            "PeopleEntry** (line 1243-1245) — `Set tRstPeopleEntry "
            "= CurrentDb.OpenRecordset(\"ZZ_SCRATCH_ENTRY\", "
            "dbOpenDynaset)` followed unguarded by `.MoveFirst`.  "
            "The same missing-guard pattern also exists in the "
            "immediately-following **block #10 EntryCode** (line "
            "1383-1385) and in other ungated tail blocks; "
            "block #10 doesn't currently surface as a separate "
            "symptom only because the chain bails at block #9 "
            "first.\n\n"
            "Note on broader code scope: blocks #1-#8 share the "
            "same `Set ... = OpenRecordset(...)` followed by "
            "`.MoveFirst` shape, but their feeder scratch tables "
            "(ZZ_SCRATCH_STATUS, ZZ_SCRATCH_OFFICE, ZZ_PLACE, "
            "ZZ_SCRATCH_P_TEXT, ZZ_ADDRESSES) are non-empty "
            "under any normal user enable scope, so the empty-"
            "feeder failure mode is NOT user-reachable on those "
            "blocks today.  Block #11 InstitutionCodes (line "
            "1485-1487) is correctly gated by `If tRecDeleted > "
            "0 Then` upstream and is NOT part of the bug.  Only "
            "blocks #9 and #10 need a guard added to fix the "
            "user-reachable symptom.\n\n"
            "This is **distinct from Issue #6**: Issue #6 is a "
            "column-typo (`ENTRY_DATA.c_parental_status` vs "
            "`c_parental_status_code`) in `queryEntry` that "
            "prevents `ZZ_SCRATCH_ENTRY` from being populated at "
            "all when ChkEntry is on.  Issue #21 is a separate "
            "downstream missing-guard bug in `CmdNeo4j_Click` "
            "that fires whenever `ZZ_SCRATCH_ENTRY` happens to "
            "be empty — including via the upstream Issue #6 "
            "path, but also independently when the user simply "
            "doesn't tick ChkEntry.  Two different code-level "
            "defects; should be filed and fixed separately."
        ),
        "summary_zh": (
            "`Form_LookAtGroupData.CmdNeo4j_Click` 的多个 CSV 汇出"
            "区块在打开暂存记录集后直接呼叫 `.MoveFirst`，没有先"
            "检查记录集是否为空（缺少 `.EOF` 或 `.RecordCount > "
            "0` 防护）。如果用户查询的群体在某个类别没有数据"
            "（例如没有 Entry 数据，导致 `ZZ_SCRATCH_ENTRY` 为空"
            "），`.MoveFirst` 呼叫会立刻抛出 DAO 3021「No current "
            "record」，弹出报错框并中断整个 Neo4j 汇出。\n\n"
            "在当前 dump 和典型小 fixture 下，**用户能首先触发"
            "的失败是 block #9 PeopleEntry**（line 1243-1245）"
            "—— `Set tRstPeopleEntry = CurrentDb.OpenRecordset("
            "\"ZZ_SCRATCH_ENTRY\", dbOpenDynaset)` 后紧接无防护"
            "的 `.MoveFirst`。同样的缺少防护模式也存在于紧随其"
            "后的 **block #10 EntryCode**（line 1383-1385）以及"
            "其他没有上游 gate 的 tail block；block #10 之所以"
            "目前没单独表现为故障，只是因为链条已经在 block #9 "
            "中断了。\n\n"
            "关于代码范围的说明：block #1-#8 共用相同的 `Set "
            "... = OpenRecordset(...)` + `.MoveFirst` 写法，但"
            "其上游暂存表（ZZ_SCRATCH_STATUS, ZZ_SCRATCH_OFFICE,"
            " ZZ_PLACE, ZZ_SCRATCH_P_TEXT, ZZ_ADDRESSES）在任何"
            "正常用户启用范围下都不会为空，所以 empty-feeder 失"
            "败模式 在这些 block 上 用户无法触达。Block #11 "
            "InstitutionCodes（line 1485-1487）在上游已经被 "
            "`If tRecDeleted > 0 Then` 正确 gate 住，不属于本 "
            "issue 范围。要修复用户可见的症状，只需在 block #9 "
            "和 #10 加 guard。\n\n"
            "这与 Issue #6 **不同**：Issue #6 是 `queryEntry` 里"
            "的列名笔误（`ENTRY_DATA.c_parental_status` 应为 "
            "`c_parental_status_code`），导致 ChkEntry 勾选时 "
            "`ZZ_SCRATCH_ENTRY` 写不进数据。Issue #21 是 "
            "`CmdNeo4j_Click` 里独立的下游缺少防护的 bug，只要 "
            "`ZZ_SCRATCH_ENTRY` 为空就触发 —— 既可能是上游 "
            "Issue #6 的连带影响，也可能是用户单纯没勾 ChkEntry。"
            "两个不同层级的代码缺陷，应分别归档与修复。"
        ),
        "steps_en": [
            "In **LookAtGroupData**, populate the import list "
            "with c_personid = 1 (An Dun 安惇) — he has 2 "
            "STATUS_DATA / 2 ENTRY_DATA / ~12 POSTED_TO_OFFICE "
            "rows so the chain has substantive feeder data for "
            "Status / Office but Entry data won't appear in "
            "ZZ_SCRATCH_ENTRY unless ChkEntry is ticked (which "
            "would also trigger Issue #6 separately).",
            "Tick **Status**, **Office**, **Addr** and the "
            "matching **GIS** sister checkboxes — but leave "
            "**Entry** unticked (ZZ_SCRATCH_ENTRY stays empty).  "
            "Click **Run**.",
            "After CmdRun finishes (ZZ_SCRATCH_STATUS gets 2 "
            "rows, ZZ_SCRATCH_OFFICE gets 12), click the "
            "**Neo4j** export button.",
            "The chain produces 8 CSVs cleanly (People, Places, "
            "PeoplePlaces, PersonPlaceCodes, PeopleStatus, "
            "StatusCode, PeopleOffice, OfficeCodes), then a "
            "`Run-time error 3021 — No current record` popup "
            "appears when the chain reaches block #9 "
            "(PeopleEntry).  The remaining 2-3 expected files "
            "(PeopleEntry, EntryCode, optional InstitutionCodes) "
            "are never written.",
            "Verified end-to-end via probe at "
            "`analysis/probe_groupdata_cmdneo4j.py` and "
            "`analysis/probe_groupdata_cmdneo4j_tail.py` — the "
            "tail probe's iter 3 split-then-seed iteration "
            "manually inserts one row into ZZ_SCRATCH_ENTRY, at "
            "which point the chain produces 10 files with no "
            "ERR (proves the trigger is the empty source "
            "recordset, not anything else).",
        ],
        "steps_zh": [
            "在 **LookAtGroupData** 上把汇入清单设为 c_personid "
            "= 1（安惇 An Dun）——他有 2 条 STATUS_DATA / 2 条 "
            "ENTRY_DATA / 约 12 条 POSTED_TO_OFFICE，链条针对 "
            "Status / Office 有实际数据，但只要不勾 ChkEntry，"
            "`ZZ_SCRATCH_ENTRY` 就会保持为空（勾上 ChkEntry 会"
            "另外触发 Issue #6）。",
            "勾 **Status**、**Office**、**Addr** 及对应的 "
            "**GIS** 三个子项——**Entry 不要勾**（让 "
            "ZZ_SCRATCH_ENTRY 维持为空）。点 **Run**。",
            "CmdRun 完成后（ZZ_SCRATCH_STATUS 写入 2 行，"
            "ZZ_SCRATCH_OFFICE 写入 12 行），点选 **Neo4j** "
            "汇出按钮。",
            "链条会先顺利产出 8 份 CSV（People / Places / "
            "PeoplePlaces / PersonPlaceCodes / PeopleStatus / "
            "StatusCode / PeopleOffice / OfficeCodes），然后在"
            "进入第 9 个 block（PeopleEntry）时弹出 `执行时错误"
            " 3021 —— No current record` 对话框。余下的 2-3 份"
            "预期档案（PeopleEntry / EntryCode / 可选的 "
            "InstitutionCodes）不会写出。",
            "已在 `analysis/probe_groupdata_cmdneo4j.py` 与 "
            "`analysis/probe_groupdata_cmdneo4j_tail.py` 端到端"
            "验证 —— tail probe 的 iter 3 split-then-seed 手动"
            "向 ZZ_SCRATCH_ENTRY 插入一行后，链条产出 10 份档案"
            "且没有 ERR（证明触发条件就是空记录集，没有其他变量）。",
        ],
        "fix_en": (
            "Add an `.EOF` (or `.RecordCount > 0`) guard before "
            "the `.MoveFirst` call in **block #9 PeopleEntry "
            "(line ~1245) and block #10 EntryCode (line ~1385)** "
            "of `Form_LookAtGroupData.CmdNeo4j_Click`.  These are "
            "the two user-reachable blocks; the other tail blocks "
            "either have an upstream gate (block #11 "
            "InstitutionCodes — `If tRecDeleted > 0 Then`) or "
            "their feeder scratch tables are non-empty under any "
            "normal user enable scope (blocks #1-#8).  Suggested "
            "shape:\n\n"
            "```vb\n"
            "Set tRstPeopleEntry = CurrentDb.OpenRecordset("
            "\"ZZ_SCRATCH_ENTRY\", dbOpenDynaset)\n"
            "With tRstPeopleEntry\n"
            "    If Not .EOF Then\n"
            "        .MoveFirst\n"
            "        Do While Not .EOF\n"
            "            ' ... existing per-row write ...\n"
            "            .MoveNext\n"
            "        Loop\n"
            "    End If\n"
            "End With\n"
            "```\n\n"
            "Defensive scope option (not required to close this "
            "issue): adding the same guard to blocks #1-#8 is "
            "harmless and would future-proof against a new enable "
            "path leaving any of those feeder tables empty.  Not "
            "required because no such path exists today."
        ),
        "fix_zh": (
            "在 `Form_LookAtGroupData.CmdNeo4j_Click` 的 **block "
            "#9 PeopleEntry（约 line 1245）和 block #10 "
            "EntryCode（约 line 1385）** 上，于 `.MoveFirst` "
            "呼叫前加上 `.EOF`（或 `.RecordCount > 0`）防护。"
            "这两段是用户能触发的两块；其他 tail block 要么上游"
            "有 gate（block #11 InstitutionCodes 上游有 "
            "`If tRecDeleted > 0 Then`），要么其上游暂存表在任"
            "何正常用户启用范围下都非空（block #1-#8）。建议写法"
            "：\n\n"
            "```vb\n"
            "Set tRstPeopleEntry = CurrentDb.OpenRecordset("
            "\"ZZ_SCRATCH_ENTRY\", dbOpenDynaset)\n"
            "With tRstPeopleEntry\n"
            "    If Not .EOF Then\n"
            "        .MoveFirst\n"
            "        Do While Not .EOF\n"
            "            ' ... 原有的逐行写出 ...\n"
            "            .MoveNext\n"
            "        Loop\n"
            "    End If\n"
            "End With\n"
            "```\n\n"
            "防御性范围选项（关闭本 issue 不需要）：把同样的 "
            "guard 加到 block #1-#8 也无害，可对未来出现新的"
            "启用路径让其上游表为空的情况做防御。当前没有这种"
            "路径，所以不必做。"
        ),
        "screenshots": [],
        "severity_en": (
            "P1 — Visible crash on a normal user click "
            "(any GroupData CmdNeo4j export where Entry data is "
            "absent for the queried person, which is the common "
            "case for figures with sparse entry records)"
        ),
        "severity_zh": (
            "P1 — 正常用户点击下的可见报错（只要查询的人没有 "
            "Entry 资料，GroupData 的 Neo4j 汇出就会触发 —— "
            "这对入仕记录稀薄的人物属于常见情况）"
        ),
    },
    {
        "id": 22,
        "tier": "P1_visible_crash",
        "form": "Form_LookAtAssociations.CmdUCINet_Click",
        "title_en": "LookAtAssociations.CmdUCINet crashes with 'Invalid procedure call or argument' on networks containing CJK Han characters in c_name",
        "title_zh": "LookAtAssociations.CmdUCINet 在被汇出的人物网络含有 c_name 中带 CJK 汉字时崩溃报「Invalid procedure call or argument」",
        "summary_en": (
            "`Form_LookAtAssociations.CmdUCINet_Click` writes "
            "the `.vna` export via "
            "`Scripting.FileSystemObject.CreateTextFile"
            "(tFileName, True)` (line ~2575).  The 3rd "
            "argument (`Unicode`) is omitted, so it defaults "
            "to FALSE — the file is opened in the system "
            "default ANSI code page (cp1252 on en-US "
            "Windows).\n\n"
            "Inside the `*node properties` section, the body "
            "writes `tQuote + !c_name + tQuote` for each "
            "row.  When `c_name` contains a character that "
            "has no cp1252 representation AND no FSO "
            "substitution mapping (CJK Han ideographs in "
            "particular — e.g. U+7A1C 稜), `tVNA.WriteLine` "
            "raises VBA error 5 ('Invalid procedure call or "
            "argument') and the whole CmdUCINet export "
            "aborts.  The partial `.vna` file is left on "
            "disk — `*node data` complete, `*node "
            "properties` truncated, `*tie data` never "
            "written.\n\n"
            "User-visible symptom: a Run-time error 5 popup "
            "blocks the user; the exported `.vna` file is "
            "incomplete and unusable in UCINET / Visone.\n\n"
            "**Affected forms (per current evidence on this "
            "dump):**\n\n"
            "- **LookAtAssociations** — directly "
            "canonicalized.  Static + runtime pinned via "
            "`tests/test_known_bugs.py::test_bug22_"
            "associations_cmducinet_createtextfile_no_"
            "unicode_arg` and `tests/test_vba_bug_behaviors."
            "py::test_bug22_associations_cmducinet_fires_"
            "invalid_procedure_call`.  Original investigation "
            "evidence in `analysis/probe_associations_"
            "cmducinet_error5.md`.\n"
            "- **LookAtKinship** — **runtime-confirmed "
            "sibling form sharing the same root cause.**  "
            "`Form_LookAtKinship.CmdUCINet_Click` uses the "
            "same `CreateTextFile(tFileName, True)` 2-arg "
            "pattern at line ~2510.  Reproduced via probe "
            "`investigate/kinship-cmducinet-sibling-risk` "
            "(commit 154bb4b) using picker = pid 152930 (He "
            "Jing 何淨) whose sole 1-hop kin is pid 140733 "
            "(He Mou 取, U+53D6 = same CJK Han ideograph "
            "trigger class as Associations' 稜 = U+7A1C).  "
            "Probe outcome: same `:ERR Invalid procedure "
            "call or argument`, same partial-file shape "
            "(full `*node data` + truncated `*node "
            "properties` + missing `*tie data`).  Static "
            "marker `tests/test_known_bugs.py::test_bug22_"
            "associations_cmducinet_createtextfile_no_"
            "unicode_arg` extended in this PR to also "
            "assert the same 2-arg pattern in "
            "`Form_LookAtKinship.vb`; runtime pin for "
            "Kinship is deferred (see Coverage caveat "
            "below).\n"
            "- **LookAtPlace** — possible separate risk; "
            "**NOT covered by this issue's confirmation.**  "
            "`Form_LookAtPlace.CmdUCINet_Click` uses ADO "
            "Stream (`tStream.WriteText`) rather than FSO "
            "(`tVNA.WriteLine`), so its encoding behaviour "
            "is potentially different and would need its "
            "own per-form probe before any same-bug-family "
            "claim.  Place CmdUCINet stays `gap` in the "
            "inventory.\n\n"
            "**Coverage caveat:** the existing Kinship × "
            "CmdUCINet coverage test "
            "(`tests/test_vba_cmducinet_kinship.py`) "
            "remains `covered` in the inventory but is now "
            "**known fixture-fragile** — it passes only "
            "because the matrix-supplied person 3211's "
            "network happens to contain no Han-character "
            "c_name values.  Switching that fixture to one "
            "whose network reaches a Han-name person (the "
            "sibling probe demonstrates this directly) "
            "reproduces the same crash there.  Documented "
            "in the test's docstring + the inventory "
            "manifest's notes field."
        ),
        "summary_zh": (
            "`Form_LookAtAssociations.CmdUCINet_Click` 透過 "
            "`Scripting.FileSystemObject.CreateTextFile"
            "(tFileName, True)`（约 line 2575）写出 `.vna` "
            "档案。第 3 个参数（`Unicode`）省略掉了，预设值"
            "为 FALSE，所以档案以系统預設的 ANSI 码页（en-US "
            "Windows 上是 cp1252）开启。\n\n"
            "在 `*node properties` 区段，每一行写出 `tQuote + "
            "!c_name + tQuote`。当 `c_name` 含有 cp1252 编码"
            "无法对应、且 FSO 也没有替代映射的字元（特别是 "
            "CJK 汉字，例如 U+7A1C 稜），`tVNA.WriteLine` 就"
            "会丢出 VBA 5 错误（「Invalid procedure call or "
            "argument」），整个 CmdUCINet 汇出中断。残破的 "
            "`.vna` 档案会留在硬盘上 —— `*node data` 完整，"
            "`*node properties` 截断，`*tie data` 完全没写。\n"
            "\n"
            "用户可见症状：弹出 Run-time error 5 对话框；汇"
            "出的 `.vna` 档案不完整，UCINET / Visone 无法读"
            "取。\n\n"
            "**受影响表單（按当前 dump 的证据）：**\n\n"
            "- **LookAtAssociations** — 直接 canonicalize。"
            "由 `tests/test_known_bugs.py::test_bug22_"
            "associations_cmducinet_createtextfile_no_"
            "unicode_arg`（静态）和 `tests/test_vba_bug_"
            "behaviors.py::test_bug22_associations_cmducinet"
            "_fires_invalid_procedure_call`（运行时）双重"
            "钉死。原始调查证据见 `analysis/probe_"
            "associations_cmducinet_error5.md`。\n"
            "- **LookAtKinship** — **同 root cause 的 "
            "runtime-confirmed sibling form。**`Form_LookAt"
            "Kinship.CmdUCINet_Click` 在约 line 2510 用"
            "同样的 `CreateTextFile(tFileName, True)` 2-arg "
            "模式。透过 probe `investigate/kinship-cmducinet-"
            "sibling-risk`（commit 154bb4b）以 picker = pid "
            "152930（He Jing 何淨，唯一 1-hop kin 是 pid "
            "140733 He Mou 取，U+53D6 = 与 Associations 的 "
            "稜 = U+7A1C 同属 CJK Han ideograph 触发类）"
            "复现。Probe 结果：同样的 `:ERR Invalid "
            "procedure call or argument`，同样的残破档案"
            "形状（`*node data` 完整 + `*node properties` "
            "截断 + `*tie data` 完全没写）。本 PR 已扩展"
            "静态 marker `tests/test_known_bugs.py::test_"
            "bug22_associations_cmducinet_createtextfile_"
            "no_unicode_arg`，让它同时检查 `Form_LookAt"
            "Kinship.vb` 的同样 2-arg 模式；Kinship 的"
            "运行时 pin 暂缓（见下方 Coverage caveat）。\n"
            "- **LookAtPlace** — 可能存在的独立风险；"
            "**本 issue 的确认范围不包含 Place。**`Form_"
            "LookAtPlace.CmdUCINet_Click` 用的是 ADO "
            "Stream（`tStream.WriteText`），不是 FSO "
            "（`tVNA.WriteLine`），编码行为可能不同，需要"
            "单独的 per-form probe 才能下同 bug-family 的"
            "结论。Place CmdUCINet 在 inventory 仍维持 "
            "`gap`。\n\n"
            "**Coverage caveat：** 现有的 Kinship × "
            "CmdUCINet 覆盖测试（`tests/test_vba_cmducinet"
            "_kinship.py`）在 inventory 上仍是 `covered`，"
            "但已 **明确标注为 fixture-fragile** —— 它能"
            "通过只是因为 matrix 提供的 person 3211 网络"
            "刚好没有 Han 字符 c_name。换成一个网络能触达"
            "Han 名字的 fixture（sibling probe 直接示范了"
            "这一点）就会在同一段 .vna 写出路径上崩溃。"
            "已在测试的 docstring 与 inventory manifest 的 "
            "notes 字段同步备注。"
        ),
        "steps_en": [
            "Open CBDB_BJ_User.mdb in Microsoft Access.",
            "Open the **LookAtAssociations** form (F11 → "
            "navigation pane → forms → double-click "
            "`LookAtAssociations`).",
            "Use the person picker to select **c_personid = "
            "437 (Jia Zhaoming 賈昭明)** — one of the "
            "people whose 1st-order association network "
            "contains a person with a Han ideograph in "
            "their `c_name` field on the current dump "
            "(specifically pid 445395, c_name = `Hu Fa稜`).",
            "Click **Run** (CmdQuery) and wait for it to "
            "finish populating ZZ_SCRATCH_ASSOC + "
            "ZZ_SCRATCH_P_ASSOC.",
            "Click the **UCINet** export button (CmdUCINet) "
            "and choose any save location for the `.vna` "
            "file.",
            "A popup appears: `Run-time error '5': Invalid "
            "procedure call or argument`.  The export "
            "aborts.  The partial `.vna` file on disk has "
            "the complete `*node data` section but a "
            "truncated `*node properties` section and no "
            "`*tie data` section at all — unusable as a "
            "UCINET / Visone import.",
            "Verified end-to-end via probe at "
            "`analysis/probe_associations_cmducinet_error5."
            "py` — reproduces deterministically in ~15 s "
            "(see `analysis/probe_associations_cmducinet_"
            "error5.md` for the full evidence chain "
            "including the row-position localisation).",
        ],
        "steps_zh": [
            "在 Microsoft Access 里打开 CBDB_BJ_User.mdb。",
            "打开 **LookAtAssociations** 表单（F11 → 导航窗"
            "格 → 表单 → 双击 `LookAtAssociations`）。",
            "用人物 picker 选 **c_personid = 437（賈昭明 "
            "Jia Zhaoming）** —— 此人在当前 dump 上的 1 阶"
            "关联网络含有 c_name 带汉字的人（具体是 pid "
            "445395，c_name = `Hu Fa稜`）。",
            "点 **Run**（CmdQuery），等它把 ZZ_SCRATCH_ASSOC "
            "和 ZZ_SCRATCH_P_ASSOC 填好。",
            "点 **UCINet** 汇出按钮（CmdUCINet），随便选一"
            "个 `.vna` 档案的存档位置。",
            "弹出对话框：`Run-time error '5': Invalid "
            "procedure call or argument`。汇出中断，硬盘上"
            "只剩残破的 `.vna` 档：`*node data` 区段完整，"
            "`*node properties` 区段被截断，`*tie data` 区"
            "段完全没写 —— UCINET / Visone 都没法当成 import "
            "档案使用。",
            "已透过 `analysis/probe_associations_cmducinet"
            "_error5.py` 端到端验证 —— 约 15 秒可稳定复"
            "现（完整证据链与崩溃定位见 `analysis/probe_"
            "associations_cmducinet_error5.md`）。",
        ],
        "fix_en": (
            "Add `True` as the 3rd argument of "
            "`CreateTextFile` to open the file in Unicode "
            "(UTF-16LE) mode:\n\n"
            "```vb\n"
            "' before (Form_LookAtAssociations.vb:2575)\n"
            "Set tVNA = tFileSystem.CreateTextFile("
            "tFileName, True)\n"
            "\n"
            "' after — 3rd arg = Unicode = True\n"
            "Set tVNA = tFileSystem.CreateTextFile("
            "tFileName, True, True)\n"
            "```\n\n"
            "This should make `tVNA.WriteLine` write all "
            "characters as UTF-16LE and prevent the "
            "cp1252 write crash on non-cp1252 c_name "
            "values.  Downstream UCINET / Visone "
            "compatibility with UTF-16 `.vna` is NOT "
            "verified by this PR's evidence and should be "
            "verified on the fixed build before declaring "
            "the bug closed.\n\n"
            "Alternative (less recommended): strip / "
            "transliterate non-cp1252 chars from `c_name` "
            "before the `WriteLine` call.  Loses real data "
            "from the export and is more code; the Unicode "
            "flag is the right fix.\n\n"
            "Same one-line fix (with the same 3rd-arg "
            "addition) is also required for "
            "`Form_LookAtKinship.CmdUCINet_Click` (line "
            "~2510) — see the Affected-forms section above.  "
            "Kinship is a runtime-confirmed sibling form "
            "of THIS issue (same root cause, same failure "
            "class, different host form / fixture trigger), "
            "so a single upstream patch should add the "
            "Unicode flag to BOTH CreateTextFile call "
            "sites in the same change.  Place "
            "(LookAtPlace.CmdUCINet) is NOT in scope for "
            "this issue — it uses ADO Stream rather than "
            "FSO and would need its own per-form probe "
            "before any same-bug-family claim or fix "
            "coordination."
        ),
        "fix_zh": (
            "在 `CreateTextFile` 第 3 个参数加上 `True`，让"
            "档案以 Unicode（UTF-16LE）模式开启：\n\n"
            "```vb\n"
            "' 修改前（Form_LookAtAssociations.vb:2575）\n"
            "Set tVNA = tFileSystem.CreateTextFile("
            "tFileName, True)\n"
            "\n"
            "' 修改后 —— 第 3 个参数 = Unicode = True\n"
            "Set tVNA = tFileSystem.CreateTextFile("
            "tFileName, True, True)\n"
            "```\n\n"
            "这样 `tVNA.WriteLine` 应能用 UTF-16LE 写入所有"
            "字元，避免 cp1252 写出时在非 cp1252 c_name 上的"
            "崩溃。下游 UCINET / Visone 对 UTF-16 `.vna` 的"
            "相容性 **不在本 PR 证据范围内**，仍应在修补版上"
            "再验证一次，才能宣告本 issue 关闭。\n\n"
            "替代方案（不太推荐）：在 `WriteLine` 之前把 "
            "`c_name` 里的非 cp1252 字元剥掉或转写。会丢失"
            "汇出的真实资料，而且代码量更大；Unicode flag "
            "才是正解。\n\n"
            "`Form_LookAtKinship.CmdUCINet_Click`（约 line "
            "2510）也需要套用同样的一行修改 —— 详见上方"
            "「受影响表單」段。Kinship 是本 issue 的 "
            "runtime-confirmed sibling form（同 root cause、"
            "同 failure class，只是宿主表单与触发 fixture "
            "不同），所以一次上游修补应同时给两个 "
            "CreateTextFile 加上 Unicode flag。Place "
            "（LookAtPlace.CmdUCINet）**不在** 本 issue "
            "的范围内 —— 它用的是 ADO Stream 而非 FSO，"
            "需要单独的 per-form probe 才能下同 bug-family "
            "的结论或加入修补范围。"
        ),
        "screenshots": [],
        "severity_en": (
            "P1 — Visible crash on a normal user click.  "
            "Any user attempting `LookAtAssociations × "
            "CmdUCINet` whose 1st-order association network "
            "happens to include a person with a CJK Han "
            "ideograph in their `c_name` will hit a "
            "Run-time error 5 popup and lose the export.  "
            "On the current dump that's at least the "
            "8087-row scratch table for person 437 (the "
            "verified fixture); the broader prevalence "
            "across CBDB persons depends on how many "
            "BIOG_MAIN rows have Han ideographs in their "
            "ostensibly-Latin `c_name` field — at minimum 2 "
            "such rows reach person 437's network, and any "
            "person with even one such 1st-order neighbour "
            "is affected."
        ),
        "severity_zh": (
            "P1 — 正常用户点击下的可见报错。任何使用者只要"
            "在 `LookAtAssociations × CmdUCINet` 上选的人"
            "在当前 dump 的 1 阶关联网络中含有 c_name 带 "
            "CJK 汉字的人，就会遇到 Run-time error 5 对话"
            "框，汇出全部失败。当前 dump 上至少 person 437"
            "（已验证 fixture）的 8087 行暂存表会触发；更"
            "大范围的影响人数取决于 BIOG_MAIN 中 c_name "
            "（理论上是 Latin / Pinyin）含汉字的行数 —— "
            "至少 2 行落在 person 437 的网络里，只要选的"
            "人 1 阶邻居含其中任一行都会受影响。"
        ),
    },

    {
        "id": 14,
        "tier": "P5_dormant_or_latent",
        "form": "Form_KIN_DATA_Subform",
        "title_en": "KIN_DATA Subform's CmdPickKinRel calls a missing picker (frmPickKINSHIP_CODES) — but the host sub-form is currently an orphan (LATENT)",
        "title_zh": "KIN_DATA 子表單的 CmdPickKinRel 呼叫不存在的 picker（frmPickKINSHIP_CODES）——但目前該子表單在主表中無入口（LATENT）",
        "summary_en": (
            "**Static defect is real, runtime trigger is not currently "
            "reachable.** The Sub `CmdPickKinRel_Click` in "
            "`Form_KIN_DATA_Subform` (line 52) calls `DoCmd.OpenForm "
            "\"frmPickKINSHIP_CODES\"` and references "
            "`Forms!frmPickKINSHIP_CODES!frmKINSHIP_CODES.Form!c_kincode`. "
            "Neither form exists in the current .mdb — same shape as "
            "Issue #13.\n\n"
            "**Why LATENT.** The host sub-form `KIN_DATA Subform` (which "
            "owns the `CmdPickKinRel` button) is not contained by any "
            "active form in the current `control_inventory.json`. "
            "`BIOG_MAIN_2_Subform` (the kinship surface users actually "
            "navigate to via CBDB_Browser_2) embeds `KIN_DATA_2 Subform` "
            "instead — and that variant has no `CmdPickKinRel` button "
            "(its 15 controls are all read-only fields). The only place "
            "the embedding still appears is `Form__TMPCLP487951` (a "
            "design-time backup snapshot, not a navigable form).\n\n"
            "Because no user-facing navigation reaches the picker button, "
            "users cannot trigger the popup from normal use. The latent "
            "code path will resurface the moment a developer re-embeds "
            "`KIN_DATA Subform` somewhere reachable, so the underlying "
            "fix is still worth applying."
        ),
        "summary_zh": (
            "**靜態缺陷確實存在，但執行時的觸發路徑當前不可達。**"
            "`Form_KIN_DATA_Subform` 第 52 行的 Sub `CmdPickKinRel_Click` "
            "呼叫 `DoCmd.OpenForm \"frmPickKINSHIP_CODES\"`，並引用 "
            "`Forms!frmPickKINSHIP_CODES!frmKINSHIP_CODES.Form!c_kincode`。"
            "這兩個表單在目前的 .mdb 都不存在——形狀與 Issue #13 相同。\n\n"
            "**為何 LATENT。**承載該按鈕的子表單 `KIN_DATA Subform` "
            "（即擁有 `CmdPickKinRel` 按鈕者）在當前的 "
            "`control_inventory.json` 裡沒有被任何 active 表單包含。"
            "用戶實際從 CBDB_Browser_2 進入的親屬介面是 "
            "`BIOG_MAIN_2_Subform`，而它包含的是 `KIN_DATA_2 Subform`"
            "（另一個版本，15 個控件全是唯讀欄位，沒有 `CmdPickKinRel`"
            "按鈕）。唯一仍然嵌入該子表單的位置是 "
            "`Form__TMPCLP487951`，那是設計時的備份快照，不是可導航的表單。\n\n"
            "因為沒有任何使用者介面能走到那個 picker 按鈕，"
            "正常使用下不會彈出錯誤。但只要日後有人把 "
            "`KIN_DATA Subform` 重新嵌進可達的位置，這條潛在錯誤路徑"
            "就會立刻浮現，所以底層的修復仍值得做。"
        ),
        "steps_en": [
            "Verification path is **static-only** — the runtime click "
            "cannot be reproduced in the current .mdb because no parent "
            "form embeds the affected sub-form.",
            "Static evidence (1): open `analysis/dump/vba/"
            "Form_KIN_DATA_Subform.vb` line 52 — confirms the Sub calls "
            "`DoCmd.OpenForm \"frmPickKINSHIP_CODES\"`.",
            "Static evidence (2): open `analysis/dump/"
            "control_inventory.json` and search for `\"frmPickKINSHIP_CODES\"`"
            " as a key — absent. The picker form does not exist.",
            "Reachability evidence: in the same JSON, search for "
            "`\"KIN_DATA Subform\"` as a `source_object` or sub-form "
            "control name — only `Form__TMPCLP487951` (a design backup) "
            "references it. `BIOG_MAIN_2_Subform` embeds "
            "`KIN_DATA_2 Subform` instead, which has no "
            "`CmdPickKinRel` button.",
        ],
        "steps_zh": [
            "驗證路徑**只能靜態驗證**——當前 .mdb 中並無主表單嵌入"
            "受影響的子表單，因此無法在執行時重現點擊。",
            "靜態證據 (1)：打開 `analysis/dump/vba/"
            "Form_KIN_DATA_Subform.vb` 第 52 行——可見該 Sub 確實呼叫 "
            "`DoCmd.OpenForm \"frmPickKINSHIP_CODES\"`。",
            "靜態證據 (2)：打開 `analysis/dump/"
            "control_inventory.json`，以 `\"frmPickKINSHIP_CODES\"` "
            "為鍵搜索——不存在。該 picker 表單已不存在於 .mdb 中。",
            "可達性證據：在同一份 JSON 中搜索 `\"KIN_DATA Subform\"` "
            "作為 `source_object` 或子表單控件名——只有 "
            "`Form__TMPCLP487951`（設計時備份快照）引用它。"
            "用戶實際走到的 `BIOG_MAIN_2_Subform` 嵌入的是 "
            "`KIN_DATA_2 Subform`，那一版沒有 `CmdPickKinRel` 按鈕。",
        ],
        "fix_en": (
            "Same as Issue #13: restore the picker form (or update the "
            "caller to its replacement). Even though the runtime path is "
            "not currently reachable, the static defect should be cleaned "
            "up so it doesn't resurface when `KIN_DATA Subform` is "
            "re-embedded."
        ),
        "fix_zh": (
            "與 Issue #13 相同：把 picker 表單恢復，或把呼叫方改成指向"
            "替代的 picker。雖然目前執行路徑不可達，靜態缺陷仍應清理，"
            "以免日後 `KIN_DATA Subform` 被重新嵌入時又冒出來。"
        ),
        "screenshots": [],
        "severity_en": (
            "P5 — Latent (would be P1 if `KIN_DATA Subform` were "
            "re-embedded somewhere users can reach)"
        ),
        "severity_zh": (
            "P5 — Latent（若日後把 `KIN_DATA Subform` 重新嵌入"
            "使用者可達的位置，會回到 P1）"
        ),
    },
    # ========== Tier 3: silent display (data shown wrong/missing) ==========
    {
        "id": 10,
        "tier": "P2_silent_display",
        "form": "EVENT_ADDR_2 Subform",
        "title_en": "EVENT_ADDR_2 Subform address columns silently render blank (wrong ControlSource)",
        "title_zh": "EVENT_ADDR_2 子表单的地址列默默地显示为空（ControlSource 写错了）",
        "summary_en": (
            "On the EVENT_ADDR_2 sub-form (events with addresses), the two "
            "address controls are bound as follows:\n\n"
            "  • `TxtAddrCHN`.ControlSource = `c_name_chn`\n"
            "  • `TxtAddrPY`.ControlSource = `c_name`\n\n"
            "But the form's RecordSource is the saved query "
            "`View_EventAddrData`, which aliases ADDR_CODES.c_name_chn as "
            "`c_event_addr_chn` and ADDR_CODES.c_name as `c_event_addr_name`. "
            "Neither `c_name` nor `c_name_chn` is in the projection, so "
            "both controls silently render blank for every row."
        ),
        "summary_zh": (
            "在 EVENT_ADDR_2 子表单（带地址的事件）上，两个地址控件的绑定如下：\n\n"
            "  • `TxtAddrCHN`.ControlSource = `c_name_chn`\n"
            "  • `TxtAddrPY`.ControlSource = `c_name`\n\n"
            "但该表单的 RecordSource 是存档查询 `View_EventAddrData`，里面把 "
            "ADDR_CODES.c_name_chn 起别名成 `c_event_addr_chn`、把 "
            "ADDR_CODES.c_name 起别名成 `c_event_addr_name`。投影里既没有 "
            "`c_name` 也没有 `c_name_chn`，所以这两个控件每一行都默默地显示"
            "为空。"
        ),
        "steps_en": [
            "Open CBDB_Browser_2 and navigate to **c_personid = 44872 "
            "(Sun Cai 孫才)** — picked because he has 1 EVENTS_DATA "
            "row with an associated EVENT_ADDR row pointing at "
            "`c_addr_id = 12603` (Anfeng / 安豐 in ADDR_CODES).  "
            "Switch to the **Events** sub-tab.",
            "Look at the small EVENT_ADDR_2 sub-form embedded inside "
            "the event row (it's nested below the main event line).  "
            "The two address textboxes there — `TxtAddrCHN` and "
            "`TxtAddrPY` — render blank.",
            "**Note:** the parent EVENTS_DATA_2 sub-form has its own "
            "address controls (also called TxtAddrCHN / TxtAddrPY but "
            "bound to `c_addr_chn` / `c_addr_name`, which `View_EventsData` "
            "DOES project) — those work and show '安豐' / 'Anfeng'.  "
            "Bug #10 is specifically about the inner EVENT_ADDR_2 "
            "sub-form's controls being blank, not about the visible "
            "address values on the parent row.",
            "SQL verification (no Access needed): "
            "`SELECT c_name_chn FROM View_EventAddrData` raises "
            "`Too few parameters. Expected 2.` — JET treats the "
            "unknown identifier as a parameter, confirming the column "
            "is not in the projection.",
        ],
        "steps_zh": [
            "打開 CBDB_Browser_2，導航到 **c_personid = 44872"
            "（孫才 Sun Cai）**——選他是因為他有 1 條 EVENTS_DATA 記錄，"
            "對應 1 條 EVENT_ADDR 指向 `c_addr_id = 12603`（ADDR_CODES "
            "裡是 Anfeng / 安豐）。切到 **Events** 子分頁。",
            "看事件那行內嵌的 EVENT_ADDR_2 子表單"
            "（在主事件那行下方的一小條）。"
            "那裡的兩個地址欄位 `TxtAddrCHN` 與 `TxtAddrPY` 都是空白。",
            "**注意：**外層的 EVENTS_DATA_2 子表單也有自己的"
            "地址欄位（也叫 TxtAddrCHN / TxtAddrPY，但是綁到 "
            "`c_addr_chn` / `c_addr_name`，這兩個 `View_EventsData` "
            "確實有 project）——這兩個欄位是正常的，會顯示「安豐 / Anfeng」。"
            "Bug #10 講的是內層 EVENT_ADDR_2 那兩個空欄位，"
            "不是父層那條看得見的地址值。",
            "SQL 驗證（不需開 Access）："
            "`SELECT c_name_chn FROM View_EventAddrData` 會拋 "
            "`Too few parameters. Expected 2.`——JET 把未知識別字當作"
            "參數對待，這就確認了該欄位不在投影裡。",
        ],
        "fix_en": (
            "In the form designer for `EVENT_ADDR_2 Subform`, change "
            "`TxtAddrCHN`.ControlSource from `c_name_chn` to "
            "`c_event_addr_chn`, and `TxtAddrPY`.ControlSource from "
            "`c_name` to `c_event_addr_name` (the actual aliases in "
            "View_EventAddrData)."
        ),
        "fix_zh": (
            "在 `EVENT_ADDR_2 Subform` 的表單設計檢視裡，"
            "把 `TxtAddrCHN`.ControlSource 由 `c_name_chn` 改成 "
            "`c_event_addr_chn`；把 `TxtAddrPY`.ControlSource 由 "
            "`c_name` 改成 `c_event_addr_name`（這才是 "
            "View_EventAddrData 裡真實的別名）。"
        ),
        "screenshots": [
            ("bug10_subform_annotated.png",
             "Runtime view of CBDB_Browser_2 → BIOG_MAIN_2 → Events tab "
             "with c_personid=44872 (孫才) loaded.  **The visible '安豐' / "
             "address values come from the parent EVENTS_DATA_2 sub-form's "
             "TxtAddrCHN (correctly bound to `c_addr_chn`).**  Bug #10's "
             "blank controls live in the smaller EVENT_ADDR_2 sub-form "
             "nested inside the event row — those two controls "
             "(TxtAddrCHN / TxtAddrPY bound to `c_name_chn` / `c_name`, "
             "neither in `View_EventAddrData`'s projection) render "
             "empty.  COM probe confirms both are Visible=True with "
             "widths 2340 / 2100 twips (≈4cm / 3.5cm) — i.e. real "
             "user-visible blank columns, just smaller than the parent "
             "row's address display.  Verification scripts: "
             "`analysis/probe_bug_10_11_12_visibility.py` (control "
             "visibility) + the SQL probe in the steps above."),
        ],
        "severity_en": "P2 — Silent display (EVENT_ADDR_2's TxtAddrCHN / TxtAddrPY render blank for every row)",
        "severity_zh": "P2 — 靜默顯示問題（EVENT_ADDR_2 的 TxtAddrCHN / TxtAddrPY 每一列都空白）",
    },
    {
        "id": 11,
        "tier": "P5_dormant_or_latent",
        "form": "EVENTS_DATA_2 Subform",
        "title_en": "EVENTS_DATA_2's c_event_record_id control bound to a non-existent column — but the control is hidden (LATENT)",
        "title_zh": "EVENTS_DATA_2 上 c_event_record_id 控件綁到不存在的欄位——但該控件本身是隱藏的（LATENT）",
        "summary_en": (
            "**Static defect is real, runtime symptom is not user-visible.** "
            "The EVENTS_DATA_2 sub-form has a control named "
            "`c_event_record_id` whose ControlSource is also "
            "`c_event_record_id`.  Neither EVENTS_DATA nor "
            "`View_EventsData` projects a column of that name (SQL probe "
            "confirms — `SELECT c_event_record_id FROM View_EventsData` "
            "raises `Too few parameters. Expected 1.`), so the control "
            "would render blank if shown.\n\n"
            "**Why LATENT.**  A live COM probe of the rendered form "
            "(`analysis/probe_bug_10_11_12_visibility.py`) reports the "
            "control as `Visible = False`, with width = 240 twips "
            "(~4mm) and height = 270 twips — i.e. a hidden internal "
            "control, almost certainly a leftover join-key field that "
            "was never meant to be shown.  Real users won't see a blank "
            "column because they don't see the control at all.  Reclassed "
            "from P2 to P5 on 2026-05-03."
        ),
        "summary_zh": (
            "**靜態缺陷確實存在，但執行時用戶看不到。**"
            "EVENTS_DATA_2 子表單上有一個叫 `c_event_record_id` 的控件，"
            "ControlSource 也是 `c_event_record_id`。EVENTS_DATA 與 "
            "`View_EventsData` 都沒有這個欄位（SQL 驗證："
            "`SELECT c_event_record_id FROM View_EventsData` 會拋 "
            "`Too few parameters. Expected 1.`）。所以如果該控件顯示出來，"
            "確實會空白。\n\n"
            "**為何 LATENT。**對 runtime 表單做 COM 探測"
            "（`analysis/probe_bug_10_11_12_visibility.py`）"
            "顯示該控件 `Visible = False`，寬 240 twips（~4mm）、"
            "高 270 twips——這就是一個隱藏的內部控件，"
            "幾乎可以肯定是早期殘留的 join-key 欄位，"
            "本來就不打算給用戶看。"
            "用戶不會看到空白欄位，因為根本看不到這個控件。"
            "2026-05-03 從 P2 降到 P5。"
        ),
        "steps_en": [
            "Verification path is **static + COM probe only** — there's "
            "no UI symptom to demonstrate.",
            "Static evidence: `SELECT c_event_record_id FROM "
            "View_EventsData` against the user mdb raises `Too few "
            "parameters. Expected 1.`, confirming the column is not in "
            "the projection.",
            "Visibility evidence: run `python "
            "analysis/probe_bug_10_11_12_visibility.py` and look at the "
            "entry for bug #11 in `analysis/dump/"
            "bug_10_11_12_visibility.json` — `control_summary.visible` "
            "is `False` and `width` is 240 twips.",
        ],
        "steps_zh": [
            "驗證路徑**只能靜態 + COM 探測**——沒有 UI 上的可見徵狀。",
            "靜態證據：對 user mdb 跑 `SELECT c_event_record_id FROM "
            "View_EventsData` 會拋 `Too few parameters. Expected 1.`，"
            "確認該欄位不在投影裡。",
            "可見性證據：跑 `python "
            "analysis/probe_bug_10_11_12_visibility.py`，看 "
            "`analysis/dump/bug_10_11_12_visibility.json` 中 bug #11 "
            "那筆——`control_summary.visible` 是 `False`、`width` 是 "
            "240 twips。",
        ],
        "fix_en": (
            "If the hidden control isn't needed, delete it.  If it's "
            "intentionally a hidden join-key holder, change its "
            "ControlSource to a real column (e.g. `c_event_code`) so "
            "it doesn't carry a stale binding.  Either way, the change "
            "is invisible to users; this is code-hygiene only."
        ),
        "fix_zh": (
            "若這個隱藏控件用不到了，直接刪除即可；若原意是隱藏的 "
            "join-key 容器，把 ControlSource 改成真實的欄位"
            "（例如 `c_event_code`），免得帶著一個失效的綁定。"
            "無論怎麼改，使用者都看不到差別——這純粹是程式碼整潔。"
        ),
        "screenshots": [],
        "severity_en": (
            "P5 — Latent (would be P2 if the control were ever made "
            "Visible=True or widened past its current 240-twip "
            "width)"
        ),
        "severity_zh": (
            "P5 — Latent（若日後把該控件改成 Visible=True 或加寬到 "
            "240 twips 以上，就會回到 P2）"
        ),
    },
    {
        "id": 12,
        "tier": "P5_dormant_or_latent",
        "form": "POSTED_TO_OFFICE_DATA_2 Subform",
        "title_en": "POSTED_TO_OFFICE_DATA_2's c_appt_type_code control bound to a non-projected column — but the control is hidden AND the user-facing appointment-type controls work (LATENT)",
        "title_zh": "POSTED_TO_OFFICE_DATA_2 上 c_appt_type_code 控件綁到沒投影的欄位——但該控件是隱藏的，且用戶實際看的任職類型欄位是正常的（LATENT）",
        "summary_en": (
            "**Static defect is real, runtime symptom is not user-visible.** "
            "The hidden internal control `c_appt_type_code` on "
            "POSTED_TO_OFFICE_DATA_2 has ControlSource "
            "`c_appt_type_code`, which `View_PostingOfficeData` doesn't "
            "project (SQL probe: raises `Too few parameters. Expected "
            "1.`).\n\n"
            "**Why LATENT.**  Two reasons:\n\n"
            "1. The COM probe (`analysis/probe_bug_10_11_12_visibility.py`) "
            "reports the control as `Visible = False`, with width = "
            "180 twips (~3mm) and height = 330 twips — a hidden "
            "internal control, almost certainly a join-key holder.\n"
            "2. The **user-facing** appointment-type controls on the "
            "same sub-form are `TxtApptType` (bound to `c_appt_desc`) "
            "and `TxtApptTypeChn` (bound to `c_appt_desc_chn`).  Both "
            "of those columns ARE in `View_PostingOfficeData`'s "
            "projection — SQL probe shows they return real values "
            "(e.g. `'Regular Appointment'` / `'正授'`).  So the "
            "appointment type IS displayed correctly on the Postings "
            "sub-tab; only the hidden `c_appt_type_code` control is "
            "broken.\n\n"
            "Reclassed from P2 to P5 on 2026-05-03 — the original P2 "
            "claim that 'the appointment-type column is blank on every "
            "row' was wrong; the user-facing appointment-type column "
            "works fine."
        ),
        "summary_zh": (
            "**靜態缺陷確實存在，但執行時用戶看不到。**"
            "POSTED_TO_OFFICE_DATA_2 上隱藏的內部控件 `c_appt_type_code` "
            "ControlSource 是 `c_appt_type_code`，而 "
            "`View_PostingOfficeData` 沒有投影這個欄位"
            "（SQL 驗證：拋 `Too few parameters. Expected 1.`）。\n\n"
            "**為何 LATENT。** 兩個理由：\n\n"
            "1. COM 探測（`analysis/probe_bug_10_11_12_visibility.py`）"
            "顯示該控件 `Visible = False`，寬 180 twips（~3mm）、"
            "高 330 twips——典型的隱藏 join-key 控件。\n"
            "2. 同一個子表單上**真正給用戶看的**任職類型欄位是 "
            "`TxtApptType`（綁 `c_appt_desc`）與 `TxtApptTypeChn`"
            "（綁 `c_appt_desc_chn`）。這兩個欄位都在 "
            "`View_PostingOfficeData` 的投影裡——SQL 探測能拿到真實值"
            "（例如 `'Regular Appointment'` / `'正授'`）。"
            "所以 Postings 分頁上的任職類型**是正常顯示的**；"
            "只有隱藏的 `c_appt_type_code` 控件壞掉。\n\n"
            "2026-05-03 從 P2 降到 P5——原本「任職類型列每一列都是空白」"
            "的 P2 說法是錯的；用戶實際看的任職類型欄位是正常的。"
        ),
        "steps_en": [
            "Verification path is **static + COM probe only** — there's "
            "no UI symptom to demonstrate.",
            "Static evidence: `SELECT c_appt_type_code FROM "
            "View_PostingOfficeData` raises `Too few parameters. "
            "Expected 1.`, confirming the column is not projected.",
            "Visibility evidence: run `python "
            "analysis/probe_bug_10_11_12_visibility.py` and look at "
            "the entry for bug #12 in `analysis/dump/"
            "bug_10_11_12_visibility.json` — `control_summary.visible` "
            "is `False` and width = 180 twips.",
            "Counter-evidence (what users actually see works fine): "
            "`SELECT TOP 1 c_appt_desc, c_appt_desc_chn FROM "
            "View_PostingOfficeData` returns real values (e.g. "
            "`'Regular Appointment'` / `'正授'`), and those are what "
            "`TxtApptType` / `TxtApptTypeChn` (the visible controls) "
            "render.",
        ],
        "steps_zh": [
            "驗證路徑**只能靜態 + COM 探測**——沒有 UI 上的可見徵狀。",
            "靜態證據：`SELECT c_appt_type_code FROM "
            "View_PostingOfficeData` 會拋 `Too few parameters. "
            "Expected 1.`，確認該欄位不在投影裡。",
            "可見性證據：跑 `python "
            "analysis/probe_bug_10_11_12_visibility.py`，看 "
            "`analysis/dump/bug_10_11_12_visibility.json` 中 bug #12 "
            "那筆——`control_summary.visible` 是 `False`、"
            "width 是 180 twips。",
            "反證（用戶實際看的欄位是正常的）："
            "`SELECT TOP 1 c_appt_desc, c_appt_desc_chn FROM "
            "View_PostingOfficeData` 能返回真實值（例如 "
            "`'Regular Appointment'` / `'正授'`），這就是 "
            "`TxtApptType` / `TxtApptTypeChn`（可見的兩個控件）"
            "渲染出來的內容。",
        ],
        "fix_en": (
            "If the hidden control isn't needed, delete it.  If it's "
            "an intentional hidden join-key holder, change its "
            "ControlSource to a real column (e.g. `c_appt_code`).  "
            "Either way the change is invisible to users; this is "
            "code-hygiene only."
        ),
        "fix_zh": (
            "若這個隱藏控件用不到了，刪除即可；若是有意為之的隱藏 "
            "join-key 容器，把 ControlSource 改成真實的欄位"
            "（例如 `c_appt_code`）。無論怎麼改，使用者都看不到差別"
            "——這純粹是程式碼整潔。"
        ),
        "screenshots": [],
        "severity_en": (
            "P5 — Latent (would be P2 if the control were ever made "
            "Visible=True or widened past its current 180-twip width)"
        ),
        "severity_zh": (
            "P5 — Latent（若日後把該控件改成 Visible=True 或加寬到 "
            "180 twips 以上，就會回到 P2）"
        ),
    },
    # ========== Tier 4: missing UI (export buttons not on the form) ==========
    {
        "id": 15,
        "tier": "P3_missing_ui",
        "form": "LookAtPlace",
        "title_en": "LookAtPlace is missing its CmdGIS button (handler exists but no UI control)",
        "title_zh": "LookAtPlace 缺少 CmdGIS 按钮（代码里有 handler 但界面上没控件）",
        "summary_en": (
            "`Form_LookAtPlace.vb` defines a fully functional "
            "`CmdGIS_Click` handler — it builds and writes a GIS .tab "
            "export, identical in shape to the GIS button on Status / "
            "Texts / Associations / Office / Kinship. But LookAtPlace's "
            "form design has NO `CmdGIS` button on it. Users on Place "
            "can use Pajek / Gephi / Neo4j export but cannot use GIS "
            "export — the handler is there, just unreachable from the UI."
        ),
        "summary_zh": (
            "`Form_LookAtPlace.vb` 里定义了一个完整可用的 `CmdGIS_Click` "
            "handler——构造并输出 GIS .tab 文件，逻辑和 Status / Texts / "
            "Associations / Office / Kinship 上的 GIS 按钮一模一样。但 "
            "LookAtPlace 的设计视图上根本没有 `CmdGIS` 按钮。用户在 Place "
            "上可以使用 Pajek / Gephi / Neo4j 导出，但用不了 GIS 导出——"
            "代码在那里，只是界面进不去。"
        ),
        "steps_en": [
            "Open **LookAtPlace**.",
            "Look at the export-buttons row at the bottom right.",
            "There's no GIS button. Compare with LookAtStatus / "
            "LookAtAssociations / LookAtOffice etc., all of which have one.",
        ],
        "steps_zh": [
            "打开 **LookAtPlace**。",
            "看右下方那一排导出按钮。",
            "没有 GIS 按钮。可以对比 LookAtStatus / LookAtAssociations / "
            "LookAtOffice 等，它们都有这个按钮。",
        ],
        "fix_en": (
            "In LookAtPlace's form design, add a CmdGIS button next to "
            "the existing CmdPajek / CmdGephi buttons, with `OnClick = "
            "[Event Procedure]` so it invokes the existing CmdGIS_Click "
            "Sub. (Also fix Issue #4 first, otherwise the button will "
            "throw 'Object required' on the first click.)"
        ),
        "fix_zh": (
            "在 LookAtPlace 的设计视图里，在已有的 CmdPajek / CmdGephi 旁边"
            "加一个 CmdGIS 按钮，把 `OnClick` 设为 `[Event Procedure]`，这样"
            "它就会调用已经写好的 CmdGIS_Click。（同时务必先修 Issue #4，"
            "否则按钮一点就会报 Object required。）"
        ),
        "screenshots": [
            ("bug15_LookAtPlace_no_CmdGIS_annotated.png",
             "LookAtPlace as it ships — no GIS button is rendered, even though `Sub CmdGIS_Click()` exists in the module."),
        ],
        "severity_en": "P3 — Missing UI (feature unavailable to users)",
        "severity_zh": "P3 — 缺失界面（用户用不到该功能）",
    },
    {
        "id": 16,
        "tier": "P3_missing_ui",
        "form": "LookAtStatus",
        "title_en": "LookAtStatus is missing its CmdPajek button",
        "title_zh": "LookAtStatus 缺少 CmdPajek 按钮",
        "summary_en": (
            "Same shape as Issue #15. `Sub CmdPajek_Click()` exists in "
            "`Form_LookAtStatus.vb` (would write a Pajek .net export of "
            "the status data) but no CmdPajek button is rendered on "
            "Status's form design.\n\n"
            "Note: even if the button is added, Issue #5 (the SQL/control "
            "defects in CmdPajek_Click itself) needs to be fixed first."
        ),
        "summary_zh": (
            "形态与 Issue #15 相同。`Sub CmdPajek_Click()` 在 "
            "`Form_LookAtStatus.vb` 里有定义（本应输出 Pajek .net 文件），"
            "但 Status 的设计视图上没有 CmdPajek 按钮。\n\n"
            "注意：即便把按钮加回去，也得先解决 Issue #5（CmdPajek_Click 本身"
            "的 SQL/控件缺陷）。"
        ),
        "steps_en": [
            "Open **LookAtStatus**. The export-buttons row has only GIS "
            "and Neo4j; there's no Pajek button.",
        ],
        "steps_zh": [
            "打开 **LookAtStatus**。导出按钮一栏只有 GIS 和 Neo4j，没有 Pajek。",
        ],
        "fix_en": (
            "Add a CmdPajek button to LookAtStatus's design (after fixing "
            "Issue #5)."
        ),
        "fix_zh": (
            "在 LookAtStatus 的设计视图里加一个 CmdPajek 按钮（先把 Issue #5 "
            "修好）。"
        ),
        "screenshots": [
            ("bug16_LookAtStatus_no_CmdPajek_annotated.png", None),
        ],
        "severity_en": "P3 — Missing UI",
        "severity_zh": "P3 — 缺失界面",
    },
    {
        "id": 17,
        "tier": "P3_missing_ui",
        "form": "LookAtStatus",
        "title_en": "LookAtStatus is missing its CmdGephi button",
        "title_zh": "LookAtStatus 缺少 CmdGephi 按钮",
        "summary_en": (
            "`Sub CmdGephi_Click()` exists in `Form_LookAtStatus.vb` but "
            "no matching button is on the form design."
        ),
        "summary_zh": (
            "`Sub CmdGephi_Click()` 在 `Form_LookAtStatus.vb` 里有定义，"
            "但表单设计视图里没有相应按钮。"
        ),
        "steps_en": [
            "Open **LookAtStatus**. There is no Gephi export button.",
        ],
        "steps_zh": [
            "打开 **LookAtStatus**。没有 Gephi 导出按钮。",
        ],
        "fix_en": (
            "Add a CmdGephi button to LookAtStatus's design."
        ),
        "fix_zh": (
            "在 LookAtStatus 的设计视图里加一个 CmdGephi 按钮。"
        ),
        "screenshots": [
            ("bug17_LookAtStatus_no_CmdGephi_annotated.png", None),
        ],
        "severity_en": "P3 — Missing UI",
        "severity_zh": "P3 — 缺失界面",
    },
    {
        "id": 18,
        "tier": "P3_missing_ui",
        "form": "LookAtStatus",
        "title_en": "LookAtStatus is missing its CmdUCINet button",
        "title_zh": "LookAtStatus 缺少 CmdUCINet 按钮",
        "summary_en": (
            "`Sub CmdUCINet_Click()` exists in `Form_LookAtStatus.vb` but "
            "no matching button is on the form design."
        ),
        "summary_zh": (
            "`Sub CmdUCINet_Click()` 在 `Form_LookAtStatus.vb` 里有定义，"
            "但表单设计视图里没有相应按钮。"
        ),
        "steps_en": [
            "Open **LookAtStatus**. There is no UCINet export button.",
        ],
        "steps_zh": [
            "打开 **LookAtStatus**。没有 UCINet 导出按钮。",
        ],
        "fix_en": (
            "Add a CmdUCINet button to LookAtStatus's design."
        ),
        "fix_zh": (
            "在 LookAtStatus 的设计视图里加一个 CmdUCINet 按钮。"
        ),
        "screenshots": [
            ("bug18_LookAtStatus_no_CmdUCINet_annotated.png", None),
        ],
        "severity_en": "P3 — Missing UI",
        "severity_zh": "P3 — 缺失界面",
    },
    {
        "id": 19,
        "tier": "P3_missing_ui",
        "form": "LookAtOffice",
        "title_en": "LookAtOffice is missing its CmdGUESS button",
        "title_zh": "LookAtOffice 缺少 CmdGUESS 按钮",
        "summary_en": (
            "`Sub CmdGUESS_Click()` exists in `Form_LookAtOffice.vb` but "
            "no CmdGUESS button is on the form design. Users on Office can "
            "use GIS / GISPeople / Neo4j export but not GUESS."
        ),
        "summary_zh": (
            "`Sub CmdGUESS_Click()` 在 `Form_LookAtOffice.vb` 里有定义，"
            "但 Office 的设计视图上没有 CmdGUESS 按钮。Office 上的用户可以"
            "使用 GIS / GISPeople / Neo4j 导出，但用不了 GUESS。"
        ),
        "steps_en": [
            "Open **LookAtOffice**. There is no GUESS export button.",
        ],
        "steps_zh": [
            "打开 **LookAtOffice**。没有 GUESS 导出按钮。",
        ],
        "fix_en": (
            "Add a CmdGUESS button to LookAtOffice's design."
        ),
        "fix_zh": (
            "在 LookAtOffice 的设计视图里加一个 CmdGUESS 按钮。"
        ),
        "screenshots": [
            ("bug19_LookAtOffice_no_CmdGUESS_annotated.png", None),
        ],
        "severity_en": "P3 — Missing UI",
        "severity_zh": "P3 — 缺失界面",
    },
    # ========== Tier 5: setup (one-time fix per machine) ==========
    {
        "id": 2,
        "tier": "P4_setup",
        "form": "VBE Project References",
        "title_en": "VBA project references the legacy dao360.dll which isn't on Office 2016+ machines",
        "title_zh": "VBA 工程引用了过时的 dao360.dll，Office 2016+ 机器上没这个文件",
        "summary_en": (
            "The shipped .mdb's VBA project carries a hard reference to "
            "`C:\\Program Files\\Common Files\\Microsoft Shared\\DAO\\"
            "dao360.dll`, which was the DAO 3.6 location used by Access "
            "2003. Modern Office (2016 onward) ships `ACEDAO.DLL` instead "
            "and does NOT install the legacy DLL. On any clean modern "
            "machine, the first attempt to open any LookAt form raises "
            "'Can't find project or library', which is opaque and scary "
            "to end users.\n\n"
            "Severity is low because it's a one-time fix per machine, but "
            "every new install hits it."
        ),
        "summary_zh": (
            ".mdb 中的 VBA 工程硬性引用了 `C:\\Program Files\\Common Files\\"
            "Microsoft Shared\\DAO\\dao360.dll`，这是 Access 2003 时代 DAO "
            "3.6 的位置。现代 Office（2016 起）改用 `ACEDAO.DLL`，并不会安"
            "装旧版 DLL。在任何全新的现代机器上，第一次尝试打开任意 LookAt "
            "表单都会报「找不到工程或库」（Can't find project or library），"
            "对终端用户来说既看不懂又吓人。\n\n"
            "严重等级较低，因为每台机器只需修一次，但每台新装都会撞上。"
        ),
        "steps_en": [
            "Install `CBDB_BJ_User.mdb` on a fresh modern Office machine.",
            "Open the file. Press Alt+F11 to enter the VBE.",
            "Tools → References. Notice an entry marked `MISSING: "
            "dao360.dll`.",
            "Open any LookAt form. A 'Can't find project or library' "
            "error appears.",
        ],
        "steps_zh": [
            "在全新的现代 Office 机器上安装 `CBDB_BJ_User.mdb`。",
            "打开文件，按 Alt+F11 进入 VBE 编辑器。",
            "工具 → 引用。可以看到一行写着 `MISSING: dao360.dll`。",
            "打开任意 LookAt 表单。会弹出「Can't find project or library」"
            "错误。",
        ],
        "fix_en": (
            "Once, on the maintainer's machine, do:\n"
            "  1. Open the .mdb in Access. Press Alt+F11.\n"
            "  2. Tools → References. Untick the MISSING dao360.dll entry.\n"
            "  3. Tick `Microsoft Office 16.0 Access Database Engine "
            "Object Library` (i.e. ACEDAO.DLL).\n"
            "  4. Save the .mdb.\n\n"
            "Then re-distribute the fixed file. Future end users won't "
            "need to do anything."
        ),
        "fix_zh": (
            "在维护者的机器上做一次：\n"
            "  1. 用 Access 打开 .mdb，按 Alt+F11。\n"
            "  2. 工具 → 引用。取消勾选标着 MISSING 的 dao360.dll。\n"
            "  3. 勾选 `Microsoft Office 16.0 Access Database Engine "
            "Object Library`（即 ACEDAO.DLL）。\n"
            "  4. 保存 .mdb。\n\n"
            "然后重新分发修好的文件。以后的终端用户什么都不用做。"
        ),
        "screenshots": [],
        "severity_en": "P4 — One-time setup hurdle on each new machine",
        "severity_zh": "P4 — 每台新机器一次性的安装障碍",
    },
]


# ---------------------------------------------------------------------
# DOCX building helpers
# ---------------------------------------------------------------------

def _add_toc(document: Document, lang: str) -> None:
    """Insert a Word TOC field; Word will offer to update it on open."""
    para = document.add_paragraph()
    run = para.add_run()
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = r'TOC \o "1-2" \h \z \u'
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "separate")
    msg = OxmlElement("w:t")
    if lang == "en":
        msg.text = "Right-click here and choose 'Update Field' to populate."
    else:
        msg.text = t("右键点这里，选「更新域」即可生成完整目录。")
    fld_char3 = OxmlElement("w:fldChar")
    fld_char3.set(qn("w:fldCharType"), "end")
    r_el = run._r
    r_el.append(fld_char1)
    r_el.append(instr)
    r_el.append(fld_char2)
    r_el.append(msg)
    r_el.append(fld_char3)


def _h(document, level, text):
    p = document.add_heading(text, level=level)
    return p


def _bullets(document, items: list[str]) -> None:
    for item in items:
        document.add_paragraph(item, style="List Bullet")


def _numbered(document, items: list[str]) -> None:
    for item in items:
        document.add_paragraph(item, style="List Number")


DRIFT_JSON = REPO / "reports" / "index_drift_examples.json"
CLASSIFICATION_JSON = REPO / "reports" / "index_drift_classification.json"
RULE_CLASSIFICATION_JSON = REPO / "reports" / "index_year_drift_rule_classification.json"
RULE_GROUPS_JSON = REPO / "reports" / "index_year_drift_rule_groups.json"
ADDR_CLASSIFICATION_JSON = REPO / "reports" / "index_addr_drift_classification.json"
CAUSE_SUMMARY_JSON = REPO / "reports" / "index_drift_cause_summary.json"
DEMO_PERSONS_JSON = REPO / "reports" / "demo_persons.json"
KNOWN_BUGS_STATUS_JSON = REPO / "reports" / "known_bugs_status.json"


def _load_demo_persons() -> dict:
    import json as _json
    if DEMO_PERSONS_JSON.exists():
        return _json.loads(DEMO_PERSONS_JSON.read_text(encoding="utf-8"))
    return {}


def _load_bug_test_status() -> dict[int, dict]:
    """Map bug-id → {outcome: 'passed'|'failed'|'mixed', tests: [...],
    when: <iso>}.

    Reads `reports/known_bugs_status.json` produced by
    `pytest tests/test_known_bugs.py --json-report
    --json-report-file=reports/known_bugs_status.json`.

    Convention:
      - test name `test_bug<N>_*` → bug N
      - test name `test_bugs_<lo>_to_<hi>_*` → bugs lo..hi inclusive
      - test name `test_bug_view_statusdata_fy_alias_swap` /
        `test_bug_view_statusdata_fy_value_equals_ly_value` → bug 1
      - test name `test_bug_dao_reference_broken_in_user_mdb` → bug 2

    SAFETY: this function only INFORMS the report; it never causes the
    generator to drop or move content.  The intent is "show the
    maintainer that test results agree with the issue" — not to
    auto-edit the report.
    """
    import json as _json, re as _re
    if not KNOWN_BUGS_STATUS_JSON.exists():
        return {}
    data = _json.loads(
        KNOWN_BUGS_STATUS_JSON.read_text(encoding="utf-8")
    )
    raw_when = data.get("created") or data.get("environment", {}).get(
        "created"
    ) or ""
    when = ""
    if raw_when:
        try:
            from datetime import datetime, timezone
            ts = float(raw_when)
            when = (datetime.fromtimestamp(ts, tz=timezone.utc)
                    .strftime("%Y-%m-%d %H:%M UTC"))
        except (TypeError, ValueError):
            when = str(raw_when)

    by_bug: dict[int, list[tuple[str, str]]] = {}

    def _record(bug_id: int, nodeid: str, outcome: str) -> None:
        by_bug.setdefault(bug_id, []).append((nodeid, outcome))

    for t in data.get("tests", []):
        nodeid = t.get("nodeid", "")
        outcome = t.get("outcome", "")
        # Strip everything before "::" so we just have the test name.
        name = nodeid.rsplit("::", 1)[-1]

        # Cluster: test_bugs_<lo>_to_<hi>_...
        m = _re.match(r"^test_bugs_(\d+)_to_(\d+)_", name)
        if m:
            lo, hi = int(m.group(1)), int(m.group(2))
            for n in range(lo, hi + 1):
                _record(n, nodeid, outcome)
            continue
        # Individual: test_bug<N>_...
        m = _re.match(r"^test_bug(\d+)_", name)
        if m:
            _record(int(m.group(1)), nodeid, outcome)
            continue
        # Special hand-mapped cases (the no-number bug names).
        if "view_statusdata" in name:
            _record(1, nodeid, outcome)
        elif "dao_reference" in name:
            _record(2, nodeid, outcome)

    out: dict[int, dict] = {}
    for bug_id, items in by_bug.items():
        outcomes = {o for _, o in items}
        if outcomes == {"passed"}:
            agg = "passed"
        elif outcomes == {"failed"}:
            agg = "failed"
        else:
            agg = "mixed"
        out[bug_id] = {"outcome": agg, "tests": items, "when": when}
    return out


def _add_index_drift_appendix(doc, is_en: bool, Z) -> None:
    """Render the 'index_year / index_addr drift, NOT a bug' chapter
    using the examples collected by collect_index_year_diffs.py."""
    import json as _json
    if not DRIFT_JSON.exists():
        return
    data = _json.loads(DRIFT_JSON.read_text(encoding="utf-8"))

    title = (
        "Appendix — `c_index_year` / `c_index_addr_id` drift "
        "vs the cbdb-online-main-server snapshot "
        "(differences need per-row classification before being filed as bugs)"
        if is_en else
        "附錄 —— `c_index_year` / `c_index_addr_id` 與 "
        "cbdb-online-main-server 快照之間的偏差"
        "（差異需要逐筆分類後才能判定是否為缺陷）"
    )
    _h(doc, 1, Z(title))

    intro = (
        "When we compare BIOG_MAIN's `c_index_year` and "
        "`c_index_addr_id` between this User MDB and the weekly "
        "cbdb-online-main-server SQLite snapshot, a small fraction "
        "of persons disagree.\n\n"
        "**The two sides are independent implementations.** The "
        "SQLite snapshot's `c_index_year` is produced by "
        "cbdb-online-main-server's PHP "
        "`IndexYearRebuildService.php` and its `c_index_addr_id` by "
        "`IndexAddressRebuildService.php` (both at "
        "<https://github.com/cbdb-project/cbdb-online-main-server>); "
        "the User MDB-side: `c_index_addr_id` is rebuilt by VBA in "
        "`Form_frmIndexAddr` in the front-end mdb; `c_index_year` "
        "is rebuilt by **37 saved QueryDefs named `BM IY Rule …`** "
        "in the linked-tables backend `data/CBDB_<YYYYMMDD>_DATA"
        ".mdb`, driven by `frmBaseMaintenance`.  Both algorithms "
        "are now extracted and committed to the repo "
        "(`analysis/dump_data/querydefs_index/*.sql`); the form / "
        "module driver VBA still needs an interactive Access "
        "`SaveAsText` pass.  PHP is "
        "intended to mirror the VBA but they are separate code "
        "paths.  Per-row "
        "differences can come from at least four sources, and a "
        "diff alone doesn't tell us which: (1) source-data "
        "snapshot drift (the online system updates source rows "
        "continuously; the User MDB ships a point-in-time "
        "snapshot); (2) algorithm / porting divergence between "
        "PHP and VBA; (3) priority / tie-break differences when "
        "multiple candidate years apply; (4) null / default "
        "handling differences (e.g. how a missing birthyear "
        "collapses).\n\n"
        "**We have not classified the steady ~575 / 657 246 "
        "diffs we currently observe.**  The examples below are a "
        "small sample (currently 13 rows across 3 buckets, "
        "from `reports/index_drift_examples.json`) — illustrative "
        "of the *shapes* of disagreement, not statistically "
        "representative.  Please don't treat them as a verdict "
        "either way; they're a starting point for whoever does "
        "the per-row triage.\n\n"
        "Why we still keep this appendix: if the *shape* of "
        "disagreement changes between releases (e.g. addresses "
        "start diverging where only years used to, or per-rule "
        "type-code distributions shift sharply), that hints the "
        "algorithm, schema, or one of the implementations moved "
        "— a stronger signal than per-row counts alone."
        if is_en else
        "我們把本 .mdb 的 BIOG_MAIN 與 cbdb-online-main-server 每週"
        "釋出的 SQLite 快照在 `c_index_year`、`c_index_addr_id` 兩個"
        "欄位上做比對，可以看到一小部分人物對不齊。\n\n"
        "**兩邊是兩套獨立的實作。**SQLite 快照中的 `c_index_year` "
        "是 cbdb-online-main-server 的 PHP "
        "`IndexYearRebuildService.php` 算出來的，`c_index_addr_id` "
        "則是 `IndexAddressRebuildService.php` 算出來的（程式碼都"
        "在 <https://github.com/cbdb-project/cbdb-online-main-server>"
        "）；User MDB 上對應的這兩個欄位則是由 Access "
        "**User MDB 那一邊**：`c_index_addr_id` 是由前端 mdb 裡的 "
        "`Form_frmIndexAddr` VBA 重建的；`c_index_year` 則是由"
        "連結表後端 `data/CBDB_<YYYYMMDD>_DATA.mdb` 裡 **37 條"
        "命名為 `BM IY Rule …` 的 QueryDef** 重建的，並由 "
        "`frmBaseMaintenance` 驅動。兩邊的演算法都已抽取並提交到 "
        "repo（`analysis/dump_data/querydefs_index/*.sql`），form / "
        "module 的驅動 VBA 仍需透過 Access 的 `SaveAsText` 互動式"
        "提取。PHP "
        "**意圖**鏡像 VBA，但兩者是兩條獨立的程式路徑。每一行"
        "差異**可能**來自下列至少四"
        "個原因，光看差異本身分不出來：(1) 源資料快照漂移"
        "（線上系統持續更新源資料；User MDB 是某個時點的快照）；"
        "(2) PHP 與 VBA 之間的演算法 / 移植差異；(3) 多個候選年份"
        "適用時，兩邊的優先序 / 平手規則不同；(4) null / 預設值"
        "處理不同（例如缺失的生年是當作 0、NULL，還是「卒年減 60」）。"
        "\n\n"
        "**我們並沒有對目前看到的 ~575 / 657 246 筆差異做完整分類。**"
        "下方列舉的樣本（目前共 13 筆、3 種分桶，"
        "來自 `reports/index_drift_examples.json`）只是**示範**這些"
        "差異**長什麼樣**，並非統計上有代表性。請不要把這些樣例"
        "當成任何方向的結論，它們只是後續逐筆分類的起點。\n\n"
        "我們仍然把這份附錄放在這裡，是因為：若差異的「形狀」在"
        "不同次發行之間發生變化（例如原本只年份對不齊，現在地點"
        "也開始對不齊；或 per-rule type-code 的分佈突然偏離），"
        "那是比逐筆計數更強的訊號，意味著演算法、schema、或某一"
        "邊的實作動過。"
    )
    for para in intro.split("\n\n"):
        doc.add_paragraph(Z(para))

    # ---- Classification summary ----
    if CLASSIFICATION_JSON.exists():
        cls = _json.loads(CLASSIFICATION_JSON.read_text(encoding="utf-8"))
        cs = cls["summary"]
        b = cs["buckets"]
        if is_en:
            _h(doc, 2, Z("Classification summary"))
            doc.add_paragraph(Z(
                f"Compared {cs['common']:,} personids common to both "
                f"databases (User MDB total {cs['user_mdb_total']:,}; "
                f"SQLite total {cs['sqlite_total']:,}; "
                f"User-only {cs['in_user_only']:,}; "
                f"SQLite-only {cs['in_sqlite_only']:,})."
            ))
            for b_key, label in [
                ("exact_match",
                 "exact match on all four compared fields"),
                ("source_drift_index_agrees",
                 "source drift but indices agreed (algorithms "
                 "tolerated the drift)"),
                ("source_drift_index_diffs_too",
                 "source drift AND at least one index field "
                 "differs (consistent with simple data-drift "
                 "hypothesis)"),
                ("index_year_only_diff",
                 "c_birthyear+c_deathyear identical, only "
                 "c_index_year differs — needs follow-up"),
                ("index_addr_only_diff",
                 "c_birthyear+c_deathyear identical, only "
                 "c_index_addr_id differs — needs follow-up"),
                ("index_both_diff",
                 "c_birthyear+c_deathyear identical, BOTH index "
                 "fields differ — strongest signal of compound "
                 "divergence, needs follow-up"),
            ]:
                doc.add_paragraph(Z(
                    f"  • {b[b_key]:>7,} ({100.0*b[b_key]/max(cs['common'],1):.3f}%)"
                    f"  — {label}"
                ))
            doc.add_paragraph(Z(
                f"Net diffs: {cs['common']-b['exact_match']:,} of "
                f"{cs['common']:,} ({100.0*(cs['common']-b['exact_match'])/max(cs['common'],1):.3f}%).  "
                f"Of those, {b['source_drift_index_agrees']+b['source_drift_index_diffs_too']:,} are "
                f"clearly attributable to source drift in birthyear/"
                f"deathyear; {b['index_year_only_diff']+b['index_addr_only_diff']+b['index_both_diff']:,} "
                f"need per-row follow-up (could be PHP↔VBA "
                f"divergence, or drift in evidence tables — "
                f"BIOG_ADDR_DATA / ENTRY_DATA / NIAN_HAO etc. — "
                f"that this classifier does not compare).  Full "
                f"output: `reports/index_drift_classification.json`; "
                f"algorithm pointers: "
                f"`analysis/index_drift_algorithm_notes.md`."
            ))
        else:
            _h(doc, 2, Z("分類匯總"))
            doc.add_paragraph(Z(
                f"比對了兩邊都有的 {cs['common']:,} 個 personid"
                f"（User MDB 共 {cs['user_mdb_total']:,} 筆；"
                f"SQLite 共 {cs['sqlite_total']:,} 筆；"
                f"僅 User MDB 有 {cs['in_user_only']:,} 筆；"
                f"僅 SQLite 有 {cs['in_sqlite_only']:,} 筆）。"
            ))
            for b_key, label in [
                ("exact_match", "四個欄位全部一致"),
                ("source_drift_index_agrees",
                 "源資料有漂移但兩邊 index 都一致"
                 "（演算法吸收了漂移）"),
                ("source_drift_index_diffs_too",
                 "源資料有漂移、且至少一個 index 不同"
                 "（與簡單資料漂移假說相符）"),
                ("index_year_only_diff",
                 "生年/卒年一致，但只有 c_index_year 不同 —— 待追查"),
                ("index_addr_only_diff",
                 "生年/卒年一致，但只有 c_index_addr_id 不同 —— 待追查"),
                ("index_both_diff",
                 "生年/卒年一致，但兩個 index 都不同 —— 複合差異"
                 "的最強單列訊號，待追查"),
            ]:
                doc.add_paragraph(Z(
                    f"  • {b[b_key]:>7,} ({100.0*b[b_key]/max(cs['common'],1):.3f}%)"
                    f"  —— {label}"
                ))
            doc.add_paragraph(Z(
                f"淨差異：{cs['common']-b['exact_match']:,} / "
                f"{cs['common']:,}（{100.0*(cs['common']-b['exact_match'])/max(cs['common'],1):.3f}%）。"
                f"其中 {b['source_drift_index_agrees']+b['source_drift_index_diffs_too']:,} "
                f"筆能明確歸因於 birthyear/deathyear 的源資料漂移；"
                f"剩下 {b['index_year_only_diff']+b['index_addr_only_diff']+b['index_both_diff']:,} "
                f"筆需要逐筆追查（可能是 PHP↔VBA 演算法差異，"
                f"也可能是本分類器沒有比較的 evidence 表"
                f"（BIOG_ADDR_DATA / ENTRY_DATA / NIAN_HAO 等）裡的漂移）。"
                f"完整輸出見 `reports/index_drift_classification.json`，"
                f"算法來源指標見 `analysis/index_drift_algorithm_notes.md`。"
            ))

    # ---- Year-drift rule classification (PR K1) ----
    if RULE_CLASSIFICATION_JSON.exists():
        rcls = _json.loads(
            RULE_CLASSIFICATION_JSON.read_text(encoding="utf-8"))
        rs = rcls["summary"]
        rb = rs["buckets"]
        if is_en:
            _h(doc, 2, Z(
                "Year-only diffs — per-row rule classification"
            ))
            doc.add_paragraph(Z(
                f"Of the {rs['total_year_diffs']} year-only diffs, "
                f"each row was bucketed against PR N's rule-level "
                f"runtime-vs-PHP comparison "
                f"(`analysis/index_year_rule_comparison.md`).  "
                f"Conservative buckets (rows count once each):"
            ))
            for k in [
                "php_returned_sentinel",
                "php_did_not_compute",
                "access_did_not_compute",
                "iteration_order_diff",
                "consistent_within_rule",
                "candidate_algorithm_divergence",
                "unclassified",
            ]:
                if rb[k] == 0:
                    continue
                doc.add_paragraph(Z(
                    f"  • {rb[k]:>3,}  {k}"
                ))
            doc.add_paragraph(Z(
                f"None of these are confirmed bugs.  Full per-row "
                f"output (with the source-data evidence we used to "
                f"bucket each one) is in "
                f"`reports/index_year_drift_rule_classification.json`."
            ))
            # ---- K2 deeper triage ----
            if RULE_GROUPS_JSON.exists():
                gcls = _json.loads(
                    RULE_GROUPS_JSON.read_text(encoding="utf-8"))
                gs = gcls["summary"]
                doc.add_paragraph(Z(
                    f"Deeper triage (PR K2, "
                    f"`analysis/triage_index_year_drift_groups.py`) "
                    f"named the leftover buckets:"
                ))
                doc.add_paragraph(Z(
                    f"  • {gs['consistent_within_rule']['rows']} "
                    f"`consistent_within_rule` rows → "
                    f"{gs['consistent_within_rule']['groups']} "
                    f"signature groups.  PR AI + AJ probes "
                    f"reversed the prior tie-break hypothesis: all "
                    f"14 rows are `source_data_drift_biog_main_or_"
                    f"kin_data_between_sides` (8 BIOG_MAIN birthyear "
                    f"drift + 6 KIN_DATA evidence-pid drift).  "
                    f"Upstream PHP-side / SQLite-snapshot data "
                    f"issue, not a CBDB algorithm divergence."
                ))
                doc.add_paragraph(Z(
                    f"  • {gs['unclassified']['total']} "
                    f"`unclassified` rows → "
                    f"{gs['unclassified']['named_after_triage']} named, "
                    f"{gs['unclassified']['blocked_by_runtime_priority_triage_pending']} "
                    f"flagged `blocked_by_runtime_priority_triage_pending` "
                    f"(PR M dumped frmBaseMaintenance, so the source is in repo; resolving each row still needs a per-row walk of the runtime priority/iteration order)."
                ))
                doc.add_paragraph(Z(
                    f"  • {gs['php_did_not_compute']['rows']} "
                    f"`php_did_not_compute` rows → "
                    f"{gs['php_did_not_compute']['groups']} groups by "
                    f"Access tcode; biggest is `access_tcode='05'` × 7 "
                    f"(`candidate_php_entry_code_mapping_gap` for jinshi)."
                ))
                doc.add_paragraph(Z(
                    f"Full per-group output: "
                    f"`reports/index_year_drift_rule_groups.json`."
                ))
            # ---- Address-drift classification (PR L) ----
            if ADDR_CLASSIFICATION_JSON.exists():
                acls = _json.loads(
                    ADDR_CLASSIFICATION_JSON.read_text(encoding="utf-8"))
                aS = acls["summary"]
                ab = aS["buckets"]
                _h(doc, 2, Z(
                    "c_index_addr_id diffs — per-row classification"
                ))
                doc.add_paragraph(Z(
                    f"Of the {aS['total_addr_diffs']} c_index_addr "
                    f"diffs (478 `index_addr_only_diff` + 10 "
                    f"`index_both_diff` from PR G's bucketing), "
                    f"each row was classified by re-simulating the "
                    f"rank-priority + MAX(c_sequence) algorithm "
                    f"against each side's BIOG_ADDR_DATA + the "
                    f"shared BIOG_ADDR_CODES rank table:"
                ))
                for k, label in [
                    ("mdb_stale_index_addr",
                     "User MDB stored value doesn't match its own "
                     "recompute; SQLite does — User MDB c_index_addr_id "
                     "is stale (re-run frmBaseMaintenance)"),
                    ("mdb_value_php_null",
                     "User MDB has a value, PHP wrote 0/null"),
                    ("same_candidates_diff_winner",
                     "identical BIOG_ADDR_DATA, different winner — "
                     "candidate algorithm divergence (all in addr_type=1)"),
                    ("both_stale_recompute_mismatch",
                     "neither side matches recompute"),
                    ("both_sides_match_recomputed",
                     "both stored values match recompute over their "
                     "own BIOG_ADDR_DATA — diff is source-data drift"),
                    ("sqlite_stale_index_addr",
                     "PHP stored value doesn't match its own recompute"),
                    ("mdb_null_php_value", ""),
                    ("unclassified", ""),
                ]:
                    n = ab.get(k, 0)
                    if n == 0:
                        continue
                    doc.add_paragraph(Z(
                        f"  • {n:>4,d}  {k}{(' — ' + label) if label else ''}"
                    ))
                doc.add_paragraph(Z(
                    f"None of these are confirmed bugs.  The 412 "
                    f"`mdb_stale_index_addr` rows are a maintenance-"
                    f"cadence diff (the User MDB needs its "
                    f"frmBaseMaintenance rebuild re-run before the "
                    f"next release).  The 10 "
                    f"`same_candidates_diff_winner` rows are the only "
                    f"candidate algorithm-divergence rows.  Full "
                    f"per-row output: "
                    f"`reports/index_addr_drift_classification.json`."
                ))
                doc.add_paragraph(Z(
                    f"PR M (`analysis/dump_data_mdb_vba.py`) extracted "
                    f"`frmBaseMaintenance.CmdIndexAddress_Click` from "
                    f"the DATA mdb.  It does NOT explicitly "
                    f"`MAX(c_sequence)`-aggregate the way PHP does — "
                    f"a candidate algorithmic divergence on top of "
                    f"the maintenance-cadence issue.  Suggested "
                    f"release-checklist mitigation: run "
                    f"`CmdIndexYear` then `CmdIndexAddress` on the "
                    f"DATA mdb before shipping a new User MDB.  See "
                    f"`analysis/index_drift_algorithm_notes.md` § "
                    f"\"Maintenance trigger path\" for the full "
                    f"write-up."
                ))

            # ---- Cause analysis appendix (PR Y) ----
            if CAUSE_SUMMARY_JSON.exists():
                cs = _json.loads(
                    CAUSE_SUMMARY_JSON.read_text(encoding="utf-8"))
                _h(doc, 2, "What currently explains the drift")
                doc.add_paragraph(Z(
                    "Per-bucket cause / supporting evidence / "
                    "confidence / next action lives in "
                    "`analysis/index_drift_cause_analysis.md`.  This "
                    "section just summarises the headline counts and "
                    "confidence per bucket; no bucket is labelled a "
                    "confirmed CBDB bug."
                ))
                # c_index_year table
                _h(doc, 3, "c_index_year cause buckets")
                tbl = doc.add_table(rows=1, cols=3)
                tbl.style = "Light Grid Accent 1"
                hdr = tbl.rows[0].cells
                hdr[0].text = "Bucket"
                hdr[1].text = "Count"
                hdr[2].text = "Confidence"
                for b in cs["c_index_year"]["buckets"]:
                    if b["count"] == 0:
                        continue
                    row = tbl.add_row().cells
                    row[0].text = b["bucket"]
                    row[1].text = str(b["count"])
                    row[2].text = b["confidence"]
                doc.add_paragraph("")
                # c_index_addr_id table
                _h(doc, 3, "c_index_addr_id cause buckets")
                tbl = doc.add_table(rows=1, cols=3)
                tbl.style = "Light Grid Accent 1"
                hdr = tbl.rows[0].cells
                hdr[0].text = "Bucket"
                hdr[1].text = "Count"
                hdr[2].text = "Confidence"
                for b in cs["c_index_addr_id"]["buckets"]:
                    if b["count"] == 0:
                        continue
                    row = tbl.add_row().cells
                    row[0].text = b["bucket"]
                    row[1].text = str(b["count"])
                    row[2].text = b["confidence"]
                doc.add_paragraph("")
                doc.add_paragraph(Z(
                    "Top suggested next investigations (full list in "
                    "the cause-analysis md):"
                ))
                for inv in cs[
                    "suggested_next_investigations_in_priority_order"
                ][:3]:
                    doc.add_paragraph(Z(
                        f"  {inv['id']}. {inv['task']} — would close "
                        f"{inv['would_close_rows']} rows; engineering "
                        f"cost: {inv['engineering_cost']}."
                    ))
        else:
            _h(doc, 2, Z("年份差異 —— 逐筆 rule 分類"))
            doc.add_paragraph(Z(
                f"在 {rs['total_year_diffs']} 筆「只有 c_index_year "
                f"不一致」的行中，逐筆比對 PR N "
                f"(`analysis/index_year_rule_comparison.md`) 的"
                f"runtime-vs-PHP 規則級差異。保守分類如下："
            ))
            label_zh = {
                "php_returned_sentinel": "PHP 寫了 sentinel／溢位值",
                "php_did_not_compute": "PHP 沒算出值（覆蓋率缺口）",
                "access_did_not_compute": "Access 沒算出值（覆蓋率缺口）",
                "iteration_order_diff": "Phase-C 迭代次數不同",
                "consistent_within_rule": "多列共享同一 (php_tcode, access_tcode, diff) — 單一規則差異訊號",
                "candidate_algorithm_divergence": "形狀符合 K1 的歷史 hypothesis probe 但無法以單筆證據重建",
                "unclassified": "尚未對上任何模式",
            }
            for k in [
                "php_returned_sentinel",
                "php_did_not_compute",
                "access_did_not_compute",
                "iteration_order_diff",
                "consistent_within_rule",
                "candidate_algorithm_divergence",
                "unclassified",
            ]:
                if rb[k] == 0:
                    continue
                doc.add_paragraph(Z(
                    f"  • {rb[k]:>3,}  {label_zh[k]}"
                ))
            doc.add_paragraph(Z(
                f"以上沒有任何一筆被視為已確認的 bug。"
                f"逐筆輸出（含分桶所依據的源資料證據）見 "
                f"`reports/index_year_drift_rule_classification.json`。"
            ))
            # ---- K2 deeper triage (zh) ----
            if RULE_GROUPS_JSON.exists():
                gcls = _json.loads(
                    RULE_GROUPS_JSON.read_text(encoding="utf-8"))
                gs = gcls["summary"]
                doc.add_paragraph(Z(
                    f"PR K2 進一步的 triage "
                    f"(`analysis/triage_index_year_drift_groups.py`) 把"
                    f"剩下的桶命名清楚："
                ))
                doc.add_paragraph(Z(
                    f"  • `consistent_within_rule` "
                    f"{gs['consistent_within_rule']['rows']} 筆 → "
                    f"{gs['consistent_within_rule']['groups']} 個 signature "
                    f"分組。PR AI + AJ 的逐筆探測推翻了原本的 tie-break "
                    f"假說：14 筆全是 `source_data_drift_biog_main_or_"
                    f"kin_data_between_sides`（8 筆 BIOG_MAIN birthyear "
                    f"漂移 + 6 筆 KIN_DATA evidence-pid 漂移）。屬於 "
                    f"PHP-side / SQLite snapshot 的上游資料漂移，並非 "
                    f"CBDB 演算法差異。"
                ))
                doc.add_paragraph(Z(
                    f"  • `unclassified` {gs['unclassified']['total']} 筆 → "
                    f"{gs['unclassified']['named_after_triage']} 筆已命名，"
                    f"{gs['unclassified']['blocked_by_runtime_priority_triage_pending']} 筆"
                    f"標為 `blocked_by_runtime_priority_triage_pending`"
                    f"（PR M 已 dump frmBaseMaintenance，源碼已在 repo；要逐筆判斷哪邊正確仍需走一遍 runtime 的 priority／iteration 順序）。"
                ))
                doc.add_paragraph(Z(
                    f"  • `php_did_not_compute` "
                    f"{gs['php_did_not_compute']['rows']} 筆 → 按 Access "
                    f"tcode 分 {gs['php_did_not_compute']['groups']} 組；"
                    f"最大的是 `access_tcode='05'` × 7（jinshi 進士類的 "
                    f"`candidate_php_entry_code_mapping_gap`）。"
                ))
                doc.add_paragraph(Z(
                    f"逐組輸出見 "
                    f"`reports/index_year_drift_rule_groups.json`。"
                ))
            # ---- Address-drift classification (PR L) — zh ----
            if ADDR_CLASSIFICATION_JSON.exists():
                acls = _json.loads(
                    ADDR_CLASSIFICATION_JSON.read_text(encoding="utf-8"))
                aS = acls["summary"]
                ab = aS["buckets"]
                _h(doc, 2, Z("c_index_addr_id 差異 —— 逐筆分類"))
                doc.add_paragraph(Z(
                    f"在 {aS['total_addr_diffs']} 筆 c_index_addr 差異"
                    f"中（PR G 分桶的 478 `index_addr_only_diff` + 10 "
                    f"`index_both_diff`），逐筆把兩邊的 BIOG_ADDR_DATA "
                    f"代入「rank-priority + MAX(c_sequence)」演算法重算，"
                    f"與實際儲存值對照分類："
                ))
                label_zh = {
                    "mdb_stale_index_addr":
                        "User MDB 儲存值與重算結果不符；SQLite 相符 "
                        "—— User MDB 的 c_index_addr_id 已 stale，"
                        "需要重新跑 frmBaseMaintenance",
                    "mdb_value_php_null": "User MDB 有值，PHP 寫了 0/null",
                    "same_candidates_diff_winner":
                        "兩邊 BIOG_ADDR_DATA 完全相同但 winner 不同 "
                        "—— 候選演算法差異（全在 addr_type=1）",
                    "both_stale_recompute_mismatch":
                        "兩邊都不符合重算結果",
                    "both_sides_match_recomputed":
                        "兩邊各自的儲存值都符合自己的重算結果 "
                        "—— 差異來自源資料漂移",
                    "sqlite_stale_index_addr":
                        "PHP 儲存值與重算結果不符",
                    "mdb_null_php_value": "",
                    "unclassified": "",
                }
                for k in [
                    "mdb_stale_index_addr",
                    "mdb_value_php_null",
                    "same_candidates_diff_winner",
                    "both_stale_recompute_mismatch",
                    "both_sides_match_recomputed",
                    "sqlite_stale_index_addr",
                    "mdb_null_php_value",
                    "unclassified",
                ]:
                    n = ab.get(k, 0)
                    if n == 0:
                        continue
                    lbl = label_zh.get(k, "")
                    doc.add_paragraph(Z(
                        f"  • {n:>4,d}  {k}{(' — ' + lbl) if lbl else ''}"
                    ))
                doc.add_paragraph(Z(
                    f"以上沒有任何一筆被視為已確認的 bug。412 筆 "
                    f"`mdb_stale_index_addr` 屬於維護週期差異（User MDB "
                    f"在下次釋出前需要重跑 frmBaseMaintenance）。10 筆 "
                    f"`same_candidates_diff_winner` 是唯一的候選演算法"
                    f"差異。逐筆輸出見 "
                    f"`reports/index_addr_drift_classification.json`。"
                ))
                doc.add_paragraph(Z(
                    f"PR M（`analysis/dump_data_mdb_vba.py`）從 DATA mdb "
                    f"抽出了 `frmBaseMaintenance.CmdIndexAddress_Click`。"
                    f"它**沒有**像 PHP 那樣明確 `MAX(c_sequence)` 聚合 "
                    f"—— 在維護週期差異之外，這還是一個候選演算法差異。"
                    f"建議的 release checklist 緩解步驟：在 User MDB "
                    f"出貨前先在 DATA mdb 上跑 `CmdIndexYear`，再跑 "
                    f"`CmdIndexAddress`。詳見 "
                    f"`analysis/index_drift_algorithm_notes.md` 中的 "
                    f"\"Maintenance trigger path\" 段。"
                ))
                doc.add_paragraph(Z(
                    f"PR S（`analysis/deep_dive_addr_same_candidates.py`）"
                    f"證實 10 筆 `same_candidates_diff_winner` 全部"
                    f"是 MAX(c_sequence) 的 tie 問題（同一(person, "
                    f"addr_type) 有多筆 BIOG_ADDR_DATA 行 c_sequence "
                    f"並列最大）。PHP、Access、以及我方重算各自挑了"
                    f"不同的 row。兩邊都遵循同一條文件規則，沒有「錯」"
                    f"的一方。候選緩解：兩邊都加一條明確的二級 "
                    f"tie-break（如 MIN(c_addr_id)）。逐筆證據見 "
                    f"`reports/index_addr_same_candidates_deep_dive.json`。"
                ))

            # ---- Cause analysis appendix (PR Y) ----
            if CAUSE_SUMMARY_JSON.exists():
                cs = _json.loads(
                    CAUSE_SUMMARY_JSON.read_text(encoding="utf-8"))
                _h(doc, 2, Z("目前能解釋的 drift 原因"))
                doc.add_paragraph(Z(
                    "每個 bucket 的成因／證據／信心度／下一步追查"
                    "都寫在 `analysis/index_drift_cause_analysis.md`。"
                    "本節只列每個 bucket 的計數和信心度摘要；目前"
                    "沒有任何 bucket 被列為已確認的 CBDB bug。"
                ))
                _h(doc, 3, Z("c_index_year 原因桶"))
                tbl = doc.add_table(rows=1, cols=3)
                tbl.style = "Light Grid Accent 1"
                hdr = tbl.rows[0].cells
                hdr[0].text = Z("Bucket")
                hdr[1].text = Z("筆數")
                hdr[2].text = Z("信心度")
                for b in cs["c_index_year"]["buckets"]:
                    if b["count"] == 0:
                        continue
                    row = tbl.add_row().cells
                    row[0].text = b["bucket"]
                    row[1].text = str(b["count"])
                    row[2].text = b["confidence"]
                doc.add_paragraph("")
                _h(doc, 3, Z("c_index_addr_id 原因桶"))
                tbl = doc.add_table(rows=1, cols=3)
                tbl.style = "Light Grid Accent 1"
                hdr = tbl.rows[0].cells
                hdr[0].text = Z("Bucket")
                hdr[1].text = Z("筆數")
                hdr[2].text = Z("信心度")
                for b in cs["c_index_addr_id"]["buckets"]:
                    if b["count"] == 0:
                        continue
                    row = tbl.add_row().cells
                    row[0].text = b["bucket"]
                    row[1].text = str(b["count"])
                    row[2].text = b["confidence"]
                doc.add_paragraph("")
                doc.add_paragraph(Z(
                    "建議優先處理的調查項目（完整列表見 cause-analysis md）："
                ))
                for inv in cs[
                    "suggested_next_investigations_in_priority_order"
                ][:3]:
                    doc.add_paragraph(Z(
                        f"  {inv['id']}. {inv['task']} —— 可消化 "
                        f"{inv['would_close_rows']} 筆；工程成本："
                        f"{inv['engineering_cost']}。"
                    ))

    # ---- Per bucket ----
    bucket_meta = {
        "year_only": {
            "title_en": "Examples where only c_index_year disagrees",
            "title_zh": "仅 `c_index_year` 不一致的样例",
            "explain_en": (
                "Same person, same source data, but the two pipelines "
                "picked different priority rules. The User MDB's "
                "`c_index_year_type_code` and `c_index_year_source_id` "
                "differ from the snapshot's, which means each pipeline "
                "selected a different upstream evidence row when "
                "deciding 'which year is most representative for this "
                "person'."
            ),
            "explain_zh": (
                "同一个人物、同一份源数据，但两套管线选了不同的优先级规则。"
                "User MDB 的 `c_index_year_type_code` 和 "
                "`c_index_year_source_id` 与快照不同，意味着两边在「为这个"
                "人物选一个最能代表生平的年份」时，挑中了不同的上游证据行。"
            ),
        },
        "addr_only": {
            "title_en": "Examples where only c_index_addr_id disagrees",
            "title_zh": "仅 `c_index_addr_id` 不一致的样例",
            "explain_en": (
                "Year agrees, address doesn't. Most commonly the User "
                "MDB still has an older `c_index_addr_id` choice while "
                "the snapshot has been re-classified to a finer-grained "
                "or higher-quality address (or to NULL when the "
                "evidence was downgraded). The two pipelines apply the "
                "same address-priority rules but to different snapshots "
                "of `BIOG_ADDR_DATA`."
            ),
            "explain_zh": (
                "年份对得上，但地点对不上。最常见的情况是：User MDB 仍保留"
                "较早一次的 `c_index_addr_id` 选择，而快照已经重新分类到"
                "更细粒度或证据等级更高的地点（或在证据被下调时改为 NULL）。"
                "两套管线用的是同一套地点优先级规则，只是依据的 "
                "`BIOG_ADDR_DATA` 快照不同。"
            ),
        },
        "both": {
            "title_en": "Examples where both fields disagree",
            "title_zh": "两个字段都不一致的样例",
            "explain_en": (
                "Both year and address moved together — usually because "
                "the snapshot got new high-priority biographical "
                "evidence (a newly-entered birth year or a more "
                "specific address from a different source text)."
            ),
            "explain_zh": (
                "年份和地点同时变动 —— 通常是因为快照在某个高优先级的传记"
                "证据行上有新增（例如新录入的生年，或来自其他原始文献的更"
                "精确地点）。"
            ),
        },
        "source_data": {
            "title_en": (
                "Examples where the SOURCE data itself differs "
                "(birthyear / deathyear)"
            ),
            "title_zh": (
                "底层 SOURCE 数据本身不同（生年 / 卒年）的样例"
            ),
            "explain_en": (
                "Here the SQLite snapshot has different `c_birthyear` / "
                "`c_deathyear` from the User MDB. This is the clearest "
                "case of pure data drift: someone updated the source "
                "row in the cbdb-online-main-server pipeline after the "
                "User MDB was last exported. There's no algorithmic "
                "disagreement here — just two different snapshots in "
                "time."
            ),
            "explain_zh": (
                "在这一组里，SQLite 快照的 `c_birthyear` / `c_deathyear` "
                "本身就和 User MDB 不一样。这是最纯粹的数据快照差异：在 "
                "User MDB 最近一次导出之后，有人在 "
                "cbdb-online-main-server 那边更新了源数据行。这里完全没有"
                "算法分歧 —— 只是两个时间点的两份快照。"
            ),
        },
    }

    for bucket_key, meta in bucket_meta.items():
        items = data.get(bucket_key, [])
        if not items:
            continue
        _h(doc, 2, Z(meta["title_en"] if is_en else meta["title_zh"]))
        for para in (meta["explain_en"] if is_en
                     else meta["explain_zh"]).split("\n\n"):
            doc.add_paragraph(Z(para))

        # Render each example as a small table.
        for ex in items[:5]:
            u = ex["user"]; s = ex["sqlite"]
            label = f"c_personid = {ex['personid']} — {ex['name_chn']} ({ex['name_py']})"
            _h(doc, 3, Z(label))
            tbl = doc.add_table(rows=1, cols=3)
            tbl.style = "Light Grid Accent 1"
            hdr = tbl.rows[0].cells
            hdr[0].text = Z("Field" if is_en else "字段")
            hdr[1].text = Z("User MDB" if is_en else "本 .mdb (User MDB)")
            hdr[2].text = Z(
                "cbdb-online-main-server snapshot"
                if is_en else
                "cbdb-online-main-server 快照"
            )
            for f_label, key in [
                ("c_index_year", "index_year"),
                ("c_index_addr_id", "index_addr_id"),
                ("c_birthyear", "birthyear"),
                ("c_deathyear", "deathyear"),
                ("c_index_year_type_code", "index_year_type_code"),
                ("c_index_year_source_id", "index_year_source_id"),
            ]:
                row = tbl.add_row().cells
                row[0].text = f_label
                row[1].text = "" if u[key] is None else str(u[key])
                row[2].text = "" if s[key] is None else str(s[key])
            doc.add_paragraph("")  # spacer

    note = (
        "These examples were collected by "
        "`reports/collect_index_year_diffs.py` from the first 20 000 "
        "person ids common to both databases. The full automated "
        "cross-check (`tests/test_index_year_xcheck.py`, ~657 246 "
        "persons) reports roughly the same shape: very small "
        "fractions disagree, with addresses being the most common "
        "drift type."
        if is_en else
        "这批样例是 `reports/collect_index_year_diffs.py` 在两边数据库共有"
        "的前 20000 个 personid 中采集的。完整的自动化对照（"
        "`tests/test_index_year_xcheck.py`，约 657246 人）给出相同的形态："
        "对不齐的比例非常小，其中地点偏差占多数。"
    )
    doc.add_paragraph(Z(note))


def _build(lang: str, out_path: Path) -> None:
    is_en = (lang == "en")

    # Apply Simplified -> Traditional Chinese conversion when emitting
    # the Chinese variant.  English passes through unchanged.
    def Z(s: str) -> str:
        return s if is_en else t(s)

    doc = Document()

    # ---- Cover page ----
    title = (
        "CBDB User MDB — Issues Report"
        if is_en else
        "CBDB 用户版 .mdb — 问题汇报"
    )
    subtitle = (
        "A respectful summary of issues uncovered during regression testing."
        if is_en else
        "测试过程中发现的问题汇总，谨呈维护团队斧正。"
    )
    h = doc.add_heading(Z(title), level=0)
    h.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p = doc.add_paragraph(Z(subtitle))
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    intro = (
        "Dear maintainer,\n\n"
        "Below is a summary of the issues we uncovered while building "
        "an automated regression-test suite for the CBDB User MDB. "
        "We hope this report is useful as you continue your wonderful "
        "stewardship of this dataset, and we sincerely thank you for "
        "the immense work that has gone into building it.\n\n"
        "The issues are ordered by severity (P0 highest). Each entry "
        "includes a concise description, step-by-step user reproduction, "
        "screenshots where the issue is visible in the Access UI, and a "
        "suggested fix. None of these are urgent; they are documented "
        "so they can be addressed at the maintainer's convenience."
        if is_en else
        "尊敬的维护者：\n\n"
        "下面是我们在为 CBDB 用户版 .mdb 编写自动化回归测试套件过程中，"
        "陆续整理出来的一些问题清单。我们希望这份报告能在您继续主持这份"
        "宝贵数据集时有所帮助；同时，对您多年来在这套数据上的辛勤付出，"
        "我们由衷地表示感谢和敬意。\n\n"
        "问题按严重程度排序（P0 最高）。每一条都包括：简明描述、用户端"
        "一步一步的复现步骤、（在界面上能看到时）相关截图，以及一份建议"
        "的修复方案。这些问题并不紧急，整理在此只是为了方便您在合适的时"
        "候逐一处理。"
    )
    for para in intro.split("\n\n"):
        doc.add_paragraph(Z(para))

    doc.add_page_break()

    # ---- Table of contents ----
    _h(doc, 1, Z("Table of Contents" if is_en else "目录"))
    _add_toc(doc, lang)
    doc.add_page_break()

    # ---- Severity legend ----
    _h(doc, 1, Z("Severity legend" if is_en else "严重等级说明"))
    legend_en = [
        "P0 — Silent data corruption: data is wrong or missing without an error popup.",
        "P1 — Visible runtime crash: a popup appears, the operation aborts.",
        "P2 — Silent display: form fields render blank when they should show data.",
        "P3 — Missing UI: a feature exists in code but no button invokes it.",
        "P4 — Setup: one-time hurdle on each new install.",
        "P5 — Dormant / latent / not currently reproducible: kept as historical "
        "record; we re-checked on the current dump and could not "
        "trigger the symptom.",
    ]
    legend_zh = [
        "P0 — 静默数据错误：数据错或缺失，但没有任何报错提示。",
        "P1 — 可见的运行时报错：弹出错误对话框，操作中断。",
        "P2 — 静默显示问题：表单字段本应有数据，却显示为空。",
        "P3 — 缺失界面：代码里实现了某功能，但界面上没有按钮去触发它。",
        "P4 — 安装设置：每台新机器需要一次性处理。",
        "P5 — 潛伏 / 不可達 / 當前無法復現：保留作为历史记录；我们在当前 dump "
        "上重新验证过，无法再触发症状。",
    ]
    _bullets(doc, [Z(s) for s in (legend_en if is_en else legend_zh)])
    doc.add_page_break()

    # ---- One section per issue ----
    by_tier: dict[str, list[dict]] = {}
    for it in ISSUES:
        by_tier.setdefault(it["tier"], []).append(it)
    tier_order = ["P0_silent_data", "P1_visible_crash",
                  "P2_silent_display", "P3_missing_ui", "P4_setup",
                  "P5_dormant_or_latent"]
    tier_titles_en = {
        "P0_silent_data": "P0 — Silent data corruption",
        "P1_visible_crash": "P1 — Visible runtime crash",
        "P2_silent_display": "P2 — Silent display",
        "P3_missing_ui": "P3 — Missing UI",
        "P4_setup": "P4 — Setup",
        "P5_dormant_or_latent":
            "P5 — Dormant / latent / not currently reproducible",
    }
    tier_titles_zh = {
        "P0_silent_data": "P0 — 静默数据错误",
        "P1_visible_crash": "P1 — 可见的运行时报错",
        "P2_silent_display": "P2 — 静默显示问题",
        "P3_missing_ui": "P3 — 缺失界面",
        "P4_setup": "P4 — 安装设置",
        "P5_dormant_or_latent": "P5 — 潛伏 / 不可達 / 當前無法復現",
    }
    demo_persons = _load_demo_persons()
    bug_status = _load_bug_test_status()

    for tier in tier_order:
        items = by_tier.get(tier, [])
        if not items:
            continue
        _h(doc, 1, Z(tier_titles_en[tier] if is_en
                      else tier_titles_zh[tier]))
        if tier == "P5_dormant_or_latent":
            preface = (
                "Items in this tier are kept as historical / latent "
                "record.  They fall into three categories: (a) DORMANT "
                "— current source data doesn't trigger the symptom; "
                "(b) NOT CURRENTLY REPRODUCIBLE — the symptom no "
                "longer surfaces even though the suspect code is "
                "still present (we have NOT confirmed an upstream "
                "source-level fix; the cause could be a JET / Office "
                "behaviour change, a fixture / driver change on our "
                "side, or the original diagnosis was a false "
                "positive); (c) LATENT — the source-code defect is "
                "real, but the user can't reach it because another "
                "issue (e.g. a missing UI button) blocks the path. "
                "None of these are user-facing today; **none have "
                "been verified as fixed upstream** — please consult "
                "before treating any of them as either urgent or "
                "closed."
                if is_en else
                "本層的條目作為歷史 / 潛伏記錄保留。可分為三類："
                "(a) DORMANT 潛伏 —— 已驗證當前源資料無法觸發該症狀；"
                "(b) 當前無法復現 —— 症狀不再出現，但可疑程式碼仍在"
                "（我們**沒有**確認上游有源碼層面的修復；原因可能是 "
                "JET / Office 的行為改變、可能是我們這邊 fixture/driver "
                "改變，也可能原本的診斷就是 false positive）；"
                "(c) LATENT 被屏蔽 —— 源碼缺陷確實存在，但因為"
                "另一個 issue（例如某個 UI 按鈕缺失）擋住了使用路徑，"
                "使用者目前碰不到。本層條目當下都不是使用者會遇到的"
                "問題，**也沒有任何一條被確認上游修復**；若要當成緊急"
                "或已關閉處理，請先諮詢。"
            )
            p = doc.add_paragraph(Z(preface))
            for run in p.runs:
                run.italic = True
        for it in items:
            title = it["title_en"] if is_en else it["title_zh"]
            _h(doc, 2, Z(f"Issue #{it['id']} — {title}"))
            _h(doc, 3, Z("Affected sub" if is_en else "涉及位置"))
            doc.add_paragraph(Z(it["form"]))
            _h(doc, 3, Z("Severity" if is_en else "严重等级"))
            doc.add_paragraph(Z(it["severity_en"] if is_en
                                  else it["severity_zh"]))
            # ---- Auto-derived test status banner (informational) ----
            # If a known-bug test exists for this issue and reports
            # 'failed', that's a hint the bug may have been fixed
            # upstream — but we never act on it automatically.  We
            # display a clearly-marked banner asking the maintainer to
            # confirm in person; the issue's full description / steps
            # / fix recommendation stay rendered as-is so nothing is
            # lost if the test gave a false-positive.
            status = bug_status.get(it["id"])
            if status:
                outcome = status["outcome"]
                if outcome == "failed":
                    banner = (
                        "⚠ Automated test status: the regression marker "
                        "for this issue currently FAILS (run on "
                        f"{status.get('when', 'unknown date')}), which "
                        "means the marker no longer reproduces.  This MAY "
                        "indicate the underlying defect was fixed "
                        "upstream, but it could equally mean that the "
                        "input fixture or Access driver changed out "
                        "from under the test, or that the original "
                        "classification was wrong.  Please investigate "
                        "in person before considering this issue closed; "
                        "this report has NOT been edited to drop the "
                        "issue.  Tests consulted:\n"
                        + "\n".join(
                            f"  • {t.rsplit('::', 1)[-1]}"
                            for t, _ in status["tests"]
                        )
                        if is_en else
                        "⚠ 自動測試狀態：本 issue 對應的回歸標記目前 "
                        f"FAIL（執行時間：{status.get('when', '未知')}），"
                        "代表標記目前無法復現。這**可能**表示上游已在源碼層"
                        "面修復，但同樣可能是輸入 fixture 或 Access "
                        "driver 在不知不覺中改變、或原本的分類就是錯的。"
                        "請務必親自調查清楚，再將此 issue 視為關閉；本報告"
                        "並未自動刪除任何 issue。對照的測試：\n"
                        + "\n".join(
                            f"  • {t.rsplit('::', 1)[-1]}"
                            for t, _ in status["tests"]
                        )
                    )
                    p = doc.add_paragraph()
                    run = p.add_run(Z(banner))
                    run.bold = True
                elif outcome == "mixed":
                    banner = (
                        "ℹ Automated test status: this issue's "
                        "regression markers report MIXED outcomes "
                        f"(run {status.get('when', 'unknown date')}). "
                        "Likely the issue's markers partially stopped "
                        "reproducing (could be a partial upstream fix, a "
                        "partial fixture/driver change, or a partial "
                        "misclassification).  Please "
                        "review the per-test breakdown:\n"
                        + "\n".join(
                            f"  • {t.rsplit('::', 1)[-1]}: {o}"
                            for t, o in status["tests"]
                        )
                        if is_en else
                        "ℹ 自動測試狀態：本 issue 對應的回歸標記呈現"
                        f"混合結果（執行時間：{status.get('when', '未知')}）。"
                        "可能是部分標記不再復現（可能是部分上游修復、"
                        "部分 fixture/driver 變化、或部分分類錯誤）。"
                        "請查看分項：\n"
                        + "\n".join(
                            f"  • {t.rsplit('::', 1)[-1]}: {o}"
                            for t, o in status["tests"]
                        )
                    )
                    p = doc.add_paragraph()
                    run = p.add_run(Z(banner))
                    run.italic = True
                # outcome == "passed" → bug confirmed present; no
                # banner (it's the default expectation; banner-spam is
                # noise).

            _h(doc, 3, Z("Description" if is_en else "问题描述"))
            for para in (it["summary_en"] if is_en
                         else it["summary_zh"]).split("\n\n"):
                doc.add_paragraph(Z(para))
            _h(doc, 3, Z("Steps to reproduce" if is_en else "复现步骤"))
            # Inject a concrete demo-person hint AHEAD of the steps so
            # the maintainer never has to guess which person id to
            # open.  For DORMANT bugs (no UI-reachable example on the
            # current data snapshot), the hint says so and offers a
            # SQL-only verification path instead.
            demo = demo_persons.get(f"bug{it['id']}")
            if demo:
                if demo.get("dormant"):
                    label = (
                        "⚠ Currently UI-dormant on this data "
                        "snapshot — see note below"
                        if is_en else
                        "⚠ 在當前資料快照下無法在 UI 上復現——請看下方說明"
                    )
                    para = doc.add_paragraph()
                    run = para.add_run(Z(label))
                    run.bold = True
                    para = doc.add_paragraph()
                    para.add_run(
                        demo["hint_en"] if is_en else Z(demo["hint_zh"])
                    )
                else:
                    label = (
                        f"Recommended demo person: c_personid="
                        f"{demo['personid']} ({demo['name_chn']}, "
                        f"{demo['name_py']})"
                        if is_en else
                        f"建議使用的範例人物：c_personid={demo['personid']}"
                        f"（{demo['name_chn']}，{demo['name_py']}）"
                    )
                    para = doc.add_paragraph()
                    run = para.add_run(Z(label))
                    run.bold = True
                    para = doc.add_paragraph()
                    para.add_run(
                        demo["hint_en"] if is_en else Z(demo["hint_zh"])
                    )
                    para.add_run(
                        Z(
                            " — picked by `reports/probe_demo_persons.py`; "
                            "a SQL probe selected this person because their "
                            "row counts genuinely satisfy the precondition "
                            "the bug needs."
                            if is_en else
                            " —— 由 `reports/probe_demo_persons.py` 透過 "
                            "SQL probe 挑選；之所以選這位，是因為其底層"
                            "記錄數確實滿足這個 bug 的觸發條件。"
                        )
                    ).italic = True
            _numbered(doc, [Z(s) for s in
                            (it["steps_en"] if is_en else it["steps_zh"])])
            # Optional concrete-reproduction block — see Issue #9
            # for the canonical use (specific personids that
            # trigger the bug on the current dump).
            cr_key = ("concrete_reproduction_en" if is_en
                      else "concrete_reproduction_zh")
            cr_text = it.get(cr_key)
            if cr_text:
                _h(doc, 3, Z("Concrete reproduction"
                              if is_en else "具體復現"))
                for chunk in str(cr_text).split("\n\n"):
                    if chunk.strip():
                        doc.add_paragraph(Z(chunk))
            shots = it.get("screenshots") or []
            if shots:
                _h(doc, 3, Z("Screenshots" if is_en else "截图"))
                for fname, caption in shots:
                    p = SHOT_DIR / fname
                    if not p.exists():
                        doc.add_paragraph(Z(
                            f"[screenshot {fname} not found]"
                            if is_en else f"[未找到截图 {fname}]"
                        ))
                        continue
                    doc.add_picture(str(p), width=Inches(6.0))
                    if caption:
                        cap = doc.add_paragraph(
                            caption if is_en else Z(caption)
                        )
                        cap.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                        for run in cap.runs:
                            run.italic = True
                            run.font.size = Pt(9)
            _h(doc, 3, Z("Suggested fix" if is_en else "建议修复方案"))
            for para in (it["fix_en"] if is_en
                         else it["fix_zh"]).split("\n\n"):
                doc.add_paragraph(Z(para))

    # ---- Appendix: index_year / index_addr drift (NOT a bug) ----
    doc.add_page_break()
    _add_index_drift_appendix(doc, is_en, Z)

    # ---- Closing ----
    doc.add_page_break()
    _h(doc, 1, Z("Closing note" if is_en else "结语"))
    closing = (
        "Thank you for taking the time to read this report. None of the "
        "items above is urgent; we hope having them all in one place "
        "makes it easy to address them at your own pace.\n\n"
        "If any of the descriptions or suggested fixes are unclear, we "
        "would be glad to discuss further. The corresponding regression "
        "tests in this repository will automatically flip from PASS to "
        "FAIL the moment any regression marker stops reproducing in "
        "the source dump — that is a signal to investigate, not an "
        "automatic confirmation that the bug is fixed (the marker "
        "could fail because of an upstream fix, a fixture / driver "
        "change on our side, or a misclassification we made earlier)."
        if is_en else
        "感谢您抽时间读完这份报告。以上各条都不紧急，我们把它们集中整理"
        "在一起，只是希望方便您在合适的时候逐一处理。\n\n"
        "如果对其中任何一条的描述或建议有疑问，欢迎随时一同讨论。本仓库"
        "里对应的回归测试，会在任何一个回归标记不再复现时自动从 PASS "
        "翻成 FAIL —— 这是「请调查一下」的讯号，而不是「问题已修复」的"
        "自动确认（因为标记不再复现也可能是 fixture / driver 变了，"
        "或者是我们当初的分类有误）。"
    )
    for para in closing.split("\n\n"):
        doc.add_paragraph(Z(para))

    doc.save(out_path)
    print(f"wrote {out_path}")


def _build_md(lang: str, out_path: Path) -> None:
    """Markdown sibling of `_build` — same content, different format,
    suitable for in-browser viewing on GitHub."""
    is_en = (lang == "en")

    def Z(s: str) -> str:
        return s if is_en else t(s)

    lines: list[str] = []
    title = (
        "CBDB User MDB — Issues Report"
        if is_en else
        "CBDB 用户版 .mdb — 问题汇报"
    )
    subtitle = (
        "A respectful summary of issues uncovered during regression testing."
        if is_en else
        "测试过程中发现的问题汇总，谨呈维护团队斧正。"
    )
    lines.append(f"# {Z(title)}")
    lines.append("")
    lines.append(f"_{Z(subtitle)}_")
    lines.append("")

    intro = (
        "Dear maintainer,\n\n"
        "Below is a summary of the issues we uncovered while building "
        "an automated regression-test suite for the CBDB User MDB. "
        "We hope this report is useful as you continue your wonderful "
        "stewardship of this dataset, and we sincerely thank you for "
        "the immense work that has gone into building it.\n\n"
        "The issues are ordered by severity (P0 highest). Each entry "
        "includes a concise description, step-by-step user reproduction, "
        "screenshots where the issue is visible in the Access UI, and a "
        "suggested fix. None of these are urgent; they are documented "
        "so they can be addressed at the maintainer's convenience."
        if is_en else
        "尊敬的维护者：\n\n"
        "下面是我们在为 CBDB 用户版 .mdb 编写自动化回归测试套件过程中，"
        "陆续整理出来的一些问题清单。我们希望这份报告能在您继续主持这份"
        "宝贵数据集时有所帮助；同时，对您多年来在这套数据上的辛勤付出，"
        "我们由衷地表示感谢和敬意。\n\n"
        "问题按严重程度排序（P0 最高）。每一条都包括：简明描述、用户端"
        "一步一步的复现步骤、（在界面上能看到时）相关截图，以及一份建议"
        "的修复方案。这些问题并不紧急，整理在此只是为了方便您在合适的时"
        "候逐一处理。"
    )
    for para in intro.split("\n\n"):
        lines.append(Z(para))
        lines.append("")

    # ---- TOC: GitHub auto-generates anchors from heading text. ----
    lines.append(f"## {Z('Table of Contents' if is_en else '目录')}")
    lines.append("")

    tier_titles_en = {
        "P0_silent_data": "P0 — Silent data corruption",
        "P1_visible_crash": "P1 — Visible runtime crash",
        "P2_silent_display": "P2 — Silent display",
        "P3_missing_ui": "P3 — Missing UI",
        "P4_setup": "P4 — Setup",
        "P5_dormant_or_latent":
            "P5 — Dormant / latent / not currently reproducible",
    }
    tier_titles_zh = {
        "P0_silent_data": "P0 — 静默数据错误",
        "P1_visible_crash": "P1 — 可见的运行时报错",
        "P2_silent_display": "P2 — 静默显示问题",
        "P3_missing_ui": "P3 — 缺失界面",
        "P4_setup": "P4 — 安装设置",
        "P5_dormant_or_latent": "P5 — 潛伏 / 不可達 / 當前無法復現",
    }
    tier_order = ["P0_silent_data", "P1_visible_crash",
                  "P2_silent_display", "P3_missing_ui", "P4_setup",
                  "P5_dormant_or_latent"]

    by_tier: dict[str, list[dict]] = {}
    for it in ISSUES:
        by_tier.setdefault(it["tier"], []).append(it)

    def _slug(s: str) -> str:
        # GitHub-flavoured anchor algorithm:
        #   1. lower-case
        #   2. drop every char that isn't [a-z0-9_-] OR a CJK
        #      ideograph OR a space (DROP IN PLACE — do not merge
        #      adjacent spaces; do not collapse runs of hyphens).
        #   3. replace each remaining space with exactly one hyphen.
        # Two subtle rules the earlier draft of this function got
        # wrong:
        #   - `_` is a word char and IS preserved by GitHub.  A blanket
        #     `[`*_~]` strip removed it and broke `View_StatusData`-
        #     style anchors.
        #   - `\s+ → -` collapsed runs of spaces into a single hyphen.
        #     GitHub doesn't.  'P0 — Silent ...' becomes
        #     'p0--silent-...' (the em-dash gets removed but the
        #     spaces around it stay, producing two hyphens).
        import re as _re
        out = s.lower().strip()
        # Drop only the markdown formatting chars that are NOT also
        # valid in identifiers (so no `_`).
        out = _re.sub(r"[`*~]", "", out)
        # Drop anything that isn't word-char / hyphen / space / CJK.
        out = _re.sub(r"[^\w一-鿿\- ]+", "", out)
        # 1-for-1 space → hyphen substitution (preserves doubles).
        out = out.replace(" ", "-")
        return out

    for tier in tier_order:
        items = by_tier.get(tier, [])
        if not items:
            continue
        tier_title = (tier_titles_en[tier] if is_en
                      else tier_titles_zh[tier])
        lines.append(f"- [{Z(tier_title)}](#{_slug(Z(tier_title))})")
        for it in items:
            t_title = it["title_en"] if is_en else it["title_zh"]
            entry = f"Issue #{it['id']} — {t_title}"
            lines.append(
                f"  - [{Z(entry)}](#{_slug(Z(entry))})"
            )
    lines.append(
        f"- [{Z('Severity legend' if is_en else '严重等级说明')}]"
        f"(#{_slug(Z('Severity legend' if is_en else '严重等级说明'))})"
    )
    appendix_title = (
        "Appendix — c_index_year / c_index_addr_id drift "
        "vs the cbdb-online-main-server snapshot "
        "(differences need per-row classification before being filed as bugs)"
        if is_en else
        "附錄 —— c_index_year / c_index_addr_id 與 "
        "cbdb-online-main-server 快照之間的偏差"
        "（差異需要逐筆分類後才能判定是否為缺陷）"
    )
    lines.append(
        f"- [{Z(appendix_title)}](#{_slug(Z(appendix_title))})"
    )
    lines.append(
        f"- [{Z('Closing note' if is_en else '结语')}]"
        f"(#{_slug(Z('Closing note' if is_en else '结语'))})"
    )
    lines.append("")

    # ---- Severity legend ----
    lines.append(f"## {Z('Severity legend' if is_en else '严重等级说明')}")
    lines.append("")
    legend_en = [
        "P0 — Silent data corruption: data is wrong or missing without an error popup.",
        "P1 — Visible runtime crash: a popup appears, the operation aborts.",
        "P2 — Silent display: form fields render blank when they should show data.",
        "P3 — Missing UI: a feature exists in code but no button invokes it.",
        "P4 — Setup: one-time hurdle on each new install.",
        "P5 — Dormant / latent / not currently reproducible: kept as historical "
        "record; we re-checked on the current dump and could not "
        "trigger the symptom.",
    ]
    legend_zh = [
        "P0 — 静默数据错误：数据错或缺失，但没有任何报错提示。",
        "P1 — 可见的运行时报错：弹出错误对话框，操作中断。",
        "P2 — 静默显示问题：表单字段本应有数据，却显示为空。",
        "P3 — 缺失界面：代码里实现了某功能，但界面上没有按钮去触发它。",
        "P4 — 安装设置：每台新机器需要一次性处理。",
        "P5 — 潛伏 / 不可達 / 當前無法復現：保留作为历史记录；我们在当前 dump "
        "上重新验证过，无法再触发症状。",
    ]
    for s in (legend_en if is_en else legend_zh):
        lines.append(f"- {Z(s)}")
    lines.append("")

    # ---- Per-issue ----
    demo_persons = _load_demo_persons()
    bug_status = _load_bug_test_status()

    for tier in tier_order:
        items = by_tier.get(tier, [])
        if not items:
            continue
        tier_title = (tier_titles_en[tier] if is_en
                      else tier_titles_zh[tier])
        lines.append(f"## {Z(tier_title)}")
        lines.append("")
        if tier == "P5_dormant_or_latent":
            lines.append(Z(
                "_Items in this tier are kept as historical / latent "
                "record.  They fall into three categories: (a) DORMANT — "
                "verified that current source data doesn't trigger the "
                "symptom; (b) NOT CURRENTLY REPRODUCIBLE — the symptom "
                "no longer surfaces even though the suspect code is "
                "still present (we have NOT confirmed an upstream "
                "source-level fix; could be a JET / Office behaviour "
                "change, a fixture / driver change on our side, or the "
                "original diagnosis was a false positive); (c) LATENT — "
                "the source-code defect is real, but the user can't "
                "reach it because another issue (e.g. a missing UI "
                "button) blocks the path.  None of these are "
                "user-facing today; **none have been verified as fixed "
                "upstream** — please consult before treating any of "
                "them as either urgent or closed._"
                if is_en else
                "_本層的條目作為歷史 / 潛伏記錄保留。可分為三類："
                "(a) DORMANT 潛伏 — 已驗證當前源資料無法觸發該症狀；"
                "(b) 當前無法復現 — 症狀不再出現，但可疑程式碼仍在"
                "（我們**沒有**確認上游有源碼層面的修復；原因可能是 "
                "JET / Office 的行為改變、可能是我們這邊 fixture/driver "
                "改變，也可能原本的診斷就是 false positive）；"
                "(c) LATENT 被屏蔽 — 源碼缺陷確實存在，但因為另一個 "
                "issue（例如某個 UI 按鈕缺失）擋住了使用路徑，使用者"
                "目前碰不到。本層條目當下都不是使用者會遇到的問題，"
                "**也沒有任何一條被確認上游修復**；若要當成緊急或"
                "已關閉處理，請先諮詢。_"
            ))
            lines.append("")
        for it in items:
            t_title = it["title_en"] if is_en else it["title_zh"]
            lines.append(f"### Issue #{it['id']} — {Z(t_title)}")
            lines.append("")
            lines.append(f"**{Z('Affected sub' if is_en else '涉及位置')}:** "
                          f"`{it['form']}`")
            lines.append("")
            lines.append(f"**{Z('Severity' if is_en else '严重等级')}:** "
                          f"{Z(it['severity_en'] if is_en else it['severity_zh'])}")
            lines.append("")
            # Auto-derived test status banner.
            status = bug_status.get(it["id"])
            if status:
                outcome = status["outcome"]
                if outcome == "failed":
                    banner = (
                        "⚠ **Automated test status: this issue's "
                        "regression marker currently FAILS** "
                        f"(run on {status.get('when', 'unknown date')}). "
                        "That usually means the underlying defect has "
                        "longer reproduces, which MAY mean upstream fixed "
                        "it but could equally be a fixture / driver "
                        "regression or a misclassification.  Please "
                        "verify in person before considering this issue "
                        "closed; this report has NOT been edited to drop "
                        "the issue.  Tests consulted: "
                        + ", ".join(
                            f"`{t.rsplit('::', 1)[-1]}`"
                            for t, _ in status["tests"]
                        )
                        if is_en else
                        "⚠ **自動測試狀態：本 issue 對應的回歸標記目前 "
                        f"FAIL**（執行時間：{status.get('when', '未知')}），"
                        "通常意味著底層缺陷已在 source dump 中被 **修復**。"
                        "請務必親自確認，再將此 issue 視為關閉；本報告"
                        "並未自動刪除任何 issue。對照的測試："
                        + "、".join(
                            f"`{t.rsplit('::', 1)[-1]}`"
                            for t, _ in status["tests"]
                        )
                    )
                    lines.append(f"> {Z(banner)}")
                    lines.append("")
                elif outcome == "mixed":
                    parts = ", ".join(
                        f"`{t.rsplit('::', 1)[-1]}`: {o}"
                        for t, o in status["tests"]
                    )
                    banner = (
                        "ℹ Automated test status: MIXED outcomes "
                        f"(run {status.get('when', 'unknown date')}). "
                        f"Per-test: {parts}"
                        if is_en else
                        f"ℹ 自動測試狀態：混合結果（執行時間："
                        f"{status.get('when', '未知')}）。分項：{parts}"
                    )
                    lines.append(f"> _{Z(banner)}_")
                    lines.append("")
            lines.append(f"#### {Z('Description' if is_en else '问题描述')}")
            lines.append("")
            for para in (it["summary_en"] if is_en
                         else it["summary_zh"]).split("\n\n"):
                lines.append(Z(para))
                lines.append("")
            lines.append(f"#### {Z('Steps to reproduce' if is_en else '复现步骤')}")
            lines.append("")
            demo = demo_persons.get(f"bug{it['id']}")
            if demo:
                if demo.get("dormant"):
                    label = (
                        "⚠ **Currently UI-dormant on this data snapshot — see note below**"
                        if is_en else
                        "⚠ **在當前資料快照下無法在 UI 上復現——請看下方說明**"
                    )
                    lines.append(Z(label))
                    lines.append("")
                    lines.append(
                        demo["hint_en"] if is_en else Z(demo["hint_zh"])
                    )
                    lines.append("")
                else:
                    label = (
                        f"**Recommended demo person:** `c_personid="
                        f"{demo['personid']}` ({demo['name_chn']}, "
                        f"{demo['name_py']})"
                        if is_en else
                        f"**建議使用的範例人物：** `c_personid="
                        f"{demo['personid']}`（{demo['name_chn']}，"
                        f"{demo['name_py']}）"
                    )
                    lines.append(Z(label))
                    lines.append("")
                    hint_text = demo["hint_en"] if is_en else Z(demo["hint_zh"])
                    extra = (
                        " _Picked by `reports/probe_demo_persons.py`; a "
                        "SQL probe selected this person because their row "
                        "counts genuinely satisfy the precondition the "
                        "bug needs._"
                        if is_en else
                        " _由 `reports/probe_demo_persons.py` 透過 SQL "
                        "probe 挑選；之所以選這位，是因為其底層記錄數確實"
                        "滿足這個 bug 的觸發條件。_"
                    )
                    lines.append(hint_text + extra)
                    lines.append("")
            for i, step in enumerate(
                it["steps_en"] if is_en else it["steps_zh"], 1
            ):
                lines.append(f"{i}. {Z(step)}")
            lines.append("")
            # Optional concrete-reproduction block — see Issue #9.
            cr_key = ("concrete_reproduction_en" if is_en
                      else "concrete_reproduction_zh")
            cr_text = it.get(cr_key)
            if cr_text:
                lines.append(f"#### {Z('Concrete reproduction' if is_en else '具體復現')}")
                lines.append("")
                lines.append(Z(str(cr_text)))
                lines.append("")
            shots = it.get("screenshots") or []
            if shots:
                lines.append(f"#### {Z('Screenshots' if is_en else '截图')}")
                lines.append("")
                for fname, caption in shots:
                    p = SHOT_DIR / fname
                    if not p.exists():
                        lines.append(
                            f"_{Z(f'[screenshot {fname} not found]' if is_en else f'[未找到截图 {fname}]')}_"
                        )
                    else:
                        rel = f"screenshots/{fname}"
                        lines.append(f"![{fname}]({rel})")
                        if caption:
                            cap = caption if is_en else Z(caption)
                            lines.append("")
                            lines.append(f"_{cap}_")
                    lines.append("")
            lines.append(
                f"#### {Z('Suggested fix' if is_en else '建议修复方案')}"
            )
            lines.append("")
            for para in (it["fix_en"] if is_en
                         else it["fix_zh"]).split("\n\n"):
                lines.append(Z(para))
                lines.append("")

    # ---- Index drift appendix ----
    import json as _json
    if DRIFT_JSON.exists():
        lines.append(f"## {Z(appendix_title)}")
        lines.append("")
        intro_drift = (
            "When we compare BIOG_MAIN's `c_index_year` and "
            "`c_index_addr_id` between this User MDB and the weekly "
            "cbdb-online-main-server SQLite snapshot, a small "
            "fraction of persons disagree.\n\n"
            "**The two sides are independent implementations.**  The "
            "SQLite snapshot's `c_index_year` is produced by "
            "cbdb-online-main-server's PHP "
            "`IndexYearRebuildService.php` and its `c_index_addr_id` "
            "by `IndexAddressRebuildService.php` (both at "
            "<https://github.com/cbdb-project/cbdb-online-main-server>"
            "); the User MDB-side: `c_index_addr_id` rebuilt by VBA "
            "in `Form_frmIndexAddr` (front-end mdb); `c_index_year` "
            "rebuilt by **37 saved QueryDefs named `BM IY Rule …`** "
            "in the linked-tables backend "
            "`data/CBDB_<YYYYMMDD>_DATA.mdb`, driven by "
            "`frmBaseMaintenance`.  Both algorithms now extracted "
            "to `analysis/dump_data/querydefs_index/*.sql`; form / "
            "module driver VBA still needs an interactive Access "
            "SaveAsText pass.  PHP is intended to mirror the "
            "VBA but they are separate code paths.  Per-row "
            "differences can come from at least four sources, and a "
            "diff alone doesn't tell us which: (1) source-data "
            "snapshot drift; (2) algorithm / porting divergence "
            "between PHP and VBA; (3) priority / tie-break "
            "differences; (4) null / default handling differences."
            "\n\n"
            "**We have not classified the steady ~575 / 657 246 "
            "diffs we currently observe.**  The examples below are a "
            "small sample (currently 13 rows across 3 buckets, "
            "from `reports/index_drift_examples.json`) — illustrative "
            "of the shapes of disagreement, not statistically "
            "representative.  They are a starting point for per-row "
            "triage, not a verdict."
            if is_en else
            "我們把本 .mdb 的 BIOG_MAIN 與 cbdb-online-main-server 每週"
            "釋出的 SQLite 快照在 `c_index_year`、`c_index_addr_id` "
            "兩個欄位上做比對，可以看到一小部分人物對不齊。\n\n"
            "**兩邊是兩套獨立的實作。**SQLite 快照中的 "
            "`c_index_year` 是 cbdb-online-main-server 的 PHP "
            "`IndexYearRebuildService.php` 算出來的，"
            "`c_index_addr_id` 則是 `IndexAddressRebuildService.php` "
            "算出來的（程式碼都在 <https://github.com/cbdb-project/"
            "cbdb-online-main-server>）；User MDB 上對應的這兩個"
            "User MDB 那一邊：`c_index_addr_id` 由前端 mdb 裡的 "
            "`Form_frmIndexAddr` VBA 重建；`c_index_year` 由連結表"
            "後端 `data/CBDB_<YYYYMMDD>_DATA.mdb` 裡 **37 條 "
            "`BM IY Rule …` 的 QueryDef** 重建，由 "
            "`frmBaseMaintenance` 驅動。兩邊算法已抽取到 "
            "`analysis/dump_data/querydefs_index/*.sql`；form / "
            "module 驅動 VBA 仍需 Access SaveAsText 互動式提取。"
            "PHP **意圖**鏡像 VBA，"
            "但兩者是兩條獨立的程式路徑。每一行差異**可能**來自下列"
            "至少四個原因，光看差異本身分不出來：(1) 源資料快照漂移；"
            "(2) PHP 與 VBA 之間的演算法 / 移植差異；(3) 優先序 / "
            "平手規則不同；(4) null / 預設值處理不同。\n\n"
            "**我們並沒有對目前看到的 ~575 / 657 246 筆差異做完整"
            "分類。**下方列舉的樣本（目前共 13 筆、3 種分桶，來自 "
            "`reports/index_drift_examples.json`）只是**示範**這些"
            "差異**長什麼樣**，並非統計上有代表性，是後續逐筆分類"
            "的起點，不是結論。"
        )
        lines.append(Z(intro_drift))
        lines.append("")

        # ---- Classification summary (markdown copy) ----
        if CLASSIFICATION_JSON.exists():
            cls = _json.loads(
                CLASSIFICATION_JSON.read_text(encoding="utf-8"))
            cs = cls["summary"]
            b = cs["buckets"]
            net = cs["common"] - b["exact_match"]
            attributable = (b["source_drift_index_agrees"]
                            + b["source_drift_index_diffs_too"])
            unclassified = (b["index_year_only_diff"]
                            + b["index_addr_only_diff"]
                            + b["index_both_diff"])
            if is_en:
                lines.append(f"### {Z('Classification summary')}")
                lines.append("")
                lines.append(Z(
                    f"Compared **{cs['common']:,}** personids common to "
                    f"both databases (User MDB total {cs['user_mdb_total']:,}; "
                    f"SQLite total {cs['sqlite_total']:,}; "
                    f"User-only {cs['in_user_only']:,}; "
                    f"SQLite-only {cs['in_sqlite_only']:,})."
                ))
                lines.append("")
                lines.append("| Bucket | Count | % of common | Meaning |")
                lines.append("|---|---:|---:|---|")
                rows = [
                    ("exact_match", "exact match on all four compared fields"),
                    ("source_drift_index_agrees",
                     "source drift but indices agreed"),
                    ("source_drift_index_diffs_too",
                     "source drift AND ≥1 index differs"),
                    ("index_year_only_diff",
                     "source matched, only c_index_year differs — needs follow-up"),
                    ("index_addr_only_diff",
                     "source matched, only c_index_addr_id differs — needs follow-up"),
                    ("index_both_diff",
                     "source matched, both indices differ — strongest signal"),
                ]
                for k, label in rows:
                    pct = 100.0 * b[k] / max(cs['common'], 1)
                    lines.append(f"| `{k}` | {b[k]:,} | {pct:.3f}% | {Z(label)} |")
                lines.append("")
                lines.append(Z(
                    f"Net diffs: **{net:,}** of {cs['common']:,} "
                    f"({100.0*net/max(cs['common'],1):.3f} %).  Of those, "
                    f"**{attributable}** are clearly attributable to source "
                    f"drift in birthyear / deathyear; **{unclassified}** need "
                    f"per-row follow-up.  These could be PHP↔VBA divergence, "
                    f"or drift in evidence tables (BIOG_ADDR_DATA / "
                    f"ENTRY_DATA / NIAN_HAO etc.) that this classifier does "
                    f"not compare.  Full output: "
                    f"`reports/index_drift_classification.json`; algorithm "
                    f"pointers: `analysis/index_drift_algorithm_notes.md`."
                ))
                lines.append("")
            else:
                lines.append(f"### {Z('分類匯總')}")
                lines.append("")
                lines.append(Z(
                    f"比對了兩邊都有的 **{cs['common']:,}** 個 personid"
                    f"（User MDB 共 {cs['user_mdb_total']:,} 筆；"
                    f"SQLite 共 {cs['sqlite_total']:,} 筆；"
                    f"僅 User MDB 有 {cs['in_user_only']:,} 筆；"
                    f"僅 SQLite 有 {cs['in_sqlite_only']:,} 筆）。"
                ))
                lines.append("")
                lines.append("| 分桶 | 筆數 | 佔比 | 含義 |")
                lines.append("|---|---:|---:|---|")
                rows = [
                    ("exact_match", "四個欄位全部一致"),
                    ("source_drift_index_agrees",
                     "源資料有漂移但兩邊 index 都一致"),
                    ("source_drift_index_diffs_too",
                     "源資料有漂移、且至少一個 index 不同"),
                    ("index_year_only_diff",
                     "生年/卒年一致，但只有 c_index_year 不同 —— 待追查"),
                    ("index_addr_only_diff",
                     "生年/卒年一致，但只有 c_index_addr_id 不同 —— 待追查"),
                    ("index_both_diff",
                     "生年/卒年一致，但兩個 index 都不同 —— 複合差異最強信號"),
                ]
                for k, label in rows:
                    pct = 100.0 * b[k] / max(cs['common'], 1)
                    lines.append(f"| `{k}` | {b[k]:,} | {pct:.3f}% | {Z(label)} |")
                lines.append("")
                lines.append(Z(
                    f"淨差異：**{net:,}** / {cs['common']:,}"
                    f"（{100.0*net/max(cs['common'],1):.3f} %）。其中 "
                    f"**{attributable}** 筆能明確歸因於 birthyear / "
                    f"deathyear 的源資料漂移；剩下 **{unclassified}** 筆"
                    f"需要逐筆追查（可能是 PHP↔VBA 演算法差異，"
                    f"也可能是本分類器沒有比較的 evidence 表"
                    f"（BIOG_ADDR_DATA / ENTRY_DATA / NIAN_HAO 等）裡的"
                    f"漂移）。完整輸出見 "
                    f"`reports/index_drift_classification.json`，"
                    f"算法來源指標見 "
                    f"`analysis/index_drift_algorithm_notes.md`。"
                ))
                lines.append("")

        # ---- Year-drift rule classification (PR K1) — markdown ----
        if RULE_CLASSIFICATION_JSON.exists():
            rcls = _json.loads(
                RULE_CLASSIFICATION_JSON.read_text(encoding="utf-8"))
            rs = rcls["summary"]
            rb = rs["buckets"]
            if is_en:
                lines.append(f"### {Z('Year-only diffs — per-row rule classification')}")
                lines.append("")
                lines.append(Z(
                    f"Of the **{rs['total_year_diffs']}** year-only "
                    f"diffs, each row was bucketed against PR N's "
                    f"rule-level runtime-vs-PHP comparison "
                    f"(`analysis/index_year_rule_comparison.md`).  "
                    f"Conservative buckets (rows count once each):"
                ))
                lines.append("")
                lines.append("| Bucket | Count |")
                lines.append("|---|---:|")
                bucket_meta_zh = None
                for k in [
                    "php_returned_sentinel",
                    "php_did_not_compute",
                    "access_did_not_compute",
                    "iteration_order_diff",
                    "consistent_within_rule",
                    "candidate_algorithm_divergence",
                    "unclassified",
                ]:
                    if rb[k] == 0:
                        continue
                    lines.append(f"| `{k}` | {rb[k]} |")
                lines.append("")
                lines.append(Z(
                    f"None of these are confirmed bugs.  Full per-row "
                    f"output is in "
                    f"`reports/index_year_drift_rule_classification.json`."
                ))
                lines.append("")
                if RULE_GROUPS_JSON.exists():
                    gcls = _json.loads(
                        RULE_GROUPS_JSON.read_text(encoding="utf-8"))
                    gs = gcls["summary"]
                    lines.append(Z(
                        f"Deeper triage (PR K2, "
                        f"`analysis/triage_index_year_drift_groups.py` "
                        f"→ `reports/index_year_drift_rule_groups.json`) "
                        f"named the leftover buckets:"
                    ))
                    lines.append("")
                    lines.append(Z(
                        f"- `consistent_within_rule` × "
                        f"{gs['consistent_within_rule']['rows']} → "
                        f"{gs['consistent_within_rule']['groups']} "
                        f"signature groups.  PR AI + AJ probes "
                        f"reversed the prior tie-break hypothesis: "
                        f"all 14 rows are `source_data_drift_biog_"
                        f"main_or_kin_data_between_sides` (8 "
                        f"BIOG_MAIN birthyear drift + 6 KIN_DATA "
                        f"evidence-pid drift).  Upstream PHP-side / "
                        f"SQLite-snapshot data issue, not a CBDB "
                        f"algorithm divergence."
                    ))
                    lines.append(Z(
                        f"- `unclassified` × {gs['unclassified']['total']} → "
                        f"{gs['unclassified']['named_after_triage']} named, "
                        f"{gs['unclassified']['blocked_by_runtime_priority_triage_pending']} "
                        f"flagged `blocked_by_runtime_priority_triage_pending` "
                        f"(PR M dumped frmBaseMaintenance, so the source is in repo; resolving each row still needs a per-row walk of the runtime priority/iteration order)."
                    ))
                    lines.append(Z(
                        f"- `php_did_not_compute` × "
                        f"{gs['php_did_not_compute']['rows']} → "
                        f"{gs['php_did_not_compute']['groups']} groups by "
                        f"Access tcode; biggest is `access_tcode='05'` × 7 "
                        f"(`candidate_php_entry_code_mapping_gap` for jinshi)."
                    ))
                    lines.append("")
            else:
                lines.append(f"### {Z('年份差異 —— 逐筆 rule 分類')}")
                lines.append("")
                lines.append(Z(
                    f"在 **{rs['total_year_diffs']}** 筆「只有 "
                    f"c_index_year 不一致」的行中，逐筆比對 PR N "
                    f"(`analysis/index_year_rule_comparison.md`) 的"
                    f"runtime-vs-PHP 規則級差異。保守分類如下："
                ))
                lines.append("")
                lines.append("| 分桶 | 筆數 |")
                lines.append("|---|---:|")
                label_zh = {
                    "php_returned_sentinel": "PHP 寫了 sentinel／溢位值",
                    "php_did_not_compute": "PHP 沒算出值（覆蓋率缺口）",
                    "access_did_not_compute": "Access 沒算出值（覆蓋率缺口）",
                    "iteration_order_diff": "Phase-C 迭代次數不同",
                    "consistent_within_rule": "多列共享同一 (php_tcode, access_tcode, diff)",
                    "candidate_algorithm_divergence": "形狀符合 K1 的歷史 hypothesis probe 但無法以單筆證據重建",
                    "unclassified": "尚未對上任何模式",
                }
                for k in [
                    "php_returned_sentinel",
                    "php_did_not_compute",
                    "access_did_not_compute",
                    "iteration_order_diff",
                    "consistent_within_rule",
                    "candidate_algorithm_divergence",
                    "unclassified",
                ]:
                    if rb[k] == 0:
                        continue
                    lines.append(f"| `{k}` ({label_zh[k]}) | {rb[k]} |")
                lines.append("")
                lines.append(Z(
                    f"以上沒有任何一筆被視為已確認的 bug。逐筆輸出見 "
                    f"`reports/index_year_drift_rule_classification.json`。"
                ))
                lines.append("")
                if RULE_GROUPS_JSON.exists():
                    gcls = _json.loads(
                        RULE_GROUPS_JSON.read_text(encoding="utf-8"))
                    gs = gcls["summary"]
                    lines.append(Z(
                        f"PR K2 進一步的 triage "
                        f"(`analysis/triage_index_year_drift_groups.py` "
                        f"→ `reports/index_year_drift_rule_groups.json`) "
                        f"把剩下的桶命名清楚："
                    ))
                    lines.append("")
                    lines.append(Z(
                        f"- `consistent_within_rule` × "
                        f"{gs['consistent_within_rule']['rows']} → "
                        f"{gs['consistent_within_rule']['groups']} 個 "
                        f"signature 分組。PR AI + AJ 的逐筆探測推翻了"
                        f"原本的 tie-break 假說：14 筆全是 "
                        f"`source_data_drift_biog_main_or_kin_data_"
                        f"between_sides`（8 筆 BIOG_MAIN birthyear 漂移 "
                        f"+ 6 筆 KIN_DATA evidence-pid 漂移）。屬於 "
                        f"PHP-side / SQLite snapshot 的上游資料漂移，"
                        f"並非 CBDB 演算法差異。"
                    ))
                    lines.append(Z(
                        f"- `unclassified` × {gs['unclassified']['total']} → "
                        f"{gs['unclassified']['named_after_triage']} 筆已命名，"
                        f"{gs['unclassified']['blocked_by_runtime_priority_triage_pending']} 筆標為 "
                        f"`blocked_by_runtime_priority_triage_pending`"
                        f"（PR M 已 dump frmBaseMaintenance，源碼已在 repo；要逐筆判斷哪邊正確仍需走一遍 runtime 的 priority／iteration 順序）。"
                    ))
                    lines.append(Z(
                        f"- `php_did_not_compute` × "
                        f"{gs['php_did_not_compute']['rows']} → 按 Access "
                        f"tcode 分 {gs['php_did_not_compute']['groups']} 組；"
                        f"最大的是 `access_tcode='05'` × 7（jinshi 進士類的 "
                        f"`candidate_php_entry_code_mapping_gap`）。"
                    ))
                    lines.append("")

        # ---- Address-drift classification (PR L) ----
        if ADDR_CLASSIFICATION_JSON.exists():
            acls = _json.loads(
                ADDR_CLASSIFICATION_JSON.read_text(encoding="utf-8"))
            aS = acls["summary"]
            ab = aS["buckets"]
            addr_bucket_order = [
                "mdb_stale_index_addr",
                "mdb_value_php_null",
                "same_candidates_diff_winner",
                "both_stale_recompute_mismatch",
                "both_sides_match_recomputed",
                "sqlite_stale_index_addr",
                "mdb_null_php_value",
                "unclassified",
            ]
            if is_en:
                lines.append(f"### {Z('c_index_addr_id diffs — per-row classification')}")
                lines.append("")
                lines.append(Z(
                    f"Of the **{aS['total_addr_diffs']}** "
                    f"c_index_addr diffs (478 `index_addr_only_diff` "
                    f"+ 10 `index_both_diff` from PR G), each row "
                    f"was classified by re-simulating the rank-"
                    f"priority + MAX(c_sequence) algorithm against "
                    f"each side's BIOG_ADDR_DATA + the shared "
                    f"BIOG_ADDR_CODES rank table."
                ))
                lines.append("")
                lines.append("| Bucket | Count |")
                lines.append("|---|---:|")
                for k in addr_bucket_order:
                    n = ab.get(k, 0)
                    if n == 0:
                        continue
                    lines.append(f"| `{k}` | {n} |")
                lines.append("")
                lines.append(Z(
                    f"None of these are confirmed bugs.  The 412 "
                    f"`mdb_stale_index_addr` rows are a maintenance-"
                    f"cadence diff (the User MDB needs its "
                    f"frmBaseMaintenance rebuild re-run before the "
                    f"next release).  The 10 "
                    f"`same_candidates_diff_winner` rows are the "
                    f"only candidate algorithm-divergence rows.  "
                    f"Full per-row output: "
                    f"`reports/index_addr_drift_classification.json`."
                ))
                lines.append("")
                lines.append(Z(
                    f"PR M (`analysis/dump_data_mdb_vba.py`) extracted "
                    f"`frmBaseMaintenance.CmdIndexAddress_Click` from "
                    f"the DATA mdb.  It does NOT explicitly "
                    f"`MAX(c_sequence)`-aggregate the way PHP does — a "
                    f"candidate algorithmic divergence on top of the "
                    f"maintenance-cadence issue.  Suggested "
                    f"release-checklist mitigation: run `CmdIndexYear` "
                    f"then `CmdIndexAddress` on the DATA mdb before "
                    f"shipping a new User MDB."
                ))
                lines.append("")
                lines.append(Z(
                    f"PR S "
                    f"(`analysis/deep_dive_addr_same_candidates.py`) "
                    f"confirmed the 10 `same_candidates_diff_winner` "
                    f"rows are all driven by MAX(c_sequence) ties "
                    f"(multiple BIOG_ADDR_DATA rows of the same "
                    f"(person, addr_type) sharing the same max "
                    f"c_sequence).  PHP, Access, and our recompute "
                    f"each pick non-deterministically.  Both sides "
                    f"follow the same documented rule; neither is "
                    f"wrong.  Candidate mitigation: add an explicit "
                    f"secondary tie-break (e.g. MIN(c_addr_id)) to "
                    f"both implementations.  Per-row evidence in "
                    f"`reports/index_addr_same_candidates_deep_dive.json`."
                ))
                lines.append("")

                # ---- Cause analysis appendix (PR Y) ----
                if CAUSE_SUMMARY_JSON.exists():
                    cs = _json.loads(
                        CAUSE_SUMMARY_JSON.read_text(encoding="utf-8"))
                    lines.append("### What currently explains the drift")
                    lines.append("")
                    lines.append(Z(
                        "Per-bucket cause / supporting evidence / "
                        "confidence / next action lives in "
                        "`analysis/index_drift_cause_analysis.md`.  "
                        "This section just summarises headline counts "
                        "and confidence; no bucket is labelled a "
                        "confirmed CBDB bug."
                    ))
                    lines.append("")
                    lines.append("**c_index_year cause buckets**")
                    lines.append("")
                    lines.append("| Bucket | Count | Confidence |")
                    lines.append("|---|---:|---|")
                    for b in cs["c_index_year"]["buckets"]:
                        if b["count"] == 0:
                            continue
                        lines.append(
                            f"| `{b['bucket']}` | {b['count']} "
                            f"| {b['confidence']} |"
                        )
                    lines.append("")
                    lines.append("**c_index_addr_id cause buckets**")
                    lines.append("")
                    lines.append("| Bucket | Count | Confidence |")
                    lines.append("|---|---:|---|")
                    for b in cs["c_index_addr_id"]["buckets"]:
                        if b["count"] == 0:
                            continue
                        lines.append(
                            f"| `{b['bucket']}` | {b['count']} "
                            f"| {b['confidence']} |"
                        )
                    lines.append("")
                    lines.append(Z(
                        "Top suggested next investigations (full list "
                        "in the cause-analysis md):"
                    ))
                    lines.append("")
                    for inv in cs[
                        "suggested_next_investigations_in_priority_order"
                    ][:3]:
                        lines.append(Z(
                            f"{inv['id']}. {inv['task']} — would close "
                            f"{inv['would_close_rows']} rows; "
                            f"engineering cost: {inv['engineering_cost']}."
                        ))
                    lines.append("")
            else:
                lines.append(f"### {Z('c_index_addr_id 差異 —— 逐筆分類')}")
                lines.append("")
                lines.append(Z(
                    f"在 **{aS['total_addr_diffs']}** 筆 c_index_addr "
                    f"差異中（PR G 的 478 `index_addr_only_diff` + 10 "
                    f"`index_both_diff`），逐筆把兩邊的 BIOG_ADDR_DATA "
                    f"代入「rank-priority + MAX(c_sequence)」演算法重算，"
                    f"與實際儲存值對照分類："
                ))
                lines.append("")
                lines.append("| 分桶 | 筆數 |")
                lines.append("|---|---:|")
                for k in addr_bucket_order:
                    n = ab.get(k, 0)
                    if n == 0:
                        continue
                    lines.append(f"| `{k}` | {n} |")
                lines.append("")
                lines.append(Z(
                    f"以上沒有任何一筆被視為已確認的 bug。412 筆 "
                    f"`mdb_stale_index_addr` 屬於維護週期差異（User MDB "
                    f"在下次釋出前需要重跑 frmBaseMaintenance）。10 筆 "
                    f"`same_candidates_diff_winner` 是唯一的候選演算法"
                    f"差異。逐筆輸出見 "
                    f"`reports/index_addr_drift_classification.json`。"
                ))
                lines.append("")
                lines.append(Z(
                    f"PR M（`analysis/dump_data_mdb_vba.py`）從 DATA mdb "
                    f"抽出了 `frmBaseMaintenance.CmdIndexAddress_Click`。"
                    f"它**沒有**像 PHP 那樣明確 `MAX(c_sequence)` 聚合 "
                    f"—— 在維護週期差異之外，這還是一個候選演算法差異。"
                    f"建議的 release checklist 緩解步驟：在 User MDB "
                    f"出貨前先在 DATA mdb 上跑 `CmdIndexYear`，再跑 "
                    f"`CmdIndexAddress`。詳見 "
                    f"`analysis/index_drift_algorithm_notes.md` 中的 "
                    f"\"Maintenance trigger path\" 段。"
                ))
                lines.append("")

                # ---- Cause analysis appendix (PR Y) ----
                if CAUSE_SUMMARY_JSON.exists():
                    cs = _json.loads(
                        CAUSE_SUMMARY_JSON.read_text(encoding="utf-8"))
                    lines.append(f"### {Z('目前能解釋的 drift 原因')}")
                    lines.append("")
                    lines.append(Z(
                        "每個 bucket 的成因／證據／信心度／下一步追查"
                        "都寫在 `analysis/index_drift_cause_analysis.md`。"
                        "本節只列每個 bucket 的計數和信心度摘要；目前"
                        "沒有任何 bucket 被列為已確認的 CBDB bug。"
                    ))
                    lines.append("")
                    lines.append(f"**{Z('c_index_year 原因桶')}**")
                    lines.append("")
                    lines.append(f"| Bucket | {Z('筆數')} | {Z('信心度')} |")
                    lines.append("|---|---:|---|")
                    for b in cs["c_index_year"]["buckets"]:
                        if b["count"] == 0:
                            continue
                        lines.append(
                            f"| `{b['bucket']}` | {b['count']} "
                            f"| {b['confidence']} |"
                        )
                    lines.append("")
                    lines.append(f"**{Z('c_index_addr_id 原因桶')}**")
                    lines.append("")
                    lines.append(f"| Bucket | {Z('筆數')} | {Z('信心度')} |")
                    lines.append("|---|---:|---|")
                    for b in cs["c_index_addr_id"]["buckets"]:
                        if b["count"] == 0:
                            continue
                        lines.append(
                            f"| `{b['bucket']}` | {b['count']} "
                            f"| {b['confidence']} |"
                        )
                    lines.append("")
                    lines.append(Z(
                        "建議優先處理的調查項目（完整列表見 cause-"
                        "analysis md）："
                    ))
                    lines.append("")
                    for inv in cs[
                        "suggested_next_investigations_in_priority_order"
                    ][:3]:
                        lines.append(Z(
                            f"{inv['id']}. {inv['task']} —— 可消化 "
                            f"{inv['would_close_rows']} 筆；工程成本："
                            f"{inv['engineering_cost']}。"
                        ))
                    lines.append("")

        data = _json.loads(DRIFT_JSON.read_text(encoding="utf-8"))
        bucket_meta = {
            "year_only": ("Examples where only c_index_year disagrees",
                          "仅 c_index_year 不一致的样例"),
            "addr_only": ("Examples where only c_index_addr_id disagrees",
                          "仅 c_index_addr_id 不一致的样例"),
            "both": ("Examples where both fields disagree",
                     "两个字段都不一致的样例"),
            "source_data": (
                "Examples where the SOURCE data itself differs (birthyear / deathyear)",
                "底层 SOURCE 数据本身不同（生年 / 卒年）的样例"),
        }
        for bucket_key, titles in bucket_meta.items():
            items = data.get(bucket_key, [])
            if not items:
                continue
            lines.append(f"### {Z(titles[0] if is_en else titles[1])}")
            lines.append("")
            for ex in items[:5]:
                u = ex["user"]; s = ex["sqlite"]
                head = (f"`c_personid = {ex['personid']}` — "
                        f"{ex['name_chn']} ({ex['name_py']})")
                lines.append(f"**{Z(head)}**")
                lines.append("")
                lines.append(
                    f"| {Z('Field' if is_en else '字段')} "
                    f"| {Z('User MDB' if is_en else '本 .mdb (User MDB)')} "
                    f"| {Z('cbdb-online-main-server snapshot' if is_en else 'cbdb-online-main-server 快照')} |"
                )
                lines.append("|---|---|---|")
                for f_label, key in [
                    ("c_index_year", "index_year"),
                    ("c_index_addr_id", "index_addr_id"),
                    ("c_birthyear", "birthyear"),
                    ("c_deathyear", "deathyear"),
                    ("c_index_year_type_code", "index_year_type_code"),
                    ("c_index_year_source_id", "index_year_source_id"),
                ]:
                    uv = "" if u[key] is None else str(u[key])
                    sv = "" if s[key] is None else str(s[key])
                    lines.append(f"| `{f_label}` | {uv} | {sv} |")
                lines.append("")

    # ---- Closing ----
    lines.append(f"## {Z('Closing note' if is_en else '结语')}")
    lines.append("")
    closing = (
        "Thank you for taking the time to read this report. None of the "
        "items above is urgent; we hope having them all in one place "
        "makes it easy to address them at your own pace.\n\n"
        "If any of the descriptions or suggested fixes are unclear, we "
        "would be glad to discuss further. The corresponding regression "
        "tests in this repository will automatically flip from PASS to "
        "FAIL the moment any regression marker stops reproducing in "
        "the source dump — that is a signal to investigate, not an "
        "automatic confirmation that the bug is fixed (the marker "
        "could fail because of an upstream fix, a fixture / driver "
        "change on our side, or a misclassification we made earlier)."
        if is_en else
        "感谢您抽时间读完这份报告。以上各条都不紧急，我们把它们集中整理"
        "在一起，只是希望方便您在合适的时候逐一处理。\n\n"
        "如果对其中任何一条的描述或建议有疑问，欢迎随时一同讨论。本仓库"
        "里对应的回归测试，会在任何一个回归标记不再复现时自动从 PASS "
        "翻成 FAIL —— 这是「请调查一下」的讯号，而不是「问题已修复」的"
        "自动确认（因为标记不再复现也可能是 fixture / driver 变了，"
        "或者是我们当初的分类有误）。"
    )
    for para in closing.split("\n\n"):
        lines.append(Z(para))
        lines.append("")

    out_path.write_text("\n".join(lines), encoding="utf-8")
    print(f"wrote {out_path}")


def main() -> int:
    # Always emit all four formats together so the docx and md never
    # drift apart.
    _build("en", OUT_EN)
    _build("zh", OUT_ZH)
    _build_md("en", OUT_EN_MD)
    _build_md("zh", OUT_ZH_MD)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
