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
        "tier": "P5_resolved_or_dormant",
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
        "id": 3,
        "tier": "P5_resolved_or_dormant",
        "form": "Form_LookAtEntry.CmdQuery_Click",
        "title_en": "LookAtEntry.CmdQuery backfill UPDATE — historical Bug #3, NOT reproducible on the current dump",
        "title_zh": "LookAtEntry.CmdQuery 回填 UPDATE — 歷史 Bug #3，當前 dump 上已無法復現",
        "summary_en": (
            "Historical context: an earlier dump of `Form_LookAtEntry.vb` "
            "(line 1778-1789, a single UPDATE joining seven+ lookup tables "
            "to backfill `c_entry_desc` / `c_addr_name` / `c_kin_name` "
            "etc. into `ZZ_SCRATCH_ENTRY`) was reported to silently leave "
            "those columns NULL on result sets above ~30 000 rows.\n\n"
            "**RE-VERIFIED on the current dump (2026-05-02): cannot "
            "reproduce.**  We fired CmdQuery on the same fixture (entry "
            "code 36 jinshi general, no year filter, 92,514 rows in "
            "`ZZ_SCRATCH_ENTRY`) and counted exactly **0 rows** with "
            "`c_entry_code IS NOT NULL` AND `c_entry_desc IS NULL`; same "
            "for `c_addr_id > 0` AND `c_addr_name IS NULL`.  The maintainer "
            "also confirmed the UI shows correct desc / addr columns.\n\n"
            "The giant multi-table UPDATE statement is still in the VBA "
            "module (the source code wasn't rewritten), so structurally "
            "the SQL pattern that was suspect remains.  But its runtime "
            "behaviour now produces correct backfills on this dump — "
            "likely because of a JET / Office update changing how it "
            "schedules complex UPDATE plans, or because the original "
            "diagnosis was a false positive.\n\n"
            "**Recommendation:** treat as resolved unless someone can "
            "produce a fresh repro on a current dump.  Verification "
            "script: `analysis/verify_bug3.py`."
        ),
        "summary_zh": (
            "歷史背景：早期某次 dump 的 `Form_LookAtEntry.vb`（第 1778-1789 "
            "行，用一條 UPDATE JOIN 七張以上 lookup 表回填 `c_entry_desc` / "
            "`c_addr_name` / `c_kin_name` 到 `ZZ_SCRATCH_ENTRY`）被報告"
            "在結果集 ~30000 行以上時靜默地讓這些欄位保持 NULL。\n\n"
            "**在當前 dump 上重新驗證（2026-05-02）：無法復現。** 我們用"
            "同樣的 fixture（入仕途徑 36 進士及第，不限年份，`ZZ_SCRATCH_"
            "ENTRY` 共 92,514 行）觸發 CmdQuery，精確統計 `c_entry_code "
            "IS NOT NULL` 且 `c_entry_desc IS NULL` 的行數為 **0**；"
            "`c_addr_id > 0` 且 `c_addr_name IS NULL` 也是 0。維護者本人"
            "也確認 UI 上 desc / addr 欄位顯示正確。\n\n"
            "那條龐大的多表 UPDATE SQL 仍在 VBA 模組裡（源碼沒被重寫），"
            "所以結構上當時被懷疑的 SQL 寫法仍在；但在當前 dump 上，其"
            "執行時行為已能正確回填 —— 可能是某次 JET / Office 更新改善"
            "了複雜 UPDATE 的執行計畫，或當時的診斷本身就是假陽性。\n\n"
            "**建議：** 視為已解決，除非有人能在當前 dump 上重新提出可"
            "復現的反例。驗證腳本：`analysis/verify_bug3.py`。"
        ),
        "steps_en": [
            "Run `python analysis/verify_bug3.py` from the repo root.",
            "It opens LookAtEntry, fires CmdQuery on entry code 36 with "
            "no year filter, and reports the count of rows whose "
            "`c_entry_desc` is NULL despite a non-null `c_entry_code`.",
            "On the current dump the count is 0 — the bug is no longer "
            "observable.  If a future dump regresses, this same script "
            "will report a non-zero count.",
        ],
        "steps_zh": [
            "在 repo 根目錄執行 `python analysis/verify_bug3.py`。",
            "腳本會開啟 LookAtEntry，對入仕 36 不加年份篩選觸發 CmdQuery，"
            "並回報 `c_entry_code` 非空但 `c_entry_desc` 為 NULL 的行數。",
            "在當前 dump 上這個數字是 0 —— bug 已不再可見。若未來 dump "
            "回歸退化，同一個腳本會報非零行數。",
        ],
        "fix_en": (
            "No action required for this dump.  If a future regression "
            "is observed: split the giant multi-table UPDATE into "
            "several smaller ones — one per lookup join — matching the "
            "pattern Status / Texts / Associations already use."
        ),
        "fix_zh": (
            "在當前 dump 上不需要任何動作。若未來再次回歸：把那條龐大的"
            "多表 UPDATE 拆成若干條小 UPDATE（每條只 JOIN 一張 lookup 表），"
            "與 Status / Texts / Associations 已使用的寫法一致。"
        ),
        "screenshots": [],
        "severity_en": "P5 — Resolved / not reproducible on current dump",
        "severity_zh": "P5 — 已解決 / 當前 dump 上無法復現",
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
            "well-attested address — for example **c_addr_id = 7213** "
            "(Kaifeng 開封) — so the resulting query has plenty of "
            "people to feed the People-CSV loop.  Click **Run Query**.",
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
            "地址——例如 **c_addr_id = 7213（開封）**——這樣查詢結果有"
            "足夠人物餵給 People-CSV 迴圈。點 **Run Query**。",
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
            ("bug7_step1_annotated.png", None),
            ("bug7_step2_faux_popup.png", "The popup users see (re-rendered for the report; the real popup blocks the COM thread our test driver runs in)."),
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
        "tier": "P0_silent_data",
        "form": "Form_LookAtEntry.CmdNeo4j_Click",
        "title_en": "LookAtEntry.CmdNeo4j Institutions block uses the wrong recordset variable",
        "title_zh": "LookAtEntry.CmdNeo4j 的机构 (Institutions) 部分用错了记录集变量",
        "summary_en": (
            "Line 1415 of `Form_LookAtEntry.vb` opens the institutions "
            "recordset as `tRstInstitutions = CurrentDb.OpenRecordset("
            "tQueryStr)`. Ten lines later, line 1425 says `With "
            "tRstAssocCodes` and the loop reads `!c_inst_code`, "
            "`!c_inst_name_code`, etc. against THAT recordset — which "
            "was bound much earlier to the AssocCodes SELECT and doesn't "
            "have `c_inst_*` columns. Same `Item not found` symptom; "
            "InstitutionCodes file is never written.\n\n"
            "Note: triggering this path requires entries with "
            "`c_inst_code > 0` (i.e. social-institution-bearing entries). "
            "Not every fixture reaches this With block."
        ),
        "summary_zh": (
            "`Form_LookAtEntry.vb` 第 1415 行打开 institutions 记录集："
            "`Set tRstInstitutions = CurrentDb.OpenRecordset(tQueryStr)`。"
            "十行之后，第 1425 行写的是 `With tRstAssocCodes`，循环又读 "
            "`!c_inst_code`、`!c_inst_name_code` 等，依据的却是早先绑定到 "
            "AssocCodes SELECT 的那个 tRstAssocCodes —— 那里没有 `c_inst_*`"
            "列。症状与 Issue #7 相同，InstitutionCodes 文件永远不会写出。\n\n"
            "注意：这条路径只在结果集中有 `c_inst_code > 0`（即带"
            "社会机构编码的入仕记录）时才会触发，并非每个 fixture 都会进入"
            "这个 With 块。"
        ),
        "steps_en": [
            "Open **LookAtEntry** with a query that produces entries with "
            "social institution codes (those are uncommon — most entries "
            "don't trigger this).",
            "Click **Neo4j**, accept all the SaveAs dialogs.",
            "When the export reaches the InstitutionCodes file, the same "
            "`Item not found` popup appears.",
        ],
        "steps_zh": [
            "用一组会产生「带社会机构编码的入仕」的查询条件打开 **LookAtEntry**"
            "（这种入仕较少见，大多数查询触发不到）。",
            "点 **Neo4j**，依次确认每个保存对话框。",
            "走到 InstitutionCodes 文件这一步时，同样的「Item not found」"
            "对话框弹出来。",
        ],
        "fix_en": (
            "Change `With tRstAssocCodes` on line 1425 to `With "
            "tRstInstitutions`. Single-character class of fix; the "
            "underlying recordset variable was simply mis-named."
        ),
        "fix_zh": (
            "把第 1425 行的 `With tRstAssocCodes` 改成 `With tRstInstitutions`。"
            "属于一字之差的笔误，底层记录集变量只是写错了。"
        ),
        "screenshots": [
            ("bug9_form_annotated.png",
             "Step 1 — open LookAtEntry, run any query, click "
             "**Neo4j**.  (Note: this code path only fires for "
             "queries whose result includes entries with "
             "`c_inst_code > 0` — see the summary's "
             "REAL_BUT_GATED note.)"),
            ("bug9_faux_popup.png",
             "Step 2 — the popup users see when the With block on "
             "line 1425 reads `!c_inst_code` against the "
             "wrong-named recordset.  Reconstructed in PIL because "
             "the real popup would block the COM test driver; the "
             "error code (DAO 3265) and message text are JET's "
             "standard response to a recordset field that doesn't "
             "exist."),
        ],
        "severity_en": "P0 — Silent data corruption (export silently produces nothing)",
        "severity_zh": "P0 — 静默数据缺失（导出无声地什么都没生成）",
    },
    # ========== Tier 2: visible runtime crash (popup blocks user) ==========
    {
        "id": 4,
        "tier": "P5_resolved_or_dormant",
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
            ("bug4_step1_annotated.png", None),
            ("bug4_step2_annotated.png", None),
            ("bug4_step3_faux_popup.png", "Re-rendered popup — exact runtime error users would see if the button were present."),
        ],
        "severity_en": "P5 — Latent (would be P1 if Issue #15 fixed without first fixing this)",
        "severity_zh": "P5 — 潛伏（若先修了 Issue #15 而沒同時修本條，會變成 P1）",
    },
    {
        "id": 5,
        "tier": "P5_resolved_or_dormant",
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
        "id": 14,
        "tier": "P5_resolved_or_dormant",
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
        "tier": "P5_resolved_or_dormant",
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
        "tier": "P5_resolved_or_dormant",
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
        "vs the cbdb-online-main-server snapshot (not bugs)"
        if is_en else
        "附录 —— `c_index_year` / `c_index_addr_id` 与 "
        "cbdb-online-main-server 快照之间的偏差（非缺陷）"
    )
    _h(doc, 1, Z(title))

    intro = (
        "When we compare BIOG_MAIN's `c_index_year` and `c_index_addr_id` "
        "between this User MDB and the weekly cbdb-online-main-server "
        "SQLite snapshot, a small fraction of persons disagree. We want "
        "to be very clear that these are NOT regressions — both pipelines "
        "run the same `IndexYearRebuildService.php` algorithm, but on "
        "different snapshots of source data and with different downstream "
        "decisions. We list a handful of representative examples below "
        "with the underlying field values so you can see exactly what "
        "kind of drift this is.\n\n"
        "Why we still document this: when the SAME person disagrees in "
        "the same way across many releases, it tells us nothing new; "
        "but if the SHAPE of disagreement changes (e.g. addresses start "
        "diverging where only years used to), that hints the algorithm "
        "or schema shifted, and we'd want to look at it."
        if is_en else
        "我们把本 .mdb 的 BIOG_MAIN 与 cbdb-online-main-server 每周发布的 "
        "SQLite 快照在 `c_index_year`、`c_index_addr_id` 两个字段上做比对，"
        "可以看到一小部分人物对不齐。我们希望明确说明：这并不是缺陷 —— "
        "两套管线跑的都是同一段 `IndexYearRebuildService.php` 算法，只是依据"
        "的源数据快照不一样，下游某些择优规则也略有出入。下面列举若干典型"
        "样例，并展示底层字段值，方便您一眼看出这是哪种偏差。\n\n"
        "我们仍然把它放进报告，是因为：如果同一个人物在多次发布里都按同样"
        "方式对不齐，那不增加任何信息；但若偏差的「类型」发生变化（例如"
        "原本只是年份不同，现在地点也开始不同），那说明算法或 schema 有过"
        "变动，值得我们关注。"
    )
    for para in intro.split("\n\n"):
        doc.add_paragraph(Z(para))

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
        "P5 — Resolved / not currently reproducible: kept as historical "
        "record; we re-checked on the current dump and could not "
        "trigger the symptom.",
    ]
    legend_zh = [
        "P0 — 静默数据错误：数据错或缺失，但没有任何报错提示。",
        "P1 — 可见的运行时报错：弹出错误对话框，操作中断。",
        "P2 — 静默显示问题：表单字段本应有数据，却显示为空。",
        "P3 — 缺失界面：代码里实现了某功能，但界面上没有按钮去触发它。",
        "P4 — 安装设置：每台新机器需要一次性处理。",
        "P5 — 已解决 / 当前无法复现：保留作为历史记录；我们在当前 dump "
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
                  "P5_resolved_or_dormant"]
    tier_titles_en = {
        "P0_silent_data": "P0 — Silent data corruption",
        "P1_visible_crash": "P1 — Visible runtime crash",
        "P2_silent_display": "P2 — Silent display",
        "P3_missing_ui": "P3 — Missing UI",
        "P4_setup": "P4 — Setup",
        "P5_resolved_or_dormant":
            "P5 — Resolved / not currently reproducible",
    }
    tier_titles_zh = {
        "P0_silent_data": "P0 — 静默数据错误",
        "P1_visible_crash": "P1 — 可见的运行时报错",
        "P2_silent_display": "P2 — 静默显示问题",
        "P3_missing_ui": "P3 — 缺失界面",
        "P4_setup": "P4 — 安装设置",
        "P5_resolved_or_dormant": "P5 — 已解决 / 当前无法复现",
    }
    demo_persons = _load_demo_persons()
    bug_status = _load_bug_test_status()

    for tier in tier_order:
        items = by_tier.get(tier, [])
        if not items:
            continue
        _h(doc, 1, Z(tier_titles_en[tier] if is_en
                      else tier_titles_zh[tier]))
        if tier == "P5_resolved_or_dormant":
            preface = (
                "Items in this tier are kept as historical / latent "
                "record.  They fall into three categories: (a) DORMANT "
                "— current source data doesn't trigger the symptom; "
                "(b) RESOLVED — the symptom no longer occurs even "
                "though the suspect code is still present (likely "
                "fixed by some Office / JET update or a previous "
                "iteration); (c) LATENT — the source-code defect is "
                "real, but the user can't reach it because another "
                "issue (e.g. a missing UI button) blocks the path. "
                "None of these are user-facing today; please consult "
                "before treating any of them as urgent."
                if is_en else
                "本層的條目作為歷史 / 潛伏記錄保留。可分為三類："
                "(a) DORMANT 潛伏 —— 已驗證當前源資料無法觸發該症狀；"
                "(b) RESOLVED 已解決 —— 症狀不再出現，雖然可疑程式碼"
                "仍在（可能是某次 Office / JET 更新或更早一次修補解決"
                "的）；(c) LATENT 被屏蔽 —— 源碼缺陷確實存在，但因為"
                "另一個 issue（例如某個 UI 按鈕缺失）擋住了使用路徑，"
                "使用者目前碰不到。本層條目當下都不是使用者會遇到的"
                "問題；若要當成緊急問題處理，請先諮詢。"
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
                        "usually means the underlying defect has been "
                        "FIXED in the source dump.  Please verify in "
                        "person before considering this issue closed; "
                        "this report has NOT been edited to drop the "
                        "issue.  Tests consulted:\n"
                        + "\n".join(
                            f"  • {t.rsplit('::', 1)[-1]}"
                            for t, _ in status["tests"]
                        )
                        if is_en else
                        "⚠ 自動測試狀態：本 issue 對應的回歸標記目前 "
                        f"FAIL（執行時間：{status.get('when', '未知')}），"
                        "通常意味著底層缺陷已在 source dump 中被修復。"
                        "請務必親自確認，再將此 issue 視為關閉；本報告"
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
                        "Likely the issue is partially fixed; please "
                        "review the per-test breakdown:\n"
                        + "\n".join(
                            f"  • {t.rsplit('::', 1)[-1]}: {o}"
                            for t, o in status["tests"]
                        )
                        if is_en else
                        "ℹ 自動測試狀態：本 issue 對應的回歸標記呈現"
                        f"混合結果（執行時間：{status.get('when', '未知')}）。"
                        "可能是部分修復，請查看分項：\n"
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
        "FAIL the moment any issue is fixed in the source dump — so you "
        "can use them as a confirmation signal."
        if is_en else
        "感谢您抽时间读完这份报告。以上各条都不紧急，我们把它们集中整理"
        "在一起，只是希望方便您在合适的时候逐一处理。\n\n"
        "如果对其中任何一条的描述或建议有疑问，欢迎随时一同讨论。本仓库"
        "里对应的回归测试，会在您修好任意一个问题、并重新导出 dump 之后"
        "自动从 PASS 翻成 FAIL —— 可以作为修复完成的信号使用。"
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
        "P5_resolved_or_dormant":
            "P5 — Resolved / not currently reproducible",
    }
    tier_titles_zh = {
        "P0_silent_data": "P0 — 静默数据错误",
        "P1_visible_crash": "P1 — 可见的运行时报错",
        "P2_silent_display": "P2 — 静默显示问题",
        "P3_missing_ui": "P3 — 缺失界面",
        "P4_setup": "P4 — 安装设置",
        "P5_resolved_or_dormant": "P5 — 已解决 / 当前无法复现",
    }
    tier_order = ["P0_silent_data", "P1_visible_crash",
                  "P2_silent_display", "P3_missing_ui", "P4_setup",
                  "P5_resolved_or_dormant"]

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
        "vs the cbdb-online-main-server snapshot (not bugs)"
        if is_en else
        "附录 —— c_index_year / c_index_addr_id 与 "
        "cbdb-online-main-server 快照之间的偏差（非缺陷）"
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
        "P5 — Resolved / not currently reproducible: kept as historical "
        "record; we re-checked on the current dump and could not "
        "trigger the symptom.",
    ]
    legend_zh = [
        "P0 — 静默数据错误：数据错或缺失，但没有任何报错提示。",
        "P1 — 可见的运行时报错：弹出错误对话框，操作中断。",
        "P2 — 静默显示问题：表单字段本应有数据，却显示为空。",
        "P3 — 缺失界面：代码里实现了某功能，但界面上没有按钮去触发它。",
        "P4 — 安装设置：每台新机器需要一次性处理。",
        "P5 — 已解决 / 当前无法复现：保留作为历史记录；我们在当前 dump "
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
        if tier == "P5_resolved_or_dormant":
            lines.append(Z(
                "_Items in this tier are kept as historical / latent "
                "record.  They fall into three categories: (a) DORMANT — "
                "verified that current source data doesn't trigger the "
                "symptom; (b) RESOLVED — the symptom no longer occurs "
                "even though the suspect code is still present (likely "
                "fixed by some Office / JET update or a previous "
                "iteration); (c) LATENT — the source-code defect is "
                "real, but the user can't reach it because another "
                "issue (e.g. a missing UI button) blocks the path.  "
                "None of these are user-facing today; please consult "
                "before treating any of them as urgent._"
                if is_en else
                "_本層的條目作為歷史 / 潛伏記錄保留。可分為三類："
                "(a) DORMANT 潛伏 — 已驗證當前源資料無法觸發該症狀；"
                "(b) RESOLVED 已解決 — 症狀不再出現，雖然可疑程式碼仍"
                "在（可能是某次 Office / JET 更新或更早一次修補解決的）；"
                "(c) LATENT 被屏蔽 — 源碼缺陷確實存在，但因為另一個 "
                "issue（例如某個 UI 按鈕缺失）擋住了使用路徑，使用者"
                "目前碰不到。本層條目當下都不是使用者會遇到的問題；"
                "若要當成緊急問題處理，請先諮詢。_"
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
                        "been **FIXED** in the source dump.  Please "
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
            "cbdb-online-main-server SQLite snapshot, a small fraction "
            "of persons disagree. We want to be very clear that these "
            "are NOT regressions — both pipelines run the same "
            "`IndexYearRebuildService.php` algorithm, but on different "
            "snapshots of source data and with different downstream "
            "decisions."
            if is_en else
            "我们把本 .mdb 的 BIOG_MAIN 与 cbdb-online-main-server 每周"
            "发布的 SQLite 快照在 `c_index_year`、`c_index_addr_id` 两个"
            "字段上做比对，可以看到一小部分人物对不齐。我们希望明确说明："
            "这并不是缺陷 —— 两套管线跑的都是同一段 "
            "`IndexYearRebuildService.php` 算法，只是依据的源数据快照"
            "不一样，下游某些择优规则也略有出入。"
        )
        lines.append(Z(intro_drift))
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
        "FAIL the moment any issue is fixed in the source dump — so you "
        "can use them as a confirmation signal."
        if is_en else
        "感谢您抽时间读完这份报告。以上各条都不紧急，我们把它们集中整理"
        "在一起，只是希望方便您在合适的时候逐一处理。\n\n"
        "如果对其中任何一条的描述或建议有疑问，欢迎随时一同讨论。本仓库"
        "里对应的回归测试，会在您修好任意一个问题、并重新导出 dump 之后"
        "自动从 PASS 翻成 FAIL —— 可以作为修复完成的信号使用。"
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
