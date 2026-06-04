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
    # ---- build-20260604 / data-20260602 test run findings ----
    # Only entries confirmed by test failures in THIS run.
    # No carry-forward from previous builds.

    {
        "id": 21,
        "tier": "P0_silent_data",
        "form": "Form_LookAtOffice.CmdGIS_Click",
        "title_en": "LookAtOffice: CmdGIS output IndexYear column is nearly empty (0.2% fill rate) — likely silent column-bind regression",
        "title_zh": "LookAtOffice：CmdGIS 輸出的 IndexYear 欄幾乎為空（填充率 0.2%）——疑似靜默欄位綁定退化",
        "summary_en": (
            "When CmdGIS runs for LookAtOffice with person 80944 (unfiltered), the GIS output "
            "file is produced but the IndexYear column contains non-empty values in only 64 of "
            "36,602 rows (0.2%), well below the 80% threshold expected for a correctly-populated "
            "GIS export.  This pattern is consistent with the silent column-bind regressions "
            "documented in Bugs #10, #11, and #12 — a column name in the CmdGIS SELECT is "
            "mismatched against the actual ZZ_SCRATCH table schema.\n\n"
            "Detected by: test_cmd_gis_produces_file[office_80944_unfiltered] — assertion "
            "[LookAtOffice] CmdGIS column 'IndexYear' is non-empty in only 64/36602 rows (0.2%), "
            "below 80% threshold."
        ),
        "summary_zh": (
            "以人物 80944（無篩選）執行 LookAtOffice CmdGIS 時，GIS 輸出檔案雖已產生，"
            "但 IndexYear 欄僅在 36,602 列中的 64 列有非空值（0.2%），遠低於正確 GIS "
            "輸出預期的 80% 閾值。此模式與 Bug #10、#11、#12 記錄的靜默欄位綁定退化一致——"
            "CmdGIS SELECT 中的欄位名稱與 ZZ_SCRATCH 表格的實際 schema 不符。\n\n"
            "由 test_cmd_gis_produces_file[office_80944_unfiltered] 偵測到，斷言 "
            "[LookAtOffice] CmdGIS 欄 IndexYear 僅 64/36602 列非空（0.2%），低於 80% 閾值。"
        ),
        "steps_en": [
            "Open CBDB_BJ_User.mdb in Microsoft Access.",
            "Open the LookAtOffice form.",
            "Set person ID to 80944 and leave all other filters blank.",
            "Click CmdGIS.  The file is produced without an error popup.",
            "Open the GIS output file and inspect the IndexYear column: "
            "the vast majority of rows will be empty.",
        ],
        "steps_zh": [
            "以 Microsoft Access 開啟 CBDB_BJ_User.mdb。",
            "開啟 LookAtOffice 表單。",
            "將人物 ID 設為 80944，其餘篩選器留空。",
            "點擊 CmdGIS。檔案產生時不會出現錯誤彈出視窗。",
            "開啟 GIS 輸出檔並檢查 IndexYear 欄：絕大多數列將為空白。",
        ],
        "screenshots": [],
        "severity_en": (
            "P0 — Silent data corruption: the GIS export appears to succeed but IndexYear "
            "data is missing for 99.8% of rows.  Downstream GIS workflows that depend on "
            "year-based filtering will silently receive null years."
        ),
        "severity_zh": (
            "P0 — 靜默資料損毀：GIS 匯出看似成功，但 99.8% 列的 IndexYear 資料遺失。"
            "依賴年份篩選的下游 GIS 工作流程將靜默地收到空白年份。"
        ),
        "fix_en": (
            "Inspect Form_LookAtOffice.CmdGIS_Click: locate the SELECT that populates "
            "the IndexYear column in the GIS output and verify the source column name "
            "matches the actual schema (check ZZ_SCRATCH_OFFICE or the equivalent table)."
        ),
        "fix_zh": (
            "檢查 Form_LookAtOffice.CmdGIS_Click：找到填充 GIS 輸出 IndexYear 欄的 SELECT 語句，"
            "確認來源欄位名稱與實際 schema 一致（檢查 ZZ_SCRATCH_OFFICE 或對應資料表）。"
        ),
    },
    {
        "id": 26,
        "tier": "P0_silent_data",
        "form": "BIOG_MAIN (c_index_addr_id)",
        "title_en": "c_index_addr_id disagreement between User MDB and cbdb-online snapshot exceeds 0.5% threshold",
        "title_zh": "User MDB 與 cbdb-online 快照的 c_index_addr_id 不一致率超過 0.5% 閾值",
        "summary_en": (
            "A cross-check of BIOG_MAIN.c_index_addr_id between the User MDB and the "
            "cbdb-online-main-server SQLite snapshot found a disagreement rate of 0.500%, "
            "exactly at the maximum acceptable threshold.  At the default 5,000-row sample "
            "this means approximately 25 persons have a different c_index_addr_id in the "
            "two systems, indicating that either the User MDB has not fully applied recent "
            "upstream address assignments or the snapshot is ahead of the current data export.\n\n"
            "Detected by: test_index_year_addr_xcheck_sample — assertion "
            "c_index_addr_id disagreement 0.500% exceeds 0.5% threshold."
        ),
        "summary_zh": (
            "對 User MDB 與 cbdb-online-main-server SQLite 快照之間的 BIOG_MAIN.c_index_addr_id "
            "進行交叉核查，發現不一致率為 0.500%，恰好達到最大可接受閾值。以預設 5,000 列樣本計，"
            "約有 25 名人物在兩個系統中的 c_index_addr_id 不同，表明 User MDB 可能尚未完全套用"
            "近期上游地址指定，或快照領先於當前資料匯出。\n\n"
            "由 test_index_year_addr_xcheck_sample 偵測到，斷言 "
            "c_index_addr_id 不一致率 0.500% 超過 0.5% 閾值。"
        ),
        "steps_en": [
            "Run: python reports/collect_index_year_diffs.py",
            "Inspect reports/index_drift_examples.json for rows where the bucket is "
            "'addr_only' — these are persons where c_index_addr_id differs between "
            "the User MDB and the online snapshot.",
            "For each differing person, query BIOG_MAIN.c_index_addr_id and compare "
            "against the online server to determine which value is authoritative.",
        ],
        "steps_zh": [
            "執行：python reports/collect_index_year_diffs.py",
            "檢查 reports/index_drift_examples.json 中 bucket 為 'addr_only' 的列——"
            "這些是 User MDB 與線上快照之間 c_index_addr_id 不同的人物。",
            "對每個不一致的人物，查詢 BIOG_MAIN.c_index_addr_id 並與線上伺服器比較，"
            "確認哪個值為權威值。",
        ],
        "screenshots": [],
        "severity_en": (
            "P0 — Silent data drift: ~25 persons have a different primary address ID than "
            "the online system.  Geographic analyses and GIS exports that use c_index_addr_id "
            "will silently place these persons at the wrong location."
        ),
        "severity_zh": (
            "P0 — 靜默資料漂移：約 25 名人物的主要地址 ID 與線上系統不同。使用 c_index_addr_id "
            "的地理分析和 GIS 匯出將靜默地將這些人物置於錯誤位置。"
        ),
        "fix_en": (
            "Apply the latest c_index_addr_id assignments from the cbdb-online server to "
            "BIOG_MAIN in the User MDB.  The differing rows are enumerated in "
            "reports/index_drift_examples.json (bucket: 'addr_only')."
        ),
        "fix_zh": (
            "將 cbdb-online 伺服器最新的 c_index_addr_id 指定套用至 User MDB 中的 BIOG_MAIN。"
            "不一致的列已列舉於 reports/index_drift_examples.json（bucket: 'addr_only'）中。"
        ),
    },
    {
        "id": 23,
        "tier": "P0_silent_data",
        "form": "Form_LookAtAssociations.CmdPajek_Click",
        "title_en": "LookAtAssociations: CmdPajek vertex section has off-by-N count — header declares 501 vertices but exports 8,093 rows",
        "title_zh": "LookAtAssociations：CmdPajek 頂點區段數量錯誤——標頭宣告 501 個頂點，但實際匯出 8,093 列",
        "summary_en": (
            "The Pajek .net file produced by Form_LookAtAssociations.CmdPajek_Click declares "
            "'*Vertices 501' in the header but the actual vertex section contains 8,093 rows "
            "before the next `*` section marker.  Pajek and other network analysis tools that "
            "rely on the vertex count header will either truncate the vertex list after 501 "
            "rows or raise a parse error, silently discarding the remaining ~7,592 vertices "
            "from the network.\n\n"
            "Detected by: test_export_button_produces_file[LookAtAssociations_CmdPajek] — "
            "assertion header declared 501 vertices but found 8093 vertex rows before the next "
            "`*` section."
        ),
        "summary_zh": (
            "Form_LookAtAssociations.CmdPajek_Click 產生的 Pajek .net 檔案在標頭宣告 "
            "'*Vertices 501'，但實際頂點區段在下一個 `*` 標記之前包含 8,093 列。Pajek "
            "及其他依賴頂點計數標頭的網路分析工具將在 501 列後截斷頂點列表或觸發解析錯誤，"
            "靜默地丟棄其餘約 7,592 個頂點。\n\n"
            "由 test_export_button_produces_file[LookAtAssociations_CmdPajek] 偵測到，"
            "斷言標頭宣告 501 個頂點但在下一個 `*` 區段前找到 8,093 列頂點。"
        ),
        "steps_en": [
            "Open CBDB_BJ_User.mdb in Microsoft Access.",
            "Open LookAtAssociations, select a query that returns a large association network.",
            "Click CmdPajek.  The .net file is written without an error popup.",
            "Open the .net file and count the lines in the *Vertices section: "
            "the count exceeds the number declared in the '*Vertices N' header.",
        ],
        "steps_zh": [
            "以 Microsoft Access 開啟 CBDB_BJ_User.mdb。",
            "開啟 LookAtAssociations，選擇一個回傳大型關聯網絡的查詢。",
            "點擊 CmdPajek。.net 檔案寫入時不出現錯誤視窗。",
            "開啟 .net 檔案並計算 *Vertices 區段的列數：實際數量超過 '*Vertices N' 標頭中宣告的數字。",
        ],
        "screenshots": [],
        "severity_en": (
            "P0 — Silent data corruption: the exported Pajek file is structurally invalid. "
            "Network analyses that ingest the file will operate on a truncated vertex set, "
            "producing incorrect centrality / community detection results without any warning."
        ),
        "severity_zh": (
            "P0 — 靜默資料損毀：匯出的 Pajek 檔案在結構上無效。讀取該檔案的網路分析將在"
            "截斷的頂點集上運行，在沒有任何警告的情況下產生不正確的中心性/社群偵測結果。"
        ),
        "fix_en": (
            "In Form_LookAtAssociations.CmdPajek_Click, locate where the '*Vertices N' header "
            "is written and ensure N is derived from the actual count of vertex rows written, "
            "not a pre-computed estimate or a separate query result."
        ),
        "fix_zh": (
            "在 Form_LookAtAssociations.CmdPajek_Click 中，找到寫入 '*Vertices N' 標頭的位置，"
            "確保 N 來自實際寫入的頂點列數，而非預先計算的估計值或單獨的查詢結果。"
        ),
    },
    {
        "id": 24,
        "tier": "P0_silent_data",
        "form": "Form_LookAtKinship.CmdGUESS_Click",
        "title_en": "LookAtKinship: CmdGUESS Gephi output has wrong field count per node row (nodedef declares 15 columns)",
        "title_zh": "LookAtKinship：CmdGUESS Gephi 輸出每個節點列的欄位數錯誤（nodedef 宣告 15 欄）",
        "summary_en": (
            "The Gephi .gdf file produced by Form_LookAtKinship.CmdGUESS_Click declares 15 "
            "columns in the nodedef header but the actual node data rows contain a different "
            "number of fields (column/value misalignment).  Gephi and downstream tools will "
            "either fail to load the file or silently map node attributes to the wrong columns.\n\n"
            "Detected by: test_cmd_guess_produces_file[kinship_person_3211] — assertion "
            "[LookAtKinship] Gephi: node rows with bad field count (nodedef has 15 cols)."
        ),
        "summary_zh": (
            "Form_LookAtKinship.CmdGUESS_Click 產生的 Gephi .gdf 檔案在 nodedef 標頭宣告 "
            "15 個欄位，但實際節點資料列包含不同數量的欄位（欄位/值錯位）。Gephi 及下游工具"
            "將無法載入該檔案，或靜默地將節點屬性對應到錯誤的欄位。\n\n"
            "由 test_cmd_guess_produces_file[kinship_person_3211] 偵測到，斷言 "
            "[LookAtKinship] Gephi: node rows with bad field count (nodedef has 15 cols)。"
        ),
        "steps_en": [
            "Open CBDB_BJ_User.mdb in Microsoft Access.",
            "Open LookAtKinship.  Set a person ID that returns a kinship network.",
            "Click CmdGUESS.  The .gdf file is written without error.",
            "Open the .gdf file: count the columns declared in the 'nodedef>' header "
            "and compare with the number of comma-separated values in the first node data row.",
        ],
        "steps_zh": [
            "以 Microsoft Access 開啟 CBDB_BJ_User.mdb。",
            "開啟 LookAtKinship，設定一個可回傳親屬網絡的人物 ID。",
            "點擊 CmdGUESS。.gdf 檔案寫入時不出現錯誤。",
            "開啟 .gdf 檔案：計算 'nodedef>' 標頭宣告的欄數，與第一個節點資料列中逗號分隔值的數量進行比較。",
        ],
        "screenshots": [],
        "severity_en": (
            "P0 — Silent data corruption: the Gephi file is structurally invalid. "
            "Node attributes are silently misaligned, making all imported node metadata unreliable."
        ),
        "severity_zh": (
            "P0 — 靜默資料損毀：Gephi 檔案在結構上無效。節點屬性靜默錯位，導致所有匯入的節點元資料不可靠。"
        ),
        "fix_en": (
            "In Form_LookAtKinship.CmdGUESS_Click, ensure the nodedef header column list and "
            "the per-row value list are generated from the same ordered column definition. "
            "A mismatch typically occurs when a column is added to one list but not the other."
        ),
        "fix_zh": (
            "在 Form_LookAtKinship.CmdGUESS_Click 中，確保 nodedef 標頭欄位列表與每列值列表"
            "來自同一個有序欄位定義。不符情況通常發生在某欄位被加入一個列表但未加入另一個時。"
        ),
    },
    {
        "id": 20,
        "tier": "P5_dormant_or_latent",
        "form": "ADDR_CODES + Form_LookAt*.CmdGIS_Click",
        "title_en": "BOM-prefixed address names would produce embedded TAB delimiters in GIS exports — DORMANT (BOM data cleaned upstream, 0 affected rows in this dump)",
        "title_zh": "BOM 前綴地址名稱在 GIS 匯出中會產生嵌入的 TAB 分隔符——當前資料集靜止（BOM 資料已在上游清理，本 dump 中受影響列數為 0）",
        "summary_en": (
            "The ADDR_CODES table previously contained rows where c_name and c_name_chn "
            "carried a leading U+FEFF (BOM) prefix (almost certainly from a UTF-8-with-BOM "
            "paste at data-import time), which when passed through JET produced embedded TAB "
            "characters in the CmdGIS output.  The key reachable example was c_addr_id = 702559 "
            "(Wei Shi 尉氏), reachable from persons with c_status_code = 40 (civil office / "
            "[為官者：文]).\n\n"
            "In this build (data-20260602) the BOM data has been cleaned upstream: "
            "test_addr_codes_has_known_bom_dirty_rows now finds 0 BOM-prefixed rows, and "
            "test_known_reachable_dirty_addr_present also returns 0 rows.  The GIS unescaped-"
            "write pattern remains in the code (CmdGIS of LookAtTexts / LookAtPlace / "
            "LookAtAssociations / LookAtOffice / LookAtKinship still does no TAB escaping), "
            "so the structural risk re-activates if future data introduces another BOM row.\n\n"
            "GOLDEN_STALE: BOM golden tests now expect 0 rows; update goldens.  This issue "
            "is retained as P5_dormant_or_latent because the unescaped-write code defect is "
            "still present."
        ),
        "summary_zh": (
            "ADDR_CODES 表先前有一些 c_name 和 c_name_chn 帶有前導 U+FEFF（BOM）前綴的列"
            "（幾乎可確定是資料匯入時以帶 BOM 的 UTF-8 貼上所致），經 JET 處理後會在 CmdGIS "
            "輸出中產生嵌入的 TAB 字元。最主要的可達範例是 c_addr_id = 702559（尉氏 Wei Shi），"
            "可從 c_status_code = 40（[為官者：文] / civil office）的人物到達。\n\n"
            "在此版本（data-20260602）中，BOM 資料已在上游完成清理：test_addr_codes_has_known_"
            "bom_dirty_rows 現在找到 0 個帶 BOM 前綴的列，test_known_reachable_dirty_addr_"
            "present 也回傳 0 列。CmdGIS 中未轉義的寫入模式仍存在於程式碼中（LookAtTexts / "
            "LookAtPlace / LookAtAssociations / LookAtOffice / LookAtKinship 的 CmdGIS 仍未"
            "進行 TAB 轉義），因此若未來資料引入新的 BOM 列，結構性風險將再次啟動。\n\n"
            "GOLDEN_STALE：BOM golden 測試現在預期 0 列；請更新 goldens。此 Issue 因未轉義"
            "寫入的程式碼缺陷仍然存在，故以 P5_dormant_or_latent 保留。"
        ),
        "steps_en": [
            "In this build the bug cannot be triggered — SELECT COUNT(*) FROM ADDR_CODES "
            "WHERE Left(c_name, 1) = ChrW(65279) returns 0.",
            "The structural risk: open Form_LookAtOffice (or any LookAt form), run CmdGIS "
            "with status code **40** (civil office / [為官者：文]).  "
            "If a BOM row were present for c_addr_id = 702559 (Wei Shi 尉氏) the output "
            "file would have an extra TAB column around row 11476.",
            "(Dormant verification) Confirm: SELECT COUNT(*) FROM ADDR_CODES WHERE "
            "Left(c_name, 1) = ChrW(65279) returns 0 on this dump.",
        ],
        "steps_zh": [
            "在此版本中，Bug 無法觸發——SELECT COUNT(*) FROM ADDR_CODES WHERE "
            "Left(c_name, 1) = ChrW(65279) 回傳 0。",
            "結構性風險：開啟 Form_LookAtOffice（或任意 LookAt 表單），以狀態碼 **40**"
            "（[為官者：文] / civil office）執行 CmdGIS。若 c_addr_id = 702559（尉氏 Wei Shi）"
            "存在 BOM 列，輸出檔案在第 11476 列附近會出現多餘的 TAB 欄。",
            "（靜止驗證）確認：SELECT COUNT(*) FROM ADDR_CODES WHERE Left(c_name, 1) = "
            "ChrW(65279) 在此 dump 中回傳 0。",
        ],
        "screenshots": [],
        "severity_en": (
            "P5 — Dormant on this dump (BOM data cleaned upstream, 0 affected rows in "
            "data-20260602).  Would re-activate as P0 silent data corruption if any future "
            "ADDR_CODES row re-introduces a U+FEFF prefix in c_name or c_name_chn.  "
            "The unescaped CmdGIS write code is the second half of the fix still outstanding."
        ),
        "severity_zh": (
            "P5 — 在此 dump 中為靜止（BOM 資料已在上游清理，data-20260602 中受影響列數為 0）。"
            "若未來任何 ADDR_CODES 列在 c_name 或 c_name_chn 重新引入 U+FEFF 前綴，"
            "將重新啟動並提升為 P0 靜默資料損毀。未轉義的 CmdGIS 寫入程式碼是尚未解決的"
            "第二部分修復。"
        ),
        "fix_en": (
            "The data-side fix has been applied upstream (0 BOM rows remain).  "
            "The code-side fix is still needed: before each tStr = tStr + value + tC append "
            "in the CmdGIS bodies of all LookAt forms, replace any embedded Chr(9), Chr(10), "
            "Chr(13), or U+FEFF in value with a space.  This prevents re-occurrence if "
            "future imports bring BOM-prefixed rows."
        ),
        "fix_zh": (
            "資料端修復已在上游套用（0 個 BOM 列殘留）。程式碼端修復仍需進行：在所有 "
            "LookAt 表單 CmdGIS 主體的每個 tStr = tStr + value + tC 追加之前，將 value "
            "中嵌入的 Chr(9)、Chr(10)、Chr(13) 或 U+FEFF 替換為空格，以防未來匯入帶 BOM "
            "前綴列時再次發生。"
        ),
    },
    {
        "id": 22,
        "tier": "P1_visible_crash",
        "form": "Form_LookAtAssociations.CmdUCINet_Click / Form_LookAtKinship.CmdUCINet_Click",
        "title_en": "LookAtAssociations / LookAtKinship: CmdUCINet crashes with 'Invalid procedure call or argument' when c_name contains CJK Han characters",
        "title_zh": "LookAtAssociations / LookAtKinship：當 c_name 含有 CJK 漢字時，CmdUCINet 因「Invalid procedure call or argument」崩潰",
        "summary_en": (
            "Form_LookAtAssociations.CmdUCINet_Click and Form_LookAtKinship.CmdUCINet_Click "
            "both call CreateTextFile with the 2-argument signature (filename, overwrite).  "
            "VBA's CreateTextFile raises 'Invalid procedure call or argument' (runtime error 5) "
            "when the output path contains a CJK Han character in a c_name value — the 2-arg "
            "form does not accept a Unicode flag, so Access silently uses the system ANSI code "
            "page, which cannot encode Han characters.  The error fires as a popup and aborts "
            "the export.  Fixtures using association code c_assoc_code = 437 "
            "('Presented literary composition as gift to' / '贈詩、文') reliably trigger this "
            "because the associated persons include Han-character names.\n\n"
            "Detected by: test_bug22_associations_cmducinet_fires_invalid_procedure_call and "
            "test_bug22_kinship_cmducinet_sibling_form_fires_invalid_procedure_call."
        ),
        "summary_zh": (
            "Form_LookAtAssociations.CmdUCINet_Click 和 Form_LookAtKinship.CmdUCINet_Click "
            "均以 2 參數形式呼叫 CreateTextFile（filename, overwrite），未傳入 Unicode 旗標。"
            "當輸出路徑或 c_name 值包含 CJK 漢字時，VBA 的 CreateTextFile 觸發「Invalid "
            "procedure call or argument」（執行期錯誤 5），因系統 ANSI 字碼頁無法編碼漢字。"
            "錯誤以彈出視窗形式出現並中止匯出。使用關聯代碼 c_assoc_code = 437"
            "（'Presented literary composition as gift to' / '贈詩、文'）的固件能可靠觸發此問題，"
            "因相關人物的 c_name 包含漢字。\n\n"
            "由 test_bug22_associations_cmducinet_fires_invalid_procedure_call 及 "
            "test_bug22_kinship_cmducinet_sibling_form_fires_invalid_procedure_call 偵測到。"
        ),
        "steps_en": [
            "Open CBDB_BJ_User.mdb in Microsoft Access.",
            "Open the LookAtAssociations form.",
            "Pick association code c_assoc_code = 437 "
            "('Presented literary composition as gift to').",
            "Click CmdUCINet.  A popup appears: "
            "'Invalid procedure call or argument'.  The UCINet export file is not created.",
            "The same error occurs in LookAtKinship.CmdUCINet when the kinship network "
            "contains a person whose c_name includes CJK Han characters (e.g. '取' / 贈詩、文).",
        ],
        "steps_zh": [
            "以 Microsoft Access 開啟 CBDB_BJ_User.mdb。",
            "開啟 LookAtAssociations 表單。",
            "選取關聯代碼 c_assoc_code = 437（'Presented literary composition as gift to'）。",
            "點擊 CmdUCINet。出現彈出視窗：'Invalid procedure call or argument'。UCINet 匯出檔案未建立。",
            "在 LookAtKinship.CmdUCINet 中，當親屬網絡包含 c_name 含 CJK 漢字的人物時，同樣會出現此錯誤。",
        ],
        "screenshots": [],
        "severity_en": (
            "P1 — Visible runtime crash: a popup aborts the export.  Any UCINet workflow on "
            "an association network that includes persons with Han-character names will fail.  "
            "Most CBDB persons have CJK names, making this effectively a blanket failure for "
            "real-world LookAtAssociations / LookAtKinship → UCINet usage."
        ),
        "severity_zh": (
            "P1 — 可見的執行期崩潰：彈出視窗中止匯出。任何包含漢字姓名人物的關聯網絡 UCINet "
            "工作流程均會失敗。絕大多數 CBDB 人物具有漢字姓名，使其在實際的 LookAtAssociations / "
            "LookAtKinship → UCINet 使用中幾乎全面失敗。"
        ),
        "fix_en": (
            "Change CreateTextFile calls in Form_LookAtAssociations.CmdUCINet_Click and "
            "Form_LookAtKinship.CmdUCINet_Click to the 3-argument form: "
            "CreateTextFile(filename, True, True) — the third argument enables Unicode output.  "
            "Test with a fixture that includes a person whose c_name contains CJK Han characters."
        ),
        "fix_zh": (
            "將 Form_LookAtAssociations.CmdUCINet_Click 和 Form_LookAtKinship.CmdUCINet_Click "
            "中的 CreateTextFile 呼叫改為 3 參數形式：CreateTextFile(filename, True, True)——"
            "第三個參數啟用 Unicode 輸出。以 c_name 含 CJK 漢字的人物固件進行測試。"
        ),
    },
    {
        "id": 6,
        "tier": "P1_visible_crash",
        "form": "Form_LookAtGroupData.queryEntry",
        "title_en": "LookAtGroupData.queryEntry crashes with 'No such field' — ENTRY_DATA.c_parental_status missing _code suffix",
        "title_zh": "LookAtGroupData.queryEntry 崩潰「找不到欄位」——ENTRY_DATA.c_parental_status 缺少 _code 後綴",
        "summary_en": (
            "Form_LookAtGroupData.vb line ~2621 has an INSERT INTO whose target column "
            "list ends with c_parental_status_code but whose SELECT projection ends with "
            "ENTRY_DATA.c_parental_status (no _code suffix).  The actual column on "
            "ENTRY_DATA is c_parental_status_code, so the SQL crashes with JET error 3061 "
            "'No value given for one or more required parameters' the moment the user "
            "ticks the Entry checkbox and clicks Run.  ZZ_SCRATCH_ENTRY remains at 0 rows.  "
            "The identical query in Form_LookAtEntry.vb uses the correct name; "
            "this is a single-character drift between the two forms.\n\n"
            "Detected by: test_bug6_lookat_groupdata_query_entry_fires_no_such_field — "
            "asserts LookAtGroupData:ERR with JET 3061 signature and "
            "ZZ_SCRATCH_ENTRY stays at 0.  Also: test_bug6_groupdata_query_entry_wrong_field "
            "— static source-string assertion."
        ),
        "summary_zh": (
            "Form_LookAtGroupData.vb 第 ~2621 行的 INSERT INTO，其目標欄位列表末尾為 "
            "c_parental_status_code，但 SELECT 投影末尾卻是 ENTRY_DATA.c_parental_status"
            "（缺少 _code 後綴）。ENTRY_DATA 上的實際欄位為 c_parental_status_code，因此 "
            "SQL 在使用者勾選 Entry 核取方塊並點擊 Run 時立即崩潰，觸發 JET 錯誤 3061。"
            "ZZ_SCRATCH_ENTRY 保持 0 列。Form_LookAtEntry.vb 中完全相同的查詢使用了正確名稱；"
            "這是兩個表單之間的單字元漂移。\n\n"
            "由 test_bug6_lookat_groupdata_query_entry_fires_no_such_field 及靜態原始碼斷言偵測到。"
        ),
        "steps_en": [
            "Open CBDB_BJ_User.mdb in Microsoft Access.",
            "Open the form **LookAtGroupData**.",
            "In the import person list, enter person ID **1** (An Dun 安惇; has 2 ENTRY_DATA rows).",
            "Tick only the **Entry** checkbox; leave Status / Office / Text / Addr unchecked.",
            "Click **Run**.",
            "A popup appears reporting JET error 3061 ('No value given for one or more "
            "required parameters' or 'No such field').  ZZ_SCRATCH_ENTRY stays empty.",
        ],
        "steps_zh": [
            "以 Microsoft Access 開啟 CBDB_BJ_User.mdb。",
            "開啟 **LookAtGroupData** 表單。",
            "在匯入人員清單中，輸入 person ID **1**（安惇 An Dun；有 2 筆 ENTRY_DATA）。",
            "僅勾選 **Entry** 核取方塊；其餘 Status / Office / Text / Addr 保持未勾選。",
            "點擊 **Run**。",
            "彈出 JET 錯誤 3061 彈窗（「未提供一個或多個必要參數的值」或「找不到欄位」）。"
            "ZZ_SCRATCH_ENTRY 保持空白。",
        ],
        "screenshots": [],
        "severity_en": (
            "P1 — Visible crash on a common path: any user who ticks the Entry checkbox "
            "in LookAtGroupData will hit this error.  ZZ_SCRATCH_ENTRY stays at 0 rows, "
            "so no Entry data is available for downstream export steps (GIS, Neo4j, etc.)."
        ),
        "severity_zh": (
            "P1 — 常見操作路徑上的明顯崩潰：任何在 LookAtGroupData 中勾選 Entry 核取方塊的"
            "使用者都會遇到此錯誤。ZZ_SCRATCH_ENTRY 保持 0 列，後續匯出步驟（GIS、Neo4j 等）"
            "無法取得任何 Entry 資料。"
        ),
        "fix_en": (
            "Change ENTRY_DATA.c_parental_status to ENTRY_DATA.c_parental_status_code "
            "on line ~2621 of Form_LookAtGroupData.vb.  One-character fix; "
            "the corrected form already appears in Form_LookAtEntry.vb."
        ),
        "fix_zh": (
            "將 Form_LookAtGroupData.vb 第 ~2621 行的 ENTRY_DATA.c_parental_status "
            "改為 ENTRY_DATA.c_parental_status_code。一字元修復；"
            "Form_LookAtEntry.vb 中已有正確寫法可參照。"
        ),
    },
    {
        "id": 7,
        "tier": "P1_visible_crash",
        "form": "Form_LookAtPlace.CmdNeo4j_Click",
        "title_en": "LookAtPlace.CmdNeo4j: People-CSV loop reads c_dynasty / c_dynasty_chn / c_female not in SELECT — crashes JET 3265 'Item not found'",
        "title_zh": "LookAtPlace.CmdNeo4j：People-CSV 迴圈讀取未在 SELECT 中投影的 c_dynasty / c_dynasty_chn / c_female——觸發 JET 3265「找不到項目」",
        "summary_en": (
            "The People-CSV section of Form_LookAtPlace.CmdNeo4j_Click opens a recordset "
            "via a SELECT that projects only four ZZ_SCRATCH_P_TEXT columns, but the "
            "row-write loop reads !c_dynasty, !c_dynasty_chn, and !c_female from that "
            "recordset.  DAO's Recordset.Fields collection contains only SELECT-projected "
            "columns; the JOIN brings DYNASTIES and BIOG_MAIN into scope for filtering "
            "but does not expose their fields.  JET raises 3265 'Item not found in this "
            "collection' on the first !c_dynasty read; the error trap exits the sub "
            "before any file is written, so the user sees a popup AND the export "
            "produces 0 CSV files.\n\n"
            "Detected by: test_bug7_lookat_place_cmdneo4j_fires_item_not_found — "
            "asserts 'Item not found' in ZZ_TEST_DEBUG :ERR markers.  "
            "Also: test_bug7_lookat_place_cmdneo4j_select_missing_dynasty_female — "
            "static SELECT projection assertion."
        ),
        "summary_zh": (
            "Form_LookAtPlace.CmdNeo4j_Click 的 People-CSV 區段透過僅投影四個 "
            "ZZ_SCRATCH_P_TEXT 欄位的 SELECT 開啟記錄集，但列寫入迴圈從該記錄集讀取 "
            "!c_dynasty、!c_dynasty_chn 與 !c_female。DAO 的 Recordset.Fields 集合"
            "僅包含 SELECT 投影的欄位；JOIN 只將 DYNASTIES 和 BIOG_MAIN 帶入範圍作為篩選，"
            "並不暴露其欄位。JET 在第一次 !c_dynasty 讀取時觸發 3265「找不到項目」，"
            "錯誤處理器在任何檔案寫入前退出 Sub，使用者看到彈窗且匯出產生 0 個 CSV 檔案。\n\n"
            "由 test_bug7_lookat_place_cmdneo4j_fires_item_not_found 及靜態原始碼斷言偵測到。"
        ),
        "steps_en": [
            "Open CBDB_BJ_User.mdb in Microsoft Access.",
            "Open the form **LookAtPlace**.",
            "Use the address picker to select a substantive address (e.g. c_addr_id = 100658, "
            "Kaifeng 開封).  Click **Run Query**.",
            "Click the **Neo4j** export button and choose a save location.",
            "A Run-time error 3265 — 'Item not found in this collection' popup appears.  "
            "The chosen folder contains no Neo4j export files.",
        ],
        "steps_zh": [
            "以 Microsoft Access 開啟 CBDB_BJ_User.mdb。",
            "開啟 **LookAtPlace** 表單。",
            "使用地址選擇器選取一個有充足資料的地址（例如 c_addr_id = 100658，開封）。"
            "點擊 **Run Query**。",
            "點擊 **Neo4j** 匯出按鈕並選擇儲存位置。",
            "彈出執行時期錯誤 3265「找不到項目」。"
            "所選資料夾中沒有任何 Neo4j 匯出檔案。",
        ],
        "screenshots": [],
        "severity_en": (
            "P1 — Visible crash on a normal user click.  Any LookAtPlace → Neo4j export "
            "with a non-empty place-people result hits this deterministically.  "
            "0 CSV files are produced despite the SaveAs dialog having already fired."
        ),
        "severity_zh": (
            "P1 — 正常使用者操作中的明顯崩潰。任何非空的 LookAtPlace → Neo4j 匯出都會"
            "確定性地觸發此問題。儘管 SaveAs 對話框已觸發，仍產生 0 個 CSV 檔案。"
        ),
        "fix_en": (
            "Extend the SELECT in Form_LookAtPlace.CmdNeo4j_Click (lines ~643-647) to "
            "also project DYNASTIES.c_dynasty, DYNASTIES.c_dynasty_chn, and BIOG_MAIN.c_female.  "
            "The FROM / JOIN structure already brings those source tables into scope; "
            "three column references added to the SELECT projection is the complete fix."
        ),
        "fix_zh": (
            "將 Form_LookAtPlace.CmdNeo4j_Click（第 ~643-647 行）的 SELECT 子句擴展，"
            "加入 DYNASTIES.c_dynasty、DYNASTIES.c_dynasty_chn 及 BIOG_MAIN.c_female 的投影。"
            "FROM / JOIN 結構已將這些來源表帶入範圍；在 SELECT 中新增三個欄位引用即為完整修復。"
        ),
    },
    {
        "id": 25,
        "tier": "P1_visible_crash",
        "form": "Form_LookAtKinship.CmdImport_Click / Form_LookAtGroupData.CmdImport_Click / Form_LookAtAssociationPairs.CmdImportList_Click",
        "title_en": "LookAtKinship / LookAtGroupData / LookAtAssociationPairs: CmdImport round-trip fails — ZZ_SCRATCH_IMPORT_PEOPLE stays empty",
        "title_zh": "LookAtKinship / LookAtGroupData / LookAtAssociationPairs：CmdImport 往返失敗——ZZ_SCRATCH_IMPORT_PEOPLE 保持空白",
        "summary_en": (
            "After seeding person IDs [1, 2, 3] and clicking CmdImport (or CmdImportList for "
            "LookAtAssociationPairs), the handler is expected to populate "
            "ZZ_SCRATCH_IMPORT_PEOPLE with the seeded IDs.  In all three forms the target "
            "table remains empty (c_person_id = []) after the import completes.  No error "
            "popup is shown — the import appears to succeed silently but writes nothing.\n\n"
            "Detected by: test_cmd_import_round_trip[LookAtKinship.CmdImport], "
            "test_cmd_import_round_trip[LookAtGroupData.CmdImport], and "
            "test_cmd_import_round_trip[LookAtAssociationPairs.CmdImportList] — all assert "
            "ZZ_SCRATCH_IMPORT_PEOPLE.c_person_id = []; expected [1, 2, 3]."
        ),
        "summary_zh": (
            "在填入人物 ID [1, 2, 3] 並點擊 CmdImport（或 LookAtAssociationPairs 的 CmdImportList）"
            "後，handler 應將填入的 ID 寫入 ZZ_SCRATCH_IMPORT_PEOPLE。在全部三個表單中，"
            "匯入完成後目標表格仍為空（c_person_id = []）。不顯示錯誤彈出視窗——匯入看似成功，"
            "但未寫入任何資料。\n\n"
            "由 test_cmd_import_round_trip[LookAtKinship.CmdImport]、"
            "[LookAtGroupData.CmdImport]、[LookAtAssociationPairs.CmdImportList] 偵測到，"
            "均斷言 ZZ_SCRATCH_IMPORT_PEOPLE.c_person_id = []；預期 [1, 2, 3]。"
        ),
        "steps_en": [
            "Open CBDB_BJ_User.mdb in Microsoft Access.",
            "Open LookAtKinship (or LookAtGroupData / LookAtAssociationPairs).",
            "Enter person IDs 1, 2, 3 in the import field.",
            "Click CmdImport.  No error popup appears.",
            "Query ZZ_SCRATCH_IMPORT_PEOPLE: SELECT c_person_id FROM "
            "ZZ_SCRATCH_IMPORT_PEOPLE — the table is empty.",
        ],
        "steps_zh": [
            "以 Microsoft Access 開啟 CBDB_BJ_User.mdb。",
            "開啟 LookAtKinship（或 LookAtGroupData / LookAtAssociationPairs）。",
            "在匯入欄位中輸入人物 ID 1, 2, 3。",
            "點擊 CmdImport。不出現錯誤彈出視窗。",
            "查詢 ZZ_SCRATCH_IMPORT_PEOPLE：SELECT c_person_id FROM ZZ_SCRATCH_IMPORT_PEOPLE——資料表為空。",
        ],
        "screenshots": [],
        "severity_en": (
            "P1 — Silent import failure: the import appears to complete successfully but the "
            "target table is empty.  Any subsequent query or export that depends on the imported "
            "person list will operate on an empty dataset without warning."
        ),
        "severity_zh": (
            "P1 — 靜默匯入失敗：匯入看似成功完成，但目標資料表為空。依賴匯入人物列表的後續查詢"
            "或匯出將在沒有警告的情況下操作空資料集。"
        ),
        "fix_en": (
            "Inspect the CmdImport_Click / CmdImportList_Click handlers in each affected form.  "
            "Verify the INSERT INTO ZZ_SCRATCH_IMPORT_PEOPLE statement: check that the "
            "source control (text box or list box) is correctly read and that the INSERT "
            "executes within an active transaction that is committed."
        ),
        "fix_zh": (
            "檢查各受影響表單的 CmdImport_Click / CmdImportList_Click handler。驗證 INSERT INTO "
            "ZZ_SCRATCH_IMPORT_PEOPLE 語句：確認來源控制項（文字方塊或列表方塊）被正確讀取，"
            "且 INSERT 在已提交的活躍交易中執行。"
        ),
    },
    {
        "id": 2,
        "tier": "P2_silent_display",
        "form": "Form_LookAtGroupData.CmdRun_Click",
        "title_en": "LookAtGroupData: CmdRun does not backfill c_name from BIOG_MAIN",
        "title_zh": "LookAtGroupData：CmdRun 未從 BIOG_MAIN 回填 c_name",
        "summary_en": (
            "When the user seeds a person ID into LookAtGroupData and clicks "
            "CmdRun, the handler is expected to run an UPDATE query that joins "
            "ZZ_SCRATCH_IMPORT_PEOPLE to BIOG_MAIN and fills in c_name (and "
            "c_dynasty) for each seeded row.  In this build the UPDATE does "
            "not execute successfully: after CmdRun completes, c_name remains "
            "NULL in ZZ_SCRATCH_IMPORT_PEOPLE.\n\n"
            "The result is that the group-data import display shows empty name "
            "cells.  The user has no indication that the backfill failed — "
            "CmdRun does not surface an error.\n\n"
            "Detected by: test_hard_form_query_small_fixture[groupdata_person_1_small] — "
            "assertion 'CmdRun didn't backfill c_name for c_person_id=1', "
            "c_name is None after CmdRun completes."
        ),
        "summary_zh": (
            "當使用者在 LookAtGroupData 中填入一個 person ID 並點擊 CmdRun 時，"
            "handler 應執行 UPDATE 查詢，將 ZZ_SCRATCH_IMPORT_PEOPLE JOIN BIOG_MAIN，"
            "並為每一筆填入 c_name（及 c_dynasty）。在此版本中，UPDATE 未成功執行："
            "CmdRun 完成後，ZZ_SCRATCH_IMPORT_PEOPLE 的 c_name 仍為 NULL。\n\n"
            "結果是群組資料匯入畫面顯示空白的姓名欄位，且使用者不會看到任何錯誤訊息，"
            "CmdRun 靜默地失敗了。\n\n"
            "由 test_hard_form_query_small_fixture[groupdata_person_1_small] 偵測到："
            "斷言 'CmdRun didn't backfill c_name for c_person_id=1'，CmdRun "
            "完成後 c_name 仍為 None。"
        ),
        "steps_en": [
            "Open CBDB_BJ_User.mdb in Microsoft Access.",
            "From the Navigation Pane, open the form **LookAtGroupData**.",
            "In the import person list, enter a valid person ID (e.g. **1**).",
            "Click **Run** (CmdRun button).",
            "When CmdRun completes, inspect the result: the Name column is blank.",
            "SQL verification: `SELECT c_person_id, c_name FROM "
            "ZZ_SCRATCH_IMPORT_PEOPLE` returns (1, NULL) — "
            "c_name was not backfilled from BIOG_MAIN.",
        ],
        "steps_zh": [
            "以 Microsoft Access 開啟 CBDB_BJ_User.mdb。",
            "從導覽窗格開啟 **LookAtGroupData** 表單。",
            "在匯入人員清單中，輸入一個有效的 person ID（例如 **1**）。",
            "點擊 **Run**（CmdRun 按鈕）。",
            "CmdRun 完成後，檢視結果：姓名欄為空白。",
            "SQL 驗證：`SELECT c_person_id, c_name FROM ZZ_SCRATCH_IMPORT_PEOPLE` "
            "回傳 (1, NULL)——c_name 未從 BIOG_MAIN 回填。",
        ],
        "screenshots": [],
        "severity_en": (
            "P2 — Silent display issue: CmdRun completes without any error "
            "message, but the c_name column in the result is blank.  The user "
            "has no indication that the backfill failed."
        ),
        "severity_zh": (
            "P2 — 靜默顯示問題：CmdRun 完成時沒有任何錯誤訊息，但結果中的 "
            "c_name 欄位為空白。使用者無從得知回填已失敗。"
        ),
        "fix_en": (
            "Locate the UPDATE statement in Form_LookAtGroupData.CmdRun_Click "
            "that joins ZZ_SCRATCH_IMPORT_PEOPLE to BIOG_MAIN and sets c_name. "
            "Check that the JOIN condition matches the correct key column and "
            "that the UPDATE target column name is spelled correctly.  After "
            "the fix, running CmdRun with any valid person ID should populate "
            "c_name in ZZ_SCRATCH_IMPORT_PEOPLE."
        ),
        "fix_zh": (
            "在 Form_LookAtGroupData.CmdRun_Click 中找到將 ZZ_SCRATCH_IMPORT_PEOPLE "
            "JOIN BIOG_MAIN 並設定 c_name 的 UPDATE 語句，確認 JOIN 條件使用了正確的"
            "主鍵欄位，且 UPDATE 目標欄位名稱拼寫正確。修復後，以任意有效 person ID "
            "執行 CmdRun，c_name 應能在 ZZ_SCRATCH_IMPORT_PEOPLE 中被填入。"
        ),
    },
    {
        "id": 3,
        "tier": "P2_silent_display",
        "form": "Form_LookAtEntry.CmdQuery_Click",
        "title_en": "LookAtEntry: c_entry_desc backfill is NULL for all rows when entry_code = 36 (jinshi general)",
        "title_zh": "LookAtEntry：entry_code = 36（進士及第）時，c_entry_desc 回填全部為 NULL",
        "summary_en": (
            "When the user runs a LookAtEntry query filtered to entry code 36 "
            "(examination: jinshi general), the result table ZZ_SCRATCH_ENTRY "
            "is populated with 92,545 rows but the c_entry_desc column is NULL "
            "for every row.  The expected value is 'examination: jinshi (general)'.\n\n"
            "The CmdQuery_Click handler successfully inserts rows from ENTRY_DATA "
            "joined to ENTRY_CODES, but the c_entry_desc backfill step does not "
            "write the description for this specific entry code.  All other columns "
            "appear to be filled normally.  The missing description means the "
            "on-screen result grid shows a blank entry-type column for every record, "
            "which is misleading — the user sees results but cannot identify what "
            "type of examination each record represents.\n\n"
            "Detected by: test_vba_full_matrix[top_entry_code_36_unfiltered] — "
            "assertion 'c_entry_desc backfill wrong' with 92,545 affected rows.  "
            "Also: test_vba_full_matrix[entry_39_dy_20], [entry_36_dy_20], "
            "[entry_36_dy_15]."
        ),
        "summary_zh": (
            "當使用者在 LookAtEntry 以 entry_code = 36（進士及第）執行查詢時，"
            "結果表 ZZ_SCRATCH_ENTRY 雖然產生了 92,545 筆資料，但 c_entry_desc "
            "欄位對每一筆都是 NULL。預期值應為 'examination: jinshi (general)'。\n\n"
            "CmdQuery_Click 成功地從 ENTRY_DATA JOIN ENTRY_CODES 插入了資料，"
            "但 c_entry_desc 的回填步驟對此 entry code 並未寫入說明文字。其他欄位"
            "看起來都正常填充。因此，使用者在螢幕上看到的查詢結果中，每一筆記錄的"
            "入仕方式欄位都是空白，難以判斷是何種考試類型。\n\n"
            "由 test_vba_full_matrix[top_entry_code_36_unfiltered] 等偵測到，"
            "斷言 'c_entry_desc backfill wrong'，影響 92,545 筆。"
        ),
        "steps_en": [
            "Open CBDB_BJ_User.mdb in Microsoft Access.",
            "From the Navigation Pane, open the form **LookAtEntry**.",
            "In the Entry Code picker, select entry code **36** "
            "(label: 'examination: jinshi (general)').",
            "Leave dynasty, address, and year filters blank.",
            "Click **Run Query** (CmdQuery button).",
            "When the query completes, inspect the result grid: the "
            "entry-type description column (c_entry_desc) is blank for every row.",
            "SQL verification: `SELECT TOP 5 c_entry_code, c_entry_desc FROM "
            "ZZ_SCRATCH_ENTRY` returns (36, NULL) for all rows.",
        ],
        "steps_zh": [
            "以 Microsoft Access 開啟 CBDB_BJ_User.mdb。",
            "從導覽窗格開啟 **LookAtEntry** 表單。",
            "在 Entry Code 選擇器中，選取 entry code **36**"
            "（標籤：'examination: jinshi (general)'）。",
            "朝代、地址、年份篩選器留空。",
            "點擊 **Run Query**（CmdQuery 按鈕）。",
            "查詢完成後，檢視結果格：每一筆記錄的入仕方式說明欄（c_entry_desc）皆為空白。",
            "SQL 驗證：`SELECT TOP 5 c_entry_code, c_entry_desc FROM "
            "ZZ_SCRATCH_ENTRY` 對所有列回傳 (36, NULL)。",
        ],
        "screenshots": [],
        "severity_en": (
            "P2 — Silent display issue: 92,545 rows affected.  The user can "
            "see the blank c_entry_desc column in the result grid, but Access "
            "shows no error — making it easy to overlook.  Exports (GIS, "
            "Neo4j, KML) that reference this column will also carry the blank."
        ),
        "severity_zh": (
            "P2 — 靜默顯示問題：92,545 筆受影響。使用者可在結果格中看到空白的 "
            "c_entry_desc 欄，但 Access 不顯示錯誤——容易被忽略。參照此欄的匯出"
            "（GIS、Neo4j、KML）也會包含空白值。"
        ),
        "fix_en": (
            "Locate the backfill step in Form_LookAtEntry.CmdQuery_Click that "
            "sets c_entry_desc for ZZ_SCRATCH_ENTRY rows.  Verify that the JOIN "
            "to ENTRY_CODES on c_entry_code = 36 is not inadvertently filtered "
            "out or that the UPDATE / backfill SQL matches the column name "
            "exactly.  After the fix, "
            "`SELECT c_entry_desc FROM ZZ_SCRATCH_ENTRY WHERE c_entry_code = 36 "
            "LIMIT 1` should return 'examination: jinshi (general)'."
        ),
        "fix_zh": (
            "在 Form_LookAtEntry.CmdQuery_Click 中找到對 ZZ_SCRATCH_ENTRY 設定 "
            "c_entry_desc 的回填步驟，確認 JOIN ENTRY_CODES 的條件（c_entry_code = 36）"
            "沒有被意外篩除，且 UPDATE / 回填 SQL 使用了正確的欄位名稱。修復後，"
            "`SELECT c_entry_desc FROM ZZ_SCRATCH_ENTRY WHERE c_entry_code = 36 "
            "LIMIT 1` 應回傳 'examination: jinshi (general)'。"
        ),
    },
    {
        "id": 10,
        "tier": "P2_silent_display",
        "form": "EVENT_ADDR_2 Subform",
        "title_en": "EVENT_ADDR_2 Subform: TxtAddrCHN / TxtAddrPY bound to unaliased column names not in View_EventAddrData — render blank",
        "title_zh": "EVENT_ADDR_2 子表單：TxtAddrCHN / TxtAddrPY 繫結至 View_EventAddrData 中不存在的未別名欄位——顯示空白",
        "summary_en": (
            "The EVENT_ADDR_2 sub-form's TxtAddrCHN control has ControlSource c_name_chn "
            "and TxtAddrPY has ControlSource c_name, but the form's RecordSource is "
            "View_EventAddrData, which aliases ADDR_CODES.c_name_chn as c_event_addr_chn "
            "and ADDR_CODES.c_name as c_event_addr_name.  Neither c_name nor c_name_chn "
            "is in the projection, so both controls silently render blank for every row "
            "on the Events-with-Addresses sub-datasheet.  A SQL probe confirms: "
            "SELECT c_name_chn FROM View_EventAddrData raises 'Too few parameters. "
            "Expected 2.' — JET treats the unknown identifier as a parameter.\n\n"
            "Detected by: test_subform_control_source_unresolved[bug10_TxtAddrCHN] — "
            "opens EVENT_ADDR_2 Subform via COM and asserts TxtAddrCHN.ControlSource "
            "is still 'c_name_chn'."
        ),
        "summary_zh": (
            "EVENT_ADDR_2 子表單的 TxtAddrCHN 控制項 ControlSource 為 c_name_chn，"
            "TxtAddrPY 為 c_name，但表單的 RecordSource 是 View_EventAddrData，"
            "該查詢將 ADDR_CODES.c_name_chn 別名為 c_event_addr_chn，將 ADDR_CODES.c_name "
            "別名為 c_event_addr_name。由於投影中不包含 c_name 與 c_name_chn，"
            "兩個控制項在每一列的「事件含地址」子資料表中均靜默地顯示空白。SQL 探測確認："
            "SELECT c_name_chn FROM View_EventAddrData 拋出「參數太少，預期 2 個」——"
            "JET 將未知識別字視為參數。\n\n"
            "由 test_subform_control_source_unresolved[bug10_TxtAddrCHN] 偵測到。"
        ),
        "steps_en": [
            "Open CBDB_BJ_User.mdb in Microsoft Access.",
            "Open CBDB_Browser_2 and navigate to person **c_personid = 44872** (Sun Cai 孫才) "
            "— this person has an EVENT_ADDR row pointing to c_addr_id = 12603 (Anfeng 安豐).",
            "Switch to the **Events** sub-tab.",
            "Observe the EVENT_ADDR_2 sub-form nested inside the event row: "
            "TxtAddrCHN and TxtAddrPY render blank even though the parent row's address "
            "controls (bound to c_addr_chn / c_addr_name in View_EventsData) show '安豐'.",
        ],
        "steps_zh": [
            "以 Microsoft Access 開啟 CBDB_BJ_User.mdb。",
            "開啟 CBDB_Browser_2 並導航至 c_personid = 44872（孫才 Sun Cai）"
            "——此人有 EVENT_ADDR 列指向 c_addr_id = 12603（安豐）。",
            "切換至 **Events** 子分頁。",
            "觀察事件列中巢狀的 EVENT_ADDR_2 子表單：TxtAddrCHN 與 TxtAddrPY 顯示空白，"
            "即使父列的地址控制項（繫結至 View_EventsData 中的 c_addr_chn / c_addr_name）"
            "顯示「安豐」。",
        ],
        "screenshots": [],
        "severity_en": (
            "P2 — Silent display: both address controls on the EVENT_ADDR_2 sub-form render "
            "blank for every row, with no error popup.  Users see no indication that the "
            "address name is available; the parent row's address display is unaffected."
        ),
        "severity_zh": (
            "P2 — 靜默顯示：EVENT_ADDR_2 子表單的兩個地址控制項對每一列均顯示空白，"
            "且不彈出任何錯誤訊息。使用者無從得知地址名稱有資料；父列的地址顯示不受影響。"
        ),
        "fix_en": (
            "In the form designer for EVENT_ADDR_2 Subform, change TxtAddrCHN.ControlSource "
            "from c_name_chn to c_event_addr_chn, and TxtAddrPY.ControlSource from c_name "
            "to c_event_addr_name — the actual alias names in View_EventAddrData."
        ),
        "fix_zh": (
            "在 EVENT_ADDR_2 Subform 的表單設計師中，將 TxtAddrCHN.ControlSource 從 "
            "c_name_chn 改為 c_event_addr_chn，將 TxtAddrPY.ControlSource 從 c_name 改為 "
            "c_event_addr_name——即 View_EventAddrData 中的實際別名。"
        ),
    },
    {
        "id": 15,
        "tier": "P3_missing_ui",
        "form": "LookAtPlace",
        "title_en": "LookAtPlace is missing its CmdGIS button — handler exists but no UI control",
        "title_zh": "LookAtPlace 缺少 CmdGIS 按鈕——處理程式存在但無 UI 控制項",
        "summary_en": (
            "Form_LookAtPlace.vb defines a fully functional CmdGIS_Click handler — it "
            "builds and writes a GIS .tab export identical in shape to the GIS button on "
            "Status / Texts / Associations / Office / Kinship.  But LookAtPlace's form "
            "design has no CmdGIS button.  Users on Place can use Pajek / Gephi / Neo4j "
            "export but cannot use GIS export; the handler is there, just unreachable from "
            "the UI.  Note: if the button is added, Issue #4 (GISFrame vs CodeFrame "
            "typo in the same handler) must be fixed at the same time.\n\n"
            "Detected by: test_orphan_export_button_truly_missing[bug15_LookAtPlace_CmdGIS] "
            "— opens LookAtPlace via COM, calls Controls('CmdGIS'), and asserts the "
            "lookup raises 'Item not found'."
        ),
        "summary_zh": (
            "Form_LookAtPlace.vb 定義了功能完整的 CmdGIS_Click 處理程式——它建立並寫入"
            "與 Status / Texts / Associations / Office / Kinship 上 GIS 按鈕形狀相同的"
            "GIS .tab 匯出。但 LookAtPlace 的表單設計沒有 CmdGIS 按鈕。"
            "Place 的使用者可以使用 Pajek / Gephi / Neo4j 匯出，但無法使用 GIS 匯出；"
            "處理程式存在，只是無法從 UI 觸及。注意：若新增按鈕，Issue #4"
            "（同一處理程式中 GISFrame 與 CodeFrame 的錯字）必須同時修復。\n\n"
            "由 test_orphan_export_button_truly_missing[bug15_LookAtPlace_CmdGIS] 偵測到。"
        ),
        "steps_en": [
            "Open CBDB_BJ_User.mdb in Microsoft Access.",
            "Open the form **LookAtPlace**.",
            "Look at the export-buttons row at the bottom right.  There is no GIS button.",
            "Compare with LookAtStatus / LookAtAssociations / LookAtOffice etc., "
            "all of which have a GIS button.",
        ],
        "steps_zh": [
            "以 Microsoft Access 開啟 CBDB_BJ_User.mdb。",
            "開啟 **LookAtPlace** 表單。",
            "查看右下角的匯出按鈕列。沒有 GIS 按鈕。",
            "與 LookAtStatus / LookAtAssociations / LookAtOffice 等比較——"
            "這些表單均有 GIS 按鈕。",
        ],
        "screenshots": [],
        "severity_en": (
            "P3 — Missing UI: the GIS export feature is completely unavailable to LookAtPlace "
            "users even though the underlying handler is functional (with the Issue #4 fix applied)."
        ),
        "severity_zh": (
            "P3 — 缺少 UI：即使底層處理程式功能完整（套用 Issue #4 修復後），GIS 匯出功能"
            "對 LookAtPlace 使用者完全不可用。"
        ),
        "fix_en": (
            "In LookAtPlace's form design, add a CmdGIS button next to the existing "
            "CmdPajek / CmdGephi buttons with OnClick = [Event Procedure].  "
            "Also fix Issue #4 (GISFrame → CodeFrame typo) in the same patch."
        ),
        "fix_zh": (
            "在 LookAtPlace 的表單設計中，在現有 CmdPajek / CmdGephi 按鈕旁新增 CmdGIS "
            "按鈕，並設定 OnClick = [事件程序]。同時在同一補丁中修復 Issue #4 "
            "（GISFrame → CodeFrame 錯字）。"
        ),
    },
    {
        "id": 1,
        "tier": "P5_dormant_or_latent",
        "form": "View_StatusData",
        "title_en": "View_StatusData would display last-year range in the first-year column — DORMANT (no source rows trigger it on this dump)",
        "title_zh": "View_StatusData：c_fy_range_desc / c_fy_range_chn 引用錯誤的 YEAR_RANGE_CODES 別名——當前資料集靜止",
        "summary_en": (
            "The saved query View_StatusData joins YEAR_RANGE_CODES twice, aliasing "
            "the second copy as YEAR_RANGE_CODES_1 and joining it on STATUS_DATA.c_ly_range "
            "(the last-year range).  However the SELECT clause pulls c_fy_range_desc and "
            "c_fy_range_chn from YEAR_RANGE_CODES_1 — the wrong alias — so every "
            "status row would display the last-year range text in the first-year range "
            "column.  On the current dump no STATUS_DATA row has both c_fy_range and "
            "c_ly_range populated with different values, so the symptom is invisible "
            "in the UI today but will surface the moment future data introduces such a row.\n\n"
            "Detected by: test_bug_view_statusdata_fy_alias_swap — assertion "
            "'YEAR_RANGE_CODES_1.c_range AS c_fy_range_desc' still present in "
            "View_StatusData SQL.  Also: test_bug_view_statusdata_fy_value_equals_ly_value "
            "— asserts c_fy_range_desc == c_ly_range_desc for all rows with non-NULL "
            "range descriptions."
        ),
        "summary_zh": (
            "已儲存查詢 View_StatusData 將 YEAR_RANGE_CODES 聯結兩次，將第二個副本別名為 "
            "YEAR_RANGE_CODES_1 並以 STATUS_DATA.c_ly_range（末年範圍）聯結。但 SELECT "
            "子句從 YEAR_RANGE_CODES_1 中取出 c_fy_range_desc 與 c_fy_range_chn——使用了錯誤的"
            "別名——導致每筆狀態記錄在「首年範圍」欄顯示的實為末年範圍文字。目前資料集中，"
            "沒有任何 STATUS_DATA 列同時填入不同的 c_fy_range 與 c_ly_range，因此症狀在 UI "
            "上目前不可見，但一旦未來資料中出現此類列即會浮現。\n\n"
            "由 test_bug_view_statusdata_fy_alias_swap 及 "
            "test_bug_view_statusdata_fy_value_equals_ly_value 偵測到。"
        ),
        "steps_en": [
            "Open CBDB_BJ_User.mdb in Microsoft Access.",
            "Press F11 to show the Navigation Pane, then double-click query **View_StatusData**.",
            "Inspect the SELECT clause: both c_fy_range_desc and c_fy_range_chn reference "
            "YEAR_RANGE_CODES_1, but the FROM clause joins YEAR_RANGE_CODES_1 on "
            "STATUS_DATA.c_ly_range — not c_fy_range.",
            "(Dormant verification) Run: SELECT c_personid, c_fy_range_desc, c_ly_range_desc "
            "FROM View_StatusData WHERE c_fy_range > 0 OR c_ly_range > 0.  "
            "On the current dump the result is empty, confirming the bug is latent.",
        ],
        "steps_zh": [
            "以 Microsoft Access 開啟 CBDB_BJ_User.mdb。",
            "按 F11 顯示導覽窗格，然後雙擊查詢 **View_StatusData**。",
            "檢視 SELECT 子句：c_fy_range_desc 與 c_fy_range_chn 皆引用 YEAR_RANGE_CODES_1，"
            "但 FROM 子句是以 STATUS_DATA.c_ly_range（而非 c_fy_range）聯結該別名。",
            "（靜止驗證）執行：SELECT c_personid, c_fy_range_desc, c_ly_range_desc "
            "FROM View_StatusData WHERE c_fy_range > 0 OR c_ly_range > 0。"
            "在當前資料集中結果為空，確認此 Bug 目前為靜止狀態。",
        ],
        "screenshots": [],
        "severity_en": (
            "P5 — Dormant on this dump (would be P2 silent display if any STATUS_DATA row "
            "has both c_fy_range and c_ly_range set differently).  The SQL alias swap is a "
            "confirmed source-level defect; the symptom simply has no trigger row today."
        ),
        "severity_zh": (
            "P5 — 在當前資料集為靜止（若有任何 STATUS_DATA 列同時填入不同的 c_fy_range 與 "
            "c_ly_range，即提升為 P2 靜默顯示）。SQL 別名錯誤為確認的原始碼層級缺陷；"
            "當前只是缺少觸發列。"
        ),
        "fix_en": (
            "In View_StatusData, change YEAR_RANGE_CODES_1.c_range AS c_fy_range_desc "
            "and YEAR_RANGE_CODES_1.c_range_chn AS c_fy_range_chn to reference the "
            "un-aliased YEAR_RANGE_CODES copy (which the FROM clause already joins on "
            "STATUS_DATA.c_fy_range).  One-line fix per column."
        ),
        "fix_zh": (
            "在 View_StatusData 中，將 YEAR_RANGE_CODES_1.c_range AS c_fy_range_desc 及 "
            "YEAR_RANGE_CODES_1.c_range_chn AS c_fy_range_chn 改為引用未別名的 "
            "YEAR_RANGE_CODES（FROM 子句已以 STATUS_DATA.c_fy_range 聯結該副本）。"
            "每欄一行修復。"
        ),
    },
    {
        "id": 4,
        "tier": "P5_dormant_or_latent",
        "form": "Form_LookAtPlace.CmdGIS_Click",
        "title_en": "LookAtPlace.CmdGIS_Click references non-existent control GISFrame — latent, masked by missing button (Issue #15)",
        "title_zh": "LookAtPlace.CmdGIS_Click 引用不存在的控制項 GISFrame——潛伏，被缺少按鈕（Issue #15）遮蔽",
        "summary_en": (
            "Form_LookAtPlace.CmdGIS_Click reads GISFrame.Value on line ~1539, but "
            "LookAtPlace has no control named GISFrame — the actual encoding selector "
            "is named CodeFrame.  If the button were ever added (fixing Issue #15) "
            "without first correcting this line, every click would raise "
            "Run-time error 424 'Object required' and the GIS export would never run.  "
            "Today the bug is masked because no CmdGIS button exists on the form (Issue #15), "
            "so users cannot click it at all.\n\n"
            "Detected by: test_bug4_lookat_place_cmdgis_fires_object_required — disables the "
            "driver's GISFrame→CodeFrame patch and confirms the un-patched code raises "
            "'Object required'.  Also: test_bug4_lookat_place_cmdgis_references_nonexistent_gisframe "
            "— asserts GISFrame.Value is still present in Form_LookAtPlace.vb."
        ),
        "summary_zh": (
            "Form_LookAtPlace.CmdGIS_Click 在第 ~1539 行讀取 GISFrame.Value，但 LookAtPlace "
            "沒有名為 GISFrame 的控制項——實際的編碼選擇器名為 CodeFrame。若在未修正此行的情況下"
            "新增了按鈕（修復 Issue #15），每次點擊都會拋出執行時期錯誤 424「需要物件」，GIS "
            "匯出將無法執行。目前此 Bug 因表單上不存在 CmdGIS 按鈕（Issue #15）而被遮蔽。\n\n"
            "由 test_bug4_lookat_place_cmdgis_fires_object_required 及 "
            "test_bug4_lookat_place_cmdgis_references_nonexistent_gisframe 偵測到。"
        ),
        "steps_en": [
            "(Hypothetical, requires Issue #15 fixed first.) Open LookAtPlace.",
            "Run any query so the scratch table has data.",
            "Click the GIS button.",
            "A Run-time error 424 — Object required popup appears; the export produces no file.",
        ],
        "steps_zh": [
            "（假設情境，需先修復 Issue #15。）開啟 LookAtPlace。",
            "執行任意查詢使暫存表有資料。",
            "點擊 GIS 按鈕。",
            "彈出執行時期錯誤 424「需要物件」；匯出不產生任何檔案。",
        ],
        "screenshots": [],
        "severity_en": (
            "P5 — Latent (would be P1 visible crash if Issue #15 were fixed without "
            "also fixing this line).  The test driver works around this via a per-form "
            "GISFrame→CodeFrame substitution patch so integration tests pass; the "
            "underlying CBDB bug remains."
        ),
        "severity_zh": (
            "P5 — 潛伏（若 Issue #15 被修復但本行未一併修正，則提升為 P1 明顯崩潰）。"
            "測試驅動程式透過 GISFrame→CodeFrame 替換補丁使整合測試通過；CBDB 原始碼缺陷仍存在。"
        ),
        "fix_en": (
            "Change GISFrame.Value to CodeFrame.Value on line ~1539 of Form_LookAtPlace.vb.  "
            "The same form's CmdNeo4j_Click, CmdGephi_Click, and CmdPajek_Click already "
            "use CodeFrame correctly — this is a single-identifier drift.  Fix in the same "
            "patch as Issue #15 (adding the CmdGIS button)."
        ),
        "fix_zh": (
            "將 Form_LookAtPlace.vb 第 ~1539 行的 GISFrame.Value 改為 CodeFrame.Value。"
            "同一表單的 CmdNeo4j_Click、CmdGephi_Click、CmdPajek_Click 已正確使用 CodeFrame "
            "——這是單一識別字的漂移。應與 Issue #15（新增 CmdGIS 按鈕）在同一補丁中修復。"
        ),
    },
    {
        "id": 5,
        "tier": "P5_dormant_or_latent",
        "form": "Form_LookAtStatus.CmdPajek_Click",
        "title_en": "LookAtStatus.CmdPajek_Click references missing control ChkIDs and three non-existent columns — latent, masked by missing button (Issue #16)",
        "title_zh": "LookAtStatus.CmdPajek_Click 引用缺少的控制項 ChkIDs 及三個不存在的欄位——潛伏，被缺少按鈕（Issue #16）遮蔽",
        "summary_en": (
            "Form_LookAtStatus.CmdPajek_Click contains two related defects copied "
            "from LookAtAssociations without adapting names: (a) line ~2308 reads "
            "ChkIDs.Value, but LookAtStatus has no ChkIDs control; "
            "(b) the SELECT inside CmdPajek_Click references ZZ_SCRATCH_STATUS.c_person_id, "
            "c_status_id, and c_status_count — none of which exist on ZZ_SCRATCH_STATUS "
            "(the real columns are c_personid, c_status_code; there is no count column).  "
            "Both defects are moot today because LookAtStatus has no CmdPajek button "
            "(Issue #16), but adding the button without fixing these would expose both "
            "failures to users.\n\n"
            "Detected by: test_bug5_lookat_status_cmdpajek_sql_fires_field_error — "
            "pre-seeds ZZ_SCRATCH_STATUS and fires CmdPajek directly, asserting "
            "an ERR marker with an object-required or missing-field signature.  "
            "Also: test_bug5_lookat_status_cmdpajek_references_nonexistent_chkids — "
            "static source-string assertion."
        ),
        "summary_zh": (
            "Form_LookAtStatus.CmdPajek_Click 包含兩個從 LookAtAssociations 複製但未更新"
            "名稱的相關缺陷：(a) 第 ~2308 行讀取 ChkIDs.Value，但 LookAtStatus 沒有 "
            "ChkIDs 控制項；(b) CmdPajek_Click 內的 SELECT 引用 ZZ_SCRATCH_STATUS.c_person_id、"
            "c_status_id 及 c_status_count——這些欄位均不存在（實際欄位為 c_personid、"
            "c_status_code，且無計數欄）。由於 LookAtStatus 沒有 CmdPajek 按鈕（Issue #16），"
            "兩個缺陷目前均無法觸發，但新增按鈕時若未一併修復，將使兩個錯誤暴露給使用者。\n\n"
            "由 test_bug5_lookat_status_cmdpajek_sql_fires_field_error 及靜態原始碼斷言偵測到。"
        ),
        "steps_en": [
            "(Hypothetical, requires Issue #16 fixed first.) Open LookAtStatus.",
            "Run any query so ZZ_SCRATCH_STATUS has data.",
            "Click the Pajek button.",
            "First: a Run-time error 424 'Object required' popup appears (ChkIDs.Value).",
            "If worked around: a 'No such field' error fires from the SELECT referencing "
            "c_person_id / c_status_id / c_status_count.",
        ],
        "steps_zh": [
            "（假設情境，需先修復 Issue #16。）開啟 LookAtStatus。",
            "執行任意查詢使 ZZ_SCRATCH_STATUS 有資料。",
            "點擊 Pajek 按鈕。",
            "首先：彈出執行時期錯誤 424「需要物件」（ChkIDs.Value）。",
            "若繞過前者：因 SELECT 引用 c_person_id / c_status_id / c_status_count，"
            "觸發「找不到欄位」錯誤。",
        ],
        "screenshots": [],
        "severity_en": (
            "P5 — Latent (would be P1 visible crash if Issue #16 were fixed without "
            "also fixing these two defects).  The whole sub looks copy-pasted from "
            "LookAtAssociations.CmdPajek_Click without adapting column names to Status's schema."
        ),
        "severity_zh": (
            "P5 — 潛伏（若 Issue #16 被修復但本缺陷未一併修正，則提升為 P1 明顯崩潰）。"
            "整個 Sub 看起來是從 LookAtAssociations.CmdPajek_Click 複製而未更新欄位名稱。"
        ),
        "fix_en": (
            "Two fixes required: (a) replace ChkIDs.Value with False (or add a real "
            "ChkIDs control to LookAtStatus) and (b) rewrite the SELECT to use "
            "ZZ_SCRATCH_STATUS.c_personid and ZZ_SCRATCH_STATUS.c_status_code, dropping "
            "or computing c_status_count differently.  In practice the entire sub likely "
            "needs a thoughtful rewrite rather than spot fixes."
        ),
        "fix_zh": (
            "需要兩項修復：(a) 將 ChkIDs.Value 替換為 False（或在 LookAtStatus 新增真正的 "
            "ChkIDs 控制項）；(b) 將 SELECT 改寫為使用 ZZ_SCRATCH_STATUS.c_personid 及 "
            "ZZ_SCRATCH_STATUS.c_status_code，並視情況刪除或重新計算 c_status_count。"
            "實際上整個 Sub 可能需要徹底改寫而非局部修補。"
        ),
    },
    {
        "id": 9,
        "tier": "P5_dormant_or_latent",
        "form": "Form_LookAtEntry.CmdNeo4j_Click",
        "title_en": "LookAtEntry.CmdNeo4j Institutions block uses wrong recordset variable tRstAssocCodes — latent (no ENTRY_DATA row has c_inst_code > 0)",
        "title_zh": "LookAtEntry.CmdNeo4j Institutions 區塊使用錯誤的記錄集變數 tRstAssocCodes——潛伏（當前無 ENTRY_DATA 列有 c_inst_code > 0）",
        "summary_en": (
            "Form_LookAtEntry.vb line ~1415 opens an institutions recordset as "
            "Set tRstInstitutions = CurrentDb.OpenRecordset(tQueryStr), but line ~1425 "
            "says With tRstAssocCodes — referencing a recordset that was already Close'd "
            "in the AssocCodes block upstream.  If executed, .MoveFirst would raise DAO 3021 "
            "'No current record'.  The entire block sits inside If tRecDeleted > 0 Then at "
            "line ~1389, where tRecDeleted is the count of ENTRY_DATA rows with "
            "c_inst_code > 0.  On the current dump 0 of 263,454 ENTRY_DATA rows have "
            "c_inst_code > 0, so the gate evaluates false and the buggy block is unreachable "
            "from any LookAtEntry fixture today.  Test test_bug9_lookat_entry_cmdneo4j PASSED "
            "because the gate was shut, confirming the latent state.\n\n"
            "Representative fixtures: c_entry_code = 36 (examination: jinshi general / 進士) "
            "and c_entry_code = 101 (recommendation / 薦舉).  Both yield ENTRY_DATA rows "
            "with c_inst_code = 0, confirming the gate remains shut on this dump.\n\n"
            "Included per MANIFEST requirement (test_report_code_labels_audit_clean)."
        ),
        "summary_zh": (
            "Form_LookAtEntry.vb 第 ~1415 行以 Set tRstInstitutions = CurrentDb.OpenRecordset"
            "(tQueryStr) 開啟機構記錄集，但第 ~1425 行卻寫 With tRstAssocCodes——引用了已在"
            "上游 AssocCodes 區塊中 Close 的記錄集。若執行，.MoveFirst 將觸發 DAO 3021"
            "「沒有當前記錄」。整個區塊位於第 ~1389 行的 If tRecDeleted > 0 Then 之內，"
            "其中 tRecDeleted 為 c_inst_code > 0 的 ENTRY_DATA 列數。當前資料集中，"
            "263,454 筆 ENTRY_DATA 中有 0 筆 c_inst_code > 0，因此閘道評估為 false，"
            "test_bug9_lookat_entry_cmdneo4j PASSED（閘道關閉，確認潛伏狀態）。\n\n"
            "代表性固件：c_entry_code = 36（進士 jinshi general）和 c_entry_code = 101"
            "（薦舉 recommendation）。兩者的 ENTRY_DATA 列 c_inst_code = 0，確認閘道"
            "在當前資料集中保持關閉。\n\n"
            "依 MANIFEST 要求保留（test_report_code_labels_audit_clean）。"
        ),
        "steps_en": [
            "On the current dump this bug cannot be triggered through the UI — the "
            "If tRecDeleted > 0 Then gate at Form_LookAtEntry.vb:~1389 is false for "
            "every possible LookAtEntry fixture (0 of 263,454 ENTRY_DATA rows have "
            "c_inst_code > 0).",
            "Verify the source-level typo statically: open analysis/dump/vba/"
            "Form_LookAtEntry.vb and inspect lines ~1415-1425.  "
            "Line ~1415: Set tRstInstitutions = CurrentDb.OpenRecordset(tQueryStr).  "
            "Line ~1425: With tRstAssocCodes (intended: With tRstInstitutions).",
            "(Optional) Confirm the gate condition: SELECT COUNT(*) FROM ENTRY_DATA "
            "WHERE c_inst_code > 0 returns 0.",
            "Representative fixtures: c_entry_code = 36 (jinshi / 進士) and "
            "c_entry_code = 101 (recommendation / 薦舉).  Both confirm c_inst_code = 0.",
        ],
        "steps_zh": [
            "在當前資料集中，此 Bug 無法透過 UI 觸發——Form_LookAtEntry.vb 第 ~1389 行的 "
            "If tRecDeleted > 0 Then 對所有可能的 LookAtEntry 固件均評估為 false（263,454 "
            "筆 ENTRY_DATA 中有 0 筆 c_inst_code > 0）。",
            "靜態驗證原始碼層級錯字：開啟 analysis/dump/vba/Form_LookAtEntry.vb 並"
            "檢視第 ~1415-1425 行。第 ~1415 行：Set tRstInstitutions = CurrentDb.OpenRecordset"
            "(tQueryStr)。第 ~1425 行：With tRstAssocCodes（應為：With tRstInstitutions）。",
            "（可選）確認閘道條件：SELECT COUNT(*) FROM ENTRY_DATA WHERE c_inst_code > 0 "
            "返回 0。",
            "代表性固件：c_entry_code = 36（進士 jinshi general）和 c_entry_code = 101"
            "（薦舉 recommendation）。兩者均確認 c_inst_code = 0。",
        ],
        "screenshots": [],
        "severity_en": (
            "P5 — Latent source-level typo (would re-promote to P1 if any future ENTRY_DATA "
            "row has c_inst_code > 0).  Test test_bug9_lookat_entry_cmdneo4j PASSED this build "
            "(gate shut).  The missing InstitutionCodes CSV is not a user-visible bug today."
        ),
        "severity_zh": (
            "P5 — 潛伏的原始碼層級錯字（若未來任何 ENTRY_DATA 列有 c_inst_code > 0，"
            "則重新提升為 P1）。test_bug9_lookat_entry_cmdneo4j 在此版本中 PASSED（閘道關閉）。"
            "當前缺少 InstitutionCodes CSV 並非使用者可見的 Bug。"
        ),
        "fix_en": (
            "Change With tRstAssocCodes on line ~1425 to With tRstInstitutions.  "
            "Single-identifier fix; the correct variable was opened just a few lines above."
        ),
        "fix_zh": (
            "將第 ~1425 行的 With tRstAssocCodes 改為 With tRstInstitutions。"
            "單一識別字修復；正確的變數就在上方幾行剛被開啟。"
        ),
    },
    {
        "id": 11,
        "tier": "P5_dormant_or_latent",
        "form": "EVENTS_DATA_2 Subform",
        "title_en": "EVENTS_DATA_2 Subform: c_event_record_id control bound to non-existent column — hidden, so latent",
        "title_zh": "EVENTS_DATA_2 子表單：c_event_record_id 控制項繫結至不存在的欄位——已隱藏，因此為潛伏",
        "summary_en": (
            "The EVENTS_DATA_2 sub-form has a control named c_event_record_id whose "
            "ControlSource is also c_event_record_id.  Neither EVENTS_DATA nor "
            "View_EventsData projects a column of that name.  If the control were visible, "
            "it would render blank.  A live COM probe confirms the control is Visible=False "
            "with width 240 twips (~4 mm) — a hidden internal control, almost certainly "
            "a leftover join-key field never meant to be shown.  Real users see no blank "
            "column because the control is not displayed.\n\n"
            "Detected by: test_subform_control_source_unresolved[bug11_c_event_record_id] "
            "— opens EVENTS_DATA_2 Subform via COM and asserts the ControlSource is still "
            "c_event_record_id."
        ),
        "summary_zh": (
            "EVENTS_DATA_2 子表單有一個名為 c_event_record_id 的控制項，其 ControlSource "
            "同樣為 c_event_record_id。EVENTS_DATA 和 View_EventsData 均未投影該名稱的欄位。"
            "若控制項可見，將顯示空白。即時 COM 探測確認控制項 Visible=False，寬度 240 "
            "twips（約 4 mm）——一個隱藏的內部控制項，幾乎可確定是從未打算顯示的遺留聯結鍵欄位。"
            "實際使用者不會看到空白欄，因為控制項未顯示。\n\n"
            "由 test_subform_control_source_unresolved[bug11_c_event_record_id] 偵測到。"
        ),
        "steps_en": [
            "Verification is static + COM probe only — there is no UI symptom.",
            "Static evidence: SELECT c_event_record_id FROM View_EventsData raises "
            "'Too few parameters. Expected 1.' — confirming the column is absent from the projection.",
            "Visibility evidence: the COM probe confirms Visible=False and width=240 twips "
            "for the c_event_record_id control on EVENTS_DATA_2 Subform.",
        ],
        "steps_zh": [
            "驗證僅限靜態與 COM 探測——無 UI 症狀。",
            "靜態證據：SELECT c_event_record_id FROM View_EventsData 拋出「參數太少，"
            "預期 1 個」——確認該欄位不在投影中。",
            "可見性證據：COM 探測確認 EVENTS_DATA_2 Subform 上的 c_event_record_id "
            "控制項 Visible=False，寬度 240 twips。",
        ],
        "screenshots": [],
        "severity_en": (
            "P5 — Latent (would be P2 silent display if the control were ever made "
            "Visible=True or widened).  Code-hygiene only; no user-visible impact today."
        ),
        "severity_zh": (
            "P5 — 潛伏（若控制項被設為 Visible=True 或加寬，則提升為 P2 靜默顯示）。"
            "僅為程式碼衛生問題；目前對使用者無可見影響。"
        ),
        "fix_en": (
            "Either delete the hidden c_event_record_id control, or change its ControlSource "
            "to a real column (e.g. c_event_code) so it does not carry a stale binding.  "
            "Either change is invisible to users; this is code-hygiene only."
        ),
        "fix_zh": (
            "刪除隱藏的 c_event_record_id 控制項，或將其 ControlSource 改為真實欄位"
            "（如 c_event_code），使其不帶有過時的繫結。任何一種修改對使用者均不可見；"
            "僅為程式碼衛生。"
        ),
    },
    {
        "id": 12,
        "tier": "P5_dormant_or_latent",
        "form": "POSTED_TO_OFFICE_DATA_2 Subform",
        "title_en": "POSTED_TO_OFFICE_DATA_2 Subform: c_appt_type_code control bound to non-projected column — hidden, so latent",
        "title_zh": "POSTED_TO_OFFICE_DATA_2 子表單：c_appt_type_code 控制項繫結至未投影的欄位——已隱藏，因此為潛伏",
        "summary_en": (
            "The POSTED_TO_OFFICE_DATA_2 sub-form has a control c_appt_type_code with "
            "ControlSource c_appt_type_code, but View_PostingOfficeData projects c_appt_code "
            "(no _type infix) — not c_appt_type_code.  A live COM probe confirms the control "
            "is Visible=False, so the blank rendering is not user-visible today.  "
            "The user-facing appointment-type controls on the same form work correctly.  "
            "This is a code-hygiene issue only.\n\n"
            "Detected by: test_subform_control_source_unresolved[bug12_c_appt_type_code] "
            "— opens POSTED_TO_OFFICE_DATA_2 Subform via COM and asserts the ControlSource "
            "is still c_appt_type_code."
        ),
        "summary_zh": (
            "POSTED_TO_OFFICE_DATA_2 子表單有一個 c_appt_type_code 控制項，其 ControlSource "
            "為 c_appt_type_code，但 View_PostingOfficeData 投影的是 c_appt_code（無 _type "
            "中綴）——而非 c_appt_type_code。即時 COM 探測確認控制項 Visible=False，"
            "因此空白渲染目前對使用者不可見。同一表單上面向使用者的任職類型控制項工作正常。"
            "此為純粹的程式碼衛生問題。\n\n"
            "由 test_subform_control_source_unresolved[bug12_c_appt_type_code] 偵測到。"
        ),
        "steps_en": [
            "Verification is static + COM probe only — there is no UI symptom.",
            "Static evidence: in control_inventory.json, POSTED_TO_OFFICE_DATA_2 Subform "
            "has a control with control_source = 'c_appt_type_code', but View_PostingOfficeData "
            "projects c_appt_code.",
            "Visibility evidence: the COM probe confirms Visible=False for c_appt_type_code.",
        ],
        "steps_zh": [
            "驗證僅限靜態與 COM 探測——無 UI 症狀。",
            "靜態證據：在 control_inventory.json 中，POSTED_TO_OFFICE_DATA_2 Subform "
            "有一個 control_source = 'c_appt_type_code' 的控制項，但 View_PostingOfficeData "
            "投影的是 c_appt_code。",
            "可見性證據：COM 探測確認 c_appt_type_code 控制項 Visible=False。",
        ],
        "screenshots": [],
        "severity_en": (
            "P5 — Latent (would be P2 silent display if the control were made visible).  "
            "The user-facing appointment-type controls on the form work correctly.  "
            "Code-hygiene only."
        ),
        "severity_zh": (
            "P5 — 潛伏（若控制項被設為可見，則提升為 P2 靜默顯示）。"
            "表單上面向使用者的任職類型控制項工作正常。僅為程式碼衛生。"
        ),
        "fix_en": (
            "Either delete the hidden c_appt_type_code control, or change its ControlSource "
            "to c_appt_code (the actual column projected by View_PostingOfficeData)."
        ),
        "fix_zh": (
            "刪除隱藏的 c_appt_type_code 控制項，或將其 ControlSource 改為 c_appt_code"
            "（View_PostingOfficeData 實際投影的欄位）。"
        ),
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
COVERAGE_MATRIX_JSON = REPO / "reports" / "coverage_matrix.json"
SCHEMA_DIFF_JSON = REPO / "reports" / "schema_diff.json"

# Cell display labels — used by both docx and markdown renderers
_CELL_DISPLAY = {
    "PASS":    ("✓ PASS",   "✓"),
    "FAIL":    ("✗ FAIL",   "✗ FAIL"),
    "ERROR":   ("⚠ ERROR",  "⚠ ERR"),
    "SKIP":    ("~ SKIP",   "~ SKIP"),
    "NOT_RUN": ("(not run)", "—?"),
    "N/A":     ("—",        "—"),
}


def _load_coverage_matrix() -> dict | None:
    """Return parsed coverage_matrix.json, or None if not generated / unreadable."""
    import json as _json
    if not COVERAGE_MATRIX_JSON.exists():
        return None
    try:
        data = _json.loads(COVERAGE_MATRIX_JSON.read_text(encoding="utf-8"))
        # Basic schema check so a truncated/stale file doesn't crash rendering
        _ = data["forms"]; _ = data["buttons"]; _ = data["matrix"]
        return data
    except Exception:
        return None


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


def _add_coverage_matrix_docx(doc, is_en: bool, Z, data: dict | None) -> None:
    """Render the form × button coverage matrix as a Word table.

    Always present: if data is None (builder not run yet), shows a
    placeholder message directing the user to run step 5b.
    """
    title = (
        "Coverage Matrix — Form × Button Test Results"
        if is_en else
        "覆蓋矩陣 —— 表單 × 按鈕測試結果"
    )
    _h(doc, 1, Z(title))

    if data is None:
        placeholder = (
            "Coverage matrix not yet generated.  Run step 5b of the "
            "build-test workflow:\n"
            "    python analysis/build_coverage_matrix.py "
            "--report reports/pytest_report_<build>.json"
            if is_en else
            "覆蓋矩陣尚未生成。請執行 build-test 工作流程第 5b 步：\n"
            "    python analysis/build_coverage_matrix.py "
            "--report reports/pytest_report_<build>.json"
        )
        doc.add_paragraph(Z(placeholder))
        return

    forms = data["forms"]
    buttons = data["buttons"]
    matrix = data["matrix"]

    # Build table: header row + one row per form
    tbl = doc.add_table(rows=1 + len(forms), cols=1 + len(buttons))
    tbl.style = "Table Grid"

    # Header row
    hdr = tbl.rows[0].cells
    hdr[0].text = Z("Form" if is_en else "表單")
    for j, btn in enumerate(buttons):
        hdr[j + 1].text = btn

    # Data rows
    for i, form in enumerate(forms):
        row_cells = tbl.rows[i + 1].cells
        row_cells[0].text = Z(form)
        for j, btn in enumerate(buttons):
            outcome = matrix.get(form, {}).get(btn, "NOT_RUN")
            label = _CELL_DISPLAY.get(outcome, (outcome, outcome))[0]
            row_cells[j + 1].text = Z(label)

    doc.add_paragraph("")
    summary = data.get("summary", {})
    note = (
        f"PASS: {summary.get('PASS', 0)}  |  "
        f"FAIL: {summary.get('FAIL', 0)}  |  "
        f"ERROR: {summary.get('ERROR', 0)}  |  "
        f"SKIP: {summary.get('SKIP', 0)}  |  "
        f"NOT RUN: {summary.get('NOT_RUN', 0)}  |  "
        f"N/A: {summary.get('N/A', 0)}"
    )
    p = doc.add_paragraph(Z(note))
    p.runs[0].italic = True


def _add_coverage_matrix_md(lines: list, is_en: bool, Z, data: dict | None) -> None:
    """Append the coverage matrix as a GitHub-flavoured markdown table.

    Always appended: placeholder shown when builder has not been run.
    """
    title = (
        "Coverage Matrix — Form × Button Test Results"
        if is_en else
        "覆蓋矩陣 —— 表單 × 按鈕測試結果"
    )
    lines.append(f"## {Z(title)}")
    lines.append("")

    if data is None:
        lines.append(
            "> Coverage matrix not yet generated.  "
            "Run step 5b: `python analysis/build_coverage_matrix.py "
            "--report reports/pytest_report_<build>.json`"
            if is_en else
            "> 覆蓋矩陣尚未生成。請執行第 5b 步：`python "
            "analysis/build_coverage_matrix.py "
            "--report reports/pytest_report_<build>.json`"
        )
        lines.append("")
        return

    forms = data["forms"]
    buttons = data["buttons"]
    matrix = data["matrix"]

    # Header
    header = "| Form | " + " | ".join(buttons) + " |"
    sep = "| --- |" + " --- |" * len(buttons)
    lines.append(Z(header))
    lines.append(sep)

    for form in forms:
        cells = []
        for btn in buttons:
            outcome = matrix.get(form, {}).get(btn, "NOT_RUN")
            short = _CELL_DISPLAY.get(outcome, (outcome, outcome))[1]
            cells.append(Z(short))
        lines.append(f"| {Z(form)} | " + " | ".join(cells) + " |")

    lines.append("")
    summary = data.get("summary", {})
    note = (
        f"_PASS: {summary.get('PASS',0)}"
        f" · FAIL: {summary.get('FAIL',0)}"
        f" · ERROR: {summary.get('ERROR',0)}"
        f" · SKIP: {summary.get('SKIP',0)}"
        f" · NOT RUN: {summary.get('NOT_RUN',0)}"
        f" · N/A: {summary.get('N/A',0)}_"
    )
    lines.append(Z(note))
    lines.append("")


def _add_index_drift_appendix(doc, is_en: bool, Z) -> None:
    """Render the 'index_year / index_addr drift, NOT a bug' chapter.

    Always renders: shows a placeholder directing the user to run
    collect_index_year_diffs.py when the source JSON files are missing.
    """
    import json as _json

    title = (
        "Appendix A — `c_index_year` / `c_index_addr_id` drift "
        "vs the cbdb-online-main-server snapshot "
        "(differences need per-row classification before being filed as bugs)"
        if is_en else
        "附錄 A —— `c_index_year` / `c_index_addr_id` 與 "
        "cbdb-online-main-server 快照之間的偏差"
        "（差異需要逐筆分類後才能判定是否為缺陷）"
    )
    _h(doc, 1, Z(title))

    if not DRIFT_JSON.exists():
        placeholder = (
            "Appendix data not yet generated.  Run step 5c of the "
            "build-test workflow:\n"
            "    python reports/collect_index_year_diffs.py\n"
            "This queries the DATA mdb via pyodbc and emits "
            "reports/index_drift_*.json."
            if is_en else
            "附錄數據尚未生成。請執行 build-test 工作流程第 5c 步：\n"
            "    python reports/collect_index_year_diffs.py\n"
            "此腳本透過 pyodbc 查詢 DATA mdb，並輸出 "
            "reports/index_drift_*.json。"
        )
        doc.add_paragraph(Z(placeholder))
        return

    try:
        data = _json.loads(DRIFT_JSON.read_text(encoding="utf-8"))
    except Exception as e:
        doc.add_paragraph(Z(
            f"Appendix data could not be loaded ({e}). "
            "Re-run: python reports/collect_index_year_diffs.py"
            if is_en else
            f"附錄數據載入失敗（{e}）。請重新執行：python reports/collect_index_year_diffs.py"
        ))
        return

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


_VALID_TIERS = frozenset([
    "P0_silent_data", "P1_visible_crash", "P2_silent_display",
    "P3_missing_ui", "P4_setup", "P5_dormant_or_latent",
])
# NOTE: "resolved" is intentionally excluded.  Per AGENTS.md mandatory rule,
# ISSUES is rebuilt from scratch each build.  Bugs that no longer reproduce
# simply don't appear — they are not marked "resolved".  Using "resolved"
# would pass validation but silently vanish from all rendered output (it is
# not in tier_order), which is a silent-corruption trap.


def _validate_issues() -> None:
    """Fail loudly if any ISSUES entry has an unknown tier.

    A typo in 'tier' silently drops the issue from all rendered outputs
    (it won't appear in tier_order).  Run this once before building so
    the error surfaces immediately instead of producing a quietly
    incomplete report.
    """
    for it in ISSUES:
        if it.get("tier") not in _VALID_TIERS:
            raise ValueError(
                f"ISSUES id={it.get('id')}: unknown tier {it.get('tier')!r}. "
                f"Valid tiers: {sorted(_VALID_TIERS)}"
            )


def _add_schema_diff_appendix_docx(doc, is_en: bool, Z) -> None:
    """Render Appendix B (TablesFields) and Appendix C (ForeignKeys)
    schema-diff sections using the python-docx API.
    Consumes reports/schema_diff.json produced by collect_schema_diffs.py."""
    import json as _json

    def _table_section(title_a: str, title_b: str,
                       only_cur_hdr: str, only_reg_hdr: str,
                       mismatch_hdr: str,
                       block_key: str) -> None:
        _h(doc, 1, Z(title_a if is_en else title_b))

        if not SCHEMA_DIFF_JSON.exists():
            p = doc.add_paragraph()
            run = p.add_run(Z(
                "Run `python reports/collect_schema_diffs.py` to populate "
                "this section."
                if is_en else
                "請先執行 `python reports/collect_schema_diffs.py` 以生成本節內容。"
            ))
            run.italic = True
            return

        diff = _json.loads(SCHEMA_DIFF_JSON.read_text(encoding="utf-8"))
        mdb_label = diff.get("data_mdb", "CBDB_*_DATA.mdb")
        blk = diff[block_key]
        is_tf = (block_key == "tables_fields")

        intro = (
            (
                f"This section compares the contents of the `TablesFields` table "
                f"in `{mdb_label}` against the database schema "
                "reconstructed from Access DAO (TableDefs) by "
                "`reports/collect_schema_diffs.py`. Discrepancies indicate the "
                "documentation table may be out of date."
            ) if is_tf else (
                "This section covers the `ForeignKeys` table and the FK "
                "relationships it documents."
            )
        )
        intro_zh = (
            (
                f"本節將 `{mdb_label}` 中 `TablesFields` 表的內容與"
                "`reports/collect_schema_diffs.py` 透過 Access DAO（TableDefs）重建的資料"
                "庫結構進行比對。若存在差異，表示文檔表可能已過時。"
            ) if is_tf else (
                "本節涵蓋 `ForeignKeys` 表及其所記錄的外鍵關係。"
            )
        )
        doc.add_paragraph(Z(intro if is_en else intro_zh))

        if not is_tf and not blk.get("fk_introspection_available"):
            doc.add_paragraph(Z(
                f"The `ForeignKeys` table ({blk['total_current']} rows) documents FK "
                "relationships in the database. All referenced table/column pairs have "
                "been verified to exist in the current dump. A catalog-level diff "
                "(documented FK vs. all actual FK constraints) is not available for "
                "Access databases and is omitted here."
                if is_en else
                f"`ForeignKeys` 表共 {blk['total_current']} 筆，記錄了資料庫中的外鍵關係。"
                "我們已驗證所有參照的表名與欄位名均存在於當前 dump 中。"
                "由於 Access 資料庫不支援通過標準 API 枚舉外鍵約束，"
                "此處無法提供文件記載 FK 與實際 FK 約束的完整對比。"
            ))
            doc.add_paragraph(Z(
                f"Reconstructed FK list: reports/foreign_keys_regen.csv"
                if is_en else
                "重建結果：reports/foreign_keys_regen.csv"
            ))
            return

        doc_name = "TablesFields" if is_tf else "ForeignKeys"
        regen_src = (
            "Reconstructed from DB" if is_tf
            else "Reconstructed from DB (via Access.Application DAO)"
        )
        regen_src_zh = (
            "從資料庫重建" if is_tf
            else "從資料庫重建（透過 Access.Application DAO）"
        )
        doc.add_paragraph(Z(
            f"Total rows in {doc_name}: {blk['total_current']}. "
            f"{regen_src}: {blk['total_regen']}."
            if is_en else
            f"{doc_name} 共 {blk['total_current']} 筆。"
            f"{regen_src_zh}：{blk['total_regen']} 筆。"
        ))
        _regen_csv = (
            "tables_fields_regen.csv" if is_tf else "foreign_keys_regen.csv"
        )
        doc.add_paragraph(Z(
            f"Reconstructed schema: reports/{_regen_csv}"
            if is_en else
            f"重建結果：reports/{_regen_csv}"
        ))

        only_cur = blk["only_in_current"]
        only_reg = blk["only_in_regen"]
        mismatches = blk.get("mismatches", [])

        if only_cur:
            _h(doc, 2, Z(only_cur_hdr))
            n = len(only_cur)
            if is_tf:
                tbl = doc.add_table(rows=n + 1, cols=2)
                tbl.style = "Table Grid"
                hdr = tbl.rows[0].cells
                hdr[0].text = "AccessTblNm"; hdr[1].text = "AccessFldNm"
                for i, row in enumerate(only_cur):
                    r = tbl.rows[i + 1].cells
                    r[0].text = row["AccessTblNm"]
                    r[1].text = row["AccessFldNm"]
            else:
                tbl = doc.add_table(rows=n + 1, cols=4)
                tbl.style = "Table Grid"
                hdr = tbl.rows[0].cells
                for ci, h in enumerate(["AccessTblNm", "AccessFldNm", "ForeignKey", "ForeignKeyBaseField"]):
                    hdr[ci].text = h
                for i, row in enumerate(only_cur):
                    r = tbl.rows[i + 1].cells
                    r[0].text = row["AccessTblNm"]
                    r[1].text = row["AccessFldNm"]
                    r[2].text = row.get("ForeignKey") or ""
                    r[3].text = row.get("ForeignKeyBaseField") or ""

        if only_reg:
            _h(doc, 2, Z(only_reg_hdr))
            n = len(only_reg)
            if is_tf:
                tbl = doc.add_table(rows=n + 1, cols=4)
                tbl.style = "Table Grid"
                hdr = tbl.rows[0].cells
                for ci, h in enumerate(["AccessTblNm", "AccessFldNm", "DataFormat", "NULL_allowed"]):
                    hdr[ci].text = h
                for i, row in enumerate(only_reg):
                    r = tbl.rows[i + 1].cells
                    r[0].text = row["AccessTblNm"]
                    r[1].text = row["AccessFldNm"]
                    r[2].text = str(row.get("DataFormat") or "")
                    r[3].text = str(row.get("NULL_allowed") or "")
            else:
                tbl = doc.add_table(rows=n + 1, cols=4)
                tbl.style = "Table Grid"
                hdr = tbl.rows[0].cells
                for ci, h in enumerate(["AccessTblNm", "AccessFldNm", "ForeignKey", "ForeignKeyBaseField"]):
                    hdr[ci].text = h
                for i, row in enumerate(only_reg):
                    r = tbl.rows[i + 1].cells
                    r[0].text = row["AccessTblNm"]
                    r[1].text = row["AccessFldNm"]
                    r[2].text = row.get("ForeignKey") or ""
                    r[3].text = row.get("ForeignKeyBaseField") or ""

        if mismatches:
            _h(doc, 2, Z(mismatch_hdr))
            csv_file = (
                "reports/schema_diff_tables_fields_mismatches.csv" if is_tf
                else "reports/schema_diff_foreign_keys_mismatches.csv"
            )
            doc.add_paragraph(Z(
                f"Full list: `{csv_file}` ({len(mismatches)} rows)"
                if is_en else
                f"完整清單：`{csv_file}`（{len(mismatches)} 筆）"
            ))

        if not only_cur and not only_reg and not mismatches:
            doc.add_paragraph(Z(
                f"✓ No discrepancies found — {doc_name} is in sync with the actual schema."
                if is_en else
                f"✓ 未發現差異 —— {doc_name} 與實際資料庫結構一致。"
            ))

    # ---- Appendix B ----
    _table_section(
        title_a="Appendix B — TablesFields: documentation vs. actual structure",
        title_b="附錄 B —— TablesFields：文檔表與實際資料庫結構對比",
        only_cur_hdr=(
            "Rows in TablesFields not found in actual DB (stale)"
            if is_en else
            "TablesFields 中有但實際資料庫中不存在的欄位（過時）"
        ),
        only_reg_hdr=(
            "Columns in actual DB not documented in TablesFields"
            if is_en else
            "實際資料庫中有但 TablesFields 未記錄的欄位"
        ),
        mismatch_hdr="Attribute mismatches" if is_en else "屬性不一致",
        block_key="tables_fields",
    )

    doc.add_page_break()

    # ---- Appendix C ----
    _table_section(
        title_a="Appendix C — ForeignKeys: documentation vs. actual structure",
        title_b="附錄 C —— ForeignKeys：文檔表與實際資料庫結構對比",
        only_cur_hdr=(
            "Rows in ForeignKeys not found in actual DB (stale)"
            if is_en else
            "ForeignKeys 中有但實際資料庫中不存在的外鍵（過時）"
        ),
        only_reg_hdr=(
            "FK relationships in actual DB not documented in ForeignKeys"
            if is_en else
            "實際資料庫中有但 ForeignKeys 未記錄的外鍵關係"
        ),
        mismatch_hdr="Attribute mismatches" if is_en else "屬性不一致",
        block_key="foreign_keys",
    )


def _add_schema_diff_appendix_md(
    lines: list,
    block_key: str,
    is_en: bool,
    Z,
    _slug,
) -> None:
    """Render one schema-diff appendix section (Appendix B or C) into
    the markdown lines list.

    block_key: 'tables_fields' for Appendix B, 'foreign_keys' for Appendix C.
    """
    import json as _json

    is_tf = (block_key == "tables_fields")

    heading = (
        ("Appendix B — TablesFields: documentation vs. actual structure"
         if is_tf else
         "Appendix C — ForeignKeys: documentation vs. actual structure")
        if is_en else
        ("附錄 B —— TablesFields：文檔表與實際資料庫結構對比"
         if is_tf else
         "附錄 C —— ForeignKeys：文檔表與實際資料庫結構對比")
    )
    lines.append(f"## {Z(heading)}")
    lines.append("")

    if not SCHEMA_DIFF_JSON.exists():
        lines.append(Z(
            "*Run `python reports/collect_schema_diffs.py` to populate "
            "this section.*"
            if is_en else
            "*請先執行 `python reports/collect_schema_diffs.py` 以生成本節內容。*"
        ))
        lines.append("")
        return

    diff = _json.loads(SCHEMA_DIFF_JSON.read_text(encoding="utf-8"))
    mdb_label = diff.get("data_mdb", "CBDB_*_DATA.mdb")
    blk = diff[block_key]

    intro = (
        (
            f"This section compares the contents of the `TablesFields` table "
            f"in `{mdb_label}` against the database schema "
            "reconstructed from Access DAO (TableDefs) by "
            "`reports/collect_schema_diffs.py`. Discrepancies indicate the "
            "documentation table may be out of date."
        ) if is_tf else (
            "This section covers the `ForeignKeys` table and the FK "
            "relationships it documents."
        )
    )
    intro_zh = (
        (
            f"本節將 `{mdb_label}` 中 `TablesFields` 表的內容與 "
            "`reports/collect_schema_diffs.py` 透過 Access DAO（TableDefs）重建的資料"
            "庫結構進行比對。若存在差異，表示文檔表可能已過時。"
        ) if is_tf else (
            "本節涵蓋 `ForeignKeys` 表及其所記錄的外鍵關係。"
        )
    )
    lines.append(Z(intro if is_en else intro_zh))
    lines.append("")

    if not is_tf and not blk.get("fk_introspection_available"):
        lines.append(Z(
            f"The `ForeignKeys` table ({blk['total_current']} rows) documents FK "
            "relationships in the database. All referenced table/column pairs have "
            "been verified to exist in the current dump. A catalog-level diff is not "
            "available for Access databases and is omitted here."
            if is_en else
            f"`ForeignKeys` 表共 {blk['total_current']} 筆，記錄了資料庫中的外鍵關係。"
            "我們已驗證所有參照的表名與欄位名均存在於當前 dump 中。"
            "由於 Access 資料庫不支援通過標準 API 枚舉外鍵約束，"
            "此處無法提供文件記載 FK 與實際 FK 約束的完整對比。"
        ))
        lines.append("")
        lines.append(Z(
            "Reconstructed FK list: [foreign_keys_regen.csv](foreign_keys_regen.csv)"
            if is_en else
            "重建結果：[foreign_keys_regen.csv](foreign_keys_regen.csv)"
        ))
        lines.append("")
        return

    doc_name = "TablesFields" if is_tf else "ForeignKeys"
    regen_src = (
        "Reconstructed from DB" if is_tf
        else "Reconstructed from DB (via Access.Application DAO)"
    )
    regen_src_zh = (
        "從資料庫重建" if is_tf
        else "從資料庫重建（透過 Access.Application DAO）"
    )
    lines.append(Z(
        f"Total rows in {doc_name}: {blk['total_current']}. "
        f"{regen_src}: {blk['total_regen']}."
        if is_en else
        f"{doc_name} 共 {blk['total_current']} 筆。"
        f"{regen_src_zh}：{blk['total_regen']} 筆。"
    ))
    lines.append("")

    _regen_csv = "tables_fields_regen.csv" if is_tf else "foreign_keys_regen.csv"
    _regen_label = (
        ("Reconstructed schema" if is_tf else "Reconstructed FK list")
        if is_en else "重建結果"
    )
    lines.append(Z(f"{_regen_label}: [{_regen_csv}]({_regen_csv})"))
    lines.append("")

    only_cur = blk["only_in_current"]
    only_reg = blk["only_in_regen"]
    mismatches = blk.get("mismatches", [])

    if only_cur:
        stale_hdr = (
            f"Rows in {doc_name} not found in actual DB (stale)"
            if is_en else
            f"{doc_name} 中有但實際資料庫中不存在的記錄（過時）"
        )
        lines.append(f"### {Z(stale_hdr)}")
        lines.append("")
        if is_tf:
            lines.append("| AccessTblNm | AccessFldNm |")
            lines.append("|---|---|")
            for row in only_cur:
                lines.append(f"| {row['AccessTblNm']} | {row['AccessFldNm']} |")
        else:
            lines.append("| AccessTblNm | AccessFldNm | ForeignKey | ForeignKeyBaseField |")
            lines.append("|---|---|---|---|")
            for row in only_cur:
                lines.append(
                    f"| {row['AccessTblNm']} | {row['AccessFldNm']} "
                    f"| {row.get('ForeignKey', '')} | {row.get('ForeignKeyBaseField', '')} |"
                )
        lines.append("")

    if only_reg:
        undoc_hdr = (
            f"Columns in actual DB not documented in {doc_name}"
            if is_en else
            f"實際資料庫中有但 {doc_name} 未記錄的欄位"
        )
        lines.append(f"### {Z(undoc_hdr)}")
        lines.append("")
        if is_tf:
            lines.append("| AccessTblNm | AccessFldNm | DataFormat | NULL_allowed |")
            lines.append("|---|---|---|---|")
            for row in only_reg:
                lines.append(
                    f"| {row['AccessTblNm']} | {row['AccessFldNm']} "
                    f"| {row.get('DataFormat', '')} | {row.get('NULL_allowed', '')} |"
                )
        else:
            lines.append("| AccessTblNm | AccessFldNm | ForeignKey | ForeignKeyBaseField |")
            lines.append("|---|---|---|---|")
            for row in only_reg:
                lines.append(
                    f"| {row['AccessTblNm']} | {row['AccessFldNm']} "
                    f"| {row.get('ForeignKey', '')} | {row.get('ForeignKeyBaseField', '')} |"
                )
        lines.append("")

    if mismatches:
        mis_hdr = "Attribute mismatches" if is_en else "屬性不一致"
        lines.append(f"### {Z(mis_hdr)}")
        lines.append("")
        csv_file = (
            "reports/schema_diff_tables_fields_mismatches.csv" if is_tf
            else "reports/schema_diff_foreign_keys_mismatches.csv"
        )
        lines.append(Z(
            f"Full list: `{csv_file}` ({len(mismatches)} rows)"
            if is_en else
            f"完整清單：`{csv_file}`（{len(mismatches)} 筆）"
        ))
        lines.append("")

    if not only_cur and not only_reg and not mismatches:
        lines.append(Z(
            f"✓ No discrepancies found — {doc_name} is in sync with the actual schema."
            if is_en else
            f"✓ 未發現差異 —— {doc_name} 與實際資料庫結構一致。"
        ))
        lines.append("")


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

    # ---- Coverage matrix (always present) ----
    _add_coverage_matrix_docx(doc, is_en, Z, _load_coverage_matrix())
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
        "resolved": "Resolved — fixed in this build",
    }
    tier_titles_zh = {
        "P0_silent_data": "P0 — 静默数据错误",
        "P1_visible_crash": "P1 — 可见的运行时报错",
        "P2_silent_display": "P2 — 静默显示问题",
        "P3_missing_ui": "P3 — 缺失界面",
        "P4_setup": "P4 — 安装设置",
        "P5_dormant_or_latent": "P5 — 潛伏 / 不可達 / 當前無法復現",
        "resolved": "已解決 — 當前 build 已修復",
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

    # ---- Appendix A: index_year / index_addr drift (NOT a bug) ----
    doc.add_page_break()
    _add_index_drift_appendix(doc, is_en, Z)

    # ---- Appendix B & C: schema diff ----
    doc.add_page_break()
    _add_schema_diff_appendix_docx(doc, is_en, Z)

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

    # ---- Coverage matrix (always present) ----
    _add_coverage_matrix_md(lines, is_en, Z, _load_coverage_matrix())

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
        "resolved": "Resolved — fixed in this build",
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
    appendix_a_toc_title = (
        "Appendix A — c_index_year / c_index_addr_id drift "
        "vs the cbdb-online-main-server snapshot "
        "(differences need per-row classification before being filed as bugs)"
        if is_en else
        "附錄 A —— c_index_year / c_index_addr_id 與 "
        "cbdb-online-main-server 快照之間的偏差"
        "（差異需要逐筆分類後才能判定是否為缺陷）"
    )
    lines.append(
        f"- [{Z(appendix_a_toc_title)}](#{_slug(Z(appendix_a_toc_title))})"
    )
    schema_appendix_b_title = (
        "Appendix B — TablesFields: documentation vs. actual structure"
        if is_en else
        "附錄 B —— TablesFields：文檔表與實際資料庫結構對比"
    )
    schema_appendix_c_title = (
        "Appendix C — ForeignKeys: documentation vs. actual structure"
        if is_en else
        "附錄 C —— ForeignKeys：文檔表與實際資料庫結構對比"
    )
    lines.append(f"- [{Z(schema_appendix_b_title)}](#{_slug(Z(schema_appendix_b_title))})")
    lines.append(f"- [{Z(schema_appendix_c_title)}](#{_slug(Z(schema_appendix_c_title))})")
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

    # ---- Index drift appendix (always present) ----
    import json as _json
    appendix_a_title = (
        "Appendix A — c_index_year / c_index_addr_id drift "
        "vs the cbdb-online-main-server snapshot "
        "(differences need per-row classification before being filed as bugs)"
        if is_en else
        "附錄 A —— c_index_year / c_index_addr_id 與 "
        "cbdb-online-main-server 快照之間的偏差"
        "（差異需要逐筆分類後才能判定是否為缺陷）"
    )
    lines.append(f"## {Z(appendix_a_title)}")
    lines.append("")
    if not DRIFT_JSON.exists():
        lines.append(
            "> Appendix data not yet generated.  "
            "Run step 5c: `python reports/collect_index_year_diffs.py`"
            if is_en else
            "> 附錄數據尚未生成。請執行第 5c 步：`python reports/collect_index_year_diffs.py`"
        )
        lines.append("")
    if DRIFT_JSON.exists():
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

    # ---- Appendix B & C: schema diff ----
    _add_schema_diff_appendix_md(lines, "tables_fields", is_en, Z, _slug)
    _add_schema_diff_appendix_md(lines, "foreign_keys", is_en, Z, _slug)

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
    _validate_issues()  # fail fast on unknown tier before writing any files
    # Always emit all four formats together so the docx and md never
    # drift apart.
    _build("en", OUT_EN)
    _build("zh", OUT_ZH)
    _build_md("en", OUT_EN_MD)
    _build_md("zh", OUT_ZH_MD)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
