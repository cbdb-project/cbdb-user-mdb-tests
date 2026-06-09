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

import re
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

ISSUES: list[dict] = [
    # CLEAN SLATE — cleared per AGENTS.md "each build is a fresh assessment".
    # Rebuild from scratch ONLY after a real test run, and ONLY for findings
    # that pass the triage gate in this file (see _validate_issues / the
    # report-triage contract in docs/skills/issue-report-maintainer.md).
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


# ---------------------------------------------------------------------
# Report-triage contract (the "user-perceptible validation gate").
# Full prose: docs/skills/issue-report-maintainer.md § "Report-triage
# contract"; summary: AGENTS.md § Confirmed bugs.
#
# WHY THIS EXISTS.  Severity must reflect what a USER perceives, not
# which test went red.  A regression run surfaces three kinds of red
# that are NOT user-perceptible on their own and were historically
# mis-filed as P0/P1:
#   - a cross-check threshold tripping vs an external snapshot
#     (c_index_year / c_index_addr_id drift) — a maintenance-cadence or
#     classification question, not a UI bug;
#   - a structural metric parsed out of an export file (vertex-count
#     header vs row count, nodedef column count) — may be a real bug or
#     a test-measurement artifact;
#   - an injected harness marker firing (a ZZ_TEST_DEBUG ":ERR" written
#     under MsgBox suppression) that the real error handler swallows so
#     the user never sees it.
# These are LEADS, not confirmed user bugs.  P0/P1/P2 are reserved for
# findings a human can reproduce through the Access UI and observe a
# concrete symptom (a popup, blank-or-wrong on-screen data, or a file
# the user asked for that is missing or corrupt).  A lead may only be
# rated P0/P1/P2 after the symptom is confirmed in the real UI
# (evidence.ui_verified == True); otherwise it is P5 pending
# verification, or (for drift) belongs in Appendix A.
# ---------------------------------------------------------------------

_FINDING_CLASSES = frozenset([
    "user_facing_bug",    # a human reproduces it via the UI and sees a symptom
    "cross_check_drift",  # divergence vs an external snapshot (index year/addr)
    "structural_metric",  # derived by parsing an export file's structure
    "internal_marker",    # detected only via an injected harness marker/state
    "latent_code",        # real source defect with no current user-facing trigger
])

# Tiers that assert the symptom is user-perceptible on the current dump.
_USER_PERCEPTIBLE_TIERS = frozenset([
    "P0_silent_data", "P1_visible_crash", "P2_silent_display",
])


def _issue_violations(it: dict) -> list[str]:
    """Return a list of triage-contract violations for one ISSUES entry.

    Empty list == the entry satisfies the contract.  Collecting all
    violations (rather than raising on the first) lets _validate_issues
    report every problem in one pass.
    """
    out: list[str] = []
    tier = it.get("tier")
    if tier not in _VALID_TIERS:
        # A typo in 'tier' silently drops the issue from all rendered
        # outputs (it won't appear in tier_order) — a silent-corruption
        # trap.  Report and stop; later checks can't reason about tier.
        out.append(
            f"unknown tier {tier!r} (valid: {sorted(_VALID_TIERS)})"
        )
        return out

    ev = it.get("evidence")
    if not isinstance(ev, dict):
        out.append(
            "missing 'evidence' block.  Every issue needs "
            "evidence={'finding_class': ..., ...}; P0/P1/P2 additionally "
            "need vba_ref, fixture, user_symptom."
        )
        return out

    fclass = ev.get("finding_class")
    if fclass not in _FINDING_CLASSES:
        out.append(
            f"evidence.finding_class {fclass!r} invalid "
            f"(valid: {sorted(_FINDING_CLASSES)})"
        )
        return out

    # NOTE — the gate trusts the author-supplied `finding_class`; it
    # enforces structure + class->tier routing, but it cannot divine
    # whether a finding labelled `user_facing_bug` really is one.  That
    # judgement is the reviewer's (the issue-report skill, the AI
    # reviewer, codex).  A deliberate mis-label (drift filed as
    # user_facing_bug to dodge routing) is caught in review, not here.
    # Keep this boundary explicit rather than bolting on brittle
    # keyword heuristics that would false-positive and create a false
    # sense of completeness.

    # `ui_verified` is a human attestation that the symptom was seen in
    # the real Access UI.  Require literal True — a truthy string like
    # "pending" must NOT unlock a user-perceptible tier.
    ui_verified = ev.get("ui_verified") is True

    # --- finding_class -> tier routing ---
    if fclass == "user_facing_bug":
        # Reproducible-now by a user → P0–P4.  A defect that is real but
        # NOT user-triggerable on this dump is latent_code (→ P5), so a
        # user_facing_bug parked at P5 is a contradiction.
        if tier == "P5_dormant_or_latent":
            out.append(
                "user_facing_bug must be P0–P4.  A defect that is real but "
                "not user-triggerable on this dump is latent_code (→ P5), "
                "not user_facing_bug."
            )
    elif fclass == "cross_check_drift":
        # Drift belongs in Appendix A, not a standalone P-tier issue.
        # Only admissible at all as P5 AND only when a per-row
        # classifier has judged it a real algorithm bug (not a
        # maintenance-cadence / source-snapshot diff).
        if tier != "P5_dormant_or_latent" or not str(
            ev.get("classification_ref") or ""
        ).strip():
            out.append(
                "cross_check_drift findings belong in Appendix A, not a "
                "P-tier issue.  Admissible only as P5_dormant_or_latent "
                "WITH evidence.classification_ref pointing to a "
                "classify_*_drift output that judged it a real algorithm "
                "bug (not maintenance-cadence / source-snapshot drift)."
            )
    elif fclass == "latent_code":
        if tier != "P5_dormant_or_latent":
            out.append(
                f"latent_code (real defect, no user-facing trigger today) "
                f"must be tier P5_dormant_or_latent, not {tier}."
            )
    elif fclass in ("structural_metric", "internal_marker"):
        # Not user-perceptible on its own.  Until a human confirms the
        # symptom in the real UI (ui_verified), it may ONLY be P5 —
        # parking it at P3/P4 must not dodge that requirement either.
        if not ui_verified and tier != "P5_dormant_or_latent":
            out.append(
                f"{fclass} rated {tier} but evidence.ui_verified is not "
                f"True.  A metric parsed from an export file / an injected "
                f"harness marker is not user-perceptible on its own — open "
                f"the artifact in the real Access UI, confirm the symptom, "
                f"and set evidence.ui_verified=True, or file it as "
                f"P5_dormant_or_latent pending verification."
            )

    # --- required evidence for user-perceptible tiers ---
    if tier in _USER_PERCEPTIBLE_TIERS:
        for field in ("vba_ref", "fixture", "user_symptom"):
            if not str(ev.get(field) or "").strip():
                out.append(
                    f"tier {tier} requires a non-empty evidence.{field} "
                    f"(source location / concrete reproducer / the symptom "
                    f"a user observes)."
                )
        symptom = str(ev.get("user_symptom") or "").strip().lower()
        # Lint (a BACKSTOP, not a proof) for the common anti-pattern of
        # pasting test/assertion output in place of a user-observable
        # symptom.  Patterns are high-precision — comparison operators,
        # pytest-style "expected N … got M" count assertions, this repo's
        # test-id prefixes and harness markers — chosen so they don't
        # flag legitimate symptom prose ("the user got an empty folder").
        # It cannot prove a symptom is genuinely user-facing; that is the
        # reviewer's call (see the finding_class note above).
        _restate_prefixes = ("detected by", "assertion", "assert ")
        _restate_substrings = (
            "zz_test_debug", "detected by test", "test_cmd", "test_export",
            "test_bug", "test_index", "test_vba", ":err marker",
            "!=", "==", "assertion failed", "assertion error",
        )
        if (symptom.startswith(_restate_prefixes)
                or any(s in symptom for s in _restate_substrings)
                or re.search(r"expected\b.{0,40}\bgot\b.*\d", symptom)):
            out.append(
                "evidence.user_symptom must describe what the USER observes "
                "(a popup, blank-or-wrong on-screen data, a missing/corrupt "
                "file), not restate the test assertion ('Detected by …' / "
                "'expected N … got M')."
            )

    return out


def _validate_issues() -> None:
    """Enforce the report-triage contract before any file is written.

    Fails loudly (raises ValueError listing every violation) so a
    mis-tiered or under-evidenced issue surfaces immediately instead of
    shipping in the report.  See _issue_violations / the module banner
    above for the rules.  An empty ISSUES list passes (clean slate).
    """
    problems: list[str] = []
    for it in ISSUES:
        for msg in _issue_violations(it):
            problems.append(f"ISSUES id={it.get('id')!r}: {msg}")
    if problems:
        raise ValueError(
            f"Report-triage contract violations ({len(problems)}):\n  - "
            + "\n  - ".join(problems)
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
    # Build stamp (B7 part 2): record which DATA build this report describes.
    try:
        import sys as _sys
        _ana = str(REPO / "analysis")
        if _ana not in _sys.path:
            _sys.path.insert(0, _ana)
        from build_stamp import current_build as _cur_build
        _build = _cur_build(REPO)
    except Exception:
        _build = None
    _build_label = ("Data build: " if is_en else "数据构建：") + (_build or "unknown")
    lines.append(f"_{Z(_build_label)}_")
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
