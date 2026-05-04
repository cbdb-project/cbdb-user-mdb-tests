"""Static delimiter-risk audit across export-bound text fields.

Extends Issue #20 (BOM-prefixed ADDR_CODES rows that JET mangles
into embedded TAB during scratch-table staging, breaking GIS
.tab exports).  PR Z's reach analysis showed only 1 of the 315
ADDR_CODES BOM rows is actually surfacing in user exports today;
this audit asks the broader question: which OTHER text columns
across export-bound source tables carry the same risk?

Scan rules:
  - TAB (U+0009)             — would split tab-separated GIS rows
  - LF (U+000A) / CR (U+000D) — would split line-based exports
  - U+FEFF (BOM)             — JET mangles to TAB on UPDATE/INSERT
  - U+0000 (NUL)             — would truncate C-string consumers
  - U+000B / U+000C          — vertical tab / form feed
  - U+000F (SI)              — appears post-mangling but worth tracking
  - U+001A (SUB)             — Access internal EOF marker
  - comma (",")              — risk for unquoted CSV-like exports
                                 (Pajek node labels, GUESS, etc.)

Pure pyodbc.  No Access COM.  Read-only.

Outputs:
  - reports/export_delimiter_risk_audit.json
  - analysis/export_delimiter_risk_audit.md
"""
from __future__ import annotations

import json
import sys
from collections import defaultdict
from pathlib import Path

import pyodbc

ROOT = Path(__file__).resolve().parent.parent
USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
OUT_JSON = ROOT / "reports" / "export_delimiter_risk_audit.json"
OUT_MD = ROOT / "analysis" / "export_delimiter_risk_audit.md"

PROBLEM_CHARS: dict[str, str] = {
    "\t":     "TAB (U+0009) — splits tab-separated GIS .tab rows",
    "\n":     "LF (U+000A) — splits line-based exports",
    "\r":     "CR (U+000D) — splits line-based exports",
    "﻿": "BOM (U+FEFF) — JET mangles to TAB on UPDATE/INSERT (Issue #20)",
    "\x00":   "NUL (U+0000) — truncates C-string consumers",
    "\x0b":   "VT (U+000B) — vertical tab",
    "\x0c":   "FF (U+000C) — form feed",
    "\x0f":   "SI (U+000F) — typically post-BOM-mangle artefact",
    "\x1a":   "SUB (U+001A) — Access internal EOF marker",
}

COMMA = ","   # tracked separately; less severe (only Pajek-family + a few)

# Per-table column scope.  Conservative — keep to text-shaped
# columns most likely to surface in exports.
TABLES_AND_COLS: list[tuple[str, list[str]]] = [
    ("ADDR_CODES",          ["c_name", "c_name_chn", "c_notes", "c_alt_names"]),
    ("BIOG_MAIN",           ["c_name", "c_name_chn", "c_surname",
                              "c_surname_chn", "c_mingzi_chn",
                              "c_surname_proper", "c_name_proper",
                              "c_surname_rm", "c_name_rm",
                              "c_fl_ey_notes", "c_fl_ly_notes", "c_notes"]),
    ("ENTRY_CODES",         ["c_entry_desc", "c_entry_desc_chn"]),
    ("ASSOC_CODES",         ["c_assoc_desc", "c_assoc_desc_chn"]),
    ("STATUS_CODES",        ["c_status_desc", "c_status_desc_chn"]),
    ("TEXT_CODES",          ["c_title_chn", "c_title_alt_chn", "c_notes"]),
    ("TEXT_BIBLCAT_CODES",  ["c_text_cat_desc", "c_text_cat_desc_chn",
                              "c_text_cat_pinyin"]),
    ("OFFICE_CODES",        ["c_office_pinyin", "c_office_chn",
                              "c_office_pinyin_alt", "c_office_chn_alt",
                              "c_office_trans_alt", "c_notes"]),
]


def _open() -> pyodbc.Connection:
    return pyodbc.connect(
        "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={USER_MDB};", autocommit=True, readonly=True)


def _scan_table(cur, table: str, cols: list[str]) -> dict:
    cols_sql = ", ".join(["c_" + table.split("_")[0].lower() + "_id"
                          if False else cols[0]] + cols)
    # Just SELECT all listed columns; iterate.
    try:
        select = ", ".join(cols)
        cur.execute(f"SELECT {select} FROM [{table}]")
    except pyodbc.Error as e:
        return {"_error": str(e), "table": table}
    rows = cur.fetchall()
    n_total = len(rows)
    findings_per_col: dict[str, dict[str, int]] = defaultdict(
        lambda: defaultdict(int))
    findings_per_col_comma: dict[str, int] = defaultdict(int)
    samples_per_col_char: dict[tuple[str, str], list[dict]] = defaultdict(list)

    for r in rows:
        for j, c in enumerate(cols):
            v = r[j]
            if v is None:
                continue
            sv = str(v)
            for ch, label in PROBLEM_CHARS.items():
                if ch in sv:
                    findings_per_col[c][f"U+{ord(ch):04X}"] += 1
                    if len(samples_per_col_char[(c, ch)]) < 3:
                        samples_per_col_char[(c, ch)].append({
                            "value_repr": repr(sv)[:160],
                            "len": len(sv),
                        })
            if COMMA in sv:
                findings_per_col_comma[c] += 1

    return {
        "table": table,
        "n_rows": n_total,
        "scanned_columns": cols,
        "findings_problem_chars": {
            c: dict(d) for c, d in findings_per_col.items()
        },
        "findings_comma": dict(findings_per_col_comma),
        "samples": [
            {"column": c, "char_codepoint": f"U+{ord(ch):04X}",
             "char_label": PROBLEM_CHARS[ch],
             "examples": samples_per_col_char[(c, ch)]}
            for (c, ch) in samples_per_col_char
        ],
    }


def _classify(table_results: list[dict]) -> dict:
    headline = []
    for r in table_results:
        if "_error" in r:
            continue
        for col, char_counts in r["findings_problem_chars"].items():
            for cp, n in char_counts.items():
                # exclude the well-known Issue #20 ADDR_CODES BOM rows
                # from the headline since they are already documented
                # — but still surface them for completeness.
                known = (r["table"] == "ADDR_CODES"
                         and col in {"c_name", "c_name_chn"}
                         and cp == "U+FEFF")
                headline.append({
                    "table": r["table"],
                    "column": col,
                    "char_codepoint": cp,
                    "rows_affected": n,
                    "known_issue_20": known,
                })
    headline.sort(key=lambda h: -h["rows_affected"])
    return {
        "headline_findings_sorted_by_severity": headline,
        "n_new_candidate_findings_excluding_issue_20": sum(
            1 for h in headline if not h["known_issue_20"]),
    }


def main() -> int:
    print(f"opening {USER_MDB}")
    conn = _open()
    cur = conn.cursor()
    table_results = []
    for tbl, cols in TABLES_AND_COLS:
        print(f"  scanning {tbl} ({len(cols)} cols)…", end=" ")
        try:
            r = _scan_table(cur, tbl, cols)
            print(f"{r.get('n_rows', '?')} rows")
        except pyodbc.Error as e:
            r = {"_error": str(e), "table": tbl}
            print(f"ERROR: {e}")
        table_results.append(r)
    classified = _classify(table_results)

    out = {
        "summary": {
            "tables_scanned": [t for t, _ in TABLES_AND_COLS],
            "n_total_findings": len(
                classified["headline_findings_sorted_by_severity"]),
            "n_new_candidate_findings_excluding_issue_20":
                classified["n_new_candidate_findings_excluding_issue_20"],
            "problem_chars": {
                f"U+{ord(c):04X}": label
                for c, label in PROBLEM_CHARS.items()
            },
        },
        "tables": table_results,
        "headline_findings_sorted_by_severity":
            classified["headline_findings_sorted_by_severity"],
        "is_confirmed_bug": False,
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"\nwrote {OUT_JSON}")

    # ----- markdown -----
    md = []
    md.append("# Export delimiter risk audit (PR AE)")
    md.append("")
    md.append("Static scan of export-bound text columns in the User MDB "
              "for characters that can break tab-/line-/comma-separated "
              "exports.  Extends Issue #20.")
    md.append("")
    md.append("## Headline")
    md.append("")
    md.append(f"- Tables scanned: {len(TABLES_AND_COLS)}")
    md.append(f"- Distinct (table, column, char) findings: "
              f"**{out['summary']['n_total_findings']}**")
    md.append(f"- New findings beyond Issue #20: "
              f"**{out['summary']['n_new_candidate_findings_excluding_issue_20']}**")
    md.append("")
    md.append("## Findings ranked by row count")
    md.append("")
    md.append("| Table | Column | Char | Rows | Known Issue #20 |")
    md.append("|---|---|---|---:|---|")
    for h in classified["headline_findings_sorted_by_severity"]:
        flag = "yes" if h["known_issue_20"] else "**no — new candidate**"
        md.append(
            f"| `{h['table']}` | `{h['column']}` | `{h['char_codepoint']}` "
            f"| {h['rows_affected']} | {flag} |"
        )
    md.append("")
    md.append("## Per-table detail")
    md.append("")
    for r in table_results:
        if "_error" in r:
            md.append(f"### `{r['table']}`")
            md.append(f"  - error: {r['_error']}")
            md.append("")
            continue
        md.append(f"### `{r['table']}` — {r['n_rows']} rows scanned")
        md.append("")
        if not r["findings_problem_chars"] and not r["findings_comma"]:
            md.append("(no findings)")
            md.append("")
            continue
        if r["findings_problem_chars"]:
            md.append("**Problem chars:**")
            for col, dd in r["findings_problem_chars"].items():
                pieces = ", ".join(f"{cp}×{n}"
                                    for cp, n in sorted(dd.items()))
                md.append(f"- `{col}` — {pieces}")
        if r["findings_comma"]:
            md.append("")
            md.append("**Comma occurrences (only relevant for "
                      "comma-separated / Pajek-family exports):**")
            for col, n in r["findings_comma"].items():
                md.append(f"- `{col}` — {n} rows contain `','`")
        if r["samples"]:
            md.append("")
            md.append("**Sample values:**")
            for s in r["samples"][:5]:
                ex = s["examples"][0] if s["examples"] else {"value_repr": "?"}
                md.append(
                    f"- `{s['column']}` / {s['char_codepoint']} "
                    f"({s['char_label']}): {ex['value_repr']}"
                )
        md.append("")
    md.append("## Notes")
    md.append("")
    md.append("- The 315 ADDR_CODES BOM-prefixed rows from Issue #20 "
              "are flagged here for completeness; they are already "
              "documented in the maintainer report.")
    md.append("- Comma findings are tracked but unranked because their "
              "severity depends on which export uses comma-quoting.  "
              "Pajek node labels and GUESS exports do quote, so commas "
              "in `c_name_chn` etc. should be safe.  GIS .tab and "
              "Neo4j CSVs use tab/comma + quote depending on writer.")
    md.append("- This audit does NOT examine `BIOG_ADDR_DATA`, "
              "`POSTED_TO_OFFICE_DATA`, etc.  Those are fact tables "
              "with foreign-key references to the code tables here; "
              "they inherit the same risk via JOIN.  Add separately "
              "if there's value.")
    md.append("- No bucket is labelled a confirmed CBDB bug.")
    OUT_MD.write_text("\n".join(md), encoding="utf-8")
    print(f"wrote {OUT_MD}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
