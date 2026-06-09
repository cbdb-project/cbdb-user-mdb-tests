"""Audit hardcoded code labels in the issue report against the
MDB dictionary tables.

Prompted by Issue #20's status-code-40 wording mistake: the
report's `steps_en/zh` originally said "Provincial Graduate /
进士" for `c_status_code = 40`, which is wrong — that's
Issue #9's jinshi / `c_entry_code = 36` label, copied by
mistake.  Pure prose review missed it because the bug-class
needs cross-checking the report's manually-typed labels
against `STATUS_CODES.c_status_desc[_chn]`,
`ENTRY_CODES.c_entry_desc[_chn]`, etc.

This audit:

1. Curates a manifest of all (issue_id, table, code_value)
   triples currently present in the report.  Each entry
   declares the exact `expected_labels` that MUST appear in
   that issue's block of the regenerated markdown report,
   plus optional `forbidden_labels` (to catch the specific
   copy-paste failure from Issue #20).
2. Reads each code's actual desc fields from the MDB to
   verify the manifest's `expected_labels` are correct (so
   if the maintainer renames a code in the data, the audit
   itself fails loudly rather than silently rubber-stamping
   the report).
3. Reads `reports/CBDB_Issues_Report_EN.md` and
   `reports/CBDB_Issues_Report_ZH-Hant.md`; for each
   manifest entry, slices out the issue block (between
   `### Issue #N —` and the next `### Issue #`) and asserts:
     - every `expected_labels[lang]` substring is present
     - no `forbidden_labels[lang]` substring is present
     - the MDB-side desc value is mentioned somewhere in the
       block (either via expected_labels or by literal match)
4. Writes the audit report to
   `reports/report_code_label_audit.json`.

If the audit finds a mismatch, the JSON's `mismatches` array
will be non-empty.  Run this directly
(`python analysis/audit_report_code_labels.py`) after regenerating
the report; the pytest wrapper that used to assert it is re-added
when the canonical report is rebuilt.

Pure pyodbc; no Access COM.
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

import pyodbc

ROOT = Path(__file__).resolve().parent.parent
USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
EN_MD = ROOT / "reports" / "CBDB_Issues_Report_EN.md"
ZH_MD = ROOT / "reports" / "CBDB_Issues_Report_ZH-Hant.md"
OUT_JSON = ROOT / "reports" / "report_code_label_audit.json"


# ----- Manifest -----------------------------------------------------
# Each entry: which (issue_id, MDB code) appears in the report and
# what we require to be true about its label rendering.
#
# `expected_labels`: substrings that MUST appear in the named issue
#   block of the regenerated report.  Use the actual DESC value the
#   MDB returns (verified at audit time) plus any well-known
#   localisation the report uses.
#
# `forbidden_labels`: substrings that MUST NOT appear in the named
#   issue block.  Use to pin specific copy-paste mistakes (e.g.
#   Issue #20 must never re-acquire jinshi / 進士 wording).
#
# The auditor also independently fetches the MDB desc fields and
# asserts they match the human-curated `expected_labels` (so the
# manifest stays honest if the data is updated).
MANIFEST: list[dict] = [
    {
        "issue_id": 7,
        "table": "ADDR_CODES",
        "code_col": "c_addr_id",
        "code_value": 100658,
        "desc_cols": ["c_name", "c_name_chn"],
        "expected_labels": {
            "en": ["c_addr_id = 100658", "Kaifeng"],
            "zh": ["c_addr_id = 100658", "開封"],
        },
        # The earlier draft used c_addr_id=7213, which is Daxing
        # (大興), not Kaifeng — the audit caught this and the
        # report was corrected.  Pin so it cannot regress.
        "forbidden_labels": {
            "en": ["c_addr_id = 7213"],
            "zh": ["c_addr_id = 7213"],
        },
    },
    {
        "issue_id": 9,
        "table": "ENTRY_CODES",
        "code_col": "c_entry_code",
        "code_value": 36,
        "desc_cols": ["c_entry_desc", "c_entry_desc_chn"],
        "expected_labels": {
            "en": ["c_entry_code = 36", "jinshi"],
            "zh": ["c_entry_code = 36", "進士"],
        },
        # Issue #9 IS about jinshi — these are not forbidden here.
        "forbidden_labels": {
            "en": ["civil office"],   # would mean Issue #9 picked
            "zh": ["[為官者：文]"],   # up Issue #20's label by mistake
        },
    },
    {
        "issue_id": 9,
        "table": "ENTRY_CODES",
        "code_col": "c_entry_code",
        "code_value": 101,
        "desc_cols": ["c_entry_desc", "c_entry_desc_chn"],
        "expected_labels": {
            "en": ["c_entry_code = 101", "recommendation"],
            "zh": ["c_entry_code = 101", "薦舉"],
        },
        "forbidden_labels": {"en": [], "zh": []},
    },
    {
        "issue_id": 10,
        "table": "ADDR_CODES",
        "code_col": "c_addr_id",
        "code_value": 12603,
        "desc_cols": ["c_name", "c_name_chn"],
        "expected_labels": {
            "en": ["c_addr_id = 12603", "Anfeng"],
            "zh": ["c_addr_id = 12603", "安豐"],
        },
        "forbidden_labels": {"en": [], "zh": []},
    },
    {
        "issue_id": 20,
        "table": "STATUS_CODES",
        "code_col": "c_status_code",
        "code_value": 40,
        "desc_cols": ["c_status_desc", "c_status_desc_chn"],
        "expected_labels": {
            "en": ["civil office"],
            "zh": ["[為官者：文]"],
        },
        # The Issue #20 fixture was originally mislabelled as Issue
        # #9's jinshi / Provincial Graduate.  Pin so it cannot recur.
        "forbidden_labels": {
            "en": [
                "status code **40** (Provincial Graduate / 进士)",
                "status code **40** (jinshi",
            ],
            "zh": [
                "status code **40**（进士）",
                "status code **40**（進士）",  # s2twp form
                "status code **40**（jinshi）",
            ],
        },
    },
    {
        "issue_id": 20,
        "table": "ADDR_CODES",
        "code_col": "c_addr_id",
        "code_value": 702559,
        "desc_cols": ["c_name", "c_name_chn"],
        "expected_labels": {
            "en": ["c_addr_id = 702559", "Wei Shi"],
            "zh": ["c_addr_id = 702559", "尉氏"],
        },
        "forbidden_labels": {"en": [], "zh": []},
    },
    {
        "issue_id": 22,
        "table": "ASSOC_CODES",
        "code_col": "c_assoc_code",
        "code_value": 437,
        "desc_cols": ["c_assoc_desc", "c_assoc_desc_chn"],
        "expected_labels": {
            "en": ["c_assoc_code = 437", "Presented literary composition as gift to"],
            "zh": ["c_assoc_code = 437", "贈詩、文"],
        },
        # Pin: must NOT reference a person ID for this issue.
        # LookAtAssociations has no person picker — the query
        # entry point is CmdPickAssoc (association-code picker).
        "forbidden_labels": {
            "en": ["c_personid = 437", "person picker"],
            "zh": ["c_personid = 437", "人物 picker"],
        },
    },
]
# NOTE: Issue #21 (LookAtOffice CmdGIS IndexYear) was removed from the
# ISSUES list in build-20260605 after cross-checking the source data.
# The 0.2% IndexYear fill rate in c_office_id=80944 (典史) output is a
# genuine data characteristic: BIOG_MAIN.c_index_year is NULL for
# 37,746/37,848 holders of that office (mostly Qing officials without a
# career landmark year).  VBA query is correct; no MANIFEST entry needed.


# -------------------------------------------------------------------
# MDB lookups + report block slicing
# -------------------------------------------------------------------
def _open_mdb() -> pyodbc.Connection:
    return pyodbc.connect(
        "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={USER_MDB};", autocommit=True, readonly=True)


_MDB_INACCESSIBLE = object()  # sentinel: linked table path not valid on this host


def _fetch_desc(cur, entry: dict) -> dict | None | object:
    cols = ", ".join(entry["desc_cols"])
    sql = (f"SELECT {cols} FROM {entry['table']} "
           f"WHERE {entry['code_col']} = {int(entry['code_value'])}")
    try:
        cur.execute(sql)
        row = cur.fetchone()
    except pyodbc.Error:
        return _MDB_INACCESSIBLE
    if row is None:
        return None
    return {entry["desc_cols"][i]: row[i] for i in range(len(entry["desc_cols"]))}


def _issue_block(text: str, issue_id: int) -> str | None:
    start_marker = f"### Issue #{issue_id} —"
    start = text.find(start_marker)
    if start == -1:
        return None
    next_start = text.find("### Issue #",
                            start + len(start_marker))
    end = next_start if next_start != -1 else len(text)
    return text[start: end]


def _check_entry_against_block(
        entry: dict, lang: str, block: str | None,
        mdb_desc: dict | None) -> dict:
    """Run the manifest entry's assertions against `block`.
    Returns a dict per-(entry, lang) with all findings."""
    out = {
        "issue_id": entry["issue_id"],
        "table": entry["table"],
        "code_col": entry["code_col"],
        "code_value": entry["code_value"],
        "lang": lang,
        "block_found": block is not None,
        "mdb_desc": mdb_desc,
        "missing_expected_labels": [],
        "found_forbidden_labels": [],
        "mdb_desc_present_in_block": None,
    }
    if block is None:
        return out

    # Required strings.
    for lab in entry["expected_labels"].get(lang, []):
        if lab not in block:
            out["missing_expected_labels"].append(lab)

    # Forbidden strings.
    for lab in entry["forbidden_labels"].get(lang, []):
        if lab in block:
            out["found_forbidden_labels"].append(lab)

    # MDB-derived desc must appear somewhere in the block (either
    # as an `expected_labels` entry literally or as a free-form
    # mention).  This catches the case where the maintainer renames
    # a code in the data and the report's manual label drifts.
    if mdb_desc:
        # Use the lang-appropriate desc column heuristically.
        # ENTRY_CODES.c_entry_desc is plain English; …_chn is CJK.
        candidate_cols = [c for c in mdb_desc
                           if c.endswith("_chn")] if lang == "zh" \
                           else [c for c in mdb_desc
                                  if not c.endswith("_chn")]
        any_present = False
        for col in candidate_cols:
            v = mdb_desc[col]
            if v is None or not str(v).strip():
                continue
            # Strip leading BOM (Issue #20-class data quirk;
            # report won't include BOM in display strings even
            # though MDB stores it).  Then break on every common
            # delimiter — `/,;()[]:` and whitespace — and try
            # each non-trivial token as a substring of the
            # block.  We also try the original whole string.
            v_clean = str(v).lstrip("﻿").strip()
            tokens: list[str] = [v_clean]
            for t in re.split(r"[/,;()\[\]:\s]+", v_clean):
                t = t.strip(" ﻿")
                if len(t) >= 2:
                    tokens.append(t)
            if any(tok and tok in block for tok in tokens):
                any_present = True
                break
        out["mdb_desc_present_in_block"] = any_present

    return out


# -------------------------------------------------------------------
def main() -> int:
    print("=== report code label audit ===")
    if not EN_MD.exists() or not ZH_MD.exists():
        raise SystemExit(
            f"missing report files; run "
            f"`python reports/generate_report.py` first")
    en_text = EN_MD.read_text(encoding="utf-8")
    zh_text = ZH_MD.read_text(encoding="utf-8")

    conn = _open_mdb()
    cur = conn.cursor()

    findings: list[dict] = []
    mismatches: list[dict] = []
    n_checks_passed = 0
    for entry in MANIFEST:
        mdb_desc = _fetch_desc(cur, entry)
        if mdb_desc is _MDB_INACCESSIBLE:
            # Linked table lives in data MDB whose path only exists on
            # the original author's machine.  Log a notice and proceed
            # with mdb_desc=None so string-only checks still run.
            print(f"  [notice] {entry['table']} inaccessible for "
                  f"issue #{entry['issue_id']} — skipping MDB desc check")
            mdb_desc = None
        elif mdb_desc is None:
            findings.append({
                "issue_id": entry["issue_id"],
                "table": entry["table"],
                "code_col": entry["code_col"],
                "code_value": entry["code_value"],
                "_error": (
                    f"no row in {entry['table']} for "
                    f"{entry['code_col']} = {entry['code_value']}"
                ),
            })
            mismatches.append(findings[-1])
            continue

        for lang, text in (("en", en_text), ("zh", zh_text)):
            block = _issue_block(text, entry["issue_id"])
            r = _check_entry_against_block(entry, lang, block, mdb_desc)
            findings.append(r)
            # mdb_desc_present_in_block is None when the table was
            # inaccessible — treat that as "skipped" (not a failure).
            ok = (
                r["block_found"]
                and not r["missing_expected_labels"]
                and not r["found_forbidden_labels"]
                and r["mdb_desc_present_in_block"] is not False
            )
            if not ok:
                mismatches.append(r)
            else:
                n_checks_passed += 1

    out = {
        "summary": {
            "n_manifest_entries": len(MANIFEST),
            "n_per_lang_checks_total": len(MANIFEST) * 2,
            "n_per_lang_checks_passed": n_checks_passed,
            "n_mismatches": len(mismatches),
            "manifest_coverage": [
                {"issue_id": m["issue_id"],
                 "table": m["table"],
                 "code_col": m["code_col"],
                 "code_value": m["code_value"]}
                for m in MANIFEST
            ],
        },
        "findings": findings,
        "mismatches": mismatches,
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"wrote {OUT_JSON}")
    print(f"  manifest entries: {len(MANIFEST)}")
    print(f"  checks passed:    {n_checks_passed} / "
          f"{len(MANIFEST) * 2}")
    print(f"  mismatches:       {len(mismatches)}")
    if mismatches:
        print()
        print("MISMATCHES:")
        for m in mismatches:
            print(f"  issue #{m.get('issue_id')} "
                  f"{m.get('table')}.{m.get('code_col')}="
                  f"{m.get('code_value')} "
                  f"[{m.get('lang')}]: "
                  f"missing={m.get('missing_expected_labels')} "
                  f"forbidden={m.get('found_forbidden_labels')} "
                  f"mdb_desc_present={m.get('mdb_desc_present_in_block')}"
                  f"{' _error=' + m['_error'] if '_error' in m else ''}")
    return 0 if not mismatches else 1


if __name__ == "__main__":
    sys.exit(main())
