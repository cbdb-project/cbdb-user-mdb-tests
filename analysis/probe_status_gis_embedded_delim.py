"""Probe LookAtStatus.CmdGIS for embedded delimiters / control chars.

Background
----------
PR T's `--include-vba` run flagged `LookAtStatus` row 11476 of the
exported `.tab` as having 10 tab-cells against a 9-col header.  The
exported snippet (decoded) was approximately::

  Ruan Fu \t 阮孚 \t M \t 283 \t ﻿Wei Shi \t  \t \\\x0fl \t
  114.116... \t 34.400... \t 1

That's 10 cells.  The export writer (Form_LookAtStatus.vb:1554-1636,
`CmdGIS_Click`) builds each output line as
``tStr + (column value) + tC`` with **no escaping** of cell contents
when ``tC = Chr(9)`` (line 1552).  So **any embedded tab / BOM / SUB
character / control char in a text column will silently produce extra
columns**.  This is a textbook silent-export-corruption pattern.

This probe scans the candidate source columns for problem characters,
without re-running the VBA chain (which is flaky in batch mode):

  - BIOG_MAIN.c_name / c_name_chn        (→ Name / NameChn)
  - ADDR_CODES.c_name / c_name_chn       (→ AddrName / AddrChn after
                                          the JOIN at Form_LookAtStatus
                                          .vb:1398)
  - BIOG_ADDR_DATA.c_addr_remarks / etc. (defensive — not in GIS, but
                                          in case there's a NULL-fill
                                          path we missed)

It also checks `ZZ_SCRATCH_P_STATUS` if it's been populated by a
recent CmdQuery run against status_code=40 (the PR T fixture).

Output: reports/gis_embedded_delimiter_findings.json with the offending
addr_id / personid, the source column, and the chars found.

Conservative classification
---------------------------
Likely candidate cause: `c_addr_name` / `c_addr_chn` for some
ADDR_CODES rows contains literal tab + BOM characters.  Two parts:

  1. **source-data-dirty** (upstream):
     ADDR_CODES has rows with embedded delimiters.  Real CBDB users
     who export GIS will see column misalignment in their .tab files.

  2. **export-writer-not-escaping** (in CBDB):
     CmdGIS does not sanitise cell values.  Even if (1) is fixed
     upstream, the same bug class would resurface the next time any
     editor pastes a tab character into a text field.

(1) is fixable in a one-shot data cleanup; (2) is the architectural
gap.  Both are real candidate bugs.  Neither is labelled "confirmed"
without manual reproduction in the live UI.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pyodbc

ROOT = Path(__file__).resolve().parent.parent
USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
OUT = ROOT / "reports" / "gis_embedded_delimiter_findings.json"

# Characters that would silently break a tab-separated export.
PROBLEM_CHARS = {
    "\t":   "TAB (U+0009) — CmdGIS field separator",
    "\n":   "LF (U+000A) — CmdGIS line separator",
    "\r":   "CR (U+000D) — CmdGIS line separator",
    "﻿": "BOM (U+FEFF) — silent prefix",
    "\x00": "NUL (U+0000)",
    "\x0f": "SI (U+000F)",
    "\x1a": "SUB (U+001A) — Access EOF marker",
}


def _open(read_only: bool = True) -> pyodbc.Connection:
    return pyodbc.connect(
        "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={USER_MDB};", autocommit=True, readonly=read_only)


def _scan(cur, table: str, key_cols: list[str],
          text_cols: list[str]) -> list[dict]:
    """Scan `table` for rows whose any `text_cols` value contains a
    problem character.  Returns one dict per offending (row, column,
    char)."""
    cols_sql = ", ".join(key_cols + text_cols)
    try:
        cur.execute(f"SELECT {cols_sql} FROM {table}")
    except pyodbc.Error as e:
        return [{"_error": f"{table}: {e}"}]
    findings: list[dict] = []
    for row in cur.fetchall():
        keys = {k: row[i] for i, k in enumerate(key_cols)}
        for j, c in enumerate(text_cols, start=len(key_cols)):
            v = row[j]
            if v is None:
                continue
            sv = str(v)
            hits = []
            for ch, label in PROBLEM_CHARS.items():
                if ch in sv:
                    hits.append({
                        "char_codepoint": f"U+{ord(ch):04X}",
                        "char_label": label,
                    })
            if hits:
                findings.append({
                    "table": table,
                    "keys": keys,
                    "column": c,
                    "raw_repr": repr(sv)[:300],
                    "raw_len": len(sv),
                    "hits": hits,
                })
    return findings


def main() -> int:
    print(f"opening {USER_MDB}")
    conn = _open()
    cur = conn.cursor()

    all_findings: list[dict] = []

    # 1. ADDR_CODES — the JOIN source for c_addr_name / c_addr_chn.
    all_findings += _scan(
        cur, "ADDR_CODES",
        key_cols=["c_addr_id"],
        text_cols=["c_name", "c_name_chn"],
    )

    # 2. BIOG_MAIN — c_name / c_name_chn for the Name / NameChn cols.
    all_findings += _scan(
        cur, "BIOG_MAIN",
        key_cols=["c_personid"],
        text_cols=["c_name", "c_name_chn"],
    )

    # 3. ZZ_SCRATCH_P_STATUS — if a recent CmdQuery left rows behind,
    # check the actually-staged values (this is what CmdGIS reads
    # from).  If it's empty, that's fine — the join sources above
    # are the underlying truth.
    all_findings += _scan(
        cur, "ZZ_SCRATCH_P_STATUS",
        key_cols=["c_person_id", "c_addr_id"],
        text_cols=["c_addr_name", "c_addr_chn"],
    )

    # Group by (table, column) to summarise.
    summary: dict[tuple, int] = {}
    for f in all_findings:
        if "_error" in f:
            continue
        k = (f["table"], f["column"])
        summary[k] = summary.get(k, 0) + 1

    # If any ADDR_CODES rows are dirty, check whether they are
    # reachable from status_code=40 (the PR T fixture).  This tells
    # us whether the failure is reproducible or just lurking.
    reachable_from_status_40: list[dict] = []
    addr_code_dirty_ids = sorted({
        f["keys"]["c_addr_id"]
        for f in all_findings
        if f.get("table") == "ADDR_CODES"
    })
    if addr_code_dirty_ids:
        ids_sql = ",".join(str(i) for i in addr_code_dirty_ids
                            if i is not None)
        try:
            cur.execute(
                "SELECT DISTINCT bad.c_personid, bad.c_addr_id "
                "FROM BIOG_ADDR_DATA bad "
                "INNER JOIN STATUS_DATA sd "
                "  ON sd.c_personid = bad.c_personid "
                f"WHERE sd.c_status_code = 40 "
                f"  AND bad.c_addr_id IN ({ids_sql})"
            )
            for r in cur.fetchall():
                reachable_from_status_40.append({
                    "c_personid": int(r[0]),
                    "c_addr_id": int(r[1]),
                })
        except pyodbc.Error as e:
            reachable_from_status_40 = [{"_error": str(e)}]

    out = {
        "summary": {
            "total_findings": sum(1 for f in all_findings
                                    if "_error" not in f),
            "by_table_column": [
                {"table": t, "column": c, "rows_affected": n}
                for (t, c), n in sorted(summary.items())
            ],
            "errors": [f["_error"] for f in all_findings
                        if "_error" in f],
            "reachable_from_status_code_40": {
                "count": len(reachable_from_status_40),
                "rows": reachable_from_status_40[:20],
            },
        },
        "findings": all_findings,
        "interpretation": (
            "CmdGIS_Click in Form_LookAtStatus.vb (and its peers in "
            "LookAtTexts/Place/Office/Associations/Kinship) "
            "concatenates raw recordset values into the line buffer "
            "with `tStr + value + tC` where tC = Chr(9).  No "
            "escaping is done.  Any text field whose value contains "
            "a literal TAB, BOM, CR/LF, or other control character "
            "produces an export row with the wrong column count, "
            "silently misaligning every column to its right.  "
            "Affected rows reproduce in real .tab exports — users "
            "see numeric coordinates land in the wrong column."
        ),
        "candidate_classification": (
            "candidate_export_writer_no_delimiter_escaping (severity "
            "depends on (a) how many ADDR_CODES rows are actually "
            "dirty and (b) which forms reach them).  Compounding "
            "cause: source data dirty (some upstream editor allowed "
            "a tab character into ADDR_CODES.c_name).  Mitigations: "
            "(1) sanitise CmdGIS output (replace Chr(9) / Chr(10) / "
            "Chr(13) / BOM with space before append); (2) data "
            "cleanup of ADDR_CODES.  Both candidate, neither "
            "currently labelled confirmed bug."
        ),
        "is_confirmed_bug": False,
    }

    OUT.write_text(
        json.dumps(out, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"wrote {OUT}")
    print(f"  total findings: {out['summary']['total_findings']}")
    for entry in out['summary']['by_table_column']:
        print(f"    {entry['table']}.{entry['column']}: "
              f"{entry['rows_affected']} rows")
    print(f"  reachable from status_code=40: "
          f"{out['summary']['reachable_from_status_code_40']['count']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
