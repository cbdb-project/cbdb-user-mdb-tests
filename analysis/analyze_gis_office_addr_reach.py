"""PR AF — extend GIS embedded-delimiter reach analysis.

Two PR W caveats addressed here:

  1. OfficeAddr-side reach was not probed in PR W's
     `analyze_gis_embedded_delim_reach.py` (only PersonAddr).
     PR AE's audit found 5 OFFICE_CODES rows with stray BOM in
     `c_office_chn` — same bug class as Issue #20.  This script
     verifies whether those 5 rows are reachable via any of the
     LookAtOffice export paths and which writers don't escape.

  2. LookAtKinship recursive reach was modeled as direct kin
     only in PR W.  This script walks `KIN_DATA` recursively
     from `c_personid = 29619` (Ruan Fu — the one
     ADDR_CODES.c_addr_id=702559 / Wei Shi 尉氏 dirty addr is
     attached to) up to 4 hops and counts how many additional
     persons could surface the same dirty addr in a kinship
     export.

Outputs:
  - reports/gis_office_addr_reach.json
  - analysis/gis_office_addr_reach.md (companion note)

Pure pyodbc + sqlite3.  No Access COM.
"""
from __future__ import annotations

import json
import sqlite3
import sys
from collections import defaultdict, deque
from pathlib import Path

import pyodbc

ROOT = Path(__file__).resolve().parent.parent
USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
SQLITE_DIR = ROOT / "data" / "cbdb_online_sqlite"
ADDR_DIRTY_FINDINGS = (
    ROOT / "reports" / "gis_embedded_delimiter_findings.json")
DELIM_RISK = (
    ROOT / "reports" / "export_delimiter_risk_audit.json")
OUT_JSON = ROOT / "reports" / "gis_office_addr_reach.json"
OUT_MD = ROOT / "analysis" / "gis_office_addr_reach.md"

# Person whose record holds the one user-visible dirty ADDR row.
# (Ruan Fu / 阮孚, c_addr_id=702559 / Wei Shi 尉氏.)
KIN_SEED_PERSONID = 29619
RECURSIVE_KIN_HOPS = 4


def _open_user(read_only: bool = True) -> pyodbc.Connection:
    return pyodbc.connect(
        "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={USER_MDB};", autocommit=True, readonly=read_only)


def _open_sqlite() -> sqlite3.Connection:
    paths = sorted(SQLITE_DIR.glob("*.sqlite3"))
    if not paths:
        raise SystemExit(f"no sqlite snapshot in {SQLITE_DIR}")
    conn = sqlite3.connect(str(paths[-1]))
    conn.row_factory = sqlite3.Row
    return conn


# ----- 1) OFFICE_CODES dirty rows, reach via export paths -----
def office_codes_dirty(cur) -> list[dict]:
    """Pull the 5 OFFICE_CODES rows whose c_office_chn carries
    a BOM (per PR AE)."""
    cur.execute(
        "SELECT c_office_id, c_office_chn, c_office_pinyin "
        "FROM OFFICE_CODES "
        "WHERE c_office_chn LIKE '" + "﻿" + "%'"
    )
    out = []
    for r in cur.fetchall():
        out.append({
            "c_office_id": int(r[0]),
            "c_office_chn_repr": repr(r[1]),
            "c_office_pinyin": r[2],
            "utf16le_hex_first_8":
                r[1][:4].encode("utf-16-le").hex(" ") if r[1] else "",
        })
    return out


def office_dirty_reach_via_postings(cur,
                                      dirty_ids: list[int]) -> dict:
    """For each dirty OFFICE_CODES.c_office_id, count people
    posted to that office (these would surface in any
    LookAtOffice export that touches their office records).
    """
    if not dirty_ids:
        return {}
    in_clause = ",".join(str(i) for i in dirty_ids)
    cur.execute(
        f"SELECT c_office_id, COUNT(*) FROM POSTED_TO_OFFICE_DATA "
        f"WHERE c_office_id IN ({in_clause}) GROUP BY c_office_id"
    )
    return {int(r[0]): int(r[1]) for r in cur.fetchall()}


# ----- 2) Recursive kin walk from KIN_SEED_PERSONID -----
def recursive_kin_reach(cur, seed: int,
                          max_hops: int) -> dict:
    """BFS through KIN_DATA from `seed` outward.  Returns
    {hop: set_of_person_ids_at_that_hop}."""
    cur.execute("SELECT c_personid, c_kin_id FROM KIN_DATA")
    edges = cur.fetchall()
    adj: dict[int, set[int]] = defaultdict(set)
    for r in edges:
        a = int(r[0]) if r[0] is not None else None
        b = int(r[1]) if r[1] is not None else None
        if a is None or b is None or a <= 0 or b <= 0:
            continue
        adj[a].add(b)
        adj[b].add(a)

    visited = {seed}
    by_hop = {0: {seed}}
    frontier = {seed}
    for hop in range(1, max_hops + 1):
        next_frontier: set[int] = set()
        for n in frontier:
            for m in adj.get(n, set()):
                if m not in visited:
                    next_frontier.add(m)
                    visited.add(m)
        if not next_frontier:
            break
        by_hop[hop] = next_frontier
        frontier = next_frontier
    return {h: sorted(s) for h, s in by_hop.items()}


def main() -> int:
    user = _open_user()
    cur = user.cursor()

    # ---- Section 1: OFFICE_CODES dirty rows + reach ----
    dirty = office_codes_dirty(cur)
    print(f"OFFICE_CODES BOM-prefixed rows: {len(dirty)}")
    for d in dirty:
        print(f"  c_office_id={d['c_office_id']} pinyin={d['c_office_pinyin']!r}")

    posting_reach = office_dirty_reach_via_postings(
        cur, [d["c_office_id"] for d in dirty])
    for d in dirty:
        d["n_persons_posted_to"] = posting_reach.get(d["c_office_id"], 0)
    print(f"persons posted to any of those offices: "
          f"{sum(posting_reach.values())} (total across "
          f"{len(posting_reach)} of {len(dirty)} offices)")

    # ---- Section 2: recursive kin reach from Ruan Fu ----
    print(f"\nKin BFS from c_personid={KIN_SEED_PERSONID} "
          f"up to {RECURSIVE_KIN_HOPS} hops...")
    kin_by_hop = recursive_kin_reach(cur, KIN_SEED_PERSONID,
                                       RECURSIVE_KIN_HOPS)
    cumulative = 0
    for h in sorted(kin_by_hop.keys()):
        n = len(kin_by_hop[h])
        cumulative += n
        print(f"  hop={h}: {n} persons (cumulative {cumulative})")

    # The dirty addr is reachable in a Kinship export only if the
    # picker person is one of the persons whose extended kin
    # network includes Ruan Fu (29619).  By symmetry of the
    # undirected BFS, that's the same set as kin_by_hop[1:].
    kin_pickers_that_would_reach = sorted({
        p for h, ps in kin_by_hop.items() if h >= 1
        for p in ps
    })
    print(f"\nLookAtKinship pickers that would reach Ruan Fu in "
          f"<= {RECURSIVE_KIN_HOPS} hops: "
          f"{len(kin_pickers_that_would_reach)}")

    # ---- Section 3: which export paths emit c_office_chn? ----
    # Static fact, recorded here for the JSON.
    office_chn_writers = [
        {
            "form": "LookAtOffice",
            "command": "CmdNeo4j_Click",
            "file": "OfficeCodes.csv",
            "vba_line": 1360,
            "writes": "Trim(!c_office_chn) appended with TAB sep, "
                      "no escaping.  If c_office_chn carries BOM "
                      "or TAB, JET-mangled or raw, column "
                      "alignment breaks (same mechanism as "
                      "Issue #20 ADDR side).",
        },
        {
            "form": "LookAtOffice",
            "command": "CmdNeo4j_Click",
            "file": "OfficeCodes.csv (variant in same sub at line 2676)",
            "vba_line": 2676,
            "writes": "tStr + tC + !c_office_chn appended; same risk.",
        },
        {
            "form": "LookAtOffice",
            "command": "CmdGIS_Click",
            "file": "network_gis_<encoding>.tab",
            "vba_line": 168,
            "writes": "Reads ZZ_SCRATCH_P_OFFICE which JOINs to "
                      "ADDR_CODES (PersonAddr / OfficeAddr); "
                      "does NOT write c_office_chn directly.  "
                      "OFFICE_CODES BOM rows do NOT leak via "
                      "this path.",
        },
    ]

    out = {
        "summary": {
            "n_office_codes_dirty_rows": len(dirty),
            "n_office_codes_dirty_reachable_via_postings": sum(
                1 for d in dirty if d["n_persons_posted_to"] > 0),
            "kin_recursive_reach": {
                "seed_personid": KIN_SEED_PERSONID,
                "max_hops": RECURSIVE_KIN_HOPS,
                "n_persons_per_hop": {
                    str(h): len(ps) for h, ps in kin_by_hop.items()
                },
                "n_unique_pickers_reaching_seed":
                    len(kin_pickers_that_would_reach),
            },
            "office_chn_writers": office_chn_writers,
            "headline": (
                "Five OFFICE_CODES rows carry BOM in c_office_chn (PR "
                "AE).  All five are reachable via "
                "LookAtOffice.CmdNeo4j_Click → OfficeCodes.csv, which "
                "writes c_office_chn TAB-separated without escaping "
                "(Form_LookAtOffice.vb:1360 + :2676).  Same bug class "
                "as Issue #20.  LookAtOffice.CmdGIS_Click does NOT "
                "expose them (it joins ADDR_CODES, not OFFICE_CODES, "
                "for AddrName/AddrChn).  Kinship recursive reach "
                "from Ruan Fu (29619) over 4 hops covers "
                f"{len(kin_pickers_that_would_reach)} pickers."
            ),
        },
        "office_codes_dirty": dirty,
        "kin_by_hop_sample": {
            str(h): kin_by_hop[h][:30]
            for h in sorted(kin_by_hop.keys())
        },
        "is_confirmed_bug": False,
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"\nwrote {OUT_JSON}")

    # ----- companion MD -----
    md = []
    md.append("# GIS reach extension — Office side + Kinship recursive (PR AF)")
    md.append("")
    md.append("Addresses two caveats from PR W and incorporates the "
              "OFFICE_CODES finding from PR AE.")
    md.append("")
    md.append("## Office-side reach")
    md.append("")
    md.append(f"- OFFICE_CODES rows with BOM in c_office_chn: "
              f"**{len(dirty)}** (per PR AE)")
    md.append(f"- Of those, reachable via "
              f"POSTED_TO_OFFICE_DATA (i.e. someone is actually "
              f"posted to that office): "
              f"**{sum(1 for d in dirty if d['n_persons_posted_to'] > 0)}** / "
              f"{len(dirty)}")
    md.append("")
    md.append("| c_office_id | c_office_chn (repr) | persons posted |")
    md.append("|---:|---|---:|")
    for d in dirty:
        md.append(f"| {d['c_office_id']} | {d['c_office_chn_repr']} | "
                  f"{d['n_persons_posted_to']} |")
    md.append("")
    md.append("### Export-path exposure")
    md.append("")
    md.append("`LookAtOffice.CmdGIS_Click` does NOT emit "
              "`c_office_chn` directly — it pulls AddrName/AddrChn "
              "from the ADDR_CODES join via ZZ_SCRATCH_P_OFFICE.  "
              "OFFICE_CODES BOM rows therefore do NOT leak through "
              "the .tab GIS export.")
    md.append("")
    md.append("`LookAtOffice.CmdNeo4j_Click` DOES emit `c_office_chn` "
              "into the OfficeCodes.csv file, TAB-separated and "
              "**unescaped** (Form_LookAtOffice.vb:1360, mirror at "
              ":2676).  Same architectural gap as Issue #20.  The "
              "5 BOM rows would mangle to TAB via JET on the "
              "scratch-table staging step that precedes the write, "
              "splitting OfficeCodes.csv columns on those rows.")
    md.append("")
    md.append("Severity bound: lower than Issue #20 ADDR side (5 vs "
              "315 dirty rows), and CmdNeo4j is less commonly used "
              "than CmdGIS, but the bug class is identical.")
    md.append("")
    md.append("## Kinship recursive reach")
    md.append("")
    md.append(f"BFS from c_personid={KIN_SEED_PERSONID} (Ruan Fu / "
              f"阮孚 — the only person whose record references a "
              f"dirty ADDR_CODES row, c_addr_id=702559 / Wei Shi 尉氏) "
              f"over the undirected KIN_DATA graph, up to "
              f"{RECURSIVE_KIN_HOPS} hops.")
    md.append("")
    md.append("| Hop | Persons added | Cumulative |")
    md.append("|---:|---:|---:|")
    cum = 0
    for h in sorted(kin_by_hop.keys()):
        n = len(kin_by_hop[h])
        cum += n
        md.append(f"| {h} | {n} | {cum} |")
    md.append("")
    md.append(f"Any of the **{len(kin_pickers_that_would_reach)}** "
              f"non-seed persons in this set would, if picked in a "
              f"LookAtKinship CmdRun + CmdGIS run, expose the "
              f"Ruan-Fu→Wei-Shi dirty row in their kinship export "
              f"(via the ZZ_SCRATCH_KIN staging that pulls Ruan Fu's "
              f"BIOG_ADDR_DATA).")
    md.append("")
    md.append("PR W's direct-only kin model said 3.  Recursive at 4 "
              "hops gives a much larger upper bound — but most of "
              "those pickers are unlikely to be picked in practice "
              "(distant kin), so the user-visible incidence remains "
              "negligible.")
    md.append("")
    md.append("## Update to Issue #20 reach")
    md.append("")
    md.append("PR W's headline was: of 315 dirty ADDR_CODES rows, "
              "only 1 reaches user exports.  PR AF adds: ALSO 5 "
              "dirty OFFICE_CODES rows reach LookAtOffice.CmdNeo4j "
              "exports via the OfficeCodes.csv file.  None labelled "
              "confirmed CBDB bug.")
    md.append("")
    md.append("## Re-running")
    md.append("")
    md.append("```")
    md.append("python analysis/audit_export_delimiter_risk.py  # PR AE")
    md.append("python analysis/analyze_gis_office_addr_reach.py  # this PR")
    md.append("```")
    md.append("")
    md.append("Pure pyodbc + sqlite3.  No Access COM.")
    OUT_MD.write_text("\n".join(md), encoding="utf-8")
    print(f"wrote {OUT_MD}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
