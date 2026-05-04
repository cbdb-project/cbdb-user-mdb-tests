"""LookAtNetworks anchor candidate inventory (PR AQ).

Goal: pick small-but-representative anchor persons for a future
LookAtNetworks CmdRun fixture experiment, so the next Access COM
attempt isn't blind.  Pure pyodbc; no Access COM.

Approach
--------
ASSOC_DATA edges are bidirectional in CBDB convention (the same
relationship typically appears as both (X→Y, code=K) and
(Y→X, code=reverse(K))).  But not every relationship is mirrored,
so we count both directions and union per-person.

KIN_DATA is similar.

For each `c_personid`:
  - assoc_out          = COUNT(*) FROM ASSOC_DATA WHERE c_personid = X
  - assoc_in           = COUNT(*) FROM ASSOC_DATA WHERE c_assoc_id = X
  - assoc_neighbors    = distinct {c_assoc_id ∪ c_personid_of_inbound} \ {X}
  - kin_out            = COUNT(*) FROM KIN_DATA   WHERE c_personid = X
  - kin_in             = COUNT(*) FROM KIN_DATA   WHERE c_kin_id   = X
  - kin_neighbors      = distinct {c_kin_id ∪ c_personid_of_inbound} \ {X}

`assoc_neighbors` is the directly-relevant size for CmdRun's
1-hop expansion.  We then sample per-candidate the `assoc_neighbors`
set to estimate the 1-hop expansion total
(`est_1_hop_assoc_total = sum(assoc_neighbors_count[n] for n in
candidate's assoc_neighbors)`).  This is the rough cost CmdRun
pays when it walks the network at depth 2 with default settings.

Filter / classify
-----------------
We rank candidates by:
  - direct_assoc_neighbor_count in [5, 50]
  - kin_neighbor_count <= 30 (avoid huge kin trees)
  - est_1_hop_assoc_total <= 500
  - has Chinese name (signal of real well-attested record)
  - has c_index_year > 0
  - is in current test_inputs.json's known list (preferred — already
    discovered as fresh by the auto-discoverer)

Flag bands:
  - too_sparse        — <  5 direct assoc neighbors
  - likely_safe_under_120s — 5..20 direct, est_1_hop <= 200
  - medium            — 20..50 direct, est_1_hop 200..500
  - too_large         — > 50 direct OR est_1_hop > 500

Output
------
- analysis/dump/lookatnetworks_anchor_candidates.json (ranked list)
- analysis/lookatnetworks_anchor_candidates.md (top 5–10 with prose)

Pure pyodbc.  No Access COM.  ~30 s on the 658k-person mdb.
"""
from __future__ import annotations

import json
import sys
from collections import defaultdict
from pathlib import Path

import pyodbc

ROOT = Path(__file__).resolve().parent.parent
USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
TEST_INPUTS = ROOT / "analysis" / "dump" / "test_inputs.json"
OUT_JSON = ROOT / "analysis" / "dump" / "lookatnetworks_anchor_candidates.json"
OUT_MD = ROOT / "analysis" / "lookatnetworks_anchor_candidates.md"

# Hard caps.  Keep CmdRun bounded.
MAX_DIRECT_ASSOC_NEIGHBORS = 50      # too_large above this
MAX_KIN_NEIGHBORS = 30
MAX_EST_1_HOP_ASSOC_TOTAL = 500
MIN_DIRECT_ASSOC_NEIGHBORS = 5       # too_sparse below this
SAFE_DIRECT_ASSOC_NEIGHBORS = 20
SAFE_EST_1_HOP_ASSOC_TOTAL = 200

TOP_N_FOR_MD = 10


def _open() -> pyodbc.Connection:
    return pyodbc.connect(
        "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        f"DBQ={USER_MDB};", autocommit=True, readonly=True)


def build_neighbor_maps(cur) -> dict:
    """Single-pass scan of ASSOC_DATA + KIN_DATA → per-person
    neighbor sets (deduped across direction)."""
    print("scanning ASSOC_DATA...")
    assoc_neighbors: dict[int, set[int]] = defaultdict(set)
    cur.execute("SELECT c_personid, c_assoc_id FROM ASSOC_DATA")
    for r in cur.fetchall():
        a = int(r[0]) if r[0] is not None else 0
        b = int(r[1]) if r[1] is not None else 0
        if a <= 0 or b <= 0 or a == b:
            continue
        assoc_neighbors[a].add(b)
        assoc_neighbors[b].add(a)
    print(f"  {len(assoc_neighbors)} persons with ≥1 assoc neighbor")

    print("scanning KIN_DATA...")
    kin_neighbors: dict[int, set[int]] = defaultdict(set)
    cur.execute("SELECT c_personid, c_kin_id FROM KIN_DATA")
    for r in cur.fetchall():
        a = int(r[0]) if r[0] is not None else 0
        b = int(r[1]) if r[1] is not None else 0
        if a <= 0 or b <= 0 or a == b:
            continue
        kin_neighbors[a].add(b)
        kin_neighbors[b].add(a)
    print(f"  {len(kin_neighbors)} persons with ≥1 kin neighbor")

    return {"assoc": assoc_neighbors, "kin": kin_neighbors}


def biog_main_lookup(cur,
                      pids: list[int]) -> dict[int, dict]:
    """Pull name/dynasty/index_year/index_addr for the given pids."""
    if not pids:
        return {}
    out: dict[int, dict] = {}
    # Chunked IN clauses; Access has a parser cap on huge IN lists.
    CHUNK = 800
    for i in range(0, len(pids), CHUNK):
        chunk = pids[i: i + CHUNK]
        in_clause = ",".join(str(p) for p in chunk)
        cur.execute(
            f"SELECT c_personid, c_name, c_name_chn, c_dy, "
            f"c_index_year, c_index_addr_id "
            f"FROM BIOG_MAIN WHERE c_personid IN ({in_clause})"
        )
        for r in cur.fetchall():
            out[int(r[0])] = {
                "name": r[1],
                "name_chn": r[2],
                "c_dy": int(r[3]) if r[3] is not None else None,
                "c_index_year": int(r[4]) if r[4] is not None else None,
                "c_index_addr_id": int(r[5]) if r[5] is not None else None,
            }
    return out


def main() -> int:
    print(f"opening {USER_MDB}")
    conn = _open()
    cur = conn.cursor()

    maps = build_neighbor_maps(cur)
    assoc_n = maps["assoc"]
    kin_n = maps["kin"]

    # Seed candidate set: anyone with 5..50 direct assoc neighbors.
    print("filtering candidates...")
    candidates: list[int] = [
        pid for pid, ns in assoc_n.items()
        if MIN_DIRECT_ASSOC_NEIGHBORS <= len(ns)
            <= MAX_DIRECT_ASSOC_NEIGHBORS
    ]
    print(f"  {len(candidates)} candidates with assoc-degree "
          f"{MIN_DIRECT_ASSOC_NEIGHBORS}..{MAX_DIRECT_ASSOC_NEIGHBORS}")

    # Compute 1-hop assoc total for each candidate.
    rows = []
    for pid in candidates:
        neigh = assoc_n[pid]
        est_1_hop = sum(len(assoc_n.get(n, set())) for n in neigh)
        kn = kin_n.get(pid, set())
        if est_1_hop > MAX_EST_1_HOP_ASSOC_TOTAL:
            flag = "too_large"
        elif len(neigh) > SAFE_DIRECT_ASSOC_NEIGHBORS \
                or est_1_hop > SAFE_EST_1_HOP_ASSOC_TOTAL:
            flag = "medium"
        elif len(kn) > MAX_KIN_NEIGHBORS:
            flag = "too_large_kin"
        else:
            flag = "likely_safe_under_120s"
        rows.append({
            "c_personid": pid,
            "assoc_neighbor_count": len(neigh),
            "kin_neighbor_count": len(kn),
            "est_1_hop_assoc_total": est_1_hop,
            "flag": flag,
        })

    # Pull BIOG_MAIN supplementary info for all candidates +
    # add "well-known" flag from test_inputs if available.
    print("pulling BIOG_MAIN supplementary info...")
    biog = biog_main_lookup(cur, [r["c_personid"] for r in rows])
    for r in rows:
        info = biog.get(r["c_personid"], {})
        r.update(info)

    # Add Zhu Xi (and other known-large) for reference / comparison.
    KNOWN_LARGE = [3767]  # Zhu Xi; verify
    extra_pids = [p for p in KNOWN_LARGE if p not in {r["c_personid"] for r in rows}]
    extra = biog_main_lookup(cur, extra_pids)
    for pid in extra_pids:
        info = extra.get(pid, {})
        rows.append({
            "c_personid": pid,
            "assoc_neighbor_count": len(assoc_n.get(pid, set())),
            "kin_neighbor_count": len(kin_n.get(pid, set())),
            "est_1_hop_assoc_total":
                sum(len(assoc_n.get(n, set())) for n in assoc_n.get(pid, set())),
            "flag": "too_large_known",
            **info,
        })

    # Look up test_inputs.json to flag overlapping pids.
    test_pids: set[int] = set()
    if TEST_INPUTS.exists():
        try:
            d = json.loads(TEST_INPUTS.read_text(encoding="utf-8"))
            # test_inputs is a nested structure; walk it for ints
            def _walk(obj):
                if isinstance(obj, dict):
                    for v in obj.values():
                        yield from _walk(v)
                elif isinstance(obj, list):
                    for v in obj:
                        yield from _walk(v)
                elif isinstance(obj, int):
                    yield obj
            for v in _walk(d):
                if 1 <= v <= 1_000_000:  # plausible personid range
                    test_pids.add(v)
        except Exception:
            pass

    for r in rows:
        r["in_test_inputs"] = r["c_personid"] in test_pids
        # Best-effort name surface
        r["has_name_chn"] = bool(r.get("name_chn"))
        r["has_index_year"] = bool(r.get("c_index_year")) and \
            r.get("c_index_year", 0) > 0

    # Rank: prefer in_test_inputs, then likely_safe, then having
    # name_chn + index_year, then sort by (assoc count near 10).
    def _rank(r: dict) -> tuple:
        bucket = {
            "likely_safe_under_120s": 0,
            "medium": 1,
            "too_large_kin": 2,
            "too_large": 3,
            "too_large_known": 4,
        }.get(r["flag"], 5)
        return (
            bucket,
            0 if r["in_test_inputs"] else 1,
            0 if r["has_name_chn"] else 1,
            0 if r["has_index_year"] else 1,
            abs(r["assoc_neighbor_count"] - 10),  # prefer near 10
            r["c_personid"],
        )

    rows.sort(key=_rank)

    # Pick top N for the human MD.
    safe_top = [r for r in rows
                 if r["flag"] == "likely_safe_under_120s"][:TOP_N_FOR_MD]

    # Counts for summary.
    by_flag: dict[str, int] = defaultdict(int)
    for r in rows:
        by_flag[r["flag"]] += 1

    out = {
        "summary": {
            "n_candidates": len(rows),
            "by_flag": dict(by_flag),
            "n_in_test_inputs": sum(1 for r in rows if r["in_test_inputs"]),
            "thresholds": {
                "MIN_DIRECT_ASSOC_NEIGHBORS": MIN_DIRECT_ASSOC_NEIGHBORS,
                "MAX_DIRECT_ASSOC_NEIGHBORS": MAX_DIRECT_ASSOC_NEIGHBORS,
                "SAFE_DIRECT_ASSOC_NEIGHBORS": SAFE_DIRECT_ASSOC_NEIGHBORS,
                "MAX_KIN_NEIGHBORS": MAX_KIN_NEIGHBORS,
                "MAX_EST_1_HOP_ASSOC_TOTAL": MAX_EST_1_HOP_ASSOC_TOTAL,
                "SAFE_EST_1_HOP_ASSOC_TOTAL": SAFE_EST_1_HOP_ASSOC_TOTAL,
            },
        },
        "rows": rows[:200],   # cap to keep JSON small
        "recommended_top_n": safe_top,
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"\nwrote {OUT_JSON}")
    print(f"  by flag: {dict(by_flag)}")
    print(f"  in test_inputs: {out['summary']['n_in_test_inputs']}")
    print(f"  recommended (top {TOP_N_FOR_MD}):")
    for r in safe_top:
        print(f"    pid={r['c_personid']:>7d} {r.get('name_chn', '')!r}"
              f" assocs={r['assoc_neighbor_count']:>3d}"
              f" kin={r['kin_neighbor_count']:>3d}"
              f" est_1hop={r['est_1_hop_assoc_total']:>4d}")

    # ---- markdown ----
    md = []
    md.append("# LookAtNetworks anchor candidate inventory (PR AQ)")
    md.append("")
    md.append("Pure-pyodbc inventory of candidate `c_personid` "
              "anchors for a future LookAtNetworks CmdRun fixture "
              "experiment.  The current matrix uses Zhu Xi (2 471 "
              "associations) which times out; this list ranks "
              "smaller candidates by likely 1-hop expansion cost.")
    md.append("")
    md.append("Inputs: ASSOC_DATA + KIN_DATA + BIOG_MAIN, "
              f"User MDB at `{USER_MDB.relative_to(ROOT)}`.  No "
              "Access COM.  No Networks-form static knowledge "
              "of the actual depth/loop caps was used — these "
              "are conservative estimates of 1-hop reach.")
    md.append("")
    md.append("## Headline counts")
    md.append("")
    md.append(f"- Candidate persons (assoc-degree "
              f"{MIN_DIRECT_ASSOC_NEIGHBORS}..{MAX_DIRECT_ASSOC_NEIGHBORS}): "
              f"**{len(rows)}**")
    for f in ("likely_safe_under_120s", "medium", "too_large_kin",
              "too_large", "too_large_known"):
        md.append(f"- flag `{f}`: {by_flag.get(f, 0)}")
    md.append(f"- Of all candidates, in current test_inputs.json: "
              f"{out['summary']['n_in_test_inputs']}")
    md.append("")
    md.append("## Top recommended anchors (likely_safe_under_120s)")
    md.append("")
    md.append("Sorted by: in_test_inputs (preferred), has_name_chn, "
              "has_index_year, then closeness to assoc-degree 10.  "
              "Pick 1–3 from this table for the next Access COM "
              "experiment.")
    md.append("")
    md.append("| # | c_personid | name (chn) | name (py) | assocs | kin | est 1-hop | dyn | index_year | in_test_inputs |")
    md.append("|---:|---:|---|---|---:|---:|---:|---:|---:|---|")
    for i, r in enumerate(safe_top, start=1):
        md.append(
            f"| {i} | {r['c_personid']} "
            f"| {r.get('name_chn') or '—'} "
            f"| {r.get('name') or '—'} "
            f"| {r['assoc_neighbor_count']} "
            f"| {r['kin_neighbor_count']} "
            f"| {r['est_1_hop_assoc_total']} "
            f"| {r.get('c_dy') or '—'} "
            f"| {r.get('c_index_year') or '—'} "
            f"| {'yes' if r['in_test_inputs'] else 'no'} |"
        )
    md.append("")
    md.append("## Reference: known-large anchors")
    md.append("")
    too_large = [r for r in rows
                 if r["flag"] in ("too_large_known", "too_large")][:10]
    if too_large:
        md.append("| c_personid | name (chn) | assocs | kin | est 1-hop | flag |")
        md.append("|---:|---|---:|---:|---:|---|")
        for r in too_large:
            md.append(
                f"| {r['c_personid']} | {r.get('name_chn') or '—'} "
                f"| {r['assoc_neighbor_count']} | {r['kin_neighbor_count']} "
                f"| {r['est_1_hop_assoc_total']} | `{r['flag']}` |"
            )
    md.append("")
    md.append("## Caveats")
    md.append("")
    md.append("- ASSOC_DATA edges are deduplicated bidirectionally "
              "by the script; the actual LookAtNetworks CmdRun may "
              "walk them differently (with kin / depth filters).  "
              "These counts are upper bounds for the 1-hop reach.")
    md.append("- `est_1_hop_assoc_total` sums the degrees of every "
              "1-hop neighbor — this is a worst-case for a "
              "depth-2 walk that doesn't dedupe second-hop "
              "neighbors.  CmdRun does dedupe (per "
              "`Form_LookAtNetworks.vb`'s ZZ_SCRATCH_PEOPLE write "
              "pattern), so true cost will be lower.")
    md.append("- We did NOT filter by `gMaxFilterTotal=29` etc. (the "
              "default checkbox state from Form_Open).  The default "
              "Networks UI applies association-type filters that "
              "would shrink the candidate set further.")
    md.append("- The 1-hop estimator assumes uniform degree; high-"
              "variance neighbours (e.g. one neighbour is a Zhu-Xi-"
              "scale hub) can blow past `est_1_hop_assoc_total`.  "
              "When picking from the table, consider whether the "
              "anchor's named neighbors are likely hubs.")
    md.append("")
    md.append("## How to use the recommended list")
    md.append("")
    md.append("Pick 1–3 anchors from the recommended table.  For "
              "each, the next Access COM experiment should:")
    md.append("")
    md.append("1. Open LookAtNetworks (Form_Open is fine per PR AA).")
    md.append("2. Set the picker to that c_personid via "
              "`set_picker_codes(\"ZZ_SCRATCH_IMPORT_PEOPLE\", [pid])`.")
    md.append("3. Set `gMaxNodes` and `gMaxLoops` to small values "
              "(e.g. 20 nodes / 1 loop) before firing CmdRun.")
    md.append("4. Use `Form_Timer` trigger with a 120 s budget.")
    md.append("5. If CmdRun completes, capture the resulting "
              "ZZ_SCRATCH_PEOPLE / ZZ_SOCIAL_NETWORK row counts.")
    md.append("")
    md.append("Out-of-scope for this PR.  See `analysis/lookatnetworks_"
              "form_open_hang.md` for what we already know about the "
              "form's behaviour under COM.")
    OUT_MD.write_text("\n".join(md), encoding="utf-8")
    print(f"wrote {OUT_MD}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
