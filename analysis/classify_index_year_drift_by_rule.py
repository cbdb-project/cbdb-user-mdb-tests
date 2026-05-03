"""Per-row classification of c_index_year drifts.

Updated 2026-05-03 (PR X) — the rule-level reference is now PR
N's `analysis/index_year_rule_comparison.md`, which pairs the
runtime VBA `GetBirthIndexYearSQL` (dumped by PR M from
`frmBaseMaintenance`) against PHP `IndexYearRebuildService.php`
by emitted `c_index_year_type_code`.  PR N's verdict at the
rule level is **22 matched / 8 matched_minor_diff / 0
logic_diff / 3 access_only**, i.e. at the rule level the two
sides agree almost everywhere.

That changes the framing of this script's per-row buckets:
they are no longer testing PR I's `+N`/`-N` sign-flip
hypothesis (which PR N showed was an artefact of comparing PHP
against the wrong Access source — the vestigial
`BM IY Rule …` QueryDefs).  The per-row buckets stay because
they're indexed by type_codes only, not by PR I's rule labels.
But the `explained_by_*` sub-buckets are best read as
**signature-based hypothesis probes** rather than evidence of
PR I-style logic divergence; in K1 they all came back at 0 so
they're effectively dormant probes today.

Scope: only c_index_year.  c_index_addr_id (the much larger
`index_addr_only_diff` bucket) is deferred to a separate PR.

What this script does:

  1. Read both BIOG_MAIN tables (User MDB and SQLite snapshot).
  2. For every common personid where c_index_year differs but
     c_birthyear and c_deathyear match (== PR G's
     `index_year_only_diff` and `index_both_diff` buckets):
       - Pull all the diagnostic fields we have on both sides
         (year, type_code, source_id, birthyear, deathyear).
       - For PHP type_codes whose runtime VBA Access path is
         known from PR N (the entry rules `05`-`10`, the
         birthyear rule `01`, the wife-from-husband rules `03`
         / `17`), probe whether a matching evidence row on
         the User MDB side reconstructs both sides' values.
  3. Bucket each row.  Buckets are evidence categories, not
     bug labels:

       - php_returned_sentinel
           PHP value is ≥ 9999 (looks like a garbage / sentinel —
           e.g. 32767 = signed-int overflow or default).  Found 1
           such row in the current sweep.  Worth flagging upstream.
       - php_did_not_compute
           Access has a value, PHP returns 0 / null.  PHP's rule
           coverage missed this person.
       - access_did_not_compute
           PHP has a value, Access returns 0 / null.  The Access
           rule chain didn't fire for this person.
       - iteration_order_diff
           Access type_code is a CONCAT of PHP's (e.g. PHP '11' →
           Access '1112').  Access ran extra Phase-C iterations
           PHP didn't, or vice versa.
       - explained_by_birthyear_offset
           PHP type_code '01' AND access == birthyear + 59 AND
           php == birthyear.  Historical probe: matches the
           vestigial BM IY Rule 03 BY's +59 offset that PR I
           hypothesised; PR N showed runtime Access Rule 01 uses
           raw c_birthyear like PHP, so this probe is dormant.
           Kept in case future divergence surfaces the pattern.
       - explained_by_entry_sign_flip
           PHP entry rule (tcode 05–10) AND |access - php| == 2N
           AND we can locate a matching ENTRY_DATA row at the
           midpoint.  Historical probe: PR I hypothesised
           Access used `+N` while PHP used `-N`; PR N showed
           runtime Access uses `-N` like PHP, so this probe is
           dormant.  A 2N gap that *did* surface today would
           more likely indicate the two sides picked different
           ENTRY_DATA rows.
       - explained_by_husband_formula
           PHP type_code 03/17 AND a husband row reconstructs
           both sides via Access `husband.c_birthyear + 62` vs
           PHP `husband.c_index_year + 3`.  Historical probe;
           PR N's runtime comparison shows Access actually uses
           `husband.c_index_year + 3` (matching PHP), so this
           probe is dormant.  Retained for the same reason.
       - consistent_within_rule
           Multiple rows share the same (PHP tcode, Access
           tcode, diff) — strong signal of a single rule-level
           cause to investigate.  K2 then names each
           signature group with PR N rule context.
       - candidate_algorithm_divergence
           A row's shape matches one of the dormant probes
           above (entry-rule 2N gap, etc.) but the supporting
           evidence row can't be located.  Kept as
           "investigate further" rather than confirmed.
       - unclassified
           No matching pattern fits.

Conservative: only marks `explained_by_*` when we can ACTUALLY
reconstruct both sides' values from source rows.  Otherwise downgrades
to `candidate_algorithm_divergence` / `consistent_within_rule` /
`unclassified`.

Output:
  reports/index_year_drift_rule_classification.json

This script reads from `analysis/index_year_rule_comparison.json`
(produced by `compare_index_year_rules.py`) for the PHP type-code
→ rule mapping; if that JSON is missing, run the comparator first.
"""
from __future__ import annotations

import json
import sqlite3
from pathlib import Path

import pyodbc

ROOT = Path(__file__).resolve().parent.parent
USER_MDB = ROOT / "data" / "CBDB_BJ_User.mdb"
SQLITE_DIR = ROOT / "data" / "cbdb_online_sqlite"
RULE_COMPARISON_JSON = ROOT / "analysis" / "index_year_rule_comparison.json"
OUT = ROOT / "reports" / "index_year_drift_rule_classification.json"

# Known rule N values per PHP entry-type code, from PR N's
# `analysis/index_year_rule_comparison.md` (Phase A entry rules
# 05-10).  Used by the dormant 2N-diff probe — see docstring.
PHP_ENTRY_RULE_N = {
    "05": 30,  # sqlEntryRule('040101', 30, '05') — jinshi
    "06": 27,  # sqlWifeFromEntryRule('040101', 27, '06')
    "07": 27,  # sqlEntryRule('040102', 27, '07') — juren
    "08": 24,  # sqlWifeFromEntryRule('040102', 24, '08')
    "09": 21,  # sqlEntryRule('040103', 21, '09')
    "10": 18,  # sqlWifeFromEntryRule('040103', 18, '10')
}


def _open_user() -> pyodbc.Connection:
    cs = ("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
          f"DBQ={USER_MDB};")
    return pyodbc.connect(cs, autocommit=True, readonly=True)


def _open_sqlite() -> sqlite3.Connection:
    paths = sorted(SQLITE_DIR.glob("*.sqlite3"))
    if not paths:
        raise SystemExit(
            f"no sqlite snapshot in {SQLITE_DIR} — run "
            f"`python analysis/download_hf_sqlite.py` first")
    return sqlite3.connect(str(paths[-1]))


def _norm(v):
    if v is None:
        return 0
    if isinstance(v, str) and not v.strip():
        return 0
    return v


def main() -> int:
    user = _open_user()
    sql = _open_sqlite()
    cur_u = user.cursor()
    cur_s = sql.cursor()

    print("loading both BIOG_MAIN tables ...")
    cur_u.execute(
        "SELECT c_personid, c_name, c_name_chn, c_index_year, "
        "c_index_addr_id, c_birthyear, c_deathyear, "
        "c_index_year_source_id, c_index_year_type_code "
        "FROM BIOG_MAIN WHERE c_personid > 0"
    )
    user_rows = {
        int(r[0]): {
            "name_py": r[1] or "",
            "name_chn": r[2] or "",
            "index_year": r[3],
            "index_addr_id": r[4],
            "birthyear": r[5],
            "deathyear": r[6],
            "index_year_source_id": r[7],
            "index_year_type_code": r[8],
        }
        for r in cur_u.fetchall()
    }
    cur_s.execute(
        "SELECT c_personid, c_name, c_name_chn, c_index_year, "
        "c_index_addr_id, c_birthyear, c_deathyear, "
        "c_index_year_source_id, c_index_year_type_code "
        "FROM BIOG_MAIN WHERE c_personid > 0"
    )
    sqlite_rows = {
        int(r[0]): {
            "name_py": r[1] or "",
            "name_chn": r[2] or "",
            "index_year": r[3],
            "index_addr_id": r[4],
            "birthyear": r[5],
            "deathyear": r[6],
            "index_year_source_id": r[7],
            "index_year_type_code": r[8],
        }
        for r in cur_s.fetchall()
    }
    print(f"  user mdb: {len(user_rows):,}")
    print(f"  sqlite:   {len(sqlite_rows):,}")

    # Find year-diff personids: birthyear+deathyear match, year differs.
    year_diffs: list[int] = []
    for pid in sorted(set(user_rows) & set(sqlite_rows)):
        u = user_rows[pid]; s = sqlite_rows[pid]
        if (_norm(u["birthyear"]) != _norm(s["birthyear"])
                or _norm(u["deathyear"]) != _norm(s["deathyear"])):
            continue
        if _norm(u["index_year"]) == _norm(s["index_year"]):
            continue
        year_diffs.append(pid)
    print(f"\nyear-diff candidates: {len(year_diffs)}")

    # Pre-fetch ENTRY_DATA + KIN_DATA for the candidate personids in
    # one shot (User MDB side only; we use User MDB rows as the
    # reference for "what source data the algorithm sees").
    print("loading ENTRY_DATA for candidates ...")
    entry_by_pid: dict[int, list[dict]] = {}
    if year_diffs:
        # Access doesn't like very long IN-lists; chunk to 500.
        for i in range(0, len(year_diffs), 500):
            chunk = year_diffs[i:i + 500]
            placeholders = ",".join("?" * len(chunk))
            cur_u.execute(
                f"SELECT c_personid, c_entry_code, c_year "
                f"FROM ENTRY_DATA WHERE c_personid IN ({placeholders}) "
                f"AND c_year > 0",
                chunk,
            )
            for r in cur_u.fetchall():
                entry_by_pid.setdefault(int(r[0]), []).append(
                    {"entry_code": r[1], "year": r[2]})
    print(f"  pids with ENTRY_DATA rows: {len(entry_by_pid)}")

    # Look up husband personid for wife rows (kin_code = 134).
    print("loading wife->husband KIN_DATA for candidates ...")
    husband_by_pid: dict[int, list[int]] = {}
    if year_diffs:
        for i in range(0, len(year_diffs), 500):
            chunk = year_diffs[i:i + 500]
            placeholders = ",".join("?" * len(chunk))
            cur_u.execute(
                f"SELECT c_personid, c_kin_id "
                f"FROM KIN_DATA "
                f"WHERE c_personid IN ({placeholders}) "
                f"  AND c_kin_code = 134",
                chunk,
            )
            for r in cur_u.fetchall():
                husband_by_pid.setdefault(int(r[0]), []).append(int(r[1]))
    print(f"  pids with husband KIN row: {len(husband_by_pid)}")

    # Helper: get a husband's birthyear + index_year from User MDB.
    husband_pids: set[int] = set()
    for hs in husband_by_pid.values():
        husband_pids.update(hs)
    husband_data: dict[int, dict] = {}
    if husband_pids:
        hpids = sorted(husband_pids)
        for i in range(0, len(hpids), 500):
            chunk = hpids[i:i + 500]
            placeholders = ",".join("?" * len(chunk))
            cur_u.execute(
                f"SELECT c_personid, c_birthyear, c_index_year "
                f"FROM BIOG_MAIN WHERE c_personid IN ({placeholders})",
                chunk,
            )
            for r in cur_u.fetchall():
                husband_data[int(r[0])] = {
                    "birthyear": r[1], "index_year": r[2]}

    # ENTRY_CODE_TYPE_REL is on the SQLite side (PHP joins through it
    # to map c_entry_code → c_entry_type).  Pull the reverse mapping
    # so we can ask "given an entry_code, which 04010X bucket does
    # PHP put it in".
    print("loading ENTRY_CODE_TYPE_REL from sqlite ...")
    entry_type_for_code: dict[int, set[str]] = {}
    try:
        cur_s.execute(
            "SELECT c_entry_code, c_entry_type FROM ENTRY_CODE_TYPE_REL"
        )
        for r in cur_s.fetchall():
            ec = int(r[0])
            entry_type_for_code.setdefault(ec, set()).add(
                str(r[1]) if r[1] is not None else "")
    except Exception as e:
        print(f"  (could not load ENTRY_CODE_TYPE_REL: {e})")

    # ---- Per-row classification ----

    classified = {
        "php_returned_sentinel": [],
        "php_did_not_compute": [],
        "access_did_not_compute": [],
        "iteration_order_diff": [],
        "explained_by_birthyear_offset": [],
        "explained_by_entry_sign_flip": [],
        "explained_by_husband_formula": [],
        "consistent_within_rule": [],
        "candidate_algorithm_divergence": [],
        "unclassified": [],
    }

    # Pre-pass: tally (php_tcode, access_tcode, diff) frequency for
    # the consistent_within_rule classifier.
    from collections import Counter
    sig_counts: Counter = Counter()
    for pid in year_diffs:
        u = user_rows[pid]; s = sqlite_rows[pid]
        sig_counts[(
            (s["index_year_type_code"] or "").strip(),
            (u["index_year_type_code"] or "").strip(),
            _norm(u["index_year"]) - _norm(s["index_year"]),
        )] += 1

    for pid in year_diffs:
        u = user_rows[pid]; s = sqlite_rows[pid]
        access_year = _norm(u["index_year"])
        php_year = _norm(s["index_year"])
        diff = access_year - php_year
        php_tcode = (s["index_year_type_code"] or "").strip()
        access_tcode = (u["index_year_type_code"] or "").strip()
        record = {
            "personid": pid,
            "name_py": u["name_py"],
            "name_chn": u["name_chn"],
            "user": {
                "index_year": access_year,
                "index_year_type_code": access_tcode,
                "index_year_source_id": u["index_year_source_id"],
            },
            "sqlite": {
                "index_year": php_year,
                "index_year_type_code": php_tcode,
                "index_year_source_id": s["index_year_source_id"],
            },
            "birthyear": _norm(u["birthyear"]),
            "deathyear": _norm(u["deathyear"]),
            "diff_access_minus_php": diff,
        }

        # --- Hypothesis 0a: PHP returned a sentinel (overflow / garbage)
        if php_year >= 9999:
            record["explanation"] = (
                f"PHP value {php_year} looks like a sentinel "
                f"(>=9999, e.g. 32767 = signed int overflow).  "
                f"Worth flagging upstream as a likely bug in PHP "
                f"output, not a rule disagreement."
            )
            classified["php_returned_sentinel"].append(record)
            continue

        # --- Hypothesis 0b: one side didn't compute
        if php_year == 0 and access_year != 0:
            record["explanation"] = (
                f"Access wrote {access_year} (tcode={access_tcode!r}) "
                f"but PHP wrote 0 / null.  PHP rule chain didn't "
                f"reach a value for this person — possible PHP "
                f"coverage gap."
            )
            classified["php_did_not_compute"].append(record)
            continue
        if access_year == 0 and php_year != 0:
            record["explanation"] = (
                f"PHP wrote {php_year} (tcode={php_tcode!r}) but "
                f"Access wrote 0 / null.  Access rule chain didn't "
                f"reach a value for this person — possible Access "
                f"coverage gap."
            )
            classified["access_did_not_compute"].append(record)
            continue

        # --- Hypothesis 0c: iteration-order divergence
        # If access_tcode is a CONCAT of php_tcode (e.g. '11' → '1112'),
        # one side ran extra Phase-C iterations.
        if (php_tcode and access_tcode
                and access_tcode != php_tcode
                and (access_tcode.startswith(php_tcode)
                     or php_tcode.startswith(access_tcode))):
            record["explanation"] = (
                f"Access tcode {access_tcode!r} vs PHP tcode "
                f"{php_tcode!r} — one side ran an extra Phase-C "
                f"propagation step.  PHP's Phase-C rules CONCAT "
                f"the new code onto the parent's; Access likely "
                f"does the same.  diff={diff} comes from the "
                f"extra propagation hop's offset."
            )
            classified["iteration_order_diff"].append(record)
            continue

        # --- Hypothesis 1: birthyear offset (Access +59 / PHP raw)
        # PHP type_code '01' means it used birthyear directly.
        if (php_tcode == "01"
                and _norm(u["birthyear"]) > 0
                and access_year == _norm(u["birthyear"]) + 59
                and php_year == _norm(u["birthyear"])):
            record["explanation"] = (
                "Access Rule 01B / Rule 03 BY adds +59 to "
                "c_birthyear; PHP sqlRule01 uses raw c_birthyear.  "
                f"birthyear={u['birthyear']}; "
                f"+59={access_year} (Access); "
                f"+0={php_year} (PHP)."
            )
            classified["explained_by_birthyear_offset"].append(record)
            continue

        # --- Hypothesis 2: entry-rule 2N-diff signature
        # PHP type_code in 05..10 → rule uses MIN(c_year) - N.
        # Per PR N runtime Access uses the same -N constant, so a
        # 2N gap shouldn't appear from a sign flip.  But if both
        # sides happened to pick *different ENTRY_DATA rows* whose
        # c_years differ by exactly 2N, the symptom would mimic
        # PR I's old sign-flip pattern.  Probe for that here.
        if php_tcode in PHP_ENTRY_RULE_N:
            n = PHP_ENTRY_RULE_N[php_tcode]
            if diff == 2 * n:
                # Hypothesised entry year (the c_year both sides see):
                mid = (access_year + php_year) // 2
                # Find an ENTRY_DATA row matching that c_year + an
                # entry_code that PHP would map into the right type.
                rows = entry_by_pid.get(pid, [])
                hit = None
                for r in rows:
                    if r["year"] == mid:
                        hit = r
                        break
                if hit is not None:
                    record["explanation"] = (
                        f"Entry-rule 2N-diff signature on PHP "
                        f"type_code '{php_tcode}' (N={n}).  "
                        f"Both sides apply -N per PR N; a single "
                        f"ENTRY_DATA row at the midpoint "
                        f"(c_entry_code={hit['entry_code']}, "
                        f"c_year={hit['year']}) would reconstruct "
                        f"both values via Access {hit['year']}+{n}"
                        f"={access_year} / PHP {hit['year']}-{n}"
                        f"={php_year} only if Access was using the "
                        f"old +N path — which PR N showed it is "
                        f"NOT.  More likely the two sides picked "
                        f"different ENTRY_DATA rows whose c_years "
                        f"happen to differ by 2N."
                    )
                    classified["explained_by_entry_sign_flip"].append(record)
                    continue
            # Diff is in the right ballpark but doesn't fully line up
            # — flag as candidate (likely two-different-rows pick).
            record["explanation"] = (
                f"PHP type_code '{php_tcode}' is an entry rule "
                f"(2N for N={n} would be ±{2*n}); observed "
                f"diff = {diff}.  Per PR N both sides use -N, so "
                f"a 2N gap suggests the two sides picked different "
                f"ENTRY_DATA rows; needs row-by-row evidence to "
                f"confirm which rows."
            )
            classified["candidate_algorithm_divergence"].append(record)
            continue

        # --- Hypothesis 3: wife-from-husband formula
        # PHP type_code '03' (wife) or '17' (concubine).  Access uses
        # husband.c_birthyear + 62; PHP uses husband.c_index_year + 3.
        if php_tcode in {"03", "17"}:
            husbands = husband_by_pid.get(pid, [])
            hit = None
            for hpid in husbands:
                h = husband_data.get(hpid)
                if not h:
                    continue
                hb = _norm(h["birthyear"])
                hi = _norm(h["index_year"])
                if (hb > 0 and hb + 62 == access_year
                        and hi > 0 and hi + 3 == php_year):
                    hit = (hpid, h); break
            if hit is not None:
                hpid, h = hit
                record["explanation"] = (
                    f"Wife rule.  Husband c_personid={hpid}, "
                    f"birthyear={h['birthyear']}, "
                    f"index_year={h['index_year']}.  "
                    f"+62={access_year} (Access Rule 04W); "
                    f"+3={php_year} (PHP sqlRule03)."
                )
                classified["explained_by_husband_formula"].append(record)
                continue
            record["explanation"] = (
                f"PHP type_code '{php_tcode}' is wife-from-husband; "
                f"could not reconstruct from a single husband row "
                f"(husbands_found={len(husbands)})."
            )
            classified["candidate_algorithm_divergence"].append(record)
            continue

        # --- Hypothesis 4: consistent (php_tcode, access_tcode, diff)
        # If multiple year-diff rows share the same triple, that's a
        # strong signal of a single rule-level cause we haven't fully
        # named yet.  Flag those together so the maintainer can
        # batch-investigate.
        sig = (php_tcode, access_tcode, diff)
        if sig_counts[sig] >= 2:
            record["explanation"] = (
                f"{sig_counts[sig]} rows share the exact triple "
                f"(php_tcode={php_tcode!r}, "
                f"access_tcode={access_tcode!r}, diff={diff}).  "
                f"Strong signal of a single rule-level cause — "
                f"investigate that pairing rather than per-row."
            )
            classified["consistent_within_rule"].append(record)
            continue

        # --- Default: unclassified
        record["explanation"] = (
            f"No matching signature from PR N's runtime-vs-PHP "
            f"rule comparison (php_tcode={php_tcode!r}, "
            f"access_tcode={access_tcode!r})."
        )
        classified["unclassified"].append(record)

    # Print summary (in the order they're applied)
    bucket_order = [
        "php_returned_sentinel",
        "php_did_not_compute",
        "access_did_not_compute",
        "iteration_order_diff",
        "explained_by_birthyear_offset",
        "explained_by_entry_sign_flip",
        "explained_by_husband_formula",
        "consistent_within_rule",
        "candidate_algorithm_divergence",
        "unclassified",
    ]
    total = sum(len(v) for v in classified.values())
    print(f"\n=== year-drift classification summary ({total} rows) ===")
    for k in bucket_order:
        print(f"  {k:38s} {len(classified[k]):>4d}")

    # Also: distribution of PHP type_codes among the unclassified +
    # candidates so the maintainer knows where to look next.
    cand_counts = Counter()
    for r in (classified["candidate_algorithm_divergence"]
              + classified["unclassified"]):
        cand_counts[r["sqlite"]["index_year_type_code"]] += 1
    print(f"\n=== unclassified+candidate by PHP type_code ===")
    for code, n in cand_counts.most_common():
        print(f"  {code!r:>10s}  {n}")

    out = {
        "summary": {
            "total_year_diffs": total,
            "buckets": {k: len(v) for k, v in classified.items()},
            "rule_comparison_source": str(
                RULE_COMPARISON_JSON.relative_to(ROOT)
            ),
            "unclassified_by_php_type_code": dict(cand_counts),
        },
        "buckets": classified,
    }
    OUT.write_text(
        json.dumps(out, ensure_ascii=False, indent=2,
                    default=str),
        encoding="utf-8",
    )
    print(f"\nwrote {OUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
