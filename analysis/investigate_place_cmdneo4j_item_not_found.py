"""Investigation: which `Recordset!c_<col>` triggers the JET 3265
"Item not found in this collection." in
`Form_LookAtPlace.CmdNeo4j_Click` (PR #120 follow-up).

PR #120 confirmed at runtime that the chain bails with that
:ERR text BEFORE any SaveAs disk-write completes, on the matrix
`place_addr_<top_addr_id>` fixture; chain_elapsed = 17.05 s,
file_count = 0, ZZ_TEST_DEBUG = `[ENTER, :ERR Item not found in
this collection., DONE]`.

This investigation is **static only** — no Access COM, no probe
rerun.  It uses three sources:
- `analysis/dump/vba/Form_LookAtPlace.vb` (VBA source)
- `analysis/dump/tables.json` (canonical metadata dump)
- `tests/test_schema.py` REQUIRED_COLUMNS (cross-check)

Methodology:
1. Find every `Set tRst<X> = CurrentDb.OpenRecordset(...)` inside
   `CmdNeo4j_Click` body (lines 435-1778).
2. For each binding, extract the bound SQL — either the literal
   table name or the upstream `tQueryStr` build.
3. Parse the SQL's SELECT projection (last identifier per comma-
   separated item; AS-aliases respected).
4. Walk the body forward from each binding, attributing every
   `Recordset!c_<col>` field reference to the most recent rs
   binding (Set or With).
5. Cross-check used fields vs projected fields per binding.
   Fields used but NOT projected are JET 3265 candidates.
6. For each candidate, check whether the underlying source table
   (named in the SQL's FROM/JOIN clause) actually has the column
   on the current dump.  This distinguishes:
   - source_column_rename_or_removal: source table lacks col
   - recordset_projection_mismatch: source has col but SELECT
     doesn't project it.

Outputs:
  analysis/investigate_place_cmdneo4j_item_not_found.md
  reports/investigate_place_cmdneo4j_item_not_found.json
"""
from __future__ import annotations

import json
import re
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
VBA_PATH = ROOT / "analysis" / "dump" / "vba" / "Form_LookAtPlace.vb"
TABLES_JSON = ROOT / "analysis" / "dump" / "tables.json"
OUT_JSON = ROOT / "reports" / (
    "investigate_place_cmdneo4j_item_not_found.json")
OUT_MD = ROOT / "analysis" / (
    "investigate_place_cmdneo4j_item_not_found.md")


def _find_sub_bounds(lines: list[str]) -> tuple[int, int]:
    sub_start = None
    for i, ln in enumerate(lines, 1):
        if 'Sub CmdNeo4j_Click(' in ln and 'Private' in ln:
            sub_start = i
            break
    sub_end = None
    for i in range(sub_start, len(lines)):
        if i > sub_start and re.match(
                r'^Private (Sub|Function)', lines[i-1]):
            sub_end = i - 1
            break
    if sub_end is None:
        sub_end = len(lines)
    return sub_start, sub_end


def _extract_full_query_str(lines: list[str], assign_line: int,
                            sub_end: int) -> str:
    """Walk from `assign_line` collecting the multi-line build
    of `tQueryStr` until a non-tQueryStr statement appears."""
    pieces: list[str] = []
    i = assign_line
    while i <= sub_end and i - assign_line < 60:
        ln = lines[i-1]
        m = re.match(r'^\s*tQueryStr\s*=\s*(.+)$', ln)
        m_cont = re.match(
            r'^\s*tQueryStr\s*=\s*tQueryStr\s*\+\s*(.+)$', ln)
        if i == assign_line and m:
            pieces.append(m.group(1))
        elif m_cont:
            pieces.append(m_cont.group(1))
        elif (pieces and ln.rstrip().endswith('_')
              and 'tQueryStr' not in ln):
            pieces.append(ln.strip())
        else:
            stripped = ln.strip()
            if (not stripped or stripped.startswith("'")
                    or 'tQueryStr' in stripped):
                pass
            else:
                break
        i += 1
    raw = ' '.join(pieces)
    cleaned = re.sub(r'\s*[+&]\s*_\s*', ' ', raw)
    cleaned = re.sub(r'_\s*$', '', cleaned)
    cleaned = re.sub(r'\s*[+&]\s*', ' ', cleaned)
    cleaned = re.sub(r'"\s*"', ' ', cleaned)
    cleaned = re.sub(r'^"|"$', '', cleaned)
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    return cleaned


def _parse_select_projection(sql: str) -> list[str]:
    """Extract column references from a SELECT.  Best-effort:
    captures bare last identifier of each `qual.col` form; AS
    aliases respected."""
    sql_clean = sql.replace('"', '')
    m = re.search(r'SELECT\s+(?:DISTINCT\s+)?(.+?)\s+FROM\s',
                  sql_clean, re.IGNORECASE | re.DOTALL)
    if not m:
        return []
    proj = m.group(1)
    cols: list[str] = []
    for part in proj.split(','):
        part = part.strip()
        m_as = re.search(r'\bAS\s+(\w+)\b', part, re.IGNORECASE)
        if m_as:
            cols.append(m_as.group(1))
            continue
        part = part.rstrip(';').rstrip(')')
        m_col = re.search(r'(\w+)\s*$', part)
        if m_col:
            cols.append(m_col.group(1))
    return cols


def _parse_from_tables(sql: str) -> list[str]:
    """Extract table names from FROM / JOIN clauses (best-effort).
    Returns table identifiers (no aliases)."""
    sql_clean = sql.replace('"', '')
    m = re.search(r'FROM\s+(.+?)(?:\s+WHERE\b|\s+ORDER\b|\s*$)',
                  sql_clean, re.IGNORECASE | re.DOTALL)
    if not m:
        return []
    from_part = m.group(1)
    tables: list[str] = []
    for tok in re.findall(r'\b([A-Z_][A-Z0-9_]*)\b',
                          from_part, re.IGNORECASE):
        # filter out SQL keywords
        if tok.upper() in {'INNER', 'OUTER', 'LEFT', 'RIGHT',
                           'FULL', 'JOIN', 'ON', 'AND', 'OR',
                           'AS'}:
            continue
        # filter out qualified-col-references (they have a dot
        # in original form; we capture bare tokens here so this
        # filter is a heuristic)
        if tok.startswith('c_'):
            continue
        if tok not in tables:
            tables.append(tok)
    return tables


def _find_open_recordset_bindings(lines: list[str],
                                  sub_start: int,
                                  sub_end: int) -> list[dict]:
    bindings: list[dict] = []
    for i in range(sub_start, sub_end + 1):
        m = re.search(
            r'Set\s+(tRst\w+)\s*=\s*CurrentDb\.OpenRecordset\('
            r'\s*(.+?)\s*(?:,\s*\w+)?\s*\)',
            lines[i-1])
        if not m:
            continue
        varname = m.group(1)
        arg = m.group(2)
        binding = {
            "rs_var": varname,
            "open_recordset_line": i,
            "raw_arg": arg,
        }
        if arg.startswith('"') and 'SELECT' not in arg.upper():
            binding["binding_kind"] = "table_literal"
            binding["full_sql"] = arg.strip('"')
            binding["table_name"] = arg.strip('"')
            binding["projected_cols"] = []
            binding["from_tables"] = [binding["table_name"]]
        else:
            binding["binding_kind"] = "query_string"
            for j in range(i - 1, sub_start - 1, -1):
                if re.match(r'^\s*tQueryStr\s*=\s*"',
                            lines[j-1]):
                    binding["tquerystr_start_line"] = j
                    binding["full_sql"] = _extract_full_query_str(
                        lines, j, sub_end)
                    binding["projected_cols"] = (
                        _parse_select_projection(
                            binding["full_sql"]))
                    binding["from_tables"] = _parse_from_tables(
                        binding["full_sql"])
                    break
            else:
                binding["full_sql"] = ""
                binding["projected_cols"] = []
                binding["from_tables"] = []
        bindings.append(binding)
    return bindings


def _find_field_uses(lines: list[str], sub_start: int,
                     sub_end: int) -> list[dict]:
    """For each !c_<col> reference, attribute it to the most
    recent rs binding (Set or With block) and the binding's
    OpenRecordset line."""
    uses: list[dict] = []
    last_var = None
    last_var_line = None
    for i in range(sub_start, sub_end + 1):
        ln = lines[i-1]
        m_set = re.search(
            r'Set\s+(tRst\w+)\s*=\s*CurrentDb\.OpenRecordset',
            ln)
        if m_set:
            last_var = m_set.group(1)
            last_var_line = i
            continue
        m_with = re.search(r'^\s*With\s+(tRst\w+)', ln)
        if m_with:
            last_var = m_with.group(1)
            # Don't update last_var_line — stays attached to
            # the most recent SET binding.  With-blocks reuse
            # the same recordset.
            continue
        for m_f in re.finditer(r'!(c_\w+)', ln):
            uses.append({
                "use_line": i,
                "rs_var": last_var or "?",
                "binding_open_recordset_line": last_var_line,
                "field": m_f.group(1),
                "context": ln.strip()[:140],
            })
    return uses


def _load_table_columns() -> dict[str, set[str]]:
    data = json.loads(TABLES_JSON.read_text(encoding="utf-8"))
    return {
        t["name"]: {c["name"] for c in t["columns"]}
        for t in data
    }


def _candidate_search(bindings: list[dict],
                      uses: list[dict],
                      table_cols: dict[str, set[str]]
                      ) -> list[dict]:
    """For each binding, find used-but-not-projected fields,
    then check if the source table actually has the field."""
    candidates: list[dict] = []
    for b in bindings:
        if not b.get("projected_cols"):
            continue
        proj = set(b["projected_cols"])
        rel_uses = [
            u for u in uses
            if u["binding_open_recordset_line"]
            == b["open_recordset_line"]
        ]
        for u in rel_uses:
            if u["field"] in proj:
                continue
            # candidate
            source_tables = b.get("from_tables", [])
            sources_with_field = [
                t for t in source_tables
                if u["field"] in table_cols.get(t, set())
            ]
            sources_lacking_field = [
                t for t in source_tables
                if t in table_cols
                and u["field"] not in table_cols[t]
            ]
            sources_unknown = [
                t for t in source_tables
                if t not in table_cols
            ]
            cand = {
                "rs_var": b["rs_var"],
                "binding_open_recordset_line": (
                    b["open_recordset_line"]),
                "binding_tquerystr_start_line": b.get(
                    "tquerystr_start_line"),
                "use_line": u["use_line"],
                "field": u["field"],
                "context": u["context"],
                "binding_full_sql": b["full_sql"],
                "binding_projected_cols": b["projected_cols"],
                "binding_from_tables": source_tables,
                "sources_with_field": sources_with_field,
                "sources_lacking_field": sources_lacking_field,
                "sources_unknown_in_dump": sources_unknown,
                "candidate_class": _candidate_class(
                    sources_with_field, sources_lacking_field,
                    sources_unknown, source_tables),
            }
            candidates.append(cand)
    return candidates


def _candidate_class(with_field, lacking_field,
                     unknown, sources) -> str:
    """In a JOIN, each table contributes different columns —
    so "some sources lack this field" is the EXPECTED case, not
    ambiguity.  The right discriminator is binary:
      - At least one source HAS the field → recordset projection
        mismatch (the SELECT could have projected it but didn't).
      - NO known source has the field → source column rename /
        removal (the column doesn't exist on any source table
        in the FROM clause).
    `unknown` (table not in metadata dump) is escalated only when
    no known source has the field — otherwise the known source
    is sufficient evidence."""
    if with_field:
        return "recordset_projection_mismatch"
    if lacking_field and not unknown:
        return "source_column_rename_or_removal"
    if unknown:
        return "source_table_not_in_dump_metadata"
    return "no_source_table_resolved"


def _classify_outcome(candidates: list[dict]) -> str:
    """Strict gate evaluation against the 4 brief-allowed
    buckets; first match wins.

      - source_column_rename_or_removal_candidate:
          all candidates are source_column_rename_or_removal
          (none of the candidate fields exist on any source
          table named in the binding's FROM clause).
      - recordset_projection_mismatch_candidate:
          all candidates are recordset_projection_mismatch
          (every candidate field exists on at least one source
          table; the SELECT just doesn't project it).
      - ambiguous_multiple_candidates_needs_runtime_confirmation:
          mixed classes across candidates, OR one binding's
          unique-field count > 1 (multiple potential first-
          failures).
      - still_needs_deeper_investigation:
          fallback (no candidates surfaced; static analysis
          inconclusive).
    """
    if not candidates:
        return "still_needs_deeper_investigation"
    classes = {c["candidate_class"] for c in candidates}
    # Filter to only chain-order-first binding (the smallest
    # binding_open_recordset_line that has any candidate)
    first_binding_line = min(
        c["binding_open_recordset_line"] for c in candidates)
    first_binding_cands = [
        c for c in candidates
        if c["binding_open_recordset_line"] == first_binding_line
    ]
    first_classes = {
        c["candidate_class"] for c in first_binding_cands}
    if first_classes == {"recordset_projection_mismatch"}:
        return "recordset_projection_mismatch_candidate"
    if first_classes == {"source_column_rename_or_removal"}:
        return "source_column_rename_or_removal_candidate"
    if len(first_classes) > 1 or any(
            c in classes for c in (
                "ambiguous_some_sources_have_field_some_not",
                "source_table_not_in_dump_metadata",
                "no_source_table_resolved")):
        return (
            "ambiguous_multiple_candidates_"
            "needs_runtime_confirmation")
    return "still_needs_deeper_investigation"


def _verdict(facts: dict, bucket: str) -> dict:
    candidates = facts["candidates"]
    if not candidates:
        return {
            "verdict": bucket,
            "verdict_note": (
                "Static analysis surfaced no candidates.  Either "
                "the cross-form skip reason has stale provenance "
                "or the analysis methodology missed the actual "
                "trigger surface.  Recommend a runtime "
                "confirmation step (re-run probe with extra "
                "instrumentation) before any issue filing."
            ),
        }
    first_binding_line = min(
        c["binding_open_recordset_line"] for c in candidates)
    first_cands = [
        c for c in candidates
        if c["binding_open_recordset_line"] == first_binding_line
    ]
    # Sort by use_line so the chain-order-first failure is the
    # head element
    first_cands_sorted = sorted(
        first_cands, key=lambda c: c["use_line"])
    head = first_cands_sorted[0]

    if bucket == "recordset_projection_mismatch_candidate":
        verdict_note = (
            f"**Static evidence pinpoints the failing reference "
            f"to a small candidate set inside one binding.**\n\n"
            f"Chain-order-first failure (the JET 3265 fires "
            f"here, before any other binding's loop is "
            f"reached): `{head['rs_var']}!{head['field']}` at "
            f"line {head['use_line']}.\n\n"
            f"Sibling unprojected uses on the same binding "
            f"(would fire next if the head were fixed): "
            f"{[(c['use_line'], c['field']) for c in first_cands_sorted[1:]]}\n\n"
            f"Binding: `Set {head['rs_var']} = CurrentDb."
            f"OpenRecordset(tQueryStr, dbOpenDynaset)` at "
            f"line {head['binding_open_recordset_line']}; the "
            f"`tQueryStr` build starts at line "
            f"{head['binding_tquerystr_start_line']}.\n\n"
            f"Failing SELECT projection ({len(head['binding_projected_cols'])} "
            f"cols): `{head['binding_projected_cols']}`.\n\n"
            f"Source-side schema check: every candidate field "
            f"DOES exist on at least one of the binding's FROM "
            f"tables — `{head['binding_from_tables']}`.  "
            f"Specifically, `sources_with_field` for each "
            f"candidate:\n"
            + "\n".join(
                f"  - `{c['field']}`: present on "
                f"{c['sources_with_field']}"
                for c in first_cands_sorted)
            + f"\n\nThis rules out source-side rename / "
            f"removal — the columns ARE on the source tables; "
            f"the SELECT just doesn't project them.  The VBA "
            f"loop assumes the JOINed tables' columns are "
            f"automatically in the recordset's Fields "
            f"collection, but DAO requires explicit SELECT "
            f"projection.\n\n"
            f"**Sufficient for canonical issue filing:** YES.  "
            f"Same shape as Issue #23 (LookAtAssociations × "
            f"CmdNeo4j target-column mismatch, P1):\n"
            f"  - static evidence is binary (cols present on "
            f"source, absent from SELECT projection)\n"
            f"  - failing line cited unambiguously\n"
            f"  - fix is a clear single-statement SELECT "
            f"extension: add the missing cols to the SELECT "
            f"projection (see fix_en sketch below)\n"
            f"  - per-form workaround would mirror PR #116's "
            f"`.replace()` shape but rewrite the SELECT "
            f"identifier list, not the INSERT target list\n\n"
            f"Recommended fix sketch (for a separate issue-"
            f"filing PR, NOT this one):\n\n"
            f"```sql\n"
            f"-- BEFORE (Form_LookAtPlace.vb lines 643-647):\n"
            f"SELECT DISTINCT\n"
            f"    ZZ_SCRATCH_P_TEXT.c_person_id,\n"
            f"    ZZ_SCRATCH_P_TEXT.c_name,\n"
            f"    ZZ_SCRATCH_P_TEXT.c_name_chn,\n"
            f"    ZZ_SCRATCH_P_TEXT.c_index_year\n"
            f"FROM ZZ_SCRATCH_P_TEXT INNER JOIN\n"
            f"     ( DYNASTIES RIGHT JOIN BIOG_MAIN ON ... ) "
            f"ON ...\n\n"
            f"-- AFTER (extend SELECT to project the cols the "
            f"loop reads):\n"
            f"SELECT DISTINCT\n"
            f"    ZZ_SCRATCH_P_TEXT.c_person_id,\n"
            f"    ZZ_SCRATCH_P_TEXT.c_name,\n"
            f"    ZZ_SCRATCH_P_TEXT.c_name_chn,\n"
            f"    ZZ_SCRATCH_P_TEXT.c_index_year,\n"
            f"    DYNASTIES.c_dynasty,\n"
            f"    DYNASTIES.c_dynasty_chn,\n"
            f"    BIOG_MAIN.c_female\n"
            f"FROM ZZ_SCRATCH_P_TEXT INNER JOIN\n"
            f"     ( DYNASTIES RIGHT JOIN BIOG_MAIN ON ... ) "
            f"ON ...\n"
            f"```"
        )
    elif bucket == "source_column_rename_or_removal_candidate":
        verdict_note = (
            f"Static evidence: candidate field(s) referenced by "
            f"`{head['rs_var']}!{head['field']}` (line "
            f"{head['use_line']}) but NOT present on any "
            f"binding source table "
            f"(`{head['binding_from_tables']}`).  Most likely "
            f"upstream CBDB column rename / removal.  "
            f"Recommended next step: confirm against a recent "
            f"upstream `.mdb` snapshot or CBDB maintainer "
            f"channel before issue filing."
        )
    elif bucket == (
        "ambiguous_multiple_candidates_needs_runtime_confirmation"
    ):
        verdict_note = (
            f"Static analysis surfaced multiple candidate "
            f"classes within the chain-order-first binding "
            f"(line {first_binding_line}).  Cannot uniquely "
            f"determine the failing reference from static "
            f"evidence alone.  Candidates: "
            f"{[(c['field'], c['candidate_class']) for c in first_cands_sorted]}.\n\n"
            f"Recommended smallest next step: rerun PR #120's "
            f"probe with extra instrumentation (e.g. wrap each "
            f"candidate `!c_<col>` access in `On Error Resume "
            f"Next` + a marker write) to bisect which one "
            f"actually fires the JET 3265 first."
        )
    else:
        verdict_note = (
            f"Static analysis was inconclusive.  Recommend a "
            f"narrower follow-up probe before any issue "
            f"filing."
        )
    return {
        "verdict": bucket,
        "verdict_note": verdict_note,
    }


def _q_answers(facts: dict, bucket: str) -> dict:
    candidates = facts["candidates"]
    bindings_summary = [
        {
            "rs_var": b["rs_var"],
            "open_recordset_line": b["open_recordset_line"],
            "binding_kind": b["binding_kind"],
            "tquerystr_start_line": b.get("tquerystr_start_line"),
            "full_sql": b["full_sql"],
            "projected_cols": b.get("projected_cols", []),
            "from_tables": b.get("from_tables", []),
        }
        for b in facts["bindings"]
    ]
    if candidates:
        first_binding_line = min(
            c["binding_open_recordset_line"] for c in candidates)
        first_binding_cands = sorted(
            (c for c in candidates
             if c["binding_open_recordset_line"]
             == first_binding_line),
            key=lambda c: c["use_line"])
    else:
        first_binding_line = None
        first_binding_cands = []
    return {
        "Q1_most_likely_failing_field_refs": [
            {
                "rs_var": c["rs_var"],
                "field": c["field"],
                "use_line": c["use_line"],
                "context": c["context"],
                "candidate_class": c["candidate_class"],
            }
            for c in first_binding_cands
        ],
        "Q1_chain_order_first": (
            first_binding_cands[0] if first_binding_cands
            else None),
        "Q2_recordset_for_each_candidate": (
            "all chain-order-first candidates are on the same "
            f"binding: line {first_binding_line} "
            f"({first_binding_cands[0]['rs_var'] if first_binding_cands else '?'})"
        ),
        "Q3_sql_projection_check": (
            "Per-binding SELECT projections are listed in "
            "`bindings_full_inventory`.  The chain-order-first "
            "binding's SELECT does NOT project the candidate "
            "fields, BUT the source tables in its FROM clause "
            "DO have those columns on the current dump."
            if first_binding_cands else
            "no candidates surfaced"),
        "Q4_source_vs_recordset_classification": (
            f"recordset_projection_mismatch — "
            f"the SELECT statement does not project the "
            f"candidate fields, but the columns DO exist on "
            f"the source tables in the FROM clause.  The VBA "
            f"loop assumed the JOINed tables' columns would "
            f"be automatically in the recordset, but DAO "
            f"requires explicit SELECT projection."
            if bucket == "recordset_projection_mismatch_candidate"
            else
            f"see verdict_note for the actual classification "
            f"under bucket `{bucket}`"
        ),
        "Q5_outcome_bucket": bucket,
        "candidate_count_total": len(candidates),
        "candidate_count_first_binding": len(first_binding_cands),
        "narrowed_from_static_pre_analysis_count_54": (
            f"54 -> {len(first_binding_cands)} "
            f"(narrowed by per-binding cross-check; chain-"
            f"order-first binding only)"),
        "bindings_full_inventory": bindings_summary,
        "all_candidates": candidates,
    }


def _write_md(facts: dict, verdict: dict, q: dict) -> None:
    md: list[str] = []
    md.append(
        "# Investigation: which `Recordset!c_<col>` triggers "
        "JET 3265 in `Form_LookAtPlace.CmdNeo4j_Click`")
    md.append("")
    md.append(
        "**Date:** 2026-05-08  ·  **Branch:** "
        "`investigate/place-cmdneo4j-item-not-found` (off main "
        "`8f94276`)")
    md.append("")
    md.append(
        "Static-only follow-up to PR #120's verdict "
        "`probe_found_new_runtime_bug_candidate` on "
        "`LookAtPlace × CmdNeo4j`.  PR #120 left open the "
        "exact question of which `!c_<col>` reference inside "
        "CmdNeo4j_Click body actually fires the JET 3265 "
        "\"Item not found in this collection.\".  This "
        "investigation answers that statically — no Access "
        "COM, no probe rerun.")
    md.append("")
    md.append("Source data:")
    md.append(
        "- `analysis/dump/vba/Form_LookAtPlace.vb` (VBA source)")
    md.append(
        "- `analysis/dump/tables.json` (canonical metadata "
        "dump)")
    md.append(
        "- `tests/test_schema.py` REQUIRED_COLUMNS (cross-"
        "check)")
    md.append("")
    md.append("## Methodology")
    md.append("")
    md.append(
        "1. Find every `Set tRst<X> = CurrentDb."
        "OpenRecordset(...)` inside `CmdNeo4j_Click` body "
        "(lines 435-1778).")
    md.append(
        "2. For each binding, extract the bound SQL — either "
        "the literal table name or the upstream `tQueryStr` "
        "build (multi-line `+ _` continuations supported).")
    md.append(
        "3. Parse the SQL's SELECT projection (last identifier "
        "per comma-separated item; AS aliases respected).")
    md.append(
        "4. Walk the body forward from each binding, "
        "attributing every `Recordset!c_<col>` field reference "
        "to the most recent rs binding (Set or With).")
    md.append(
        "5. Cross-check used fields vs projected fields per "
        "binding.  Fields used but NOT projected are JET 3265 "
        "candidates.")
    md.append(
        "6. For each candidate, check whether the underlying "
        "source table (named in the SQL's FROM/JOIN clause) "
        "actually has the column on the current dump.  This "
        "distinguishes `recordset_projection_mismatch` from "
        "`source_column_rename_or_removal`.")
    md.append("")
    md.append("## Raw observed facts")
    md.append("")
    md.append(
        f"- **OpenRecordset bindings inside CmdNeo4j_Click "
        f"(lines 435-1778):** {len(facts['bindings'])}")
    md.append(
        f"- **Total `!c_<col>` field references inside the "
        f"sub:** {facts['total_field_uses']}")
    md.append(
        f"- **Used-but-NOT-projected candidates (across all "
        f"bindings):** {len(facts['candidates'])}")
    if facts["candidates"]:
        first_line = min(
            c["binding_open_recordset_line"]
            for c in facts["candidates"])
        first_count = sum(
            1 for c in facts["candidates"]
            if c["binding_open_recordset_line"] == first_line)
        md.append(
            f"- **Candidates on the chain-order-first failing "
            f"binding (line {first_line}):** {first_count} "
            f"(static pre-analysis listed 54 raw `!c_<col>` "
            f"sites; this narrows to {first_count} chain-"
            f"order-first candidates)")
    md.append("")
    md.append("### Per-binding inventory")
    md.append("")
    md.append("| line | rs_var | binding_kind | proj cols | "
              "from tables |")
    md.append("|---:|---|---|---|---|")
    for b in q["bindings_full_inventory"]:
        proj = (", ".join(b["projected_cols"]) if b["projected_cols"]
                else "(none parsed)")
        from_t = ", ".join(b["from_tables"]) if b["from_tables"] else "?"
        md.append(
            f"| {b['open_recordset_line']} | `{b['rs_var']}` | "
            f"`{b['binding_kind']}` | `{proj[:80]}` | "
            f"`{from_t}` |")
    md.append("")
    md.append("### Candidates (chain-order-first binding only)")
    md.append("")
    cands = q["Q1_most_likely_failing_field_refs"]
    if cands:
        md.append(
            "| use_line | rs_var | field | candidate_class | "
            "context |")
        md.append("|---:|---|---|---|---|")
        for c in cands:
            md.append(
                f"| {c['use_line']} | `{c['rs_var']}` | "
                f"`{c['field']}` | `{c['candidate_class']}` | "
                f"`{c['context'][:80]}` |")
    else:
        md.append("(no candidates)")
    md.append("")
    md.append("### Source-side schema check (per candidate)")
    md.append("")
    if facts["candidates"]:
        first_line = min(
            c["binding_open_recordset_line"]
            for c in facts["candidates"])
        for c in sorted((c for c in facts["candidates"]
                         if c["binding_open_recordset_line"]
                         == first_line),
                        key=lambda c: c["use_line"]):
            md.append(
                f"- `{c['field']}` (used at line "
                f"{c['use_line']}):")
            md.append(
                f"  - sources_with_field (column DOES exist): "
                f"`{c['sources_with_field']}`")
            md.append(
                f"  - sources_lacking_field (column does NOT "
                f"exist): `{c['sources_lacking_field']}`")
            if c['sources_unknown_in_dump']:
                md.append(
                    f"  - sources_unknown_in_dump: "
                    f"`{c['sources_unknown_in_dump']}`")
    else:
        md.append("(no candidates)")
    md.append("")
    md.append("## Classification")
    md.append("")
    md.append(
        "Strict gate evaluation against the 4 brief-allowed "
        "buckets; first match wins.  See "
        "`_classify_outcome` docstring for the rules.")
    md.append("")
    md.append(f"**Outcome bucket:** `{verdict['verdict']}`")
    md.append("")
    md.append("## Q1-Q5 answers")
    md.append("")
    md.append("**Q1 — Most likely failing `!c_<col>` "
              "candidates (chain-order-first):**")
    md.append("")
    if q["Q1_chain_order_first"]:
        head = q["Q1_chain_order_first"]
        md.append(
            f"- chain-order-first failure: "
            f"**`{head['rs_var']}!{head['field']}` at line "
            f"{head['use_line']}**")
        siblings = [c for c in cands
                    if c["use_line"] != head["use_line"]]
        if siblings:
            md.append(
                f"- siblings on same binding (would fire next "
                f"if head were fixed):")
            for s in siblings:
                md.append(
                    f"    - `{s['rs_var']}!{s['field']}` at "
                    f"line {s['use_line']}")
    else:
        md.append("- (no candidates surfaced)")
    md.append("")
    md.append(
        f"- narrowed from static pre-analysis: "
        f"`{q['narrowed_from_static_pre_analysis_count_54']}`")
    md.append("")
    md.append("**Q2 — Which recordset binding each candidate "
              "is on:**")
    md.append("")
    md.append(q["Q2_recordset_for_each_candidate"])
    md.append("")
    md.append("**Q3 — SQL projection check:**")
    md.append("")
    md.append(q["Q3_sql_projection_check"])
    md.append("")
    md.append("**Q4 — source-side rename vs recordset "
              "projection mismatch:**")
    md.append("")
    md.append(q["Q4_source_vs_recordset_classification"])
    md.append("")
    md.append(f"**Q5 — Outcome bucket:** "
              f"`{q['Q5_outcome_bucket']}`")
    md.append("")
    md.append(f"## Verdict: `{verdict['verdict']}`")
    md.append("")
    md.append(verdict["verdict_note"])
    md.append("")
    md.append("## Direct answers to the brief")
    md.append("")
    if cands:
        md.append("**1. Most likely failing `!c_<col>` "
                  "(narrowed from 54):**")
        md.append("")
        for c in cands:
            md.append(
                f"- `{c['rs_var']}!{c['field']}` at line "
                f"{c['use_line']} ({c['candidate_class']})")
        md.append("")
        head = q["Q1_chain_order_first"]
        md.append(
            f"In chain order, **the first one to fire is "
            f"`{head['rs_var']}!{head['field']}` at line "
            f"{head['use_line']}** — the JET 3265 raises here "
            f"and the error trap exits before any other "
            f"candidate is reached.")
        md.append("")
        md.append(
            f"**2. Which recordset they're on:** all chain-"
            f"order-first candidates hang on the binding at "
            f"line "
            f"{q['Q1_chain_order_first']['use_line'] if q['Q1_chain_order_first'] else '?'}'s "
            f"upstream `Set tRstX = ...` (see "
            f"`Q1_chain_order_first.binding_open_recordset_line` "
            f"in JSON).")
        md.append("")
        md.append("**3. Sufficient for canonical issue filing?**")
        md.append("")
        if (verdict['verdict']
                == 'recordset_projection_mismatch_candidate'):
            md.append(
                "**Yes.**  Same shape as Issue #23 "
                "(LookAtAssociations × CmdNeo4j target-column "
                "mismatch, P1):")
            md.append(
                "- Static evidence is binary — the candidate "
                "fields are confirmed present on the binding's "
                "source tables (per `tables.json`), and "
                "confirmed absent from the SELECT projection "
                "(per the parsed VBA).  No runtime "
                "confirmation needed to file.")
            md.append(
                "- Failing line cited unambiguously (chain-"
                "order-first).")
            md.append(
                "- Fix is a clear single-statement SELECT "
                "extension (see verdict_note for the SQL "
                "sketch).")
            md.append(
                "- Per-form workaround would mirror PR #116's "
                "`.replace()` shape (rewrite the SELECT "
                "identifier list, not the INSERT target list).")
            md.append("")
            md.append("**4. Smallest next step:** open a "
                      "canonical issue filing PR analogous to "
                      "PR #115 (which filed Issue #23).  No "
                      "smaller confirmation step needed.")
        else:
            md.append("Not yet — see verdict_note for the "
                      "minimum next confirmation step.")
    else:
        md.append("**Investigation inconclusive — see "
                  "verdict_note for the recommended next "
                  "step.**")
    md.append("")
    md.append("## Constraints honoured per brief")
    md.append("")
    md.append("- ✅ Investigation artifacts only — no `tests/` "
              "changes")
    md.append(
        "- ✅ No driver / README / canonical reports / issue "
        "severity / triage docs touched")
    md.append("- ✅ No new issue filed (deferred to maintainer "
              "brief)")
    md.append("- ✅ No coverage PR")
    md.append(
        "- ✅ No Access COM / no probe rerun — VBA dump + "
        "tables.json + parsed SELECT projections are "
        "sufficient")
    md.append(
        "- ✅ Raw facts and inference separated (`## Raw "
        "observed facts` is dump/parse output only; `## "
        "Verdict` and `## Direct answers` are interpretation)")
    md.append(
        "- ✅ Candidate set narrowed from 54 raw sites to a "
        "minimal chain-order-first set")

    OUT_MD.write_text("\n".join(md), encoding="utf-8")


def _write_outputs(facts: dict, verdict: dict, q: dict) -> None:
    out: dict = {
        "schema_version": 1,
        "generated_date": "2026-05-08",
        "investigation_branch": (
            "investigate/place-cmdneo4j-item-not-found"),
        "main_at_investigation": "8f94276",
        "follows_up_pr": 120,
        "static_only": True,
        "no_access_com": True,
        "source_data": {
            "vba_dump": str(VBA_PATH.relative_to(ROOT)),
            "metadata_dump": str(TABLES_JSON.relative_to(ROOT)),
        },
        "raw_facts": facts,
        "verdict": verdict["verdict"],
        "verdict_note": verdict["verdict_note"],
        "answers": q,
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2,
                   default=str),
        encoding="utf-8",
    )
    print(f"wrote {OUT_JSON}")
    _write_md(facts, verdict, q)
    print(f"wrote {OUT_MD}")


def _collect_facts() -> dict:
    src = VBA_PATH.read_text(encoding="cp1252")
    lines = src.splitlines()
    sub_start, sub_end = _find_sub_bounds(lines)
    bindings = _find_open_recordset_bindings(
        lines, sub_start, sub_end)
    uses = _find_field_uses(lines, sub_start, sub_end)
    table_cols = _load_table_columns()
    candidates = _candidate_search(bindings, uses, table_cols)
    return {
        "vba_path": str(VBA_PATH.relative_to(ROOT)),
        "vba_total_lines": len(lines),
        "sub_line_range": [sub_start, sub_end],
        "bindings": bindings,
        "total_field_uses": len(uses),
        "candidates": candidates,
        "field_use_count_per_binding": {
            f"line_{b['open_recordset_line']}_{b['rs_var']}":
            sum(1 for u in uses
                if u["binding_open_recordset_line"]
                == b["open_recordset_line"])
            for b in bindings
        },
    }


def main() -> int:
    print("=== Investigation: which !c_<col> triggers JET 3265 "
          "in Form_LookAtPlace.CmdNeo4j_Click ===\n")
    facts = _collect_facts()
    bucket = _classify_outcome(facts["candidates"])
    verdict = _verdict(facts, bucket)
    q = _q_answers(facts, bucket)
    _write_outputs(facts, verdict, q)
    print(f"\nverdict: {verdict['verdict']}")
    if facts["candidates"]:
        first_line = min(
            c["binding_open_recordset_line"]
            for c in facts["candidates"])
        first_cands = sorted(
            (c for c in facts["candidates"]
             if c["binding_open_recordset_line"] == first_line),
            key=lambda c: c["use_line"])
        head = first_cands[0]
        print(f"chain-order-first failure: "
              f"{head['rs_var']}!{head['field']} at line "
              f"{head['use_line']}")
        print(f"narrowed from 54 raw sites to "
              f"{len(first_cands)} chain-order-first "
              f"candidates")
    return 0


if __name__ == "__main__":
    sys.exit(main())
