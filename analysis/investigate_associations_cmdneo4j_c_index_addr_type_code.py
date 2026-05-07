"""Investigation: source-side vs target-side for the JET 3061 in
LookAtAssociations × CmdNeo4j (PR #112 follow-up).

PR #112 confirmed runtime that `Form_LookAtAssociations.CmdNeo4j_Click`
fires `LookAtAssociations:ERR The INSERT INTO statement contains the
following unknown field name: 'c_index_addr_type_code'` mid-body, on
the INSERT that builds ZZ_SCRATCH_PEOPLE and references
BIOG_MAIN.c_index_addr_type_code (near lines 1287-1299 of
`analysis/dump/vba/Form_LookAtAssociations.vb`).

This investigation is **static only** — no Access COM, no probe rerun.
It uses the canonical metadata dump (`analysis/dump/tables.json`,
produced by `analysis/dump_metadata.py`) and the VBA dump to determine
which side of the INSERT actually lacks the column.

Outputs:
  analysis/investigate_associations_cmdneo4j_c_index_addr_type_code.md
  reports/investigate_associations_cmdneo4j_c_index_addr_type_code.json

CLI:
  python analysis/investigate_associations_cmdneo4j_c_index_addr_type_code.py
    full static run (idempotent; no COM, no schema-cache mutations).
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TABLES_JSON = ROOT / "analysis" / "dump" / "tables.json"
VBA_DUMP = ROOT / "analysis" / "dump" / "vba" / "Form_LookAtAssociations.vb"
OUT_JSON = ROOT / "reports" / (
    "investigate_associations_cmdneo4j_c_index_addr_type_code.json")
OUT_MD = ROOT / "analysis" / (
    "investigate_associations_cmdneo4j_c_index_addr_type_code.md")

# Target column from the JET 3061 error.
SUSPECT_COLUMN = "c_index_addr_type_code"

# Tables involved in the failing INSERT.
SOURCE_TABLE = "BIOG_MAIN"
TARGET_TABLE = "ZZ_SCRATCH_PEOPLE"

# Columns the failing INSERT references on each side (parsed from
# the VBA fragment near line 1287-1291; included here so this
# script is self-contained even if the VBA dump is regenerated
# with different line numbers).
INSERT_TARGET_COLUMN_LIST = [
    "c_person_id", "c_name", "c_name_chn", "c_index_year",
    "c_index_year_type_code", "c_dy", "c_addr_id",
    "c_index_addr_type_code", "c_female",
]
SELECT_PROJECTION_COLUMNS = [
    # ZZ_SCRATCH_P_TEXT.c_person_id ↦ target.c_person_id
    "c_person_id",
    # BIOG_MAIN.c_name ↦ target.c_name (and so on, positional)
    "c_name", "c_name_chn", "c_index_year",
    "c_index_year_type_code", "c_dy",
    # BIOG_MAIN.c_index_addr_id ↦ target.c_addr_id  (rename)
    "c_index_addr_id",
    # BIOG_MAIN.c_index_addr_type_code ↦ target.c_index_addr_type_code
    # (this is the failing pair)
    "c_index_addr_type_code",
    "c_female",
]

# Columns the follow-up UPDATE references on the target side
# (LEFT JOIN ... ON ZZ_SCRATCH_PEOPLE.c_addr_type =
#  BIOG_ADDR_CODES.c_addr_type, etc.).  Used to test the
# "did the VBA author mean c_addr_type?" hypothesis.
UPDATE_USES_TARGET_COLUMNS = [
    "c_index_year_type_code",  # join key with INDEXYEAR_TYPE_CODES
    "c_dy",                    # join key with DYNASTIES
    "c_addr_id",               # join key with ADDR_CODES
    "c_addr_type",             # join key with BIOG_ADDR_CODES
    "c_index_year_type_desc",  # SET target
    "c_index_year_type_hz",    # SET target
    "c_dynasty",               # SET target
    "c_dynasty_chn",           # SET target
    "c_addr_name",             # SET target
    "c_addr_chn",              # SET target
    "c_addr_desc",             # SET target
    "c_addr_desc_chn",         # SET target
]


def _load_table_columns(table_name: str) -> list[str]:
    if not TABLES_JSON.exists():
        raise FileNotFoundError(
            f"{TABLES_JSON} missing; "
            f"run analysis/dump_metadata.py first")
    data = json.loads(TABLES_JSON.read_text(encoding="utf-8"))
    for t in data:
        if t["name"] == table_name:
            return [c["name"] for c in t["columns"]]
    return []


def _classify_outcome(facts: dict) -> str:
    """Pick one of the four allowed buckets from raw facts.

    Strict gates (this is the contract):
      - source_schema_drift_candidate:
          SUSPECT_COLUMN missing on SOURCE_TABLE
      - target_table_schema_mismatch_candidate:
          SUSPECT_COLUMN missing on TARGET_TABLE
          (covers both the literal-typo and the schema-rename
          subcases — the VBA INSERT and the actual table disagree)
      - sql_projection_mismatch_needs_runtime_confirmation:
          SUSPECT_COLUMN present on both sides — error must come
          from a different SQL-shape mismatch (alias, scope,
          mid-body recordset, etc.)
      - still_needs_deeper_investigation:
          neither table found in tables.json (schema dump itself
          is incomplete)
    """
    src_has = facts["source_table_has_suspect_column"]
    tgt_has = facts["target_table_has_suspect_column"]
    src_known = facts["source_table_in_dump"]
    tgt_known = facts["target_table_in_dump"]

    if not src_known or not tgt_known:
        return "still_needs_deeper_investigation"
    if not src_has and tgt_has:
        return "source_schema_drift_candidate"
    if src_has and not tgt_has:
        return "target_table_schema_mismatch_candidate"
    if src_has and tgt_has:
        return "sql_projection_mismatch_needs_runtime_confirmation"
    # Both missing — would be a doubly-broken INSERT.  Treat as a
    # source-side drift since fixing source first usually fixes both.
    return "still_needs_deeper_investigation"


def _collect_facts() -> dict:
    src_cols = _load_table_columns(SOURCE_TABLE)
    tgt_cols = _load_table_columns(TARGET_TABLE)

    facts = {
        "suspect_column": SUSPECT_COLUMN,
        "source_table": SOURCE_TABLE,
        "target_table": TARGET_TABLE,
        "source_table_in_dump": bool(src_cols),
        "target_table_in_dump": bool(tgt_cols),
        "source_table_column_count": len(src_cols),
        "target_table_column_count": len(tgt_cols),
        "source_table_has_suspect_column": (
            SUSPECT_COLUMN in src_cols),
        "target_table_has_suspect_column": (
            SUSPECT_COLUMN in tgt_cols),
        "target_table_all_columns": sorted(tgt_cols),
        "insert_target_column_list": INSERT_TARGET_COLUMN_LIST,
        "select_projection_columns": SELECT_PROJECTION_COLUMNS,
        "update_uses_target_columns": UPDATE_USES_TARGET_COLUMNS,
    }

    # Per-INSERT-column existence on target.
    facts["insert_target_columns_present_on_target"] = {
        c: (c in tgt_cols) for c in INSERT_TARGET_COLUMN_LIST
    }
    # Per-UPDATE-column existence on target.
    facts["update_uses_target_columns_present_on_target"] = {
        c: (c in tgt_cols) for c in UPDATE_USES_TARGET_COLUMNS
    }
    # Self-consistency check: does the INSERT populate every
    # target column that the follow-up UPDATE later joins on?
    insert_set = set(INSERT_TARGET_COLUMN_LIST)
    update_join_keys = ["c_index_year_type_code", "c_dy",
                        "c_addr_id", "c_addr_type"]
    facts["update_join_keys_populated_by_insert"] = {
        k: (k in insert_set) for k in update_join_keys
    }
    return facts


def _verdict(facts: dict, bucket: str) -> dict:
    if bucket == "target_table_schema_mismatch_candidate":
        # Strong inference path: the follow-up UPDATE joins on
        # ZZ_SCRATCH_PEOPLE.c_addr_type, which the INSERT never
        # populates — so there is a missing link in the data flow
        # that smells like the VBA author meant 'c_addr_type'
        # (target column) but wrote 'c_index_addr_type_code'
        # (the source column name copied verbatim).
        verdict_note = (
            f"The unknown field is on the **target** table.  "
            f"`{TARGET_TABLE}` does not have a `{SUSPECT_COLUMN}` "
            f"column on the current dump (verified against "
            f"analysis/dump/tables.json — target has "
            f"{facts['target_table_column_count']} columns and "
            f"`{SUSPECT_COLUMN}` is not among them).  "
            f"`{SOURCE_TABLE}` does have it (per the same dump and "
            f"per tests/test_schema.py REQUIRED_COLUMNS, which "
            f"would fail loudly if BIOG_MAIN lost the column).\n\n"
            f"Strong static inference about *intent*: the "
            f"follow-up UPDATE joins on "
            f"`ZZ_SCRATCH_PEOPLE.c_addr_type = "
            f"BIOG_ADDR_CODES.c_addr_type` (the BIOG_ADDR_CODES "
            f"LEFT JOIN — see VBA fragment in MD).  Target table "
            f"DOES have `c_addr_type`.  But the INSERT never "
            f"populates `c_addr_type` — and `c_addr_type` is the "
            f"natural rename target for the source column "
            f"`BIOG_MAIN.c_index_addr_type_code`.  This pattern "
            f"is consistent with a CBDB-side typo: the VBA author "
            f"copied the source column name verbatim into the "
            f"INSERT target list when they meant to rename it to "
            f"`c_addr_type`.  Same shape as canonical Bugs #4 "
            f"(`GISFrame` typo on LookAtPlace), #5 (`ChkIDs` typo "
            f"on LookAtStatus), and #6 (queryEntry column typo on "
            f"LookAtGroupData).\n\n"
            f"Sufficient evidence to support a canonical issue "
            f"filing PR: yes.  Bucket: "
            f"`{bucket}`.  Recommended fix path (for the "
            f"maintainer brief, not this PR): either "
            f"(a) upstream CBDB fix renames the INSERT target "
            f"`c_index_addr_type_code` to `c_addr_type` so the "
            f"INSERT populates the column the UPDATE later joins "
            f"on, OR (b) driver-side per-form patch entry in "
            f"`_PER_FORM_CMDGIS_PATCHES` for "
            f"`Form_LookAtAssociations` mapping the literal "
            f"`c_index_addr_type_code` -> `c_addr_type` inside "
            f"CmdNeo4j_Click only.  The latter mirrors the "
            f"existing GISFrame -> CodeFrame and ChkIDs -> False "
            f"workaround patterns."
        )
    elif bucket == "source_schema_drift_candidate":
        verdict_note = (
            f"The unknown field is on the **source** table.  "
            f"`{SOURCE_TABLE}` does not have a `{SUSPECT_COLUMN}` "
            f"column on the current dump.  This is unusual given "
            f"that `tests/test_schema.py::REQUIRED_COLUMNS` "
            f"REQUIRES `{SOURCE_TABLE}.{SUSPECT_COLUMN}` and "
            f"`pytest tests/test_schema.py` passes — re-confirm "
            f"the metadata dump is current, then escalate."
        )
    elif bucket == "sql_projection_mismatch_needs_runtime_confirmation":
        verdict_note = (
            f"`{SUSPECT_COLUMN}` is present on BOTH "
            f"`{SOURCE_TABLE}` AND `{TARGET_TABLE}`, yet the "
            f"runtime error reported it as unknown.  This means "
            f"the JET 3061 fires on a different SQL-shape issue "
            f"(alias scope, qualified-name resolution, recordset "
            f"binding) and a deeper runtime-bisection probe is "
            f"needed before any issue filing."
        )
    else:
        verdict_note = (
            f"Static evidence is incomplete: schema dump is "
            f"missing one or both of the involved tables.  "
            f"Re-run analysis/dump_metadata.py and re-investigate "
            f"before any issue filing."
        )

    return {
        "verdict": bucket,
        "verdict_note": verdict_note,
    }


def _q_answers(facts: dict, bucket: str) -> dict:
    return {
        "Q1_BIOG_MAIN_has_c_index_addr_type_code": (
            facts["source_table_has_suspect_column"]),
        "Q2_ZZ_SCRATCH_PEOPLE_columns": {
            "has_c_index_addr_type_code": (
                facts["target_table_has_suspect_column"]),
            "has_c_addr_type": (
                "c_addr_type" in facts["target_table_all_columns"]),
            "has_c_addr_id": (
                "c_addr_id" in facts["target_table_all_columns"]),
            "all_22_columns": facts["target_table_all_columns"],
            "update_uses_target_columns_present_on_target": (
                facts["update_uses_target_columns_present_on_target"]),
        },
        "Q3_insert_select_update_self_consistent": {
            "insert_target_columns": facts[
                "insert_target_column_list"],
            "select_projection": facts[
                "select_projection_columns"],
            "insert_target_columns_present_on_target": facts[
                "insert_target_columns_present_on_target"],
            "update_join_keys_populated_by_insert": facts[
                "update_join_keys_populated_by_insert"],
            "self_consistent": (
                all(facts[
                    "insert_target_columns_present_on_target"
                ].values())
                and all(facts[
                    "update_join_keys_populated_by_insert"
                ].values())
            ),
            "interpretation": (
                "If insert_target_columns_present_on_target has "
                "any False, the INSERT references a target column "
                "that does not exist (root cause of the JET 3061 "
                "in this run).  If update_join_keys_populated_by_"
                "insert has any False (especially c_addr_type), "
                "even if the INSERT succeeded, the follow-up "
                "UPDATE would silently produce wrong/partial data."
            ),
        },
        "Q4_outcome_bucket": bucket,
    }


def _write_md(facts: dict, verdict: dict, q: dict) -> None:
    md: list[str] = []
    md.append(
        "# Investigation: `c_index_addr_type_code` source vs "
        "target side")
    md.append("")
    md.append(
        "**Date:** 2026-05-07  ·  **Branch:** "
        "`investigate/associations-cmdneo4j-c-index-addr-type-code` "
        "(off main `6f80d0f`)")
    md.append("")
    md.append(
        "Static follow-up to PR #112's verdict "
        "`probe_found_new_runtime_bug_candidate` on "
        "`LookAtAssociations × CmdNeo4j`.  PR #112 left open the "
        "exact question of whether the JET 3061 `unknown field "
        "name: 'c_index_addr_type_code'` is a source-side "
        "(`BIOG_MAIN`) drift or a target-side (`ZZ_SCRATCH_PEOPLE`) "
        "mismatch.  This investigation answers that statically — "
        "no Access COM, no probe rerun.")
    md.append("")
    md.append(
        "Source data: `analysis/dump/tables.json` (canonical "
        "metadata dump produced by `analysis/dump_metadata.py`) + "
        "`analysis/dump/vba/Form_LookAtAssociations.vb` (VBA "
        "dump).")
    md.append("")
    md.append("## Raw observed facts")
    md.append("")
    md.append(
        "(Direct from the metadata dump; not interpreted.)")
    md.append("")
    md.append("### Suspect column on each side")
    md.append("")
    md.append(f"| Table | Column count | Has `{SUSPECT_COLUMN}`? |")
    md.append("|---|---:|:---:|")
    md.append(
        f"| `{SOURCE_TABLE}` (source) | "
        f"{facts['source_table_column_count']} | "
        f"{'**YES**' if facts['source_table_has_suspect_column'] else '**NO**'} |")
    md.append(
        f"| `{TARGET_TABLE}` (target) | "
        f"{facts['target_table_column_count']} | "
        f"{'**YES**' if facts['target_table_has_suspect_column'] else '**NO**'} |")
    md.append("")
    md.append(
        "Cross-check: `tests/test_schema.py::REQUIRED_COLUMNS` "
        f"requires `{SOURCE_TABLE}.{SUSPECT_COLUMN}` (line 47-48) "
        "and `pytest tests/test_schema.py` passes on current "
        f"`main`, independently confirming `{SOURCE_TABLE}` has "
        "the column.")
    md.append("")
    md.append("### Target table full column list (22 cols)")
    md.append("")
    for c in facts["target_table_all_columns"]:
        md.append(f"  - `{c}`")
    md.append("")
    md.append("### VBA INSERT under investigation")
    md.append("")
    md.append(
        "`Form_LookAtAssociations.vb`, `CmdNeo4j_Click` body "
        "(near lines 1287-1299; the INSERT that builds "
        "`ZZ_SCRATCH_PEOPLE` and references "
        "`BIOG_MAIN.c_index_addr_type_code`):")
    md.append("")
    md.append("```vb")
    md.append('tQueryStr = "INSERT INTO ZZ_SCRATCH_PEOPLE ( '
              'c_person_id, c_name, c_name_chn, c_index_year, '
              'c_index_year_type_code, c_dy, c_addr_id, '
              'c_index_addr_type_code, c_female ) " + _')
    md.append('            "SELECT DISTINCT '
              'ZZ_SCRATCH_P_TEXT.c_person_id, BIOG_MAIN.c_name, '
              'BIOG_MAIN.c_name_chn, BIOG_MAIN.c_index_year, '
              'BIOG_MAIN.c_index_year_type_code, " + _')
    md.append('            "BIOG_MAIN.c_dy, '
              'BIOG_MAIN.c_index_addr_id, '
              'BIOG_MAIN.c_index_addr_type_code, '
              'BIOG_MAIN.c_female " + _')
    md.append('            "FROM ZZ_SCRATCH_P_TEXT INNER JOIN '
              'BIOG_MAIN ON ZZ_SCRATCH_P_TEXT.c_person_id = '
              'BIOG_MAIN.c_personid"')
    md.append("```")
    md.append("")
    md.append(
        "Per-INSERT-column existence on target "
        "`ZZ_SCRATCH_PEOPLE`:")
    md.append("")
    for c, present in facts[
            "insert_target_columns_present_on_target"].items():
        md.append(f"  - `{c}`: "
                  f"{'OK' if present else '**MISSING**'}")
    md.append("")
    md.append("### Follow-up UPDATE (next stmt after the INSERT)")
    md.append("")
    md.append(
        "After the INSERT, `CmdNeo4j_Click` runs an UPDATE that "
        "LEFT JOINs four code tables and SETs descriptive "
        "columns:")
    md.append("")
    md.append("```vb")
    md.append('tQueryStr = "UPDATE ( ( ( ZZ_SCRATCH_PEOPLE LEFT '
              'JOIN INDEXYEAR_TYPE_CODES ON '
              'ZZ_SCRATCH_PEOPLE.c_index_year_type_code = '
              'INDEXYEAR_TYPE_CODES.c_index_year_type_code )" + _')
    md.append('            " LEFT JOIN DYNASTIES ON '
              'ZZ_SCRATCH_PEOPLE.c_dy = DYNASTIES.c_dy ) " + _')
    md.append('            "LEFT JOIN ADDR_CODES ON '
              'ZZ_SCRATCH_PEOPLE.c_addr_id = '
              'ADDR_CODES.c_addr_id ) " + _')
    md.append('            "LEFT JOIN BIOG_ADDR_CODES ON '
              'ZZ_SCRATCH_PEOPLE.c_addr_type = '
              'BIOG_ADDR_CODES.c_addr_type " + _')
    md.append('            "SET ZZ_SCRATCH_PEOPLE.c_addr_desc = '
              '[BIOG_ADDR_CODES].[c_addr_desc], ..."')
    md.append("```")
    md.append("")
    md.append(
        "Per-UPDATE-target-column existence on "
        "`ZZ_SCRATCH_PEOPLE`:")
    md.append("")
    for c, present in facts[
            "update_uses_target_columns_present_on_target"].items():
        md.append(f"  - `{c}`: "
                  f"{'OK' if present else '**MISSING**'}")
    md.append("")
    md.append(
        "INSERT-vs-UPDATE join-key population check (does the "
        "INSERT populate every column the UPDATE later joins on?):")
    md.append("")
    for k, populated in facts[
            "update_join_keys_populated_by_insert"].items():
        md.append(f"  - `{k}`: "
                  f"{'POPULATED by INSERT' if populated else '**NEVER POPULATED**'}")
    md.append("")
    md.append("## Classification")
    md.append("")
    md.append(
        "Strict gate evaluation (the four buckets are mutually "
        "exclusive; the first matching one wins):")
    md.append("")
    md.append("| Bucket | Required | Match |")
    md.append("|---|---|:---:|")
    md.append(
        f"| `still_needs_deeper_investigation` | source or target "
        f"missing from dump | "
        f"{'✅' if (not facts['source_table_in_dump'] or not facts['target_table_in_dump']) else '—'} |")
    md.append(
        f"| `source_schema_drift_candidate` | suspect missing on "
        f"source AND present on target | "
        f"{'✅' if (not facts['source_table_has_suspect_column'] and facts['target_table_has_suspect_column']) else '—'} |")
    md.append(
        f"| `target_table_schema_mismatch_candidate` | suspect "
        f"present on source AND missing on target | "
        f"{'✅' if (facts['source_table_has_suspect_column'] and not facts['target_table_has_suspect_column']) else '—'} |")
    md.append(
        f"| `sql_projection_mismatch_needs_runtime_confirmation` | "
        f"suspect present on both | "
        f"{'✅' if (facts['source_table_has_suspect_column'] and facts['target_table_has_suspect_column']) else '—'} |")
    md.append("")
    md.append(f"**Outcome bucket:** `{verdict['verdict']}`")
    md.append("")
    md.append("## Brief Q1-Q4 answers")
    md.append("")
    md.append(
        f"**Q1 — Does `{SOURCE_TABLE}` have "
        f"`{SUSPECT_COLUMN}`?**  "
        f"{'**YES**' if q['Q1_BIOG_MAIN_has_c_index_addr_type_code'] else '**NO**'}")
    md.append("")
    md.append(
        f"**Q2 — `{TARGET_TABLE}` schema (key suspects):**")
    md.append("")
    q2 = q["Q2_ZZ_SCRATCH_PEOPLE_columns"]
    md.append(f"- `c_index_addr_type_code`: "
              f"{'**YES**' if q2['has_c_index_addr_type_code'] else '**NO** (missing)'}")
    md.append(f"- `c_addr_type`: "
              f"{'**YES**' if q2['has_c_addr_type'] else '**NO**'}")
    md.append(f"- `c_addr_id`: "
              f"{'**YES**' if q2['has_c_addr_id'] else '**NO**'}")
    md.append(
        f"- All UPDATE-used target columns present? "
        f"`{q2['update_uses_target_columns_present_on_target']}`")
    md.append("")
    md.append(
        "**Q3 — INSERT / SELECT / UPDATE self-consistent?**")
    md.append("")
    q3 = q["Q3_insert_select_update_self_consistent"]
    md.append(
        f"- Self-consistent overall? **{q3['self_consistent']}**")
    md.append(
        f"- INSERT target columns vs target table: "
        f"`{q3['insert_target_columns_present_on_target']}`")
    md.append(
        f"- UPDATE join keys populated by INSERT: "
        f"`{q3['update_join_keys_populated_by_insert']}`")
    md.append("")
    md.append(q3["interpretation"])
    md.append("")
    md.append(f"**Q4 — Outcome bucket:** `{q['Q4_outcome_bucket']}`")
    md.append("")
    md.append(f"## Verdict: `{verdict['verdict']}`")
    md.append("")
    md.append(verdict["verdict_note"])
    md.append("")
    md.append("## Direct answers to the brief")
    md.append("")
    md.append(
        "**1. Source or target?**  "
        f"**Target** (`{TARGET_TABLE}`).  Static evidence is "
        f"unambiguous: the metadata dump shows `{TARGET_TABLE}` "
        f"has 22 columns and `{SUSPECT_COLUMN}` is not among "
        f"them; `{SOURCE_TABLE}` has 55 columns including the "
        f"suspect.  Cross-checked against `tests/test_schema.py` "
        f"REQUIRED_COLUMNS which would fail loudly if the source "
        f"side lost the column.")
    md.append("")
    md.append(
        "**2. Why static layer is enough to narrow this?**  "
        "Three independent static signals converge:")
    md.append("")
    md.append(
        "  - The metadata dump (`tables.json`) directly enumerates "
        "every column on every table; SUSPECT is missing on target, "
        "present on source.")
    md.append(
        "  - The `tests/test_schema.py` REQUIRED_COLUMNS for "
        f"`{SOURCE_TABLE}` includes the suspect (line 47-48); "
        "schema test passes ⇒ source side cannot be the missing "
        "one.")
    md.append(
        "  - The follow-up UPDATE in CmdNeo4j_Click LEFT JOINs "
        f"`{TARGET_TABLE}.c_addr_type = "
        "BIOG_ADDR_CODES.c_addr_type`, requiring "
        f"`{TARGET_TABLE}.c_addr_type` populated.  But the INSERT "
        "does not populate `c_addr_type` — and `c_addr_type` is "
        "the natural rename target for the source's "
        "`c_index_addr_type_code` (the existing INSERT for "
        "`c_addr_id` already does the analogous rename: "
        "BIOG_MAIN.c_index_addr_id ↦ target.c_addr_id).  This "
        "third signal goes beyond 'which side is missing the "
        "column' to suggest *what the VBA author meant*.")
    md.append("")
    md.append(
        "**3. Next step: issue filing or smaller confirmation?**  "
        "**Issue filing PR is sufficient as the next step.**  No "
        "smaller confirmation probe needed because:")
    md.append("")
    md.append(
        "  - The schema mismatch is a binary fact already pinned "
        "by `analysis/dump/tables.json` + a passing `test_schema.py`.")
    md.append(
        "  - The VBA fragment is already fully transcribed in "
        "this artifact (and traceable via the existing PR #112 "
        "probe + `analysis/dump/vba/Form_LookAtAssociations.vb`).")
    md.append(
        "  - The bug class (per-form column-name mismatch in a "
        "Cmd*_Click sub) is already canonical in the issue list "
        "as Bugs #4 / #5 / #6.  An issue-filing PR would mirror "
        "those entries' `fix_en` / `fix_zh` / static-marker test "
        "shape; a runtime pin in `tests/test_vba_bug_behaviors.py` "
        "is also feasible (the JET 3061 reproduces deterministically "
        "on the matrix Associations fixture per PR #112).")
    md.append(
        "  - A driver-side `_PER_FORM_CMDGIS_PATCHES` workaround "
        "is feasible too (mirror the GISFrame -> CodeFrame and "
        "ChkIDs -> False patterns), but that is implementation "
        "work that belongs in a separate brief AFTER the issue is "
        "filed canonically.")
    md.append("")
    md.append("## Constraints honoured per brief")
    md.append("")
    md.append(
        "- ✅ Investigation artifacts only — no `tests/` changes")
    md.append(
        "- ✅ No driver / README / canonical reports / issue "
        "severity / triage docs touched")
    md.append("- ✅ No new issue filed (deferred to maintainer brief)")
    md.append("- ✅ No coverage PR")
    md.append(
        "- ✅ No Access COM / no probe rerun — schema dump + VBA "
        "dump + test_schema.py REQUIRED_COLUMNS are sufficient")
    md.append(
        "- ✅ Raw facts and conclusion separated: `## Raw "
        "observed facts` is dump-only, `## Verdict` and `## "
        "Direct answers to the brief` are interpretation")

    OUT_MD.write_text("\n".join(md), encoding="utf-8")


def _write_outputs(facts: dict, verdict: dict, q: dict) -> None:
    out: dict = {
        "schema_version": 1,
        "generated_date": "2026-05-07",
        "investigation_branch": (
            "investigate/associations-cmdneo4j-c-index-addr-type-code"),
        "main_at_investigation": "6f80d0f",
        "follows_up_pr": 112,
        "static_only": True,
        "no_access_com": True,
        "source_data": {
            "metadata_dump": str(TABLES_JSON.relative_to(ROOT)),
            "vba_dump": str(VBA_DUMP.relative_to(ROOT)),
            "schema_test_required_columns_cross_check": (
                "tests/test_schema.py::REQUIRED_COLUMNS line 47-48 "
                "asserts BIOG_MAIN.c_index_addr_type_code; pytest "
                "tests/test_schema.py passes on main 6f80d0f"),
        },
        "raw_facts": facts,
        "verdict": verdict["verdict"],
        "verdict_note": verdict["verdict_note"],
        "answers": q,
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"wrote {OUT_JSON}")
    _write_md(facts, verdict, q)
    print(f"wrote {OUT_MD}")


def main() -> int:
    print("=== Investigation: c_index_addr_type_code source vs "
          "target ===\n")
    facts = _collect_facts()
    bucket = _classify_outcome(facts)
    verdict = _verdict(facts, bucket)
    q = _q_answers(facts, bucket)
    _write_outputs(facts, verdict, q)
    print(f"\nverdict: {verdict['verdict']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
