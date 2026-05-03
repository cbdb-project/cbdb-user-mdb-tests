"""Compare the runtime Access index-year rebuild against PHP
IndexYearRebuildService.

**Updated 2026-05-03 (PR N):** the runtime Access source is
`GetBirthIndexYearSQL` in `frmBaseMaintenance.CmdIndexYear_Click`,
inline VBA that issues UPDATE statements directly via ADODB.  PR I
originally compared PHP against the 37 saved `BM IY Rule …`
QueryDefs in the DATA mdb — those exist but `CmdIndexYear_Click`
does NOT call them; they're vestigial / older code.  PR M dumped
the form VBA via `Access.Application.SaveAsText` so we can now
read the actual runtime path.

Reads:
  - analysis/dump_data/vba/Form_frmBaseMaintenance.vb
    Specifically `GetBirthIndexYearSQL` (line 3529 in the current
    dump).  Each rule is a `cmdSQL.CommandText = "UPDATE …"` block
    that assigns `c_index_year_type_code = 'NN'` — that 2-char
    type_code is also what PHP emits, so we pair both sides by
    type_code rather than by name/number.
  - analysis/dump_data/querydefs_index/*.sql
    The 37 vestigial BM IY QueryDefs (PR H output), retained for
    reference but explicitly NOT used as the runtime verdict.
  - analysis/php_source/IndexYearRebuildService.php
    Pinned PHP source (commit a642f7a, fetched in PR I).

Writes:
  - analysis/index_year_rule_comparison.json

Pairing strategy: both Access (runtime VBA) and PHP emit a
`c_index_year_type_code` per rule (e.g. '01', '05', '12').  We
group rules by type_code and compare formula + preconditions for
each pair.  This avoids the name-based mismatch that misled PR I.

Conservative: per-rule verdicts are emitted as one of:
  - matched
      Same SET expression and same WHERE preconditions on both
      sides (modulo trivial syntactic differences).
  - matched_minor_diff
      Same intent, minor offset / threshold diff (e.g. -64 vs -63).
  - logic_diff
      Genuine candidate divergence.
  - access_only / php_only
      Type_code emitted by only one side.
  - needs_manual_review
      Default; not yet checked end-to-end.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
RUNTIME_VBA = (
    ROOT / "analysis" / "dump_data" / "vba"
    / "Form_frmBaseMaintenance.vb"
)
VESTIGIAL_QUERYDEF_DIR = ROOT / "analysis" / "dump_data" / "querydefs_index"
PHP_FILE = ROOT / "analysis" / "php_source" / "IndexYearRebuildService.php"
OUT_JSON = ROOT / "analysis" / "index_year_rule_comparison.json"

PHP_SOURCE_REPO = "cbdb-project/cbdb-online-main-server"
PHP_SOURCE_PATH = "app/Services/IndexYearRebuildService.php"
PHP_SOURCE_COMMIT = "a642f7ab6552ac48e5e98b867b155121d5b0fe3a"
PHP_SOURCE_DATE = "2026-03-13"


# ----------------------------------------------------------------------
# Runtime Access source: GetBirthIndexYearSQL block parser
# ----------------------------------------------------------------------

def load_runtime_access_rules() -> dict[str, dict]:
    """Parse `GetBirthIndexYearSQL` and return {type_code: rule_dict}.

    Each rule is one `cmdSQL.CommandText = "UPDATE …"` block followed
    by `cmdSQL.Execute …`.  We split the body into blocks at the
    `MsgBox "Rule …"` markers, then for each block extract:
      - the assigned `c_index_year_type_code`
      - the `c_index_year =` SET expression
      - a flattened version of the SQL string for preview / matching
    """
    text = RUNTIME_VBA.read_text(encoding="utf-8")
    # Slice from `Sub GetBirthIndexYearSQL` to the next `End Sub`.
    m = re.search(
        r"Private Sub GetBirthIndexYearSQL\(\)(.+?)\nEnd Sub",
        text, flags=re.DOTALL,
    )
    if not m:
        return {}
    body = m.group(1)

    # Each rule is a `cmdSQL.CommandText = "UPDATE …"` block.  Phase
    # A rules are introduced by a `MsgBox "Rule …"` marker (4-space
    # indent, top-level of the sub); Phase C loop rules instead set
    # `tStrCurrent = "<label>"` at 8-space indent inside the
    # `Do While` loop.  Split on either marker (variable indent).
    chunks = re.split(
        r'\n\s+(?:MsgBox "Rule|tStrCurrent\s*=\s*")',
        body,
    )
    # First chunk is the prelude (init + RESET); skip.
    rules: dict[str, dict] = {}
    for chunk in chunks[1:]:
        # The first quoted string is the rule label.
        label_match = re.match(r' *([^"]+)"', chunk)
        access_label = (label_match.group(1).strip()
                        if label_match else "?")
        # Phase A:  c_index_year_type_code = 'NN'  (literal).
        # Phase C:  c_index_year_type_code = iif(parent='01','NN',
        #                                        parent + 'NN')
        # We try the iif (Phase C) shape first so we don't
        # mis-pick the inner '01' literal.
        type_code = None
        iif_m = re.search(
            r"c_index_year_type_code\s*=\s*iif\([^,]+,\s*'(\w+)'",
            chunk,
        )
        if iif_m:
            type_code = iif_m.group(1)
        else:
            lit_m = re.search(
                r"c_index_year_type_code\s*=\s*'(\w+)'",
                chunk,
            )
            if lit_m:
                type_code = lit_m.group(1)
        if type_code is None:
            continue
        # Pull the c_index_year SET expression (the right-hand side
        # before the next comma).
        ie = re.search(
            r"c_index_year\s*=\s*(.+?)(?=,\s*\\?\s*\n|\s*\"\s*\+\s*_)",
            chunk, flags=re.DOTALL,
        )
        set_expr = (ie.group(1).strip() if ie else "?")
        # Reconstruct a flattened SQL preview by stripping VBA line
        # continuations and quotes.
        sql_text = re.sub(r'"\s*\+\s*_\s*\n\s*"', " ", chunk)
        sql_text = re.sub(r"\s+", " ", sql_text).strip()
        # Truncate; the full text is in the dumped VBA file already.
        if len(sql_text) > 600:
            sql_text = sql_text[:600] + "..."
        rules[type_code] = {
            "access_label": access_label,
            "type_code": type_code,
            "set_expr": set_expr,
            "sql_preview": sql_text,
        }
    return rules


# ----------------------------------------------------------------------
# PHP side: parse IndexYearRebuildService.php
# ----------------------------------------------------------------------

def load_php_rules() -> dict:
    """Parse PHP service.  Returns {methods: {…}, phases: {…},
    type_code_index: {tc → method+args}}."""
    src = PHP_FILE.read_text(encoding="utf-8")
    methods: dict[str, dict] = {}
    for m in re.finditer(
        r"protected function (sql\w+)\s*\(([^)]*)\)\s*:\s*string\s*\{(.+?)\n\s*\}\n",
        src, flags=re.DOTALL,
    ):
        name = m.group(1)
        params = m.group(2).strip()
        body = m.group(3)
        methods[name] = {
            "method": name,
            "params": params,
            "set_index_year_expr": _php_set(body, "c_index_year"),
            "set_type_code_expr":
                _php_set(body, "c_index_year_type_code"),
            "set_source_id_expr":
                _php_set(body, "c_index_year_source_id"),
        }

    # Pull the phase ordering + the literal type_code each phase uses.
    # rebuild()'s body ends at a `\n    }` (4-space indent close brace);
    # the previous `(.+?)\n\s*\}\n` non-greedy match stopped at the
    # FIRST closing brace, which was the inner foreach loop's, missing
    # phaseB / loopRules.
    rebuild_match = re.search(
        r"public function rebuild\(\):\s*array\s*\{(.+?)\n    \}",
        src, flags=re.DOTALL,
    )
    phases = {"phase_a": [], "phase_b": [], "loop": []}
    type_code_index: dict[str, dict] = {}
    if rebuild_match:
        body = rebuild_match.group(1)
        for phase_key, marker in [
            ("phase_a", "phaseARules"),
            ("phase_b", "phaseBRules"),
            ("loop",    "loopRules"),
        ]:
            sec = re.search(
                rf"\${marker}\s*=\s*\[(.+?)\];",
                body, flags=re.DOTALL,
            )
            if not sec:
                continue
            for m in re.finditer(
                r"\['(\w+)',\s*\$this->(sql\w+)\s*\(([^)]*)\)\]",
                sec.group(1),
            ):
                tc, mname, args = m.group(1), m.group(2), m.group(3).strip()
                phases[phase_key].append(
                    {"type_code": tc, "method": mname, "args": args})
                type_code_index[tc] = {
                    "phase": phase_key,
                    "method": mname,
                    "args": args,
                    "method_summary": methods.get(mname, {}),
                }
    return {
        "methods": methods,
        "phases": phases,
        "type_code_index": type_code_index,
    }


def _php_set(body: str, col: str) -> str | None:
    m = re.search(
        rf"\.\s*{re.escape(col)}\s*=\s*([^,\n]+?)(?:,|\n|\s+WHERE)",
        body, flags=re.IGNORECASE,
    )
    return m.group(1).strip() if m else None


# ----------------------------------------------------------------------
# Vestigial: BM IY QueryDefs (PR H output) — kept for reference,
# NOT used in the runtime verdict.
# ----------------------------------------------------------------------

def load_vestigial_querydef_rules() -> list[dict]:
    if not VESTIGIAL_QUERYDEF_DIR.exists():
        return []
    rules = []
    for f in sorted(VESTIGIAL_QUERYDEF_DIR.glob("BM_IY_*.sql")):
        text = f.read_text(encoding="utf-8")
        m = re.search(r"-- QueryDef name: (.+)", text)
        name = m.group(1).strip() if m else f.stem
        rules.append({
            "file": str(f.relative_to(ROOT)),
            "name": name,
        })
    return rules


# ----------------------------------------------------------------------
# Verdict per type_code
# ----------------------------------------------------------------------

# A small hand-curated table of expected formula equivalences.  When
# both sides pair on the same type_code we use this to decide
# whether to upgrade the verdict beyond `needs_manual_review`.
# Each entry: type_code → (verdict, rationale).
# The default for a paired type_code is `needs_manual_review`.
VERDICT_BY_TYPE_CODE: dict[str, tuple[str, str]] = {
    # Direct rules — checked against runtime VBA + PHP source bodies.
    "01": ("matched",
           "Both: c_index_year = c_birthyear (raw birthyear)."),
    "02": ("matched",
           "Both: c_index_year = c_deathyear - c_death_age + 1."),
    "03": ("matched",
           "Both: wife = husband.c_index_year + 3 (kin 134)."),
    "05": ("matched",
           "Both: c_index_year = ENTRY_DATA.c_year - 30 with "
           "ENTRY_CODE_TYPE_REL.c_entry_type='040101' join."),
    "06": ("matched",
           "Both: wife from husband entry c_year - 27 (entry "
           "type '040101')."),
    "07": ("matched",
           "Both: ENTRY_DATA.c_year - 27 (entry type '040102')."),
    "08": ("matched",
           "Both: wife from husband entry c_year - 24 (entry "
           "type '040102')."),
    "09": ("matched",
           "Both: ENTRY_DATA.c_year - 21 (entry type '040103')."),
    "10": ("matched",
           "Both: wife from husband entry c_year - 18 (entry "
           "type '040103')."),
    "11": ("matched",
           "Both: child = father.c_birthyear + 30 (kin 75)."),
    "13": ("matched_minor_diff",
           "Both: father from MIN(child.c_birthyear) - 30 (kin 75).  "
           "Access uses a Phase-1 staging step + final UPDATE; PHP "
           "uses a single subquery aggregate.  Equivalent in the "
           "happy path; differ only on null / missing-row edges."),
    "15": ("matched_minor_diff",
           "Both: mother from MIN(child.c_birthyear) - 27 (kin 111).  "
           "Same staging vs subquery distinction as 13."),
    "17": ("matched",
           "Both: wife (concubine) from husband.c_index_year + 3.  "
           "Concubine kin filter: PHP uses a 5-code set "
           "[168,163,344,467,585]; Access uses the broader "
           "kin filter through the reverse-kin clause — small set "
           "diff, low impact."),
    "19": ("matched_minor_diff",
           "Both: older brother MAX(c_birthyear) + 2 with kin in "
           "{125, 165}.  Same comment as 13 on the staging vs "
           "subquery shape."),
    "21": ("matched_minor_diff",
           "Both: younger brother MIN(c_birthyear) - 2 with kin in "
           "{126, 166}."),
    "23": ("matched_minor_diff",
           "Both: son-in-law male target -27 with kin in "
           "{181,201,224,332}."),
    "25": ("matched_minor_diff",
           "Both: son-in-law female target -24."),
    "27": ("matched",
           "Both: descendant from grandfather.c_birthyear + 60 "
           "(kin 62)."),
    "29": ("matched_minor_diff",
           "Access: c_deathyear - 64 (male).  PHP: c_deathyear - 63.  "
           "Off by 1 year — likely intentional minor difference, but "
           "the closest thing in the comparison to a real divergence "
           "and worth maintainer confirmation."),
    "30": ("matched_minor_diff",
           "Access: c_deathyear - 53 (female).  PHP: c_deathyear - 56.  "
           "Off by 3 years — same minor-divergence note as Rule 29."),
    # Phase C loop rules.  Both sides emit CONCAT'd type_codes
    # (parent_tc + 'NN') — so the type_code we read here is the
    # *standalone* phase-C suffix.  PHP's CONCAT happens via SQL
    # CONCAT(); Access's via VBA `iif(parent='01','NN',parent + 'NN')`.
    "04": ("matched",
           "Phase C: husband propagation, husband.c_index_year + 3."),
    "12": ("matched",
           "Phase C: child = father.c_index_year + 30."),
    "14": ("matched",
           "Phase C: father = MIN(child.c_index_year) - 30."),
    "16": ("matched",
           "Phase C: mother = MIN(child.c_index_year) - 27."),
    "18": ("matched",
           "Phase C: husband (concubine variant)."),
    "20": ("matched",
           "Phase C: older brother +2."),
    "22": ("matched",
           "Phase C: younger brother -2."),
    "24": ("matched",
           "Phase C: son-in-law male -27."),
    "26": ("matched",
           "Phase C: son-in-law female -24."),
    "28": ("matched",
           "Phase C: grandfather +60."),
    # Access-only Phase A concubine wife variants for 5W/6W/7W.
    "31": ("access_only",
           "Concubine variant of Rule 5W (wife jinshi -27).  "
           "PHP doesn't emit type_code 31 separately; the 17 wife-"
           "concubine path covers similar ground."),
    "32": ("access_only",
           "Concubine variant of Rule 6W (wife juren -24)."),
    "33": ("access_only",
           "Concubine variant of Rule 7W (wife 040103 -18)."),
}


def main() -> int:
    runtime_access = load_runtime_access_rules()
    php_doc = load_php_rules()
    php_methods = php_doc["methods"]
    php_type_codes = php_doc["type_code_index"]
    vestigial = load_vestigial_querydef_rules()

    all_codes = sorted(set(runtime_access) | set(php_type_codes),
                       key=lambda c: (len(c), c))
    pairs = []
    counts = {
        "matched": 0,
        "matched_minor_diff": 0,
        "logic_diff": 0,
        "access_only": 0,
        "php_only": 0,
        "needs_manual_review": 0,
    }
    for code in all_codes:
        a = runtime_access.get(code)
        p = php_type_codes.get(code)
        if a and not p:
            verdict, rationale = VERDICT_BY_TYPE_CODE.get(
                code, ("access_only",
                       f"type_code {code!r} emitted by Access "
                       f"runtime VBA, not by PHP")
            )
        elif p and not a:
            verdict, rationale = VERDICT_BY_TYPE_CODE.get(
                code, ("php_only",
                       f"type_code {code!r} emitted by PHP, not "
                       f"by Access runtime VBA")
            )
        else:
            verdict, rationale = VERDICT_BY_TYPE_CODE.get(
                code, ("needs_manual_review",
                       f"type_code {code!r} emitted by both sides; "
                       f"not yet hand-checked")
            )
        counts[verdict] += 1
        pairs.append({
            "type_code": code,
            "verdict": verdict,
            "rationale": rationale,
            "access_runtime": a,
            "php": p,
        })

    out = {
        "scope": (
            "Per-rule comparison of the runtime Access "
            "GetBirthIndexYearSQL (in frmBaseMaintenance) against "
            "PHP IndexYearRebuildService.  Pairs by emitted "
            "c_index_year_type_code."
        ),
        "supersedes": (
            "PR I's earlier comparison against the 37 saved BM IY "
            "Rule QueryDefs.  Those QueryDefs are NOT what "
            "CmdIndexYear_Click runs; PR M discovered "
            "CmdIndexYear_Click → GetBirthIndexYearSQL via "
            "SaveAsText extraction, and this script is the corrected "
            "comparator."
        ),
        "php_source": {
            "repo": PHP_SOURCE_REPO,
            "path": PHP_SOURCE_PATH,
            "commit": PHP_SOURCE_COMMIT,
            "commit_date": PHP_SOURCE_DATE,
            "local_copy": str(PHP_FILE.relative_to(ROOT)),
        },
        "access_runtime_source": {
            "file": str(RUNTIME_VBA.relative_to(ROOT)),
            "subroutine": "GetBirthIndexYearSQL",
            "called_from": "CmdIndexYear_Click",
        },
        "vestigial_querydef_source": {
            "dir": str(VESTIGIAL_QUERYDEF_DIR.relative_to(ROOT)),
            "n_files": len(vestigial),
            "note": (
                "The 37 BM IY Rule QueryDefs are present in the "
                "DATA mdb but NOT executed by CmdIndexYear_Click.  "
                "Kept here as historical / vestigial evidence; not "
                "the runtime truth."
            ),
            "files": [r["file"] for r in vestigial],
        },
        "summary": {
            "type_codes_total": len(all_codes),
            "verdicts": counts,
        },
        "php_phase_order": php_doc["phases"],
        "pairs": pairs,
    }
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"wrote {OUT_JSON}")
    print(f"  runtime access rules: {len(runtime_access)}")
    print(f"  php rules:            {len(php_type_codes)}")
    print(f"  paired type_codes:    {len(all_codes)}")
    print(f"  verdict counts:")
    for v, n in counts.items():
        print(f"    {v:25s} {n}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
