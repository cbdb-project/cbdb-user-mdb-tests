"""Cross-reference all LookAt* export buttons with their test
coverage and depth.

Static analysis only.  Reads `analysis/dump/button_taxonomy.md`
for the universe of export buttons; greps `tests/` for which
ones are referenced and at what depth.

Coverage depth ladder (lowest → highest):

  - **none**         no test references the button at all
  - **smoke**        a test calls the button but only asserts
                     "no error / file appears"
  - **structural**   per-row width / header columns / cell
                     count assertions
  - **manifest**     per-form required-column manifest
                     (PR P / PR T style)
  - **byte_golden**  byte-level diff against a frozen golden
                     file (e.g. tests/golden/exports/...)

Outputs:

  - reports/export_coverage_matrix.json
  - analysis/export_coverage_matrix.md
"""
from __future__ import annotations

import json
import re
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TAXONOMY = ROOT / "analysis" / "dump" / "button_taxonomy.md"
TESTS_DIR = ROOT / "tests"
GOLDEN_DIR = ROOT / "tests" / "golden" / "exports"
OUT_JSON = ROOT / "reports" / "export_coverage_matrix.json"
OUT_MD = ROOT / "analysis" / "export_coverage_matrix.md"

# Regex helpers
RE_FORM_HDR = re.compile(r"^## (LookAt\w+)\s*$")
RE_EXPORT_BTN = re.compile(
    r"^\| `(Cmd\w+)` \| CommandButton \| \*\*export\*\*"
)


def parse_taxonomy() -> dict[str, list[str]]:
    text = TAXONOMY.read_text(encoding="utf-8")
    out: dict[str, list[str]] = defaultdict(list)
    cur = None
    for line in text.splitlines():
        m = RE_FORM_HDR.match(line)
        if m:
            cur = m.group(1)
            continue
        m = RE_EXPORT_BTN.match(line)
        if m and cur:
            out[cur].append(m.group(1))
    return dict(out)


def scan_tests() -> dict[tuple[str, str], list[dict]]:
    """For every (form, button) pair, list test files that
    reference it together with a depth-class hint."""
    out: dict[tuple[str, str], list[dict]] = defaultdict(list)
    for path in sorted(TESTS_DIR.rglob("*.py")):
        try:
            text = path.read_text(encoding="utf-8")
        except Exception:
            continue
        for form in FORMS:
            if form not in text:
                continue
            for btn in BUTTONS_BY_FORM.get(form, []):
                # Pattern: form mentioned AND button mentioned somewhere
                # nearby.  Heuristic: any line mentioning btn in this file
                # that ALSO has the form name in the file body counts.
                if btn in text:
                    out[(form, btn)].append({
                        "test_file": str(path.relative_to(ROOT)),
                        "depth_hints": _depth_hints(text, btn),
                    })
    return out


def _depth_hints(text: str, btn: str) -> list[str]:
    hints = []
    if "byte" in text.lower() and "golden" in text.lower():
        hints.append("byte_golden_keyword")
    if "manifest" in text.lower() or "_REQUIRED_COLUMNS" in text:
        hints.append("manifest_keyword")
    if "_assert_gis_export_depth" in text or "depth check" in text.lower():
        hints.append("structural_depth_check")
    if "out_path.exists" in text or "sz > 0" in text:
        hints.append("smoke_file_exists")
    if "header" in text.lower() and "split" in text:
        hints.append("structural_header_parse")
    return hints


def classify_depth(hits: list[dict], golden_files: list[str],
                    btn: str, form: str) -> str:
    """Return single highest-rung depth class for this (form, btn)."""
    if not hits:
        return "none"
    # Look for golden file naming convention.
    for g in golden_files:
        gn = g.lower()
        if form.lower().replace("lookat", "") in gn and (
            (btn[3:].lower() in gn) or
            ("gis" in gn and btn == "CmdGIS") or
            ("neo4j" in gn and btn == "CmdNeo4j") or
            ("kml" in gn and btn == "CmdGIS")
        ):
            return "byte_golden"
    # Otherwise use depth hints.
    all_hints = [h for hit in hits for h in hit["depth_hints"]]
    if "manifest_keyword" in all_hints:
        return "manifest"
    if "structural_depth_check" in all_hints or \
            "structural_header_parse" in all_hints:
        return "structural"
    if "smoke_file_exists" in all_hints:
        return "smoke"
    return "smoke"  # default lowest


# ----------------------------------------------------------------------
# Bucket export button categories so the report groups them sensibly.
# ----------------------------------------------------------------------
BUTTON_FAMILY = {
    "CmdGIS":       "GIS (.tab spatial)",
    "CmdGISPeople": "GIS (.tab spatial)",
    "CmdNeo4j":     "Neo4j (CSVs)",
    "CmdGUESS":     "GUESS / Gephi-family",
    "CmdGephi":     "GUESS / Gephi-family",
    "CmdPajek":     "Pajek",
    "CmdUTF8Pajek": "Pajek",
    "CmdUCINet":    "UCINet",
    "CmdImport":           "Import-list",
    "CmdImportPeople":     "Import-list",
    "CmdImportPlaces":     "Import-list",
    "CmdImportEntryCodes": "Import-list",
    "CmdImportList":       "Import-list",
    "CmdImportOffices":    "Import-list",
    "CmdImportAssociations":"Import-list",
    "CmdImportStatusCodes":"Import-list",
    "CmdImportTextCategories":"Import-list",
    "CmdImportPlaceOffice":"Import-list",
    "CmdImportPlacePeople":"Import-list",
    "CmdSaveEntryCodes":   "Save-list",
    "CmdSaveOffices":      "Save-list",
    "CmdSaveStatusCodes":  "Save-list",
    "CmdSaveTextCategories":"Save-list",
    "CmdSaveAssociations": "Save-list",
}


def main() -> int:
    global BUTTONS_BY_FORM, FORMS
    BUTTONS_BY_FORM = parse_taxonomy()
    FORMS = sorted(BUTTONS_BY_FORM.keys())

    golden_files = []
    if GOLDEN_DIR.exists():
        golden_files = [f.name for f in GOLDEN_DIR.iterdir()
                         if f.is_file()]

    coverage = scan_tests()

    rows: list[dict] = []
    for form in FORMS:
        for btn in BUTTONS_BY_FORM[form]:
            hits = coverage.get((form, btn), [])
            depth = classify_depth(hits, golden_files, btn, form)
            rows.append({
                "form": form,
                "button": btn,
                "family": BUTTON_FAMILY.get(btn, "uncategorised"),
                "depth": depth,
                "n_test_files_referencing": len(hits),
                "test_files": [h["test_file"] for h in hits],
            })

    by_form_depth: dict[str, dict[str, int]] = defaultdict(
        lambda: defaultdict(int))
    by_family_depth: dict[str, dict[str, int]] = defaultdict(
        lambda: defaultdict(int))
    by_depth_total: dict[str, int] = defaultdict(int)
    for r in rows:
        by_form_depth[r["form"]][r["depth"]] += 1
        by_family_depth[r["family"]][r["depth"]] += 1
        by_depth_total[r["depth"]] += 1

    summary = {
        "total_buttons": len(rows),
        "by_depth_total": dict(by_depth_total),
        "by_form_depth": {f: dict(d) for f, d in by_form_depth.items()},
        "by_family_depth": {f: dict(d) for f, d in by_family_depth.items()},
        "depth_ladder_low_to_high": [
            "none", "smoke", "structural", "manifest", "byte_golden",
        ],
    }
    out = {"summary": summary, "rows": rows}
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"wrote {OUT_JSON}")

    # ----- markdown -----
    md: list[str] = []
    md.append("# Export coverage matrix")
    md.append("")
    md.append("Generated by `analysis/export_coverage_matrix.py`.")
    md.append("Static parse only — depth classification is best-effort "
              "based on keyword hints in the test files.")
    md.append("")
    md.append("**Caveat — over-attribution risk for `manifest`.** The "
              "heuristic classifies a (form, button) pair as `manifest` "
              "if ANY test file mentioning both has a manifest-style "
              "keyword (`_REQUIRED_COLUMNS`, etc.).  A single test file "
              "containing manifest tests for one button and smoke tests "
              "for another will over-attribute the smoke button as "
              "manifest.  When the count surprises you, drill into "
              "the listed test files to confirm the actual depth.")
    md.append("")
    md.append("## Depth ladder")
    md.append("")
    md.append("`none` < `smoke` < `structural` < `manifest` < `byte_golden`")
    md.append("")
    md.append("## Headline counts")
    md.append("")
    md.append(f"- Total LookAt* export buttons: **{summary['total_buttons']}**")
    for depth in summary["depth_ladder_low_to_high"]:
        n = by_depth_total.get(depth, 0)
        md.append(f"- depth `{depth}`: {n}")
    md.append("")
    md.append("## By export family")
    md.append("")
    md.append("| Family | Total | none | smoke | structural | manifest | byte_golden |")
    md.append("|---|---:|---:|---:|---:|---:|---:|")
    for fam, dd in sorted(by_family_depth.items()):
        total = sum(dd.values())
        md.append(
            f"| {fam} | {total} | {dd.get('none',0)} | "
            f"{dd.get('smoke',0)} | {dd.get('structural',0)} | "
            f"{dd.get('manifest',0)} | {dd.get('byte_golden',0)} |"
        )
    md.append("")
    md.append("## By form")
    md.append("")
    md.append("| Form | Total | none | smoke | structural | manifest | byte_golden |")
    md.append("|---|---:|---:|---:|---:|---:|---:|")
    for form in FORMS:
        dd = by_form_depth[form]
        total = sum(dd.values())
        md.append(
            f"| `{form}` | {total} | {dd.get('none',0)} | "
            f"{dd.get('smoke',0)} | {dd.get('structural',0)} | "
            f"{dd.get('manifest',0)} | {dd.get('byte_golden',0)} |"
        )
    md.append("")
    md.append("## Per-button detail")
    md.append("")
    for form in FORMS:
        md.append(f"### {form}")
        md.append("")
        md.append("| Button | Family | Depth | Tests |")
        md.append("|---|---|---|---|")
        for r in rows:
            if r["form"] != form:
                continue
            tests = ", ".join(f"`{Path(p).name}`"
                              for p in r["test_files"][:3]) or "—"
            if len(r["test_files"]) > 3:
                tests += f" (+{len(r['test_files']) - 3} more)"
            md.append(
                f"| `{r['button']}` | {r['family']} "
                f"| `{r['depth']}` | {tests} |"
            )
        md.append("")
    md.append("## Buttons with depth=`none` (no test reference)")
    md.append("")
    none_rows = [r for r in rows if r["depth"] == "none"]
    if not none_rows:
        md.append("(none)")
    else:
        md.append("| Form | Button | Family |")
        md.append("|---|---|---|")
        for r in none_rows:
            md.append(f"| `{r['form']}` | `{r['button']}` | {r['family']} |")
    md.append("")

    OUT_MD.write_text("\n".join(md), encoding="utf-8")
    print(f"wrote {OUT_MD}")
    print()
    print("=== summary ===")
    for depth in summary["depth_ladder_low_to_high"]:
        print(f"  {depth}: {by_depth_total.get(depth, 0)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
