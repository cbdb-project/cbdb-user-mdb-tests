"""Inventory pytest skip / xfail markers across the test suite.

Static analysis only — parses test files via regex.  Output:

  - reports/hard_form_skip_inventory.json — per-occurrence
    record (test_file, line, marker_kind, raw_text, reason
    string when extractable, surrounding form name when
    extractable, classified blocker hint).
  - analysis/hard_form_skip_inventory.md — human summary
    grouped by form and by blocker class.

No tests run.  No imports.  Pure file scan.
"""
from __future__ import annotations

import json
import re
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TESTS_DIR = ROOT / "tests"
OUT_JSON = ROOT / "reports" / "hard_form_skip_inventory.json"
OUT_MD = ROOT / "analysis" / "hard_form_skip_inventory.md"

# Marker patterns.  Each entry: regex, kind label.  Use raw r"" patterns.
MARKER_PATTERNS = [
    (re.compile(r"@pytest\.mark\.skip(?:if)?\b"),
     "decorator_skip"),
    (re.compile(r"@pytest\.mark\.xfail\b"),
     "decorator_xfail"),
    (re.compile(r"\bpytest\.skip\("),
     "call_skip"),
    (re.compile(r"\bpytest\.xfail\("),
     "call_xfail"),
    (re.compile(r"\bpytest\.mark\.skip(?:if)?\("),
     "factory_skip"),
    (re.compile(r"\bpytest\.mark\.xfail\("),
     "factory_xfail"),
]

# Try to extract a `reason="..."` or first `"..."` arg from the
# matched line plus the next ~5 lines.
REASON_RE = re.compile(r'reason\s*=\s*[fr]?["\'](?P<reason>[^"\']{1,500})["\']')
QUOTED_FIRST = re.compile(r'[fr]?["\'](?P<reason>[^"\']{3,500})["\']')

# Map a reason string to a blocker class.  Order matters — first match wins.
CLASSIFIERS: list[tuple[re.Pattern, str]] = [
    (re.compile(r"\bForm_Open hangs", re.I), "lookatnetworks_form_open_hang_legacy_label"),
    (re.compile(r"\bCmdRun times? out", re.I), "cmdrun_timeout"),
    (re.compile(r"\bCmdQuery (?:times? out|aggregat)", re.I),
        "cmdquery_timeout"),
    (re.compile(r"\bnot present|missing\b", re.I), "fixture_not_available"),
    (re.compile(r"\bblessed golden|TODO|future PR", re.I), "intentional_placeholder"),
    (re.compile(r"\bno (?:rows?|matrix fixture|matching|personids|codes|category|addr)", re.I),
        "fixture_empty"),
    (re.compile(r"\bdiscover_test_inputs|test_inputs\.json", re.I),
        "discovery_input_missing"),
    (re.compile(r"\bdump_metadata|TABLES_JSON|QUERIES_JSON", re.I),
        "metadata_dump_missing"),
    (re.compile(r"\binfra smoke|--include-vba", re.I),
        "include_vba_gate"),
]

# Forms / scope tags worth grouping by — best-effort extraction
# from the surrounding context.
FORM_NAMES = [
    "LookAtEntry", "LookAtStatus", "LookAtTexts", "LookAtAssociations",
    "LookAtOffice", "LookAtPlace", "LookAtKinship", "LookAtNetworks",
    "LookAtAssociationPairs", "LookAtGroupData",
]


def _extract_reason(line: str, lookahead: list[str]) -> str | None:
    blob = " ".join([line] + lookahead)
    m = REASON_RE.search(blob)
    if m:
        return m.group("reason").strip()
    m = QUOTED_FIRST.search(blob)
    if m:
        return m.group("reason").strip()
    return None


def _extract_form(line: str, lookahead_back: list[str]) -> str | None:
    blob = " ".join(lookahead_back + [line])
    for f in FORM_NAMES:
        if f in blob:
            return f
    return None


def _classify_blocker(reason: str | None) -> str | None:
    if not reason:
        return None
    for pat, label in CLASSIFIERS:
        if pat.search(reason):
            return label
    return "unclassified"


def _likely_next_action(blocker: str | None, form: str | None,
                         reason: str | None) -> str:
    if blocker == "lookatnetworks_form_open_hang_legacy_label":
        return ("Stale label — PR AA showed Form_Open does NOT hang.  "
                "Replace skip reason with cmdrun_timeout once a smaller "
                "fixture lets CmdRun complete (PR AB scope was subsumed "
                "by AA — see analysis/lookatnetworks_form_open_hang.md).")
    if blocker == "cmdrun_timeout" and form == "LookAtNetworks":
        return ("Smaller fixture: anchor person with 5–20 assocs "
                "instead of Zhu Xi (2471).  Then attempt unskip.")
    if blocker == "cmdquery_timeout" and form == "LookAtAssociationPairs":
        return ("Pre-filter ASSOC_DATA before the self-join; or pick "
                "two anchors whose mutual edge count is small.  See "
                "test_vba_matrix_all_forms.py:300-308 comment.")
    if blocker == "fixture_not_available":
        return "Provide a fixture (data file / picker code) to enable."
    if blocker == "fixture_empty":
        return ("Fixture-data drift — re-run discover_test_inputs.py or "
                "pick a different anchor for the affected test.")
    if blocker == "include_vba_gate":
        return "Run with --include-vba; gate is intentional."
    if blocker == "discovery_input_missing":
        return "Run analysis/discover_test_inputs.py to regenerate."
    if blocker == "metadata_dump_missing":
        return "Run analysis/dump_metadata.py."
    if blocker == "intentional_placeholder":
        return "Future-PR placeholder; no action."
    return "Read the test for context; no automated next-action heuristic."


def main() -> int:
    findings: list[dict] = []
    files = sorted(TESTS_DIR.rglob("*.py"))
    for path in files:
        try:
            text = path.read_text(encoding="utf-8")
        except Exception as e:
            findings.append({
                "test_file": str(path.relative_to(ROOT)),
                "line": 0, "kind": "_read_error",
                "raw_text": str(e),
            })
            continue
        lines = text.splitlines()
        for idx, line in enumerate(lines, start=1):
            for pat, kind in MARKER_PATTERNS:
                if pat.search(line):
                    lookahead = lines[idx: idx + 6]
                    lookback = lines[max(0, idx - 7): idx - 1]
                    reason = _extract_reason(line, lookahead)
                    form = _extract_form(line, lookback)
                    blocker = _classify_blocker(reason)
                    findings.append({
                        "test_file": str(path.relative_to(ROOT)),
                        "line": idx,
                        "kind": kind,
                        "raw_text": line.strip()[:200],
                        "reason": reason,
                        "form": form,
                        "blocker_class": blocker,
                        "likely_next_action":
                            _likely_next_action(blocker, form, reason),
                    })
                    break  # one kind per line is enough

    summary = {
        "total_findings": len(findings),
        "by_kind": dict(_count_by(findings, "kind")),
        "by_blocker_class": dict(_count_by(findings, "blocker_class")),
        "by_form": dict(_count_by(findings, "form")),
        "by_test_file_top10": dict(sorted(
            _count_by(findings, "test_file").items(),
            key=lambda kv: -kv[1])[:10]),
    }
    out = {"summary": summary, "findings": findings}
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"wrote {OUT_JSON}")

    # ----- markdown summary -----
    md: list[str] = []
    md.append("# Hard-form skip / xfail inventory")
    md.append("")
    md.append("Generated by `analysis/hard_form_skip_inventory.py` "
              "(static parse of `tests/**/*.py`).")
    md.append("")
    md.append("## Headline counts")
    md.append("")
    md.append(f"- Total occurrences: **{summary['total_findings']}**")
    md.append("- By marker kind:")
    for k, n in sorted(summary["by_kind"].items(), key=lambda kv: -kv[1]):
        md.append(f"  - `{k}`: {n}")
    md.append("- By blocker class:")
    for k, n in sorted(summary["by_blocker_class"].items(),
                        key=lambda kv: (-kv[1], str(kv[0]))):
        md.append(f"  - `{k}`: {n}")
    md.append("- By form (best-effort extraction):")
    for k, n in sorted(summary["by_form"].items(),
                        key=lambda kv: (-kv[1], str(kv[0]))):
        md.append(f"  - `{k}`: {n}")
    md.append("")
    md.append("## Top test files by occurrence count")
    md.append("")
    md.append("| Test file | Skip/xfail occurrences |")
    md.append("|---|---:|")
    for k, n in summary["by_test_file_top10"].items():
        md.append(f"| `{k}` | {n} |")
    md.append("")
    md.append("## Per-form rollup")
    md.append("")
    by_form: dict[str | None, list[dict]] = defaultdict(list)
    for f in findings:
        by_form[f["form"]].append(f)
    for form in FORM_NAMES + [None]:
        items = by_form.get(form, [])
        if not items:
            continue
        title = form if form is not None else "(no form detected)"
        md.append(f"### {title}")
        md.append("")
        for f in items:
            md.append(
                f"- `{f['test_file']}:{f['line']}` "
                f"({f['kind']}, blocker = `{f['blocker_class']}`)"
            )
            if f.get("reason"):
                md.append(f"  - reason: {f['reason']}")
            md.append(f"  - next action: {f['likely_next_action']}")
        md.append("")
    md.append("## Per-blocker rollup (cross-form)")
    md.append("")
    by_block: dict[str | None, list[dict]] = defaultdict(list)
    for f in findings:
        by_block[f["blocker_class"]].append(f)
    for blocker in sorted(by_block.keys(),
                           key=lambda b: (b is None, b)):
        items = by_block[blocker]
        md.append(f"### `{blocker}` × {len(items)}")
        md.append("")
        seen_reasons: set[str] = set()
        for f in items:
            tag = f"{f['test_file']}:{f['line']}"
            r = f.get("reason") or "(no reason string)"
            if r in seen_reasons and len(items) > 5:
                continue
            seen_reasons.add(r)
            md.append(f"- `{tag}` — {r[:160]}")
        md.append("")

    OUT_MD.write_text("\n".join(md), encoding="utf-8")
    print(f"wrote {OUT_MD}")
    print()
    print("=== summary ===")
    for k, n in sorted(summary["by_blocker_class"].items(),
                        key=lambda kv: -kv[1]):
        print(f"  {k}: {n}")
    return 0


def _count_by(items: list[dict], key: str) -> dict:
    out: dict = defaultdict(int)
    for f in items:
        out[f.get(key)] += 1
    return out


if __name__ == "__main__":
    sys.exit(main())
