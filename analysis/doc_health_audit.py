"""Static doc health audit — scan README + AGENTS + analysis/*.md
+ reports/*.md for known-stale references that earlier PRs
invalidated, and for broken local file paths.

Read-only.  Pure file scan.

Outputs:
  - reports/doc_health_audit.json
  - analysis/doc_health_audit.md
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT_JSON = ROOT / "reports" / "doc_health_audit.json"
OUT_MD = ROOT / "analysis" / "doc_health_audit.md"

# Patterns flagged as stale, with the PR that invalidated them.
STALE_PATTERNS: list[tuple[re.Pattern, str, str]] = [
    (re.compile(r"blocked_by_missing_frmBaseMaintenance_vba"),
     "PR X renamed to blocked_by_runtime_priority_triage_pending",
     "stale_label_pre_PR_X"),
    (re.compile(r"Form_Open hangs"),
     "PR AA showed Form_Open does NOT hang; the actual blocker "
     "is CmdRun (Zhu Xi 2471 assocs)",
     "stale_claim_pre_PR_AA"),
    (re.compile(r"PR I'?s? (?:logic_diff|sign-flip)"),
     "PR N invalidated PR I's logic_diff / sign-flip findings; "
     "should be framed as historical only",
     "pr_i_residue_check"),
    (re.compile(r"candidate_same_rule_tie_break_or_aggregation_diff"),
     "PR AI + AJ showed all 14 rows are upstream BIOG_MAIN / "
     "KIN_DATA drift, NOT tie-break.  Bucket label rename "
     "queued for morning.",
     "candidate_label_overtaken_by_PR_AI_AJ"),
]

# Markdown link path patterns: `[label](path)` — verify path
# resolves locally if it's a relative path under the repo.
RE_MD_LINK = re.compile(r"\[([^\]]{1,200})\]\(([^)\s]{1,200})\)")

DOC_FILES = (
    list(ROOT.glob("README.md"))
    + list(ROOT.glob("AGENTS.md"))
    + list(ROOT.glob("FINAL_STATE.md"))
    + list(ROOT.glob("OVERNIGHT_*.md"))
    + sorted(ROOT.glob("analysis/*.md"))
    + sorted(ROOT.glob("reports/*.md"))
    + sorted(ROOT.glob("tests/*.md"))
)


def scan_stale(path: Path) -> list[dict]:
    text = path.read_text(encoding="utf-8", errors="replace")
    findings = []
    for pat, note, kind in STALE_PATTERNS:
        for m in pat.finditer(text):
            line_no = text.count("\n", 0, m.start()) + 1
            line_text = text.splitlines()[line_no - 1] \
                if line_no - 1 < len(text.splitlines()) else ""
            # Skip if the line clearly notes the supersession.
            if any(s in line_text.lower() for s in
                   ("invalidated", "stale", "renamed", "supersed",
                    "historical", "no longer", "no-longer",
                    "old label", "older")):
                continue
            findings.append({
                "doc": str(path.relative_to(ROOT)),
                "line": line_no,
                "kind": kind,
                "matched_text": m.group(0),
                "line_excerpt": line_text.strip()[:200],
                "note": note,
            })
    return findings


def scan_broken_links(path: Path) -> list[dict]:
    text = path.read_text(encoding="utf-8", errors="replace")
    findings = []
    for m in RE_MD_LINK.finditer(text):
        target = m.group(2).strip()
        # Skip absolute URLs.
        if target.startswith(("http://", "https://", "mailto:",
                               "#")):
            continue
        # Skip in-doc anchors after a path.
        target_path = target.split("#", 1)[0]
        if not target_path:
            continue
        # Resolve relative to the doc's parent.
        candidate = (path.parent / target_path).resolve()
        # Also try repo-rooted resolution.
        rooted = (ROOT / target_path).resolve()
        if candidate.exists() or rooted.exists():
            continue
        line_no = text.count("\n", 0, m.start()) + 1
        findings.append({
            "doc": str(path.relative_to(ROOT)),
            "line": line_no,
            "kind": "broken_md_link",
            "label": m.group(1)[:120],
            "target": target,
        })
    return findings


def main() -> int:
    all_findings: list[dict] = []
    for path in DOC_FILES:
        if not path.exists():
            continue
        all_findings += scan_stale(path)
        all_findings += scan_broken_links(path)

    counts: dict[str, int] = {}
    for f in all_findings:
        counts[f["kind"]] = counts.get(f["kind"], 0) + 1

    out = {
        "summary": {
            "n_docs_scanned": len(DOC_FILES),
            "n_findings": len(all_findings),
            "by_kind": counts,
        },
        "findings": all_findings,
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(out, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"wrote {OUT_JSON}")
    print(f"  {len(all_findings)} findings:")
    for k, n in counts.items():
        print(f"    {k}: {n}")

    md = []
    md.append("# Doc health audit (PR AL)")
    md.append("")
    md.append("Static scan of all `.md` docs in the repo for "
              "known-stale phrasing patterns and broken local file "
              "links.  No edits made; this is a worklist for "
              "morning review.")
    md.append("")
    md.append(f"## Headline")
    md.append("")
    md.append(f"- Docs scanned: {len(DOC_FILES)}")
    md.append(f"- Findings: **{len(all_findings)}**")
    md.append("")
    if not all_findings:
        md.append("(none — clean)")
    else:
        md.append("- By kind:")
        for k, n in sorted(counts.items(), key=lambda kv: -kv[1]):
            md.append(f"  - `{k}`: {n}")
        md.append("")
        # Group by doc for actionable rollup.
        by_doc: dict[str, list[dict]] = {}
        for f in all_findings:
            by_doc.setdefault(f["doc"], []).append(f)
        for doc in sorted(by_doc):
            items = by_doc[doc]
            md.append(f"## `{doc}` × {len(items)}")
            md.append("")
            for f in items:
                md.append(f"- line {f['line']} (`{f['kind']}`)")
                if "matched_text" in f:
                    md.append(f"  - match: `{f['matched_text']}`")
                    md.append(f"  - excerpt: {f.get('line_excerpt', '')}")
                if "target" in f:
                    md.append(f"  - target: `{f['target']}` "
                              f"(label: {f['label']})")
                if "note" in f:
                    md.append(f"  - note: {f['note']}")
            md.append("")
    OUT_MD.write_text("\n".join(md), encoding="utf-8")
    print(f"wrote {OUT_MD}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
