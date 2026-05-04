"""Audit screenshot captions in `reports/generate_report.py::ISSUES`
for consistency with each issue's tier and declared latency state.

Conservative, deterministic, text-only.  No MDB.  No Access COM.  No
NLP-style inference.  Two rules, plain substring matching:

  RULE A (P5 only):
    For each P5 (dormant/latent) issue: if any screenshot's filename
    contains a trigger keyword (popup / runtime / form_open /
    annotated), its caption MUST contain at least one hedge keyword
    (Hypothetical / latent / cannot trigger / can't trigger / not
    currently triggerable / "if ... existed" / users currently
    CAN'T trigger / 潛伏 / 目前不能觸發 / 目前無法觸發 / 不可達 /
    假設).  Otherwise flag.

  RULE B (all issues):
    For ANY issue: if any caption contains an active-trigger phrase
    (users see / popup users see / appears immediately / 立即彈出 /
    立即弹出) AND the issue's title/summary/steps/severity contains
    a latency marker (LATENT / DORMANT / cannot trigger / no UI
    trigger / 潛伏 / 不可達 / 目前不能觸發 / 目前無法觸發), flag.

Out of scope by design:
  - P3 missing-UI issues are exempt from RULE A.  Their runtime
    captures are how we PROVE the button is missing.
  - P0 / P1 / P2 are exempt from RULE A.  Their faux popups +
    annotated runtime captures are the standard reporting pattern
    when the bug is user-visible today.
  - This auditor never reads the MDB / control inventory / VBA
    dump, never spawns Access COM, never decides on its own that
    a screenshot is "out of date".  All it checks is that the
    report is internally consistent: the screenshot caption text
    must agree with the surrounding tier and prose.

Outputs:
  reports/report_screenshot_audit.json   — machine-readable
  analysis/report_screenshot_audit.md    — human-readable
exit code 1 iff any mismatch.

Born 2026-05-04 from the Issue #9 reclassification (P0 -> P5
latent): the prior P0 narrative shipped with a faux 3265 popup
screenshot whose caption asserted present-tense user impact, while
the issue had no realistic UI repro.  This auditor catches that
class of self-contradiction directly from `ISSUES`.
"""
from __future__ import annotations

import datetime as _dt
import json
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "reports"))

from generate_report import ISSUES  # noqa: E402

OUT_JSON = ROOT / "reports" / "report_screenshot_audit.json"
OUT_MD = ROOT / "analysis" / "report_screenshot_audit.md"


# ---------------------------------------------------------------
# Keyword tables — deterministic, substring match.
# ---------------------------------------------------------------
FILENAME_TRIGGER_KEYWORDS = (
    "popup", "runtime", "form_open", "annotated",
)

# Case-insensitive substring match.  Both ASCII and CJK forms.
CAPTION_HEDGE_KEYWORDS = (
    # English
    "Hypothetical",
    "latent",
    "cannot trigger",
    "can't trigger",
    "not currently triggerable",
    "if ... existed",
    "if Issue #",            # narrow heuristic for "if Issue #15 fixed"
    "users currently CAN'T trigger",
    # ZH-Hant + ZH-Hans
    "潛伏", "潜伏",
    "目前不能觸發", "目前不能触发",
    "目前無法觸發", "目前无法触发",
    "不可達", "不可达",
    "假設", "假设",
)

CAPTION_ACTIVE_TRIGGER_PHRASES = (
    # English — present-tense user-impact phrases
    "users see",
    "popup users see",
    "the popup users see",
    "appears immediately",
    # ZH
    "立即彈出", "立即弹出",
)

# Markers that mean "this issue is not user-triggerable today".
# Mixed case + CJK.  The all-uppercase ones (LATENT, DORMANT) are
# matched case-sensitively so they don't accidentally fire on the
# word "latent" appearing in normal prose; everything else is
# case-insensitive substring.
ISSUE_LATENT_MARKERS_CASE_SENSITIVE = (
    "LATENT", "DORMANT",
)
ISSUE_LATENT_MARKERS_CASE_INSENSITIVE = (
    "cannot trigger",
    "can't trigger",
    "CAN'T trigger",
    "no UI trigger",
    "not currently triggerable",
    # ZH
    "潛伏", "潜伏",
    "不可達", "不可达",
    "目前不能觸發", "目前不能触发",
    "目前無法觸發", "目前无法触发",
)


# ---------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------
def _ci_contains(text: str, needles: tuple[str, ...]) -> list[str]:
    low = text.lower()
    return [n for n in needles if n.lower() in low]


def _has_filename_trigger(filename: str) -> list[str]:
    return _ci_contains(filename, FILENAME_TRIGGER_KEYWORDS)


def _caption_has_hedge(caption: str) -> list[str]:
    return _ci_contains(caption, CAPTION_HEDGE_KEYWORDS)


def _caption_has_active_trigger(caption: str) -> list[str]:
    return _ci_contains(caption, CAPTION_ACTIVE_TRIGGER_PHRASES)


def _issue_text_for_latent_check(issue: dict) -> str:
    parts = [
        issue.get("title_en", ""), issue.get("title_zh", ""),
        issue.get("summary_en", ""), issue.get("summary_zh", ""),
        issue.get("severity_en", ""), issue.get("severity_zh", ""),
    ]
    for k in ("steps_en", "steps_zh"):
        for s in issue.get(k) or []:
            parts.append(s)
    return "\n".join(parts)


def _issue_has_latent_marker(issue: dict) -> list[str]:
    text = _issue_text_for_latent_check(issue)
    found: list[str] = []
    for m in ISSUE_LATENT_MARKERS_CASE_SENSITIVE:
        if m in text:
            found.append(m)
    low = text.lower()
    for m in ISSUE_LATENT_MARKERS_CASE_INSENSITIVE:
        if m.lower() in low:
            found.append(m)
    return sorted(set(found))


# ---------------------------------------------------------------
# Main
# ---------------------------------------------------------------
def _build_findings() -> dict:
    rows: list[dict] = []
    mismatches: list[dict] = []
    for issue in ISSUES:
        screenshots = issue.get("screenshots") or []
        tier = issue["tier"]
        latent_markers = _issue_has_latent_marker(issue)
        for fn, caption in screenshots:
            # Some entries (P3 missing-UI screenshots) ship with a
            # None caption — treat as empty string for matching, but
            # preserve the fact in the row so it's visible in the MD.
            caption_str = caption if isinstance(caption, str) else ""
            row = {
                "issue_id": issue["id"],
                "tier": tier,
                "filename": fn,
                "caption_excerpt": caption_str[:200],
                "caption_is_none": caption is None,
                "filename_trigger_keywords": _has_filename_trigger(fn),
                "caption_hedge_keywords":
                    _caption_has_hedge(caption_str),
                "caption_active_trigger_phrases":
                    _caption_has_active_trigger(caption_str),
                "issue_latent_markers": latent_markers,
                "rule_a_violation": False,
                "rule_b_violation": False,
            }
            # Rule A: P5 only.
            if tier == "P5_dormant_or_latent":
                if (row["filename_trigger_keywords"]
                        and not row["caption_hedge_keywords"]):
                    row["rule_a_violation"] = True
                    mismatches.append({
                        "rule": "A",
                        "issue_id": issue["id"],
                        "tier": tier,
                        "filename": fn,
                        "filename_trigger_keywords":
                            row["filename_trigger_keywords"],
                        "caption_excerpt": caption_str[:300],
                        "explanation": (
                            "P5 issue + screenshot filename suggests "
                            "active trigger (popup / runtime / "
                            "form_open / annotated) but caption "
                            "contains none of the required hedge "
                            "keywords."
                        ),
                    })
            # Rule B: all issues.
            if (row["caption_active_trigger_phrases"]
                    and row["issue_latent_markers"]):
                row["rule_b_violation"] = True
                mismatches.append({
                    "rule": "B",
                    "issue_id": issue["id"],
                    "tier": tier,
                    "filename": fn,
                    "caption_excerpt": caption[:300],
                    "active_trigger_phrases":
                        row["caption_active_trigger_phrases"],
                    "issue_latent_markers": latent_markers,
                    "explanation": (
                        "Self-contradiction: caption asserts present-"
                        "tense user impact while issue text "
                        "(summary/steps/severity/title) declares the "
                        "bug latent or untriggerable today."
                    ),
                })
            rows.append(row)

    summary = {
        "n_issues_total": len(ISSUES),
        "n_issues_with_screenshots": sum(
            1 for i in ISSUES if i.get("screenshots")),
        "n_screenshots_total": len(rows),
        "n_mismatches": len(mismatches),
        "n_rule_a_violations": sum(
            1 for m in mismatches if m["rule"] == "A"),
        "n_rule_b_violations": sum(
            1 for m in mismatches if m["rule"] == "B"),
    }
    return {"summary": summary, "mismatches": mismatches,
            "all_rows": rows}


def _render_md(findings: dict) -> str:
    s = findings["summary"]
    mismatches = findings["mismatches"]
    rows = findings["all_rows"]

    lines: list[str] = []
    lines.append("# Report screenshot consistency audit")
    lines.append("")
    lines.append("**Generated:** "
                 f"{_dt.datetime.now(_dt.timezone.utc).isoformat(timespec='seconds')}")
    lines.append("**Generator:** "
                 "`analysis/audit_report_screenshot_consistency.py`")
    lines.append("**Companion JSON:** "
                 "`reports/report_screenshot_audit.json`")
    lines.append("")
    lines.append("This audit cross-checks every screenshot caption "
                 "in `reports/generate_report.py::ISSUES` against "
                 "the surrounding issue's tier and declared latency "
                 "state.  Conservative text-only rules; no MDB, no "
                 "Access COM, no NLP.")
    lines.append("")

    lines.append("## Rules")
    lines.append("")
    lines.append("**Rule A — P5 only.** If a screenshot filename "
                 "contains any of "
                 f"`{', '.join(FILENAME_TRIGGER_KEYWORDS)}`, its "
                 "caption must contain at least one hedge keyword: "
                 f"`{', '.join(CAPTION_HEDGE_KEYWORDS)}`.  Rationale: "
                 "the typical bug-report pattern is annotated "
                 "runtime + faux popup; for P5 issues that are "
                 "gated/dormant today, the caption MUST disclaim "
                 "that the popup is hypothetical or that the "
                 "trigger isn't currently reachable.")
    lines.append("")
    lines.append("**Rule B — all issues.** If a caption contains "
                 "any active-trigger phrase "
                 f"(`{', '.join(CAPTION_ACTIVE_TRIGGER_PHRASES)}`) "
                 "AND the issue's title/summary/steps/severity "
                 "contains any latency marker "
                 f"(`{', '.join(ISSUE_LATENT_MARKERS_CASE_SENSITIVE)}`"
                 " case-sensitive, plus "
                 f"`{', '.join(ISSUE_LATENT_MARKERS_CASE_INSENSITIVE)}`"
                 " case-insensitive), flag.  Rationale: the literal "
                 "self-contradiction pattern (Issue #9 class).")
    lines.append("")
    lines.append("**Out of scope by design:**")
    lines.append("- **P3 missing UI** is exempt from Rule A — runtime "
                 "captures are how we PROVE a button is missing.")
    lines.append("- **P0 / P1 / P2** are exempt from Rule A — faux "
                 "popups + annotated runtime are the standard "
                 "reporting pattern when the bug is user-visible "
                 "today.")
    lines.append("- No MDB / Access COM checks, no image-content "
                 "inference.")
    lines.append("")

    lines.append("## Summary")
    lines.append("")
    lines.append("| Metric | Value |")
    lines.append("|---|---:|")
    lines.append(f"| Issues total | {s['n_issues_total']} |")
    lines.append("| Issues with at least one screenshot | "
                 f"{s['n_issues_with_screenshots']} |")
    lines.append(f"| Screenshots inspected | {s['n_screenshots_total']} |")
    lines.append(f"| Rule A violations (P5) | {s['n_rule_a_violations']} |")
    lines.append(f"| Rule B violations (all) | {s['n_rule_b_violations']} |")
    lines.append(f"| **Mismatches** | **{s['n_mismatches']}** |")
    lines.append("")

    if mismatches:
        lines.append("## Mismatches")
        lines.append("")
        for m in mismatches:
            lines.append(f"### Rule {m['rule']} — Issue #{m['issue_id']} "
                         f"(`{m['tier']}`)")
            lines.append("")
            lines.append(f"- **File:** `{m['filename']}`")
            if m["rule"] == "A":
                lines.append("- **Filename trigger keywords:** "
                             f"`{m['filename_trigger_keywords']}`")
            else:
                lines.append("- **Active-trigger phrases:** "
                             f"`{m['active_trigger_phrases']}`")
                lines.append("- **Issue latent markers:** "
                             f"`{m['issue_latent_markers']}`")
            lines.append(f"- **Why:** {m['explanation']}")
            lines.append("- **Caption excerpt:**")
            lines.append("")
            lines.append("  > "
                         + m["caption_excerpt"].replace("\n", " "))
            lines.append("")
    else:
        lines.append("## Mismatches")
        lines.append("")
        lines.append("**Clean — no mismatches.** Every screenshot "
                     "caption is consistent with its issue's tier "
                     "and latency state under the active rules.")
        lines.append("")

    lines.append("## Per-screenshot inventory")
    lines.append("")
    lines.append("| Issue | Tier | Filename | "
                 "Filename trigger | Caption hedge | "
                 "Active phrase | Issue latent? | Rule A | Rule B |")
    lines.append("|---:|---|---|---|---|---|---|---|---|")
    for r in rows:
        ft = ", ".join(r["filename_trigger_keywords"]) or "—"
        ch = ", ".join(r["caption_hedge_keywords"]) or "—"
        ap = ", ".join(r["caption_active_trigger_phrases"]) or "—"
        lm = ", ".join(r["issue_latent_markers"]) or "—"
        ra = "❌" if r["rule_a_violation"] else "✓"
        rb = "❌" if r["rule_b_violation"] else "✓"
        lines.append(
            f"| #{r['issue_id']} | `{r['tier']}` | `{r['filename']}` | "
            f"{ft} | {ch} | {ap} | {lm} | {ra} | {rb} |")
    lines.append("")

    return "\n".join(lines)


def main() -> int:
    findings = _build_findings()
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(
        json.dumps(findings, indent=2, ensure_ascii=False),
        encoding="utf-8")
    OUT_MD.parent.mkdir(parents=True, exist_ok=True)
    OUT_MD.write_text(_render_md(findings), encoding="utf-8")

    s = findings["summary"]
    print("=== report screenshot consistency audit ===")
    print(f"  issues with screenshots: "
          f"{s['n_issues_with_screenshots']}/{s['n_issues_total']}")
    print(f"  screenshots inspected:   {s['n_screenshots_total']}")
    print(f"  rule A violations (P5):  {s['n_rule_a_violations']}")
    print(f"  rule B violations (all): {s['n_rule_b_violations']}")
    print(f"  total mismatches:        {s['n_mismatches']}")
    print(f"  wrote {OUT_JSON}")
    print(f"  wrote {OUT_MD}")
    return 0 if s["n_mismatches"] == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
