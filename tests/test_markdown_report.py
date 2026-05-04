from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path


REPO = Path(__file__).resolve().parent.parent
AUDIT_SCRIPT = REPO / "analysis" / "audit_report_code_labels.py"
AUDIT_JSON = REPO / "reports" / "report_code_label_audit.json"
SCREENSHOT_AUDIT_SCRIPT = (
    REPO / "analysis" / "audit_report_screenshot_consistency.py")
SCREENSHOT_AUDIT_JSON = (
    REPO / "reports" / "report_screenshot_audit.json")


def test_markdown_toc_anchors_match_github_slug_shape():
    """GitHub preserves underscores and turns each surviving space into
    one hyphen after punctuation removal, so em-dash-separated headings
    get double hyphens in their anchors.

    Updated 2026-05-02: Issue #1 was demoted to P5 (DORMANT) — the
    heading now carries the DORMANT suffix, and the anchor reflects
    that. The test is here to guard the slug algorithm itself, not the
    specific wording of any one issue.
    """
    text = (REPO / "reports/CBDB_Issues_Report_EN.md").read_text(
        encoding="utf-8"
    )

    # Tier-level anchor: em dash → double hyphen
    assert (
        "- [P0 — Silent data corruption](#p0--silent-data-corruption)"
        in text
    )

    # Issue #1 lives in P5 now. The full heading + matching TOC entry
    # both have to be present and match the GitHub slug shape (lower-
    # cased, underscores preserved, em dash → double hyphen, apostrophe
    # stripped).
    issue1_heading = (
        "### Issue #1 — View_StatusData would display last-year "
        "range in the first-year column — DORMANT (no source "
        "rows trigger it on this dump)"
    )
    assert issue1_heading in text, "Issue #1 heading wording drifted"

    issue1_toc = (
        "- [Issue #1 — View_StatusData would display last-year "
        "range in the first-year column — DORMANT (no source "
        "rows trigger it on this dump)]"
        "(#issue-1--view_statusdata-would-display-last-year-range-"
        "in-the-first-year-column--dormant-no-source-rows-trigger-"
        "it-on-this-dump)"
    )
    assert issue1_toc in text, "Issue #1 TOC anchor drifted"


def _issue_block(text: str, issue_num: int) -> str:
    """Return the substring of `text` between `### Issue #N — ...`
    and the next `### Issue #` heading (or EOF).  Used to scope
    regression assertions to a single issue block."""
    start_marker = f"### Issue #{issue_num} —"
    start = text.find(start_marker)
    assert start != -1, f"Issue #{issue_num} heading not found in report"
    next_start = text.find("### Issue #", start + len(start_marker))
    return text[start: next_start if next_start != -1 else len(text)]


def test_issue_20_status_code_40_uses_civil_office_label():
    """Issue #20 is about a LookAtStatus / `c_status_code = 40`
    fixture.  STATUS_CODES.c_status_code = 40 is `civil office /
    [為官者：文]`.  The Provincial-Graduate / 进士 wording that
    earlier drafts of the steps used was a copy-paste error from
    Issue #9 (which IS about jinshi / `c_entry_code = 36`).  This
    test pins the corrected wording so the mistake doesn't recur.

    Scoped to the Issue #20 block only — Issue #9's correct
    `c_entry_code = 36 / jinshi / 進士` mentions must not be
    affected.
    """
    for path in (
        REPO / "reports/CBDB_Issues_Report_EN.md",
        REPO / "reports/CBDB_Issues_Report_ZH-Hant.md",
    ):
        text = path.read_text(encoding="utf-8")
        block = _issue_block(text, 20)

        # Forbidden incorrect wordings (the EN and ZH variants of
        # the original copy-paste error from Issue #9).
        assert "status code **40** (Provincial Graduate / 进士)" \
            not in block, (
                f"{path.name}: Issue #20 still contains the stale "
                f"'status code **40** (Provincial Graduate / 进士)' "
                f"wording — that's Issue #9's jinshi label, not "
                f"status 40 (which is civil office)."
            )
        assert "status code **40**（进士）" not in block, (
            f"{path.name}: Issue #20 still contains the stale "
            f"'status code **40**（进士）' wording — see comment "
            f"above."
        )
        # Also block the s2twp-traditionalised form, since the ZH
        # report goes through OpenCC.
        assert "status code **40**（進士）" not in block, (
            f"{path.name}: Issue #20 still contains the stale "
            f"'status code **40**（進士）' wording (s2twp form)."
        )

        # Required correct labels — the exact phrasing comes from
        # STATUS_CODES.c_status_desc / c_status_desc_chn for
        # c_status_code = 40.
        assert "civil office" in block, (
            f"{path.name}: Issue #20 must mention 'civil office' "
            f"(STATUS_CODES.c_status_desc for status_code=40); "
            f"the corrected steps wording is required."
        )
        assert "[為官者：文]" in block, (
            f"{path.name}: Issue #20 must mention '[為官者：文]' "
            f"(STATUS_CODES.c_status_desc_chn for status_code=40); "
            f"the corrected steps wording is required."
        )


def test_report_code_labels_audit_clean():
    """Run `analysis/audit_report_code_labels.py` and assert it
    finds no mismatches between the report's hardcoded code
    labels and the MDB dictionary tables.

    Born from the Issue #20 status-code-40 wording mistake (it
    originally said `Provincial Graduate / 进士`, which is
    Issue #9's jinshi label, not the actual STATUS_CODES desc
    for status_code=40 which is `civil office / [為官者：文]`).
    The auditor's curated manifest covers every (table,
    code_value) the report currently hardcodes.

    Skipped if the user MDB isn't present (matches the
    behaviour of other MDB-touching tests in this file's
    sibling files).
    """
    import pytest
    user_mdb = REPO / "data" / "CBDB_BJ_User.mdb"
    if not user_mdb.exists():
        pytest.skip(f"{user_mdb} not present")

    # Re-run the auditor so the test reflects the live state of
    # the report + MDB, not a stale cached JSON.
    r = subprocess.run(
        [sys.executable, str(AUDIT_SCRIPT)],
        capture_output=True, text=True, timeout=60,
    )
    # Auditor exits non-zero on any mismatch; the JSON is still
    # written either way.
    assert AUDIT_JSON.exists(), (
        f"audit script did not write {AUDIT_JSON}; stderr:\n"
        f"{r.stderr[-500:]}"
    )

    audit = json.loads(AUDIT_JSON.read_text(encoding="utf-8"))
    s = audit["summary"]
    mismatches = audit["mismatches"]

    assert mismatches == [], (
        f"report code-label audit found {len(mismatches)} "
        f"mismatch(es).  Each mismatch is a place where the "
        f"report's hardcoded label disagrees with the MDB "
        f"dictionary table.  Fix either the report wording or "
        f"the manifest in analysis/audit_report_code_labels.py "
        f"(if the data changed).  Detail:\n"
        + "\n".join(
            f"  - issue #{m['issue_id']} {m['table']}."
            f"{m['code_col']}={m['code_value']} [{m['lang']}]: "
            f"missing={m.get('missing_expected_labels')} "
            f"forbidden={m.get('found_forbidden_labels')} "
            f"mdb_desc_present={m.get('mdb_desc_present_in_block')}"
            for m in mismatches
        )
    )
    # All language checks must have passed (manifest entries × 2).
    assert s["n_per_lang_checks_passed"] == s["n_per_lang_checks_total"], (
        f"audit summary disagrees with mismatches: "
        f"{s['n_per_lang_checks_passed']} / "
        f"{s['n_per_lang_checks_total']} per-lang checks passed"
    )


def test_report_screenshot_consistency_audit_clean():
    """Run `analysis/audit_report_screenshot_consistency.py` and
    assert it finds no mismatches between screenshot captions and
    each issue's tier / declared latency state.

    Born from the Issue #9 reclassification (P0 -> P5 latent on
    2026-05-04): the prior P0 narrative shipped with a faux 3265
    popup screenshot whose caption asserted present-tense user
    impact, while the issue had no realistic UI repro on the
    current dump.  This auditor's two text-only rules catch that
    self-contradiction class directly from
    `reports/generate_report.py::ISSUES`.

    Conservative scope:
      - Rule A: P5 only — filename trigger keyword (popup /
        runtime / form_open / annotated) requires caption hedge
        keyword (Hypothetical / latent / cannot trigger / 潛伏 / …).
      - Rule B: all issues — caption active-trigger phrase + issue
        latent marker = self-contradiction.
      - P3 missing-UI runtime captures and P0/P1/P2 faux popups
        are exempt from Rule A by design.

    Always runs (no MDB / no Access COM dependency).
    """
    r = subprocess.run(
        [sys.executable, str(SCREENSHOT_AUDIT_SCRIPT)],
        capture_output=True, text=True, timeout=30,
    )
    assert SCREENSHOT_AUDIT_JSON.exists(), (
        f"audit script did not write {SCREENSHOT_AUDIT_JSON}; "
        f"stderr:\n{r.stderr[-500:]}"
    )
    audit = json.loads(SCREENSHOT_AUDIT_JSON.read_text(encoding="utf-8"))
    s = audit["summary"]
    mismatches = audit["mismatches"]

    assert mismatches == [], (
        f"report screenshot audit found {len(mismatches)} "
        f"mismatch(es).  Each mismatch is a screenshot whose "
        f"caption disagrees with its issue's tier or declared "
        f"latency state.  Either re-caption the screenshot, "
        f"remove it, or update the issue's wording in "
        f"reports/generate_report.py.  Detail:\n"
        + "\n".join(
            f"  - rule {m['rule']} #{m['issue_id']} ({m['tier']}) "
            f"file={m['filename']!r} excerpt={m['caption_excerpt'][:120]!r}"
            for m in mismatches
        )
    )
    assert s["n_rule_a_violations"] == 0
    assert s["n_rule_b_violations"] == 0
