from __future__ import annotations

from pathlib import Path


REPO = Path(__file__).resolve().parent.parent


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
