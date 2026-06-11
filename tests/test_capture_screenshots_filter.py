"""Pin capture_screenshots._select_capture_keys (Step 6 should only re-shoot
bugs still in the current report, not dropped ones)."""
import sys
from pathlib import Path

import pytest

pytest.importorskip("win32gui")
pytest.importorskip("PIL")

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "reports"))

import capture_screenshots as cs  # noqa: E402


def test_select_skips_dropped_bugs():
    # the live build-20260602 report = #2,#19,#20,#22,#23,#24; only the
    # bug15_19 routine (which covers #19) should run.
    assert cs._select_capture_keys({2, 19, 20, 22, 23, 24}) == ["bug15_19"]


def test_select_keeps_only_intersecting_routines():
    assert set(cs._select_capture_keys({7, 11})) == {"bug7", "bug10_11_12"}


def test_select_empty_ids_selects_all():
    # no ids available (e.g. generate_report import failed) -> capture all
    assert cs._select_capture_keys(set()) == list(cs._ALL_CAPTURES)


def test_every_capture_key_has_a_bug_mapping():
    assert set(cs._CAPTURE_BUGS) == set(cs._ALL_CAPTURES)
