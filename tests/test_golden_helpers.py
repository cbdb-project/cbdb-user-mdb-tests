"""Unit tests for tests/golden_helpers.py (B9: strict golden comparison).

Pure pandas — no Access, no MDB.  Pins: a missing golden FAILS (no
self-bless), strict per-row/column/count diffs FAIL, and even in
drift-tolerant mode a per-cell regression on a SHARED row still FAILS.
"""
from __future__ import annotations

import importlib.util
from pathlib import Path

import pandas as pd
import pytest

REPO = Path(__file__).resolve().parent.parent
_spec = importlib.util.spec_from_file_location(
    "cbdb_golden_helpers", REPO / "tests" / "golden_helpers.py")
gh = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(gh)  # type: ignore[union-attr]

SORT = ["id"]


def _df(rows):
    return pd.DataFrame(rows, columns=["id", "val"])


# --- missing golden / regenerate ------------------------------------- #

def test_missing_golden_fails_and_writes(tmp_path):
    g = tmp_path / "g.csv"
    with pytest.raises(AssertionError, match="golden was missing"):
        gh.assert_matches_golden(_df([(1, "a")]), g, sort_by=SORT)
    assert g.exists()  # written so the author can inspect/commit


def test_regenerate_writes_and_passes(tmp_path):
    g = tmp_path / "g.csv"
    gh.assert_matches_golden(_df([(1, "a")]), g, sort_by=SORT, regenerate=True)
    assert g.exists()
    # second call (now present) compares clean
    gh.assert_matches_golden(_df([(1, "a")]), g, sort_by=SORT)


# --- strict (drift=0) ------------------------------------------------ #

def test_exact_match_passes(tmp_path):
    g = tmp_path / "g.csv"
    gh.assert_matches_golden(_df([(1, "a"), (2, "b")]), g, sort_by=SORT, regenerate=True)
    gh.assert_matches_golden(_df([(2, "b"), (1, "a")]), g, sort_by=SORT)  # order-insensitive


def test_per_row_value_diff_fails(tmp_path):
    g = tmp_path / "g.csv"
    gh.assert_matches_golden(_df([(1, "a"), (2, "b")]), g, sort_by=SORT, regenerate=True)
    with pytest.raises(AssertionError, match="diff vs golden"):
        gh.assert_matches_golden(_df([(1, "a"), (2, "CHANGED")]), g, sort_by=SORT)


def test_row_count_diff_fails_when_strict(tmp_path):
    g = tmp_path / "g.csv"
    gh.assert_matches_golden(_df([(1, "a"), (2, "b")]), g, sort_by=SORT, regenerate=True)
    with pytest.raises(AssertionError, match="row count drift"):
        gh.assert_matches_golden(_df([(1, "a")]), g, sort_by=SORT)


# --- drift-tolerant mode still value-checks shared rows -------------- #

def test_drift_tolerates_count_when_shared_rows_match(tmp_path):
    g = tmp_path / "g.csv"
    gh.assert_matches_golden(
        _df([(1, "a"), (2, "b"), (3, "c")]), g, sort_by=SORT, regenerate=True)
    # one row dropped (count drift ~33%), but shared rows unchanged -> OK
    gh.assert_matches_golden(
        _df([(1, "a"), (2, "b")]), g, sort_by=SORT, allow_count_drift=0.5)


def test_drift_still_catches_per_cell_regression_on_shared_row(tmp_path):
    g = tmp_path / "g.csv"
    gh.assert_matches_golden(
        _df([(1, "a"), (2, "b"), (3, "c")]), g, sort_by=SORT, regenerate=True)
    # count within tolerance, but a SHARED row (id=2) regressed -> must FAIL
    with pytest.raises(AssertionError, match="shared row"):
        gh.assert_matches_golden(
            _df([(1, "a"), (2, "CHANGED")]), g, sort_by=SORT, allow_count_drift=0.5)


def test_drift_with_nonunique_key_warns_count_only(tmp_path):
    # Non-unique sort_by under drift>0 can't safely value-match shared rows;
    # it must WARN (not silently skip) that the comparison was count-only.
    g = tmp_path / "g.csv"
    gh.assert_matches_golden(
        _df([(1, "a"), (1, "b"), (2, "c")]), g, sort_by=SORT, regenerate=True)
    with pytest.warns(UserWarning, match="NOT.*compared|count-only"):
        gh.assert_matches_golden(
            _df([(1, "a"), (1, "b")]), g, sort_by=SORT, allow_count_drift=0.5)
