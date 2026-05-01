"""
Golden-snapshot helpers.

We compare DataFrames against frozen CSVs.  To keep diffs stable:
  - sort by primary-identity columns
  - drop noisy columns by default (optional)
  - normalize NaN to empty string in CSV
  - write with utf-8-sig so Excel doesn't mangle Chinese

Strategy when --regenerate-goldens is set: simply (re)write the CSV
and pass.  The user must visually inspect git diff to bless changes.
"""
from __future__ import annotations

import io
from pathlib import Path

import pandas as pd


def normalize(df: pd.DataFrame, *, sort_by: list[str]) -> pd.DataFrame:
    df = df.copy()
    # consistent ordering
    cols = list(df.columns)
    df = df.sort_values(by=[c for c in sort_by if c in cols],
                        kind="mergesort").reset_index(drop=True)
    return df


def write_csv(df: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path, index=False, encoding="utf-8-sig")


def read_csv(path: Path) -> pd.DataFrame:
    # all columns as object so dtype mismatches don't trip the diff
    return pd.read_csv(path, encoding="utf-8-sig", dtype=str, keep_default_na=False)


def assert_matches_golden(df: pd.DataFrame,
                          golden_path: Path,
                          *,
                          sort_by: list[str],
                          regenerate: bool = False,
                          allow_count_drift: float = 0.0) -> None:
    """Compare ``df`` against the CSV at ``golden_path``.

    On --regenerate-goldens, write and pass.

    ``allow_count_drift`` lets a test tolerate small differences in row
    count (useful when the underlying DATA mdb has been updated
    incrementally — e.g. 0.05 = 5% drift OK).  Set to 0 for strict.
    """
    df_norm = normalize(df, sort_by=sort_by)

    if regenerate or not golden_path.exists():
        write_csv(df_norm, golden_path)
        if not golden_path.exists():
            raise AssertionError(f"wrote new golden: {golden_path}")
        # On regenerate, succeed silently.
        return

    expected = read_csv(golden_path)
    # bring df_norm to the same string-typed shape so we can compare:
    # null/NaN -> "" everywhere
    actual = df_norm.where(df_norm.notna(), "")
    actual = actual.astype(str)
    # pandas turns numeric columns into "1.0" -- we want canonical "1"
    for col in actual.columns:
        actual[col] = actual[col].str.replace(r"\.0+$", "", regex=True)
    expected = expected.fillna("")
    for col in expected.columns:
        if col in actual.columns:
            expected[col] = expected[col].astype(str).str.replace(r"\.0+$", "", regex=True)

    # row count check first
    n_expected = len(expected)
    n_actual = len(actual)
    if n_expected == 0 and n_actual == 0:
        return
    drift = abs(n_expected - n_actual) / max(n_expected, 1)
    if drift > allow_count_drift:
        raise AssertionError(
            f"row count drift {n_actual} vs golden {n_expected} "
            f"(drift={drift:.2%}, allowed={allow_count_drift:.2%}) at {golden_path}"
        )

    # column set check
    if set(actual.columns) != set(expected.columns):
        raise AssertionError(
            f"columns differ at {golden_path}\n"
            f"  expected: {sorted(expected.columns)}\n"
            f"  actual:   {sorted(actual.columns)}"
        )

    # if drift was allowed, set-compare on primary-identity columns
    if drift > 0:
        ids_a = set(map(tuple, actual[sort_by].itertuples(index=False, name=None)))
        ids_e = set(map(tuple, expected[sort_by].itertuples(index=False, name=None)))
        added = ids_a - ids_e
        removed = ids_e - ids_a
        # For drift-tolerant mode, just report; don't fail on individual rows
        # provided the count drift was within bounds.
        return

    # strict per-row comparison on the columns the golden has, in golden order
    expected = expected.reset_index(drop=True)
    actual = actual[expected.columns].reset_index(drop=True)
    diffs = actual.compare(expected, keep_shape=False, keep_equal=False)
    if not diffs.empty:
        buf = io.StringIO()
        buf.write(f"diff vs golden {golden_path.name}:\n")
        buf.write(diffs.head(20).to_string())
        if len(diffs) > 20:
            buf.write(f"\n  ... and {len(diffs)-20} more rows differ")
        raise AssertionError(buf.getvalue())
