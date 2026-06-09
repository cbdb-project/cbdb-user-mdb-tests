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
import warnings
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

    if regenerate:
        write_csv(df_norm, golden_path)
        return  # explicit bless

    if not golden_path.exists():
        # A missing golden must FAIL in a standardized run, NOT self-bless.
        # Write it so the author can inspect/commit it, but fail loudly so a
        # fresh checkout can't silently "pass" a never-compared case.
        write_csv(df_norm, golden_path)
        raise AssertionError(
            f"golden was missing -- wrote {golden_path}; inspect it, commit "
            f"it, and re-run.  (Use --regenerate-goldens to bless intentionally.)"
        )

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

    # If a count drift was allowed, tolerate the COUNT difference but STILL
    # value-check the rows present in BOTH frames (matched on sort_by) -- a
    # per-cell regression on a shared row must never be silently skipped
    # (the old behaviour returned here without any value comparison).
    if drift > 0:
        key_cols = [c for c in sort_by if c in actual.columns]
        a_idx = actual.set_index(key_cols) if key_cols else None
        e_idx = expected.set_index(key_cols) if key_cols else None
        if a_idx is not None and a_idx.index.is_unique and e_idx.index.is_unique:
            common = a_idx.index.intersection(e_idx.index)
            val_cols = [c for c in expected.columns if c not in key_cols]
            a_common = a_idx.loc[common, val_cols].sort_index()
            e_common = e_idx.loc[common, val_cols].sort_index()
            cmp = a_common.compare(e_common, keep_shape=False, keep_equal=False)
            if not cmp.empty:
                raise AssertionError(
                    f"value diffs on {len(cmp)} shared row(s) vs golden "
                    f"{golden_path.name} (count drift within tolerance, but "
                    f"matched rows differ):\n{cmp.head(20).to_string()}"
                )
        else:
            # Degenerate key (empty or non-unique sort_by): we can't safely
            # value-match shared rows, so only the count check ran.  Warn so a
            # maintainer notices the comparison was weaker than per-row strict.
            warnings.warn(
                f"{golden_path.name}: allow_count_drift>0 with a non-unique / "
                f"empty sort_by key {sort_by} -- shared-row values were NOT "
                f"compared (count-only).  Use a unique sort_by for strict drift.",
                stacklevel=2,
            )
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
