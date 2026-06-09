"""Resolve + pin the DATA build identity for a standardized run.

A standardized run must record WHICH build it tested (so a report can be
trusted to belong to a specific data build) and may optionally PIN an
expected build (so an accidental run against the wrong DATA mdb fails
loudly instead of silently producing a mislabeled report).  See
`docs/standardized-testing-remediation-plan.md` task B7 part 2.

Pure helper: no pyodbc, no Access COM.  The build id is the YYYYMMDD
token in the DATA mdb filename (CBDB_<YYYYMMDD>_DATA.mdb), resolved via
the shared `_data_mdb_finder`.

Pin source (first that is set wins):
  1. env var  CBDB_EXPECTED_BUILD
  2. file     data/EXPECTED_BUILD   (a single YYYYMMDD line)
  3. unset    -> no pin (record only)
"""
from __future__ import annotations

import os
import re
from pathlib import Path

_BUILD_RE = re.compile(r"^\d{8}$")


def build_from_name(name: str) -> str | None:
    """Extract the YYYYMMDD build token from a DATA mdb filename/stem.

    'CBDB_20260602_DATA.mdb' / 'CBDB_20260602_DATA' -> '20260602'.
    Returns None if no 8-digit token is present.
    """
    stem = name[:-4] if name.lower().endswith(".mdb") else name
    for part in stem.split("_"):
        if _BUILD_RE.match(part):
            return part
    return None


def current_build(root: Path) -> str | None:
    """Build id of the DATA mdb in root/data/, or None if none found."""
    import sys
    analysis = Path(__file__).resolve().parent
    if str(analysis) not in sys.path:
        sys.path.insert(0, str(analysis))
    try:
        from _data_mdb_finder import find_data_mdb
        data_mdb = find_data_mdb(root)
    except FileNotFoundError:
        return None
    return build_from_name(data_mdb.name)


def expected_build(root: Path) -> str | None:
    """The pinned expected build, or None if no pin is configured."""
    env = os.environ.get("CBDB_EXPECTED_BUILD", "").strip()
    if env:
        return env
    pin_file = root / "data" / "EXPECTED_BUILD"
    if pin_file.exists():
        txt = pin_file.read_text(encoding="utf-8").strip()
        if txt:
            return txt
    return None


def build_pin_error(root: Path) -> str | None:
    """Return a human message if a pin is set and the current build does
    NOT match it; None when there is no pin or it matches.

    A missing DATA mdb while a pin IS set is itself an error.
    """
    exp = expected_build(root)
    if exp is None:
        return None  # no pin -> record-only, never blocks
    cur = current_build(root)
    if cur is None:
        return (
            f"expected DATA build {exp!r} but no CBDB_*_DATA.mdb is present "
            f"in {root / 'data'}"
        )
    if cur != exp:
        return (
            f"DATA build mismatch: expected {exp!r} (pin) but data/ has "
            f"{cur!r}.  Place the expected build in data/, or update the pin "
            f"(env CBDB_EXPECTED_BUILD / data/EXPECTED_BUILD)."
        )
    return None


def build_stamp(root: Path) -> dict:
    """Machine-readable stamp for embedding in reports / pytest JSON."""
    return {
        "cbdb_data_build": current_build(root),
        "cbdb_expected_build": expected_build(root),
    }
