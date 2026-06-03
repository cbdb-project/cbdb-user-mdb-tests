"""Side-effect-free helper for locating the DATA mdb.

Imported by both tests/conftest.py and analysis/discover_test_inputs.py
so both use identical selection logic with no risk of desync.

No pyodbc, no module-level file-system access — safe to import anywhere.
"""
from __future__ import annotations

from pathlib import Path


def find_data_mdb(root: Path) -> Path:
    """Return the CBDB_*_DATA.mdb in root/data/.

    When multiple files exist (old build not cleaned up), picks the
    newest by YYYYMMDD embedded in the filename.

    Raises FileNotFoundError when data/ contains no matching file.
    """
    matches = list((root / "data").glob("CBDB_*_DATA.mdb"))
    if not matches:
        raise FileNotFoundError(f"No CBDB_*_DATA.mdb found in {root / 'data'}")
    if len(matches) == 1:
        return matches[0]

    def _date_key(p: Path) -> str:
        parts = p.stem.split("_")          # ["CBDB", "YYYYMMDD", "DATA"]
        return parts[1] if len(parts) >= 2 else p.stem

    chosen = sorted(matches, key=_date_key)[-1]
    return chosen
