"""
Shared helpers for the LookAt* replay modules.

Most LookAt forms follow the same pre-query pattern:
  1. Picker writes selected codes to a ZZ_*_CODE / ZZ_SCRATCH_<X>
     scratch table.
  2. If addresses are selected, expand them via ZZZ_BELONGS_TO
     (sub-units) and/or XY proximity in ADDR_CODES.
  3. Build a year-range WHERE clause from one of three modes:
     entry years, index years, or dynasty range.

These helpers are extracted so each LookAt replay only needs the
form-specific JOIN structure + result columns.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, Literal

import pyodbc

YearMode = Literal["none", "entry", "index", "dynasty"]


# ----------------------------------------------------------------------
# Address expansion (used by LookAtEntry, LookAtStatus, LookAtTexts,
# LookAtPlace, LookAtAssociations, LookAtOffice, LookAtNetworks)
# ----------------------------------------------------------------------

def expand_addresses(conn: pyodbc.Connection,
                     addr_ids: Iterable[int],
                     *,
                     include_subunits: bool = False,
                     use_xy_radius: bool = False,
                     xy_eps: float = 0.03) -> list[int]:
    """Return the expanded list of addr_ids that the form would have
    produced in ZZ_SCRATCH_ADDR_LIST.

    - If ``include_subunits`` is True, replace each seed addr with all
      addresses that belong to it (per ZZZ_BELONGS_TO).
    - If ``use_xy_radius`` is True, additionally include any
      ADDR_CODES rows whose (x, y) is within ±xy_eps degrees of any
      expanded seed (LookAtPlace uses ±0.03; LookAtStatus etc. let
      the user toggle ±0.03 vs ±0.06 via FrameXY).
    """
    seed = sorted({int(a) for a in addr_ids})
    if not seed:
        return []
    cur = conn.cursor()
    try:
        if include_subunits:
            in_clause = ",".join(str(x) for x in seed)
            cur.execute(
                f"SELECT DISTINCT c_addr_id FROM ZZZ_BELONGS_TO "
                f"WHERE c_belongs_to IN ({in_clause})"
            )
            expanded = {int(r[0]) for r in cur.fetchall()}
        else:
            expanded = set(seed)

        if use_xy_radius and expanded:
            in_clause = ",".join(str(x) for x in expanded)
            cur.execute(f"""
                SELECT DISTINCT A.c_addr_id
                FROM ADDR_CODES AS A, ADDR_CODES AS S
                WHERE S.c_addr_id IN ({in_clause})
                  AND A.x_coord >= S.x_coord - {xy_eps}
                  AND A.x_coord <= S.x_coord + {xy_eps}
                  AND A.y_coord >= S.y_coord - {xy_eps}
                  AND A.y_coord <= S.y_coord + {xy_eps}
            """)
            expanded |= {int(r[0]) for r in cur.fetchall()}
            cur.execute(f"""
                SELECT c_addr_id FROM ADDR_CODES
                WHERE c_addr_id IN ({in_clause})
                  AND (x_coord IS NULL OR y_coord IS NULL)
            """)
            expanded |= {int(r[0]) for r in cur.fetchall()}
    finally:
        cur.close()
    return sorted(expanded)


# ----------------------------------------------------------------------
# Year / dynasty WHERE clause
# ----------------------------------------------------------------------

@dataclass
class YearFilter:
    """Common shape for year filtering across forms."""
    mode: YearMode = "none"
    from_year: int | None = None
    to_year: int | None = None
    # used when mode == "dynasty"
    from_dynasty: int = -1
    to_dynasty: int = -1
    from_dynasty_begin: int | None = None
    to_dynasty_end: int | None = None


def years_where(yf: YearFilter,
                *,
                index_col: str = "BIOG_MAIN.c_index_year",
                entry_col: str = "ENTRY_DATA.c_year",
                dynasty_table: str = "DYNASTIES",
                dynasty_dy_col: str = "BIOG_MAIN.c_dy") -> str:
    """Build the WHERE clause fragment for the given year filter."""
    if yf.mode == "entry":
        cond = []
        if yf.from_year is not None:
            cond.append(f"{entry_col} >= {int(yf.from_year)}")
        if yf.to_year is not None:
            cond.append(f"{entry_col} <= {int(yf.to_year)}")
        return " AND ".join(cond)
    if yf.mode == "index":
        cond = []
        if yf.from_year is not None:
            cond.append(f"{index_col} >= {int(yf.from_year)}")
        if yf.to_year is not None:
            cond.append(f"{index_col} <= {int(yf.to_year)}")
        return " AND ".join(cond)
    if yf.mode == "dynasty":
        if yf.from_dynasty == -2:
            return f"{dynasty_dy_col} > 0"
        if yf.from_dynasty == -1 and yf.to_dynasty > 0:
            return f"{dynasty_table}.c_start < {int(yf.to_dynasty_end or 0)}"
        if yf.from_dynasty > 0 and yf.to_dynasty == -1:
            return f"{dynasty_table}.c_end > {int(yf.from_dynasty_begin or 0)}"
        if yf.from_dynasty == yf.to_dynasty and yf.from_dynasty > 0:
            return f"{dynasty_table}.c_dy = {int(yf.from_dynasty)}"
        if yf.from_dynasty > 0 and yf.to_dynasty > 0:
            return (f"{dynasty_table}.c_end > {int(yf.from_dynasty_begin or 0)} "
                    f"AND {dynasty_table}.c_start < {int(yf.to_dynasty_end or 0)}")
    return ""


def in_clause(values: Iterable[int]) -> str:
    """Render an IN (...) clause of ints. Empty list returns '(NULL)'
    (an always-false guard)."""
    vals = [int(v) for v in values]
    return "(" + ",".join(str(v) for v in vals) + ")" if vals else "(NULL)"
