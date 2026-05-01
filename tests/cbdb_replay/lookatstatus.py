"""
Python replay of LookAtStatus.CmdQuery_Click
(VBA: analysis/dump/vba/Form_LookAtStatus.vb:1156-1473)

Inputs:
  - status codes via ZZ_STATUS_CODE.c_status_code (picker → table)
  - addresses via ZZ_SCRATCH_ADDR_LIST.c_addr_id (picker → table),
    optionally expanded by sub-units (ChkSubUnits) and/or XY radius
    (ChkXYRef + FrameXY: 1=broad ±0.06, 2=narrow ±0.03)
  - year mode (FrameFilterYears): 1=no year filter, 2=index years,
    3=dynasties
  - year range (gFromStr / gToStr / gFromDynasty / gToDynasty / ...)

Output: rows that would be inserted into ZZ_SCRATCH_STATUS
  c_personid, c_name, c_name_chn, c_index_year, c_sex, c_addr_id,
  c_dy, c_status_code, c_status_desc, c_status_desc_chn, c_source,
  c_index_year_type_code, c_sequence
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, Literal

import pandas as pd
import pyodbc


YearMode = Literal["none", "index", "dynasty"]


@dataclass
class StatusQueryInputs:
    # Picker writes to ZZ_STATUS_CODE.c_status_code
    status_codes: list[int] | None = None

    # Picker writes to ZZ_SCRATCH_ADDR_LIST.c_addr_id (raw selections)
    addr_ids: list[int] | None = None
    include_subunits: bool = False
    use_xy_radius: bool = False
    xy_narrow: bool = True       # FrameXY 2=narrow±0.03, 1=broad±0.06

    # FrameFilterYears: 1=none, 2=index years, 3=dynasties
    year_mode: YearMode = "none"
    from_year: int | None = None
    to_year: int | None = None

    # Dynasty range (only used when year_mode == "dynasty")
    from_dynasty: int = -1
    to_dynasty: int = -1
    from_dynasty_begin: int | None = None
    to_dynasty_end: int | None = None


def _expand_addresses(conn: pyodbc.Connection,
                      addr_ids: Iterable[int],
                      include_subunits: bool,
                      use_xy_radius: bool,
                      xy_narrow: bool) -> list[int]:
    seed = sorted({int(a) for a in addr_ids})
    if not seed:
        return []
    cur = conn.cursor()
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
        eps = 0.03 if xy_narrow else 0.06
        in_clause = ",".join(str(x) for x in expanded)
        cur.execute(f"""
            SELECT DISTINCT A.c_addr_id
            FROM ADDR_CODES AS A, ADDR_CODES AS S
            WHERE S.c_addr_id IN ({in_clause})
              AND A.x_coord >= S.x_coord - {eps}
              AND A.x_coord <= S.x_coord + {eps}
              AND A.y_coord >= S.y_coord - {eps}
              AND A.y_coord <= S.y_coord + {eps}
        """)
        expanded |= {int(r[0]) for r in cur.fetchall()}
        # plus seeds with NULL coords
        cur.execute(f"""
            SELECT c_addr_id FROM ADDR_CODES
            WHERE c_addr_id IN ({in_clause})
              AND (x_coord IS NULL OR y_coord IS NULL)
        """)
        expanded |= {int(r[0]) for r in cur.fetchall()}
    cur.close()
    return sorted(expanded)


def _years_clause(inp: StatusQueryInputs) -> str:
    if inp.year_mode == "index":
        cond = []
        if inp.from_year is not None:
            cond.append(f"BIOG_MAIN.c_index_year >= {int(inp.from_year)}")
        if inp.to_year is not None:
            cond.append(f"BIOG_MAIN.c_index_year <= {int(inp.to_year)}")
        return " AND ".join(cond)
    if inp.year_mode == "dynasty":
        if inp.from_dynasty == -2:
            return "DYNASTIES.c_dy > 0"
        if inp.from_dynasty == -1 and inp.to_dynasty > 0:
            return f"DYNASTIES.c_start < {int(inp.to_dynasty_end or 0)}"
        if inp.from_dynasty > 0 and inp.to_dynasty == -1:
            return f"DYNASTIES.c_end > {int(inp.from_dynasty_begin or 0)}"
        if inp.from_dynasty == inp.to_dynasty and inp.from_dynasty > 0:
            return f"DYNASTIES.c_dy = {int(inp.from_dynasty)}"
        if inp.from_dynasty > 0 and inp.to_dynasty > 0:
            return (f"DYNASTIES.c_end > {int(inp.from_dynasty_begin or 0)} "
                    f"AND DYNASTIES.c_start < {int(inp.to_dynasty_end or 0)}")
    return ""


def run(conn: pyodbc.Connection, inp: StatusQueryInputs) -> pd.DataFrame:
    """Replay LookAtStatus.CmdQuery_Click.

    Returns the rows that would be inserted into ZZ_SCRATCH_STATUS
    (before the cosmetic UPDATE joins backfill descriptions / xy
    counts).
    """
    use_addr = bool(inp.addr_ids)
    use_dynasties = inp.year_mode == "dynasty"
    where_years = _years_clause(inp)

    # Status code constraint via ZZ_STATUS_CODE INNER JOIN STATUS_DATA
    if not inp.status_codes:
        # Without status codes the form returns nothing meaningful (the
        # picker is required); we mirror by returning empty.
        return pd.DataFrame()

    code_in = ",".join(str(int(c)) for c in inp.status_codes)
    addr_ids: list[int] = []
    if use_addr:
        addr_ids = _expand_addresses(
            conn, inp.addr_ids, inp.include_subunits,
            inp.use_xy_radius, inp.xy_narrow,
        )
        if not addr_ids:
            return pd.DataFrame()
    addr_in = ",".join(str(a) for a in addr_ids) if addr_ids else None

    # Build FROM (Access SQL needs nested joins parenthesised).
    base = (
        "STATUS_CODES INNER JOIN STATUS_DATA "
        "ON STATUS_CODES.c_status_code = STATUS_DATA.c_status_code"
    )
    biog = "BIOG_MAIN"
    if use_addr:
        # addr filter: BIOG_MAIN INNER JOIN ZZ_SCRATCH_ADDR ...
        # we inline it as a WHERE IN to avoid mutating ZZ_SCRATCH_ADDR
        # in the test session.
        pass

    if use_dynasties:
        from_clause = (
            f"({base}) INNER JOIN "
            f"(DYNASTIES RIGHT JOIN {biog} ON DYNASTIES.c_dy = {biog}.c_dy) "
            f"ON STATUS_DATA.c_personid = {biog}.c_personid"
        )
    else:
        from_clause = (
            f"({base}) INNER JOIN {biog} "
            f"ON STATUS_DATA.c_personid = {biog}.c_personid"
        )

    where_parts: list[str] = [
        f"STATUS_DATA.c_status_code IN ({code_in})"
    ]
    if addr_in is not None:
        where_parts.append(f"BIOG_MAIN.c_index_addr_id IN ({addr_in})")
    if where_years:
        where_parts.append(f"({where_years})")
    where_clause = " WHERE " + " AND ".join(where_parts)

    sql = f"""
        SELECT
            BIOG_MAIN.c_personid,
            BIOG_MAIN.c_name,
            BIOG_MAIN.c_name_chn,
            BIOG_MAIN.c_index_year,
            IIf(BIOG_MAIN.c_female, 'F', 'M') AS c_sex,
            BIOG_MAIN.c_index_addr_id      AS c_addr_id,
            BIOG_MAIN.c_dy,
            STATUS_DATA.c_status_code,
            STATUS_CODES.c_status_desc,
            STATUS_CODES.c_status_desc_chn,
            STATUS_DATA.c_source,
            BIOG_MAIN.c_index_year_type_code,
            STATUS_DATA.c_sequence
        FROM {from_clause}
        {where_clause}
    """
    return pd.read_sql(sql, conn)
