"""
Python replay of LookAtTexts.CmdQuery_Click
(VBA: analysis/dump/vba/Form_LookAtTexts.vb:977)

Inputs:
  - text-bibliography categories via ZZ_TEXT_BIBLCAT_CODES
  - addresses, year mode (same shape as LookAtStatus)

Output rows (would be inserted into ZZ_SCRATCH_BIOG_TEXT_DATA):
  c_personid, c_name, c_name_chn, c_index_year, c_index_year_type_code,
  c_dy, c_sex, c_addr_id, c_textid, c_title, c_title_chn, c_role_id,
  c_role_desc, c_role_desc_chn, c_source
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, Literal

import pandas as pd
import pyodbc

from .lookatstatus import _expand_addresses, _years_clause as _years_clause_status

YearMode = Literal["none", "index", "dynasty"]


@dataclass
class TextsQueryInputs:
    biblcat_codes: list[int] | None = None    # → ZZ_TEXT_BIBLCAT_CODES
    addr_ids: list[int] | None = None
    include_subunits: bool = False
    use_xy_radius: bool = False
    xy_narrow: bool = True
    year_mode: YearMode = "none"
    from_year: int | None = None
    to_year: int | None = None
    from_dynasty: int = -1
    to_dynasty: int = -1
    from_dynasty_begin: int | None = None
    to_dynasty_end: int | None = None


def _years_for_texts(inp: TextsQueryInputs) -> str:
    """LookAtTexts uses BIOG_MAIN.c_dy (not DYNASTIES.c_dy) for the
    'all dynasties' branch."""
    if inp.year_mode == "index":
        cond = []
        if inp.from_year is not None:
            cond.append(f"BIOG_MAIN.c_index_year >= {int(inp.from_year)}")
        if inp.to_year is not None:
            cond.append(f"BIOG_MAIN.c_index_year <= {int(inp.to_year)}")
        return " AND ".join(cond)
    if inp.year_mode == "dynasty":
        if inp.from_dynasty == -2:
            return "BIOG_MAIN.c_dy > 0"
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


def run(conn: pyodbc.Connection, inp: TextsQueryInputs) -> pd.DataFrame:
    if not inp.biblcat_codes:
        return pd.DataFrame()
    cat_in = ",".join(str(int(c)) for c in inp.biblcat_codes)

    use_addr = bool(inp.addr_ids)
    addr_in = None
    if use_addr:
        addrs = _expand_addresses(
            conn, inp.addr_ids, inp.include_subunits,
            inp.use_xy_radius, inp.xy_narrow,
        )
        if not addrs:
            return pd.DataFrame()
        addr_in = ",".join(str(a) for a in addrs)

    # Use DYNASTIES INNER JOIN BIOG_MAIN when filtering by dynasty range
    use_dyn_join = inp.year_mode == "dynasty" and inp.from_dynasty > -2

    base = (
        "TEXT_ROLE_CODES INNER JOIN ("
        "(TEXT_CODES INNER JOIN BIOG_TEXT_DATA "
        "ON TEXT_CODES.c_textid = BIOG_TEXT_DATA.c_textid) "
        "INNER JOIN BIOG_MAIN ON BIOG_MAIN.c_personid = BIOG_TEXT_DATA.c_personid"
        ") ON TEXT_ROLE_CODES.c_role_id = BIOG_TEXT_DATA.c_role_id"
    )
    if use_dyn_join:
        from_clause = (
            f"({base}) INNER JOIN DYNASTIES "
            f"ON DYNASTIES.c_dy = BIOG_MAIN.c_dy"
        )
    else:
        from_clause = base

    where_parts: list[str] = [
        f"TEXT_CODES.c_bibl_cat_code IN ({cat_in})"
    ]
    if addr_in is not None:
        where_parts.append(f"BIOG_MAIN.c_index_addr_id IN ({addr_in})")
    yc = _years_for_texts(inp)
    if yc:
        where_parts.append(f"({yc})")
    where_clause = " WHERE " + " AND ".join(where_parts)

    sql = f"""
        SELECT
            BIOG_MAIN.c_personid,
            BIOG_MAIN.c_name,
            BIOG_MAIN.c_name_chn,
            BIOG_MAIN.c_index_year,
            BIOG_MAIN.c_index_year_type_code,
            BIOG_MAIN.c_dy,
            IIf(BIOG_MAIN.c_female, 'F', 'M') AS c_sex,
            BIOG_MAIN.c_index_addr_id    AS c_addr_id,
            BIOG_TEXT_DATA.c_textid,
            TEXT_CODES.c_title,
            TEXT_CODES.c_title_chn,
            BIOG_TEXT_DATA.c_role_id,
            TEXT_ROLE_CODES.c_role_desc,
            TEXT_ROLE_CODES.c_role_desc_chn,
            BIOG_TEXT_DATA.c_source
        FROM {from_clause}
        {where_clause}
    """
    return pd.read_sql(sql, conn)
