"""
Python replay of LookAtAssociations.CmdQuery_Click
(VBA: analysis/dump/vba/Form_LookAtAssociations.vb:1670)

Inputs:
  - assoc codes via ZZ_ASSOC_CODE.c_assoc_code (picker → table)
  - addresses (optional)
  - year mode (FrameFilterYears: 1=none, 2=index, 3=dynasty)

Output: rows that would be inserted into ZZ_SCRATCH_ASSOC.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal

import pandas as pd
import pyodbc

from .common import expand_addresses, in_clause, YearFilter, years_where

YearMode = Literal["none", "index", "dynasty"]


@dataclass
class AssocQueryInputs:
    assoc_codes: list[int] | None = None
    addr_ids: list[int] | None = None
    include_subunits: bool = False
    use_xy_radius: bool = False
    xy_narrow: bool = True
    year_filter: YearFilter = field(default_factory=YearFilter)


def run(conn: pyodbc.Connection, inp: AssocQueryInputs) -> pd.DataFrame:
    if not inp.assoc_codes:
        return pd.DataFrame()
    code_in = in_clause(inp.assoc_codes)

    addr_in = None
    if inp.addr_ids:
        addrs = expand_addresses(
            conn, inp.addr_ids,
            include_subunits=inp.include_subunits,
            use_xy_radius=inp.use_xy_radius,
            xy_eps=0.03 if inp.xy_narrow else 0.06,
        )
        if not addrs:
            return pd.DataFrame()
        addr_in = in_clause(addrs)

    use_dyn_join = (inp.year_filter.mode == "dynasty"
                    and inp.year_filter.from_dynasty > -2)

    base = (
        "ASSOC_DATA INNER JOIN BIOG_MAIN "
        "ON ASSOC_DATA.c_personid = BIOG_MAIN.c_personid"
    )
    if use_dyn_join:
        from_clause = (
            f"({base}) INNER JOIN DYNASTIES "
            f"ON DYNASTIES.c_dy = BIOG_MAIN.c_dy"
        )
    else:
        from_clause = base

    where_parts = [f"ASSOC_DATA.c_assoc_code IN {code_in}"]
    if addr_in is not None:
        where_parts.append(f"BIOG_MAIN.c_index_addr_id IN {addr_in}")
    yc = years_where(inp.year_filter)
    if yc:
        where_parts.append(f"({yc})")
    where_sql = " WHERE " + " AND ".join(where_parts)

    sql = f"""
        SELECT
            ASSOC_DATA.c_personid          AS c_person_id,
            BIOG_MAIN.c_name,
            BIOG_MAIN.c_name_chn,
            BIOG_MAIN.c_index_year,
            BIOG_MAIN.c_index_year_type_code,
            IIf(BIOG_MAIN.c_female,'F','M') AS c_sex,
            BIOG_MAIN.c_index_addr_id      AS c_addr_id,
            BIOG_MAIN.c_dy,
            ASSOC_DATA.c_assoc_code,
            ASSOC_DATA.c_kin_code,
            ASSOC_DATA.c_kin_id,
            ASSOC_DATA.c_assoc_id,
            ASSOC_DATA.c_assoc_kin_code,
            ASSOC_DATA.c_assoc_kin_id,
            ASSOC_DATA.c_assoc_count,
            ASSOC_DATA.c_assoc_first_year,
            ASSOC_DATA.c_assoc_last_year,
            ASSOC_DATA.c_source,
            ASSOC_DATA.c_addr_id           AS c_assoc_place_addr_id,
            ASSOC_DATA.c_text_title,
            ASSOC_DATA.c_assoc_claimer_id
        FROM {from_clause}
        {where_sql}
    """
    return pd.read_sql(sql, conn)
