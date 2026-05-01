"""
Python replay of LookAtPlace.CmdQuery_Click (Individual mode).

VBA: analysis/dump/vba/Form_LookAtPlace.vb:1027

LookAtPlace can pull people-at-place from MANY sources controlled by
checkboxes:
  - ChkIndividual: people whose index_addr is in the selection
  - ChkKin: people whose kin's address is in the selection
  - ChkOffice: people who held office at one of the selected addresses
  - ChkStatus: people whose status was held at the address
  - ChkEntry: people whose entry happened at the address
  - ChkInstitution: people associated with an institution at the address
  - ChkAssocPerson / ChkAssocPlace: people linked via ASSOC_DATA

This replay covers the Individual mode (the simplest + most common
case).  Other modes are straightforward extensions: each adds a
UNION over a different join (BIOG_ADDR_DATA / KIN_DATA / etc).

Output: rows that would be inserted into ZZ_SCRATCH_PLACE_PEOPLE.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

import pandas as pd
import pyodbc

from .common import expand_addresses, in_clause


@dataclass
class PlaceQueryInputs:
    addr_ids: list[int] | None = None
    include_subunits: bool = False
    use_xy_radius: bool = False     # ChkXYRef ±0.03

    # Year filter (LookAtPlace uses tUseFirstYear/tUseLastYear from
    # gUseIndexYears + TxtFromYear/TxtToYear)
    use_index_years: bool = False
    from_year: int | None = None
    to_year: int | None = None

    # Source toggle (only Individual implemented for now)
    source: Literal["individual"] = "individual"


def run(conn: pyodbc.Connection, inp: PlaceQueryInputs) -> pd.DataFrame:
    if not inp.addr_ids:
        return pd.DataFrame()
    addrs = expand_addresses(
        conn, inp.addr_ids,
        include_subunits=inp.include_subunits,
        use_xy_radius=inp.use_xy_radius,
    )
    if not addrs:
        return pd.DataFrame()

    if inp.source != "individual":
        raise NotImplementedError(
            f"source={inp.source!r} not yet ported. Individual mode only."
        )

    addr_in = in_clause(addrs)
    where = [f"BIOG_MAIN.c_index_addr_id IN {addr_in}"]
    if inp.use_index_years:
        if inp.from_year is not None:
            where.append(f"BIOG_MAIN.c_index_year >= {int(inp.from_year)}")
        if inp.to_year is not None:
            where.append(f"BIOG_MAIN.c_index_year <= {int(inp.to_year)}")
    where_sql = " WHERE " + " AND ".join(where)

    sql = f"""
        SELECT DISTINCT
            BIOG_MAIN.c_personid,
            BIOG_MAIN.c_name,
            BIOG_MAIN.c_name_chn,
            BIOG_MAIN.c_index_year,
            BIOG_MAIN.c_dy,
            BIOG_MAIN.c_index_addr_id    AS c_addr_id,
            ADDR_CODES.c_name             AS c_addr_name,
            ADDR_CODES.c_name_chn         AS c_addr_chn,
            ADDR_CODES.x_coord,
            ADDR_CODES.y_coord
        FROM BIOG_MAIN
        LEFT JOIN ADDR_CODES ON BIOG_MAIN.c_index_addr_id = ADDR_CODES.c_addr_id
        {where_sql}
    """
    return pd.read_sql(sql, conn)
