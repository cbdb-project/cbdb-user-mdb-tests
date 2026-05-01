"""
Python replay of LookAtOffice.CmdQuery_Click (office-code branch).

VBA: analysis/dump/vba/Form_LookAtOffice.vb:1823

LookAtOffice has the richest input set of any LookAt form:
  - office codes (ZZ_OFFICE_CODE)
  - office addresses (ZZ_SCRATCH_ADDR_OFFICE) + sub-units + XY
  - people addresses (ZZ_SCRATCH_ADDR_PEOPLE) + sub-units + XY
  - year mode: index years / dynasties / OFFICE years
    (gUseIndexYears / gUseDynasties / gUseOfficeYears)

This replay covers the MOST COMMON case: office codes + index/dynasty
year filter, no address constraint. The other branches add WHERE
clauses on POSTED_TO_OFFICE_DATA.c_office_id ∈ ZZ_SCRATCH_ADDR ranges
and follow the same pattern as the address handling in LookAtEntry.

Output: rows that would be inserted into ZZ_SCRATCH_OFFICE.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal

import pandas as pd
import pyodbc

from .common import in_clause, YearFilter, years_where


@dataclass
class OfficeQueryInputs:
    office_codes: list[int] | None = None     # → ZZ_OFFICE_CODE
    year_filter: YearFilter = field(default_factory=YearFilter)
    # office year filter (separate from index/dynasty)
    use_office_years: bool = False
    office_from_year: int | None = None
    office_to_year: int | None = None


def run(conn: pyodbc.Connection, inp: OfficeQueryInputs) -> pd.DataFrame:
    if not inp.office_codes:
        return pd.DataFrame()
    code_in = in_clause(inp.office_codes)

    use_dyn_join = (inp.year_filter.mode == "dynasty"
                    and inp.year_filter.from_dynasty > -2)

    base = (
        "POSTED_TO_OFFICE_DATA AS POD "
        "INNER JOIN BIOG_MAIN ON POD.c_personid = BIOG_MAIN.c_personid"
    )
    if use_dyn_join:
        from_clause = (
            f"({base}) INNER JOIN DYNASTIES "
            f"ON DYNASTIES.c_dy = BIOG_MAIN.c_dy"
        )
    else:
        from_clause = base

    where_parts = [f"POD.c_office_id IN {code_in}"]
    yc = years_where(inp.year_filter,
                     index_col="BIOG_MAIN.c_index_year",
                     dynasty_dy_col="BIOG_MAIN.c_dy")
    if yc:
        where_parts.append(f"({yc})")
    if inp.use_office_years:
        if inp.office_from_year is not None:
            where_parts.append(
                f"POD.c_firstyear >= {int(inp.office_from_year)}"
            )
        if inp.office_to_year is not None:
            where_parts.append(
                f"POD.c_lastyear <= {int(inp.office_to_year)}"
            )

    where_sql = " WHERE " + " AND ".join(where_parts)

    sql = f"""
        SELECT
            POD.c_personid,
            BIOG_MAIN.c_name,
            BIOG_MAIN.c_name_chn,
            BIOG_MAIN.c_index_year,
            BIOG_MAIN.c_index_year_type_code,
            BIOG_MAIN.c_dy,
            IIf(BIOG_MAIN.c_female,'F','M') AS c_sex,
            BIOG_MAIN.c_index_addr_id      AS c_person_addr_id,
            POD.c_office_id,
            POD.c_posting_id,
            POD.c_sequence,
            POD.c_firstyear,
            POD.c_lastyear,
            POD.c_appt_code,
            POD.c_assume_office_code,
            POD.c_inst_code,
            POD.c_source
        FROM {from_clause}
        {where_sql}
    """
    return pd.read_sql(sql, conn)
