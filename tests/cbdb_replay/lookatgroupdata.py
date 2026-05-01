"""
Python replay of LookAtGroupData.CmdRun_Click
(VBA: analysis/dump/vba/Form_LookAtGroupData.vb:1551)

LookAtGroupData accepts a list of imported person IDs (in
ZZ_SCRATCH_IMPORT_PEOPLE) and optionally pulls associated data of
several kinds based on Chk* checkboxes:
  - ChkAddr → BIOG_ADDR_DATA
  - ChkEntry → ENTRY_DATA
  - ChkOffice → POSTED_TO_OFFICE_DATA
  - ChkStatus → STATUS_DATA
  - ChkText → BIOG_TEXT_DATA

Result table per category. This replay covers the BASE PEOPLE
extraction (Chk* all False = just the person records).
"""
from __future__ import annotations

from dataclasses import dataclass

import pandas as pd
import pyodbc

from .common import in_clause


@dataclass
class GroupDataQueryInputs:
    person_ids: list[int] | None = None     # → ZZ_SCRATCH_IMPORT_PEOPLE
    # category toggles — only base implemented
    include_addr: bool = False
    include_entry: bool = False
    include_office: bool = False
    include_status: bool = False
    include_text: bool = False


def run(conn: pyodbc.Connection, inp: GroupDataQueryInputs) -> pd.DataFrame:
    """Return BIOG_MAIN records for the imported person IDs.

    Per-category sub-queries (Addr/Entry/Office/Status/Text) are TODO;
    each is essentially a JOIN of the seed person list against the
    corresponding *_DATA table.
    """
    if any([inp.include_addr, inp.include_entry, inp.include_office,
            inp.include_status, inp.include_text]):
        raise NotImplementedError(
            "category sub-queries (Chk*) not yet ported; this replay "
            "returns just the base person records for the imported list."
        )
    if not inp.person_ids:
        return pd.DataFrame()

    pid_in = in_clause(inp.person_ids)
    sql = f"""
        SELECT
            BIOG_MAIN.c_personid          AS c_person_id,
            BIOG_MAIN.c_name,
            BIOG_MAIN.c_name_chn,
            BIOG_MAIN.c_index_year,
            BIOG_MAIN.c_index_year_type_code,
            BIOG_MAIN.c_dy,
            BIOG_MAIN.c_female,
            BIOG_MAIN.c_index_addr_id     AS c_addr_id,
            ADDR_CODES.c_name             AS c_addr_name,
            ADDR_CODES.c_name_chn         AS c_addr_chn,
            ADDR_CODES.x_coord,
            ADDR_CODES.y_coord
        FROM BIOG_MAIN LEFT JOIN ADDR_CODES
            ON BIOG_MAIN.c_index_addr_id = ADDR_CODES.c_addr_id
        WHERE BIOG_MAIN.c_personid IN {pid_in}
    """
    return pd.read_sql(sql, conn)
