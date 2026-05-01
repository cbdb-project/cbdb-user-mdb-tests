"""
Python replay of LookAtNetworks.CmdRun_Click
(VBA: analysis/dump/vba/Form_LookAtNetworks.vb:3324)

⚠️ MOST COMPLEX FORM — direct ASSOC subset only.

LookAtNetworks does multi-hop social-network traversal from a seed
person, gated by 28 checkboxes that filter NONKIN edge types
(political/financial/scholarship/writings/military/etc.) plus
optional kinship expansion. The full handler is ~6,000+ lines
spread across multiple subroutines.

This module provides a minimal entry point: 1-hop NONKIN edges from
a seed person, optionally filtered by selected ASSOC type codes.
Multi-hop is TODO.
"""
from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd
import pyodbc

from .common import in_clause


@dataclass
class NetworksQueryInputs:
    person_id: int = 0
    # ASSOC types to include (subset of the 28 ChkXxx flags); empty = all
    assoc_codes: list[int] | None = None
    include_kin: bool = False         # ChkKin
    max_hops: int = 1                  # NotImplemented for >1


def run(conn: pyodbc.Connection, inp: NetworksQueryInputs) -> pd.DataFrame:
    """1-hop NONKIN edges from person_id, optionally filtered by codes."""
    if inp.max_hops > 1:
        raise NotImplementedError(
            f"multi-hop network traversal (max_hops={inp.max_hops}) "
            "not yet ported; see VBA Form_LookAtNetworks.vb for the "
            "iterative ZZ_NETWORK_LIST expansion."
        )
    if inp.include_kin:
        raise NotImplementedError("ChkKin branch not yet ported")
    if inp.person_id <= 0:
        return pd.DataFrame()

    where_parts = [f"ASSOC_DATA.c_personid = {int(inp.person_id)}"]
    if inp.assoc_codes:
        where_parts.append(
            f"ASSOC_DATA.c_assoc_code IN {in_clause(inp.assoc_codes)}"
        )
    where_sql = " WHERE " + " AND ".join(where_parts)

    sql = f"""
        SELECT
            ASSOC_DATA.c_personid     AS c_person_id,
            BIOG_MAIN.c_name,
            BIOG_MAIN.c_name_chn,
            ASSOC_DATA.c_assoc_id     AS c_node_id,
            BIOG_MAIN_1.c_name        AS c_node_name,
            BIOG_MAIN_1.c_name_chn    AS c_node_chn,
            ASSOC_DATA.c_assoc_code   AS c_link_code,
            ASSOC_CODES.c_assoc_desc  AS c_link_desc,
            ASSOC_CODES.c_assoc_desc_chn AS c_link_chn,
            ASSOC_DATA.c_assoc_count  AS c_link_count,
            ASSOC_DATA.c_assoc_first_year AS c_link_first_year,
            ASSOC_DATA.c_assoc_last_year  AS c_link_last_year,
            ASSOC_DATA.c_source
        FROM ((ASSOC_DATA
            INNER JOIN BIOG_MAIN ON ASSOC_DATA.c_personid = BIOG_MAIN.c_personid)
            INNER JOIN BIOG_MAIN AS BIOG_MAIN_1
                ON ASSOC_DATA.c_assoc_id = BIOG_MAIN_1.c_personid)
            INNER JOIN ASSOC_CODES
                ON ASSOC_DATA.c_assoc_code = ASSOC_CODES.c_assoc_code
        {where_sql}
    """
    return pd.read_sql(sql, conn)
