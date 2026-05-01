"""
Python replay of LookAtAssociationPairs.CmdQuery_Click
(VBA: analysis/dump/vba/Form_LookAtAssociationPairs.vb:1588)

⚠️ COMPLEX FORM — partial replay only.

This form does 1st- and 2nd-order graph traversal between two seed
persons, gathering all intermediate nodes via four (or eight) calls
to Link1stOrder / Link2ndOrder.  Faithful replay would require
re-implementing those subroutines.

We implement only the SIMPLEST case: given two person IDs, find direct
ASSOC_DATA edges between them.  Anything that needs Link*Order is
TODO.
"""
from __future__ import annotations

from dataclasses import dataclass

import pandas as pd
import pyodbc


@dataclass
class AssocPairsQueryInputs:
    person_id_1: int = 0
    person_id_2: int = 0
    include_kinship: bool = False    # ChkKinship
    second_order: bool = False        # Chk2Nodes — NotImplemented


def run(conn: pyodbc.Connection, inp: AssocPairsQueryInputs) -> pd.DataFrame:
    """Direct ASSOC_DATA edges between person_id_1 and person_id_2 only.

    The full form additionally collects 1st-order (and optionally
    2nd-order) intermediate-node graph; that's not yet ported.
    """
    if inp.second_order:
        raise NotImplementedError(
            "second-order traversal (Chk2Nodes) not yet ported. "
            "See VBA Link2ndOrder for the 8-variant kin/non-kin matrix."
        )
    if inp.include_kinship:
        raise NotImplementedError(
            "kinship-included branch not yet ported. "
            "See VBA Link1stOrder('KIN', 'KIN') etc."
        )
    if inp.person_id_1 <= 0 or inp.person_id_2 <= 0:
        return pd.DataFrame()

    # Direct A→B and B→A association edges
    sql = f"""
        SELECT
            ASSOC_DATA.c_personid          AS c_person_id,
            BIOG_MAIN.c_name,
            BIOG_MAIN.c_name_chn,
            BIOG_MAIN.c_index_year,
            ASSOC_DATA.c_assoc_id          AS c_node_id,
            BIOG_MAIN_1.c_name             AS c_node_name,
            BIOG_MAIN_1.c_name_chn         AS c_node_chn,
            ASSOC_DATA.c_assoc_code        AS c_link_code,
            ASSOC_CODES.c_assoc_desc       AS c_link_desc,
            ASSOC_CODES.c_assoc_desc_chn   AS c_link_chn,
            ASSOC_DATA.c_assoc_first_year  AS c_link_first_year,
            ASSOC_DATA.c_assoc_last_year   AS c_link_last_year,
            ASSOC_DATA.c_source
        FROM ((ASSOC_DATA
            INNER JOIN BIOG_MAIN
                ON ASSOC_DATA.c_personid = BIOG_MAIN.c_personid)
            INNER JOIN BIOG_MAIN AS BIOG_MAIN_1
                ON ASSOC_DATA.c_assoc_id = BIOG_MAIN_1.c_personid)
            INNER JOIN ASSOC_CODES
                ON ASSOC_DATA.c_assoc_code = ASSOC_CODES.c_assoc_code
        WHERE
            (ASSOC_DATA.c_personid = {int(inp.person_id_1)}
             AND ASSOC_DATA.c_assoc_id = {int(inp.person_id_2)})
         OR (ASSOC_DATA.c_personid = {int(inp.person_id_2)}
             AND ASSOC_DATA.c_assoc_id = {int(inp.person_id_1)})
    """
    return pd.read_sql(sql, conn)
