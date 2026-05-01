"""
Python replay of LookAtKinship.CmdRun_Click
(VBA: analysis/dump/vba/Form_LookAtKinship.vb:829)

⚠️ COMPLEX FORM — direct-relationship subset only.

LookAtKinship does recursive kinship traversal: starting from one
person, walk KIN_DATA up/down/marriage/collateral edges out to
configured limits (TxtMaxUp / TxtMaxDwn / TxtMaxMar / TxtMaxCol).
Faithful replay would require porting the recursive traversal that
the VBA implements via repeated INSERT loops with a "distance"
counter.

This module provides the DIRECT-RELATIONSHIP query: given a person id,
return all immediate kin (1-hop). Multi-hop is TODO.
"""
from __future__ import annotations

from dataclasses import dataclass

import pandas as pd
import pyodbc


@dataclass
class KinshipQueryInputs:
    person_id: int = 0
    # multi-hop limits (NotImplemented at the moment)
    max_up: int = 1
    max_down: int = 1
    max_marriage: int = 1
    max_collateral: int = 1
    include_mourning: bool = False    # ChkMourning


def run(conn: pyodbc.Connection, inp: KinshipQueryInputs) -> pd.DataFrame:
    """Direct (1-hop) kin of person_id only.

    Multi-hop traversal is TODO; if any max_* > 1, raises
    NotImplementedError.
    """
    if any(v > 1 for v in (inp.max_up, inp.max_down,
                            inp.max_marriage, inp.max_collateral)):
        raise NotImplementedError(
            "multi-hop kinship traversal not yet ported; "
            "see VBA Form_LookAtKinship.vb for the recursive INSERT loops."
        )
    if inp.include_mourning:
        raise NotImplementedError(
            "ChkMourning expansion not yet ported."
        )
    if inp.person_id <= 0:
        return pd.DataFrame()

    sql = f"""
        SELECT
            KIN_DATA.c_personid,
            BIOG_MAIN.c_name,
            BIOG_MAIN.c_name_chn,
            KIN_DATA.c_kin_id,
            BIOG_MAIN_1.c_name              AS c_kin_name,
            BIOG_MAIN_1.c_name_chn          AS c_kin_chn,
            KIN_DATA.c_kin_code,
            KINSHIP_CODES.c_kinrel,
            KINSHIP_CODES.c_kinrel_chn,
            KINSHIP_CODES.c_upstep,
            KINSHIP_CODES.c_dwnstep,
            KINSHIP_CODES.c_marstep,
            KINSHIP_CODES.c_colstep,
            KIN_DATA.c_source
        FROM ((KIN_DATA
            INNER JOIN BIOG_MAIN ON KIN_DATA.c_personid = BIOG_MAIN.c_personid)
            INNER JOIN BIOG_MAIN AS BIOG_MAIN_1
                ON KIN_DATA.c_kin_id = BIOG_MAIN_1.c_personid)
            INNER JOIN KINSHIP_CODES
                ON KIN_DATA.c_kin_code = KINSHIP_CODES.c_kincode
        WHERE KIN_DATA.c_personid = {int(inp.person_id)}
    """
    return pd.read_sql(sql, conn)
