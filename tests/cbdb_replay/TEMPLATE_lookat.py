"""
TEMPLATE: copy this file to tests/cbdb_replay/lookat<formname>.py
and fill in the form-specific logic.

Recipe for porting a LookAt form's CmdQuery_Click:

1. Read  analysis/dump/vba/Form_LookAt<Formname>.vb
2. Locate Private Sub CmdQuery_Click   (use grep for "CmdQuery_Click")
3. Identify the inputs:
     * controls (text boxes, frame value, checkboxes)
     * picker tables (ZZ_SCRATCH_<XXX>) populated from picker forms
     * public form-module globals (gUseADDRID, gUse<X>Years, etc.) — these
       are usually set by other Cmd... event handlers
4. Identify the SQL templates the VBA picks between:
     * locate every  tStrFrom = "FROM ..."   branch
     * note which (use_addr × addr_field × use_codes × use_years) combo
       each branch corresponds to
5. Identify the WHERE clauses:
     * year/dynasty fragments
     * address IN (...) or join-on-c_addr_id branch
     * code IN (...) or join-on-ZZ_SCRATCH_<code-table>
6. Translate to a parameterised Python function ``run(conn, inputs) -> DataFrame``

The output DataFrame should at minimum include the primary identity columns
of the result table (e.g. c_personid, the discriminator c_<event>_id).

Then write tests in tests/test_lookat<formname>.py with at least 3 fixtures
covering different code-path branches.  Use pytest --regenerate-goldens
to write the first golden CSVs after manual inspection.
"""
from __future__ import annotations
from dataclasses import dataclass
from typing import Literal

import pandas as pd
import pyodbc


@dataclass
class XxxQueryInputs:
    # TODO: fill in form's control inputs
    pass


def run(conn: pyodbc.Connection, inp: XxxQueryInputs) -> pd.DataFrame:
    raise NotImplementedError(
        "Port the relevant CmdQuery_Click branches from "
        "analysis/dump/vba/Form_LookAt<Formname>.vb"
    )
