"""
cbdb_replay
-----------"""

# Re-export form-replay modules so callers can do
#   from cbdb_replay import lookatentry, lookatstatus, ...
from . import (
    lookatentry,
    lookatstatus, lookattexts, lookatplace,
    lookatassociations, lookatoffice,
    lookatassociationpairs, lookatkinship,
    lookatnetworks, lookatgroupdata,
    common, exports,
)

__all__ = [
    "lookatentry", "lookatstatus", "lookattexts", "lookatplace",
    "lookatassociations", "lookatoffice", "lookatassociationpairs",
    "lookatkinship", "lookatnetworks", "lookatgroupdata",
    "common", "exports",
]


_DOC = """
Python re-implementations of the SQL that each LookAt* form's
CmdQuery_Click would generate.  Each module exposes a single
``run(conn, **inputs) -> pandas.DataFrame`` callable.

The goal is NOT a 100% faithful translation of every column the form
displays — only the *selection logic* that determines which person /
event rows end up in the result table.  Display columns can be
backfilled from the original linked tables on demand.

Design rules:
 * Each replay function takes a ``pyodbc.Connection`` and the
   inputs the corresponding form would receive (entry codes, address
   ids, year range, etc.).  The connection points at the test copy of
   CBDB_BJ_User.mdb (which has the linked tables to CBDB_*_DATA.mdb).
 * No mutation of database state -- the replay returns a DataFrame
   without touching the ZZ_SCRATCH_* tables.
 * Outputs include the primary identity columns (c_personid +
   discriminators) plus enough joined columns to make goldens
   readable without being noisy.
"""
