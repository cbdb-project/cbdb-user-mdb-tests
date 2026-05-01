"""
Python replay of LookAtEntry's CmdQuery_Click.

VBA reference: analysis/dump/vba/Form_LookAtEntry.vb (lines 1478-1862)

The form builds an INSERT INTO ZZ_SCRATCH_ENTRY ... SELECT ... query whose
shape depends on:
  - whether addresses are constrained (gUseADDRID)
  - if so, whether to join via BIOG_MAIN.c_index_addr_id (FrameAddress=1)
    or ENTRY_DATA.c_entry_addr_id (FrameAddress=2)
  - whether to expand the address list with sub-units (ChkSubUnits) and/or
    XY-coord proximity (ChkXYRef, ±0.03°)
  - whether entry codes are constrained ("[All]" + empty type = no constraint)
  - which year mode is active: entry years, index years, dynasties, or none
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Iterable, Literal

import pandas as pd
import pyodbc


YearMode = Literal["entry", "index", "dynasty", "none"]
AddrField = Literal["person", "entry"]


@dataclass
class EntryQueryInputs:
    # entry codes (from picker -> ZZ_SCRATCH_ENTRY_CODE).
    # None / [] = "[All]" (no constraint).
    entry_codes: list[int] | None = None

    # addresses to constrain by. None / [] = no address filter.
    addr_ids: list[int] | None = None
    addr_field: AddrField = "person"            # FrameAddress 1=person, 2=entry
    include_subunits: bool = False              # ChkSubUnits
    use_xy_radius: bool = False                 # ChkXYRef (±0.03°)

    # year limits
    year_mode: YearMode = "none"                # FrameYears: 1/2/3/none
    from_year: int | None = None
    to_year: int | None = None

    # dynasty selection (only used when year_mode == "dynasty")
    # gFromDynasty, gToDynasty: -2 = all, -1 = unset, positive = code
    from_dynasty: int = -1
    to_dynasty: int = -1
    from_dynasty_begin: int | None = None       # cached gFromDynastyBegin
    to_dynasty_end: int | None = None           # cached gToDynastyEnd


# ----------------------------------------------------------------------
# helpers


def _expand_addresses(conn: pyodbc.Connection,
                      addr_ids: Iterable[int],
                      include_subunits: bool,
                      use_xy_radius: bool) -> list[int]:
    """Return the expanded list of addresses to filter on, mirroring the
    VBA logic that populates ZZ_SCRATCH_ADDR_LIST."""
    seed = list({int(a) for a in addr_ids})
    if not seed:
        return []
    cur = conn.cursor()
    if include_subunits:
        # SELECT DISTINCT ZZZ_BELONGS_TO.c_addr_id
        # FROM ZZ_SCRATCH_ADDR INNER JOIN ZZZ_BELONGS_TO
        #   ON ZZ_SCRATCH_ADDR.c_addr_id = ZZZ_BELONGS_TO.c_belongs_to
        in_clause = ",".join(str(x) for x in seed)
        cur.execute(
            f"SELECT DISTINCT c_addr_id FROM ZZZ_BELONGS_TO "
            f"WHERE c_belongs_to IN ({in_clause})"
        )
        expanded = {int(r[0]) for r in cur.fetchall()}
    else:
        expanded = set(seed)

    if use_xy_radius and expanded:
        in_clause = ",".join(str(x) for x in expanded)
        # All ADDR_CODES whose x/y are within ±0.03 of any seed addr
        cur.execute(f"""
            SELECT DISTINCT A.c_addr_id
            FROM ADDR_CODES AS A, ADDR_CODES AS S
            WHERE S.c_addr_id IN ({in_clause})
              AND A.x_coord >= S.x_coord - 0.03
              AND A.x_coord <= S.x_coord + 0.03
              AND A.y_coord >= S.y_coord - 0.03
              AND A.y_coord <= S.y_coord + 0.03
        """)
        expanded |= {int(r[0]) for r in cur.fetchall()}
        # plus seeds that have NULL coords
        cur.execute(f"""
            SELECT c_addr_id FROM ADDR_CODES
            WHERE c_addr_id IN ({in_clause})
              AND (x_coord IS NULL OR y_coord IS NULL)
        """)
        expanded |= {int(r[0]) for r in cur.fetchall()}
    cur.close()
    return sorted(expanded)


def _years_clause(inp: EntryQueryInputs) -> str:
    """Build the WHERE fragment for year/dynasty constraints."""
    if inp.year_mode == "entry":
        cond = []
        if inp.from_year is not None:
            cond.append(f"ENTRY_DATA.c_year >= {int(inp.from_year)}")
        if inp.to_year is not None:
            cond.append(f"ENTRY_DATA.c_year <= {int(inp.to_year)}")
        return " AND ".join(cond)
    if inp.year_mode == "index":
        cond = []
        if inp.from_year is not None:
            cond.append(f"BIOG_MAIN.c_index_year >= {int(inp.from_year)}")
        if inp.to_year is not None:
            cond.append(f"BIOG_MAIN.c_index_year <= {int(inp.to_year)}")
        return " AND ".join(cond)
    if inp.year_mode == "dynasty":
        # mirror VBA branches in CmdQuery_Click @ ~1622
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


# ----------------------------------------------------------------------
# main entry point


def run(conn: pyodbc.Connection, inp: EntryQueryInputs) -> pd.DataFrame:
    """Replay LookAtEntry's CmdQuery_Click. Returns the would-be
    contents of ZZ_SCRATCH_ENTRY (before the cosmetic UPDATE joins)
    as a DataFrame.
    """
    # --- shape: address list / address join column ---
    use_addr = bool(inp.addr_ids)
    addr_ids: list[int] = []
    if use_addr:
        addr_ids = _expand_addresses(
            conn, inp.addr_ids, inp.include_subunits, inp.use_xy_radius
        )
        if not addr_ids:
            # picker yielded an empty list -> empty result
            return pd.DataFrame()
    addr_col = ("BIOG_MAIN.c_index_addr_id"
                if inp.addr_field == "person"
                else "ENTRY_DATA.c_entry_addr_id")

    # --- entry-code constraint ---
    use_codes = bool(inp.entry_codes)

    # --- year/dynasty WHERE ---
    where_years = _years_clause(inp)
    use_dynasties = (inp.year_mode == "dynasty" and where_years)

    # --- assemble FROM (Access SQL requires nested joins to be parenthesised) ---
    base_join = ("BIOG_MAIN INNER JOIN ENTRY_DATA "
                 "ON BIOG_MAIN.c_personid = ENTRY_DATA.c_personid")
    where_extra: list[str] = []
    if use_addr:
        in_clause = ",".join(str(a) for a in addr_ids)
        where_extra.append(f"{addr_col} IN ({in_clause})")
    if use_codes:
        in_clause = ",".join(str(c) for c in inp.entry_codes)
        where_extra.append(f"ENTRY_DATA.c_entry_code IN ({in_clause})")
    if use_dynasties:
        from_clause = (
            f"({base_join}) LEFT JOIN DYNASTIES "
            f"ON BIOG_MAIN.c_dy = DYNASTIES.c_dy"
        )
    else:
        from_clause = base_join

    where = list(where_extra)
    if where_years:
        where.append(f"({where_years})")
    where_clause = (" WHERE " + " AND ".join(where)) if where else ""

    sql = f"""
        SELECT
            ENTRY_DATA.c_personid,
            BIOG_MAIN.c_name,
            BIOG_MAIN.c_name_chn,
            BIOG_MAIN.c_index_year,
            BIOG_MAIN.c_index_year_type_code,
            BIOG_MAIN.c_dy,
            ENTRY_DATA.c_entry_code,
            ENTRY_DATA.c_year,
            ENTRY_DATA.c_sequence,
            ENTRY_DATA.c_exam_rank,
            BIOG_MAIN.c_index_addr_id      AS c_addr_id,
            ENTRY_DATA.c_kin_id,
            ENTRY_DATA.c_kin_code,
            ENTRY_DATA.c_assoc_id,
            ENTRY_DATA.c_assoc_code,
            ENTRY_DATA.c_parental_status_code,
            ENTRY_DATA.c_entry_addr_id,
            ENTRY_DATA.c_source,
            ENTRY_DATA.c_inst_code,
            ENTRY_DATA.c_inst_name_code,
            BIOG_MAIN.c_index_addr_type_code AS c_addr_type
        FROM {from_clause}
        {where_clause}
    """
    return pd.read_sql(sql, conn)
