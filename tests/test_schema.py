"""
Schema invariants — these are the lowest-level tests.  If they break
after a CBDB_BJ_User.mdb update, the .mdb has changed in a way that
will likely cascade into many other failures.

The expected schema is captured in analysis/dump/tables.json (the
output of analysis/dump_metadata.py).  Re-run that script and inspect
the diff before --regenerate-goldens-ing the rest of the suite.
"""
from __future__ import annotations

import json
from pathlib import Path

import pytest


REPO = Path(__file__).resolve().parent.parent
TABLES_JSON = REPO / "analysis" / "dump" / "tables.json"


# Tables we DEPEND on across the test suite.  If any of these go
# missing, fail loudly.
REQUIRED_TABLES = {
    # core data (linked from DATA mdb)
    "BIOG_MAIN", "ENTRY_DATA", "POSTED_TO_OFFICE_DATA", "STATUS_DATA",
    "ASSOC_DATA", "KIN_DATA", "BIOG_ADDR_DATA", "BIOG_TEXT_DATA",
    "EVENTS_DATA", "POSSESSION_DATA",
    # code lookups
    "ENTRY_CODES", "DYNASTIES", "ADDR_CODES", "OFFICE_CODES",
    "KINSHIP_CODES", "STATUS_CODES", "ASSOC_CODES", "TEXT_CODES",
    "BIOG_ADDR_CODES", "INDEXYEAR_TYPE_CODES", "YEAR_RANGE_CODES",
    "NIAN_HAO", "GANZHI_CODES",
    # working/scratch
    "ZZ_SCRATCH_ENTRY", "ZZ_SCRATCH_ENTRY_CODE", "ZZ_SCRATCH_ADDR",
    "ZZ_SCRATCH_ADDR_LIST", "ZZ_SCRATCH_KIN", "ZZ_SCRATCH_OFFICE",
    "ZZ_SCRATCH_PEOPLE", "ZZ_SCRATCH_STATUS", "ZZ_SCRATCH_ASSOC",
    "ZZ_SCRATCH_PLACE_PEOPLE", "ZZ_SCRATCH_PLACE_AGG",
    "ZZ_SCRATCH_BIOG_TEXT_DATA", "ZZZ_BELONGS_TO",
}

# Columns we depend on per table, by name (we don't pin types since
# Access expresses them as DAO type-codes which can drift).
REQUIRED_COLUMNS = {
    "BIOG_MAIN": {
        "c_personid", "c_name", "c_name_chn", "c_index_year",
        "c_index_year_type_code", "c_index_addr_id",
        "c_index_addr_type_code", "c_dy", "c_birthyear", "c_deathyear",
        "c_female",
    },
    "ENTRY_DATA": {
        "c_personid", "c_entry_code", "c_year", "c_sequence",
        "c_exam_rank", "c_kin_id", "c_kin_code", "c_assoc_id",
        "c_assoc_code", "c_parental_status_code", "c_entry_addr_id",
        "c_source", "c_inst_code", "c_inst_name_code",
    },
    "ENTRY_CODES": {"c_entry_code", "c_entry_desc", "c_entry_desc_chn"},
    "ADDR_CODES": {
        "c_addr_id", "c_name", "c_name_chn", "x_coord", "y_coord",
    },
    "ZZ_SCRATCH_ENTRY_CODE": {"c_entry_code"},
    "ZZ_SCRATCH_ADDR": {"c_addr_id"},
    "ZZ_SCRATCH_ADDR_LIST": {"c_addr_id"},
    "ZZZ_BELONGS_TO": {"c_addr_id", "c_belongs_to"},
}


@pytest.fixture(scope="module")
def schema():
    if not TABLES_JSON.exists():
        pytest.skip(f"{TABLES_JSON} missing; run analysis/dump_metadata.py")
    data = json.loads(TABLES_JSON.read_text(encoding="utf-8"))
    return {t["name"]: t for t in data}


def test_required_tables_present(schema):
    missing = REQUIRED_TABLES - set(schema)
    assert not missing, f"missing tables: {sorted(missing)}"


@pytest.mark.parametrize("table,cols", REQUIRED_COLUMNS.items())
def test_required_columns_present(schema, table, cols):
    if table not in schema:
        pytest.skip(f"{table} not present (covered by required-tables test)")
    have = {c["name"] for c in schema[table]["columns"]}
    missing = cols - have
    assert not missing, (
        f"{table}: missing columns {sorted(missing)}\n"
        f"  table has: {sorted(have)}"
    )
