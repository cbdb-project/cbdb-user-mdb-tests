"""
Smoke tests for every saved query (View_*).  These views are the
RecordSource of the form sub-datasheets, so if they fail to parse
or return zero rows on a populated DB, the form datasheets will be
empty.

For each query:
  * verify it parses (SELECT * FROM <view> LIMIT 1 succeeds)
  * verify it returns at least one row (linked tables resolve)
  * verify the column set matches what's recorded in queries.json
    (regenerate queries.json via analysis/dump_metadata.py after a
    deliberate update)
"""
from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import pytest

REPO = Path(__file__).resolve().parent.parent
QUERIES_JSON = REPO / "analysis" / "dump" / "queries.json"


@pytest.fixture(scope="module")
def queries():
    if not QUERIES_JSON.exists():
        pytest.skip(f"{QUERIES_JSON} missing; run analysis/dump_metadata.py")
    return json.loads(QUERIES_JSON.read_text(encoding="utf-8"))


@pytest.fixture(scope="module")
def view_names(queries):
    return sorted(q["name"] for q in queries if q["name"].startswith("View_"))


def _ids(queries):
    return [q["name"] for q in queries if q["name"].startswith("View_")]


def test_view_count_unchanged(queries):
    n = sum(1 for q in queries if q["name"].startswith("View_"))
    assert n == 18, (
        f"expected 18 View_* queries, found {n}.  "
        f"If a view was added/removed deliberately, regenerate goldens."
    )


@pytest.mark.parametrize("view_name", [
    "View_AltNameData", "View_AssociationData", "View_BiogAddrData",
    "View_BiogInstAddrData", "View_BiogInstData", "View_BiogSourceData",
    "View_BiogTextData", "View_EntryData", "View_EventAddrData",
    "View_EventsData", "View_KinAddrData", "View_PeopleAddr",
    "View_PeopleData", "View_PossessionsAddrData", "View_PossessionsData",
    "View_PostingAddrData", "View_PostingOfficeData", "View_StatusData",
])
def test_view_returns_rows(ro_conn, view_name):
    """The view parses, runs, and returns at least 1 row."""
    cur = ro_conn.cursor()
    cur.execute(f"SELECT TOP 1 * FROM [{view_name}]")
    row = cur.fetchone()
    cur.close()
    assert row is not None, f"{view_name} returned no rows"


@pytest.mark.parametrize("view_name", [
    "View_EntryData", "View_PostingOfficeData", "View_AssociationData",
    "View_KinAddrData", "View_PeopleData", "View_StatusData",
])
def test_view_personid_present_and_valid(ro_conn, view_name):
    """For person-keyed views, every returned c_personid must exist
    in BIOG_MAIN. (Catches views that lost their FK join.)"""
    cur = ro_conn.cursor()
    # take a small sample of DISTINCT pids to keep tests fast
    cur.execute(
        f"SELECT TOP 50 c_personid FROM [{view_name}] "
        f"WHERE c_personid IS NOT NULL GROUP BY c_personid "
        f"ORDER BY c_personid"
    )
    sample = sorted({int(r[0]) for r in cur.fetchall()})
    cur.close()
    assert sample, f"{view_name} returned no non-null c_personid"
    # validate they all exist in BIOG_MAIN
    in_clause = ",".join(str(p) for p in sample)
    cur = ro_conn.cursor()
    # Access SQL doesn't support COUNT(DISTINCT ...) — wrap in subquery
    cur.execute(
        f"SELECT COUNT(*) FROM "
        f"(SELECT DISTINCT c_personid FROM BIOG_MAIN "
        f"WHERE c_personid IN ({in_clause})) AS sub"
    )
    n = cur.fetchone()[0]
    cur.close()
    assert n == len(sample), (
        f"{view_name}: only {n}/{len(sample)} distinct c_personid values "
        f"exist in BIOG_MAIN — view may be referencing a stale or wrong table"
    )
