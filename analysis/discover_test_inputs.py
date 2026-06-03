"""
Discover HIGH-DENSITY test inputs by querying the live data.

For each LookAt form, find input combinations (entry_code,
dynasty, address, etc.) that have a substantial number of matching
rows in the current data.  These get serialised to JSON for
parametrized tests to consume.

Run after every significant DATA.mdb update — the candidate set
should be re-derived since data shifts.

Output: analysis/dump/test_inputs.json
  {
    "lookatentry": {
        "top_entry_codes": [...],          # codes by row count desc
        "top_addresses": [...],            # addresses by indexed-person count
        "top_dynasties": [...],
        "high_density_combos": [
            {"entry_code": 36, "dynasty": 15, "addr_id": null,
             "expected_rows_min": 100},
            ...
        ],
    },
    "lookatkinship": {
        "top_persons_by_kin_count": [...],
        ...
    },
    ...
  }
"""
from __future__ import annotations

import json
from pathlib import Path

import pyodbc

ROOT = Path(__file__).resolve().parent.parent
OUT = Path(__file__).resolve().parent / "dump" / "test_inputs.json"


def _find_data_mdb(root: Path) -> Path:
    """Find the DATA mdb in data/.

    All discovery queries target linked tables that live in the DATA mdb.
    Connecting to the DATA mdb directly avoids stale linked-table path
    errors when the DATA mdb has been replaced with a newer build.

    If multiple CBDB_*_DATA.mdb files exist (e.g. old build not cleaned up),
    pick the one with the latest YYYYMMDD in the filename so the script is
    still usable without manual cleanup.
    """
    matches = list((root / "data").glob("CBDB_*_DATA.mdb"))
    if not matches:
        raise FileNotFoundError("No CBDB_*_DATA.mdb found in data/")
    if len(matches) == 1:
        return matches[0]
    # Multiple matches: sort by embedded date (CBDB_YYYYMMDD_DATA.mdb) and pick newest
    def _date_key(p: Path) -> str:
        parts = p.stem.split("_")  # ["CBDB", "YYYYMMDD", "DATA"]
        return parts[1] if len(parts) >= 2 else p.stem
    chosen = sorted(matches, key=_date_key)[-1]
    print(f"[discover] multiple DATA mdbs found; using newest: {chosen.name}")
    return chosen


DATA_MDB = _find_data_mdb(ROOT)
CONN_STR = (
    "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    f"DBQ={DATA_MDB};ReadOnly=True;"
)


def fetch(conn, sql, *params):
    cur = conn.cursor()
    cur.execute(sql, *params)
    cols = [c[0] for c in cur.description]
    rows = [dict(zip(cols, r)) for r in cur.fetchall()]
    cur.close()
    return rows


def discover_lookatentry(conn) -> dict:
    """Find populous (entry_code) and populous (entry_code × dynasty)
    combinations."""
    out: dict = {}

    out["top_entry_codes"] = fetch(conn, """
        SELECT TOP 15 ENTRY_DATA.c_entry_code,
               ENTRY_CODES.c_entry_desc,
               ENTRY_CODES.c_entry_desc_chn,
               COUNT(*) AS n_rows
        FROM ENTRY_DATA INNER JOIN ENTRY_CODES
          ON ENTRY_DATA.c_entry_code = ENTRY_CODES.c_entry_code
        GROUP BY ENTRY_DATA.c_entry_code,
                 ENTRY_CODES.c_entry_desc,
                 ENTRY_CODES.c_entry_desc_chn
        ORDER BY COUNT(*) DESC
    """)

    out["top_addresses"] = fetch(conn, """
        SELECT TOP 15 BIOG_MAIN.c_index_addr_id AS c_addr_id,
               ADDR_CODES.c_name,
               ADDR_CODES.c_name_chn,
               COUNT(*) AS n_persons
        FROM BIOG_MAIN INNER JOIN ADDR_CODES
          ON BIOG_MAIN.c_index_addr_id = ADDR_CODES.c_addr_id
        WHERE BIOG_MAIN.c_index_addr_id IS NOT NULL
        GROUP BY BIOG_MAIN.c_index_addr_id,
                 ADDR_CODES.c_name,
                 ADDR_CODES.c_name_chn
        ORDER BY COUNT(*) DESC
    """)

    out["top_dynasties"] = fetch(conn, """
        SELECT TOP 10 BIOG_MAIN.c_dy,
               DYNASTIES.c_dynasty,
               DYNASTIES.c_dynasty_chn,
               DYNASTIES.c_start, DYNASTIES.c_end,
               COUNT(*) AS n_persons
        FROM BIOG_MAIN INNER JOIN DYNASTIES ON BIOG_MAIN.c_dy = DYNASTIES.c_dy
        GROUP BY BIOG_MAIN.c_dy, DYNASTIES.c_dynasty,
                 DYNASTIES.c_dynasty_chn,
                 DYNASTIES.c_start, DYNASTIES.c_end
        ORDER BY COUNT(*) DESC
    """)

    # populous (entry_code × dynasty) combos — Access SQL doesn't have
    # COUNT(DISTINCT), so we wrap in a DISTINCT subquery.
    out["entry_x_dynasty_combos"] = fetch(conn, """
        SELECT TOP 20 c_entry_code, c_dy, COUNT(*) AS n_persons
        FROM (
            SELECT DISTINCT ENTRY_DATA.c_entry_code,
                   BIOG_MAIN.c_dy,
                   BIOG_MAIN.c_personid
            FROM ENTRY_DATA INNER JOIN BIOG_MAIN
              ON ENTRY_DATA.c_personid = BIOG_MAIN.c_personid
            WHERE BIOG_MAIN.c_dy IS NOT NULL
        ) AS sub
        GROUP BY c_entry_code, c_dy
        ORDER BY COUNT(*) DESC
    """)

    # populous (entry_code × address) combos — what HelpFile-style fixtures look like
    out["entry_x_address_combos"] = fetch(conn, """
        SELECT TOP 20 c_entry_code, c_index_addr_id, COUNT(*) AS n_persons
        FROM (
            SELECT DISTINCT ENTRY_DATA.c_entry_code,
                   BIOG_MAIN.c_index_addr_id,
                   BIOG_MAIN.c_personid
            FROM ENTRY_DATA INNER JOIN BIOG_MAIN
              ON ENTRY_DATA.c_personid = BIOG_MAIN.c_personid
            WHERE BIOG_MAIN.c_index_addr_id IS NOT NULL
        ) AS sub
        GROUP BY c_entry_code, c_index_addr_id
        ORDER BY COUNT(*) DESC
    """)

    return out


def discover_lookatstatus(conn) -> dict:
    return {
        "top_status_codes": fetch(conn, """
            SELECT TOP 15 STATUS_DATA.c_status_code,
                   STATUS_CODES.c_status_desc,
                   STATUS_CODES.c_status_desc_chn,
                   COUNT(*) AS n_rows
            FROM STATUS_DATA INNER JOIN STATUS_CODES
              ON STATUS_DATA.c_status_code = STATUS_CODES.c_status_code
            GROUP BY STATUS_DATA.c_status_code,
                     STATUS_CODES.c_status_desc,
                     STATUS_CODES.c_status_desc_chn
            ORDER BY COUNT(*) DESC
        """),
        "status_x_dynasty_combos": fetch(conn, """
            SELECT TOP 15 c_status_code, c_dy, COUNT(*) AS n_persons
            FROM (
                SELECT DISTINCT STATUS_DATA.c_status_code,
                       BIOG_MAIN.c_dy,
                       BIOG_MAIN.c_personid
                FROM STATUS_DATA INNER JOIN BIOG_MAIN
                  ON STATUS_DATA.c_personid = BIOG_MAIN.c_personid
            ) AS sub
            GROUP BY c_status_code, c_dy
            ORDER BY COUNT(*) DESC
        """),
    }


def discover_lookattexts(conn) -> dict:
    return {
        "top_biblcat_codes": fetch(conn, """
            SELECT TOP 15 TEXT_CODES.c_bibl_cat_code,
                   TEXT_BIBLCAT_CODES.c_text_cat_desc,
                   TEXT_BIBLCAT_CODES.c_text_cat_desc_chn,
                   COUNT(*) AS n_texts
            FROM TEXT_CODES INNER JOIN TEXT_BIBLCAT_CODES
              ON TEXT_CODES.c_bibl_cat_code = TEXT_BIBLCAT_CODES.c_text_cat_code
            WHERE TEXT_CODES.c_bibl_cat_code IS NOT NULL
            GROUP BY TEXT_CODES.c_bibl_cat_code,
                     TEXT_BIBLCAT_CODES.c_text_cat_desc,
                     TEXT_BIBLCAT_CODES.c_text_cat_desc_chn
            ORDER BY COUNT(*) DESC
        """),
        "biblcat_x_writers": fetch(conn, """
            SELECT TOP 15 c_bibl_cat_code, COUNT(*) AS n_writers
            FROM (
                SELECT DISTINCT TEXT_CODES.c_bibl_cat_code,
                       BIOG_TEXT_DATA.c_personid
                FROM TEXT_CODES INNER JOIN BIOG_TEXT_DATA
                  ON TEXT_CODES.c_textid = BIOG_TEXT_DATA.c_textid
                WHERE TEXT_CODES.c_bibl_cat_code IS NOT NULL
            ) AS sub
            GROUP BY c_bibl_cat_code
            ORDER BY COUNT(*) DESC
        """),
    }


def discover_lookatplace(conn) -> dict:
    return {
        # high-density addresses by # of indexed persons
        "top_addr_by_indexed_persons": fetch(conn, """
            SELECT TOP 20 BIOG_MAIN.c_index_addr_id,
                   ADDR_CODES.c_name,
                   ADDR_CODES.c_name_chn,
                   COUNT(*) AS n_persons
            FROM BIOG_MAIN INNER JOIN ADDR_CODES
              ON BIOG_MAIN.c_index_addr_id = ADDR_CODES.c_addr_id
            WHERE BIOG_MAIN.c_index_addr_id IS NOT NULL
            GROUP BY BIOG_MAIN.c_index_addr_id,
                     ADDR_CODES.c_name, ADDR_CODES.c_name_chn
            ORDER BY COUNT(*) DESC
        """),
    }


def discover_lookatassociations(conn) -> dict:
    return {
        "top_assoc_codes": fetch(conn, """
            SELECT TOP 20 ASSOC_DATA.c_assoc_code,
                   ASSOC_CODES.c_assoc_desc,
                   ASSOC_CODES.c_assoc_desc_chn,
                   COUNT(*) AS n_rows
            FROM ASSOC_DATA INNER JOIN ASSOC_CODES
              ON ASSOC_DATA.c_assoc_code = ASSOC_CODES.c_assoc_code
            GROUP BY ASSOC_DATA.c_assoc_code,
                     ASSOC_CODES.c_assoc_desc,
                     ASSOC_CODES.c_assoc_desc_chn
            ORDER BY COUNT(*) DESC
        """),
        "assoc_x_dynasty_combos": fetch(conn, """
            SELECT TOP 15 ASSOC_DATA.c_assoc_code,
                   BIOG_MAIN.c_dy,
                   COUNT(*) AS n_rows
            FROM ASSOC_DATA INNER JOIN BIOG_MAIN
              ON ASSOC_DATA.c_personid = BIOG_MAIN.c_personid
            GROUP BY ASSOC_DATA.c_assoc_code, BIOG_MAIN.c_dy
            ORDER BY COUNT(*) DESC
        """),
    }


def discover_lookatoffice(conn) -> dict:
    return {
        "top_office_codes": fetch(conn, """
            SELECT TOP 20 POSTED_TO_OFFICE_DATA.c_office_id,
                   OFFICE_CODES.c_office_chn,
                   OFFICE_CODES.c_office_pinyin,
                   COUNT(*) AS n_postings
            FROM POSTED_TO_OFFICE_DATA INNER JOIN OFFICE_CODES
              ON POSTED_TO_OFFICE_DATA.c_office_id = OFFICE_CODES.c_office_id
            GROUP BY POSTED_TO_OFFICE_DATA.c_office_id,
                     OFFICE_CODES.c_office_chn,
                     OFFICE_CODES.c_office_pinyin
            ORDER BY COUNT(*) DESC
        """),
    }


def discover_lookatkinship(conn) -> dict:
    return {
        # persons with the most kin (interesting graph centers)
        "top_persons_by_kin_count": fetch(conn, """
            SELECT TOP 20 KIN_DATA.c_personid,
                   BIOG_MAIN.c_name,
                   BIOG_MAIN.c_name_chn,
                   COUNT(*) AS n_kin
            FROM KIN_DATA INNER JOIN BIOG_MAIN
              ON KIN_DATA.c_personid = BIOG_MAIN.c_personid
            GROUP BY KIN_DATA.c_personid,
                     BIOG_MAIN.c_name,
                     BIOG_MAIN.c_name_chn
            ORDER BY COUNT(*) DESC
        """),
        # most-used kinship codes
        "top_kin_codes": fetch(conn, """
            SELECT TOP 15 KIN_DATA.c_kin_code,
                   KINSHIP_CODES.c_kinrel,
                   KINSHIP_CODES.c_kinrel_chn,
                   COUNT(*) AS n_rows
            FROM KIN_DATA INNER JOIN KINSHIP_CODES
              ON KIN_DATA.c_kin_code = KINSHIP_CODES.c_kincode
            GROUP BY KIN_DATA.c_kin_code,
                     KINSHIP_CODES.c_kinrel,
                     KINSHIP_CODES.c_kinrel_chn
            ORDER BY COUNT(*) DESC
        """),
    }


def discover_lookatnetworks(conn) -> dict:
    return {
        # persons with most non-kin associations (interesting network seeds)
        "top_persons_by_assoc_count": fetch(conn, """
            SELECT TOP 20 ASSOC_DATA.c_personid,
                   BIOG_MAIN.c_name,
                   BIOG_MAIN.c_name_chn,
                   COUNT(*) AS n_assocs
            FROM ASSOC_DATA INNER JOIN BIOG_MAIN
              ON ASSOC_DATA.c_personid = BIOG_MAIN.c_personid
            GROUP BY ASSOC_DATA.c_personid,
                     BIOG_MAIN.c_name,
                     BIOG_MAIN.c_name_chn
            ORDER BY COUNT(*) DESC
        """),
    }


def discover_lookatassociationpairs(conn) -> dict:
    """Find well-connected (person1, person2) pairs with direct edges."""
    return {
        "top_pairs_by_edge_count": fetch(conn, """
            SELECT TOP 15 c_personid AS person_id_1,
                   c_assoc_id AS person_id_2,
                   COUNT(*) AS n_edges
            FROM ASSOC_DATA
            WHERE c_assoc_id IS NOT NULL
            GROUP BY c_personid, c_assoc_id
            ORDER BY COUNT(*) DESC
        """),
    }


def main() -> None:
    print(f"discovering test inputs from {DATA_MDB.name} ...")
    conn = pyodbc.connect(CONN_STR, autocommit=True)
    out = {
        "lookatentry": discover_lookatentry(conn),
        "lookatstatus": discover_lookatstatus(conn),
        "lookattexts": discover_lookattexts(conn),
        "lookatplace": discover_lookatplace(conn),
        "lookatassociations": discover_lookatassociations(conn),
        "lookatoffice": discover_lookatoffice(conn),
        "lookatkinship": discover_lookatkinship(conn),
        "lookatnetworks": discover_lookatnetworks(conn),
        "lookatassociationpairs": discover_lookatassociationpairs(conn),
    }
    conn.close()

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(out, ensure_ascii=False, indent=2,
                              default=str), encoding="utf-8")
    print(f"wrote {OUT}: {OUT.stat().st_size:,} bytes")
    # quick summary
    for form, data in out.items():
        print(f"\n[{form}]")
        for key, rows in data.items():
            print(f"  {key}: {len(rows)} candidates")
            if rows:
                try:
                    summary = repr(dict(list(rows[0].items())[:5]))
                    print(f"    e.g. top: {summary}")
                except UnicodeEncodeError:
                    print(f"    e.g. top: (contains non-ASCII chars)")


if __name__ == "__main__":
    main()
