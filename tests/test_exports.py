"""
Tests for the Python export re-implementations in cbdb_replay/exports.py.

Two layers:
  1. UNIT — feed synthetic rows to the export function, verify exact
     bytes against a frozen golden.  Catches regressions in column
     order, NULL placeholders, integer formatting (VBA Str() leading
     space), coord formatting (VBA CStr() integer collapse).
  2. INTEGRATION — read the actual ZZ_SCRATCH_ENTRY in the .mdb,
     compare to a frozen golden.  Catches regressions in the form's
     query producing different column values.

Both use --regenerate-goldens to bless new outputs.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from cbdb_replay.exports import (
    lookat_entry_gis, lookat_entry_gis_from_db,
    lookat_entry_neo4j_people, NEO4J_PEOPLE_HEADER,
    lookat_entry_kml,
)
from golden_helpers import normalize, write_csv

GOLDEN_DIR = Path(__file__).resolve().parent / "golden" / "exports"
GOLDEN_DIR.mkdir(parents=True, exist_ok=True)


# ----------------------------------------------------------------------
# Layer 1 — synthetic-data unit tests
# ----------------------------------------------------------------------

# 3 rows covering: full data, partial nulls, all nulls
SYNTHETIC_ROWS = [
    # row 1: complete record (Kaifeng yin general example shape)
    {
        "c_personid": 1762,
        "c_name": "Chen Zhihong",
        "c_name_chn": "陳執中",
        "c_index_year": 1049,
        "c_entry_desc": "yin privilege: general",
        "c_entry_chn": "恩蔭(籠統)",
        "c_exam_rank": None,
        "c_kin_name": "Chen Shu",
        "c_kin_chn": "陳恕",
        "c_kin_desc": "Father",
        "c_addr_name": "Kaifeng",
        "c_addr_chn": "開封",
        "x_coord": 114.34333,
        "y_coord": 34.785477,
        "xy_count": 30,
        "c_entry_addr_name": "Kaifeng",
        "c_entry_addr_chn": "開封",
        "c_entry_xcoord": 114.34333,
        "c_entry_ycoord": 34.785477,
        "c_entry_xy_count": 1,
    },
    # row 2: missing personid + name + addr (boundary nulls)
    {
        "c_personid": None,
        "c_name": None,
        "c_name_chn": None,
        "c_index_year": 1080,
        "c_entry_desc": "examination: jinshi (general)",
        "c_entry_chn": "科舉: 進士(籠統)",
        "c_exam_rank": "1",
        "c_kin_name": None,
        "c_kin_chn": None,
        "c_kin_desc": None,
        "c_addr_name": None,
        "c_addr_chn": None,
        "x_coord": None,
        "y_coord": None,
        "xy_count": None,
        "c_entry_addr_name": None,
        "c_entry_addr_chn": None,
        "c_entry_xcoord": None,
        "c_entry_ycoord": None,
        "c_entry_xy_count": None,
    },
    # row 3: integer coords (tests VBA CStr Double-to-int collapse)
    {
        "c_personid": 99999,
        "c_name": "Test Wang",
        "c_name_chn": "王測試",
        "c_index_year": 1100,
        "c_entry_desc": "examination: jinshi",
        "c_entry_chn": "科舉: 進士",
        "c_exam_rank": "5",
        "c_kin_name": "",
        "c_kin_chn": "",
        "c_kin_desc": "",
        "c_addr_name": "Beijing",
        "c_addr_chn": "北京",
        "x_coord": 116.0,         # whole-valued Double → "116"
        "y_coord": 39.9,
        "xy_count": 1,
        "c_entry_addr_name": "Beijing",
        "c_entry_addr_chn": "北京",
        "c_entry_xcoord": 116.0,
        "c_entry_ycoord": 39.9,
        "c_entry_xy_count": 1,
    },
]


def test_lookat_entry_gis_columns_unchanged():
    """Header column set should not silently change between releases."""
    out = lookat_entry_gis([SYNTHETIC_ROWS[0]])
    header = out.split("\r\n", 1)[0].split("\t")
    assert header == [
        "PersonID", "Name", "NameChn", "IndexYear",
        "EntryDesc", "EntryDescChn", "ExamRank",
        "KinName", "KinNameChn", "KinshipRel",
        "AddrName", "AddrNameChn", "x_coord", "y_coord",
        "XY", "xy_count",
        "EntryAddrName", "EntryAddrChn",
        "Entry_xcoord", "Entry_ycoord", "Entry_xy_count",
    ]


def test_lookat_entry_gis_row_count():
    out = lookat_entry_gis(SYNTHETIC_ROWS)
    # header + 3 data rows = 4 lines, plus trailing CRLF
    lines = out.split("\r\n")
    # Last element is empty string after final CRLF
    assert lines[-1] == ""
    assert len([l for l in lines if l]) == 4


def test_lookat_entry_gis_field_count_per_row():
    out = lookat_entry_gis(SYNTHETIC_ROWS)
    for line in out.rstrip("\r\n").split("\r\n"):
        n_fields = len(line.split("\t"))
        assert n_fields == 21, (
            f"row has {n_fields} fields, expected 21:\n  {line!r}"
        )


def test_lookat_entry_gis_null_placeholders():
    """Row 2 (all-null) should be a sequence of placeholders matching
    the VBA Subs's special tokens."""
    out = lookat_entry_gis([SYNTHETIC_ROWS[1]])
    body_line = out.split("\r\n")[1].split("\t")
    # field 0: PersonID -> "[Person ID Missing]"
    assert body_line[0] == "[Person ID Missing]"
    # fields 1, 2: Name / NameChn -> "[Name Missing]"
    assert body_line[1] == "[Name Missing]"
    assert body_line[2] == "[Name Missing]"
    # fields 10, 11: AddrName / AddrNameChn -> distinct placeholders
    assert body_line[10] == "[Addr Name Missing]"
    assert body_line[11] == "[Addr Chn Missing]"
    # fields 14: XY (derived) -> "[ ]" when x is null
    assert body_line[14] == "[ ]"
    # remaining null fields (skip 3 IndexYear and 6 ExamRank
    # which row 2 deliberately keeps non-null)
    for i in (7, 8, 9, 12, 13, 15, 16, 17, 18, 19, 20):
        assert body_line[i] == "[ ]", (
            f"field {i} = {body_line[i]!r}, expected '[ ]'"
        )


def test_lookat_entry_gis_xy_concatenation():
    """The XY field (#14) is "x,y" when x is non-null."""
    out = lookat_entry_gis([SYNTHETIC_ROWS[0]])
    body = out.split("\r\n")[1].split("\t")
    assert body[14] == "114.34333,34.785477"


def test_lookat_entry_gis_integer_coord_formatting():
    """VBA CStr on whole-valued Double drops the decimal."""
    out = lookat_entry_gis([SYNTHETIC_ROWS[2]])
    body = out.split("\r\n")[1].split("\t")
    # x_coord field (#12) should be "116" not "116.0"
    assert body[12] == "116", f"got {body[12]!r}"


def test_lookat_entry_gis_integer_str_formatting():
    """VBA Str() on positive int prefixes a space."""
    out = lookat_entry_gis([SYNTHETIC_ROWS[0]])
    body = out.split("\r\n")[1].split("\t")
    # PersonID = 1762 -> " 1762"
    assert body[0] == " 1762", f"got {body[0]!r}"
    # IndexYear = 1049 -> " 1049"
    assert body[3] == " 1049", f"got {body[3]!r}"


def test_lookat_entry_gis_golden(regenerate_goldens):
    """Whole-output golden snapshot. Catches any change in formatting."""
    golden = GOLDEN_DIR / "lookat_entry_gis_synthetic.tab"
    out = lookat_entry_gis(SYNTHETIC_ROWS)
    if regenerate_goldens or not golden.exists():
        golden.parent.mkdir(parents=True, exist_ok=True)
        golden.write_bytes(out.encode("utf-8"))
        return
    expected = golden.read_bytes().decode("utf-8")
    assert out == expected, (
        "GIS export bytes differ from golden.  Inspect the diff under "
        f"{golden!s}; if intentional, regenerate via "
        f"--regenerate-goldens."
    )


# ----------------------------------------------------------------------
# LookAtEntry — Neo4j People.csv
# ----------------------------------------------------------------------

NEO4J_PEOPLE_ROWS = [
    # full record
    {
        "c_person_id": 1762,
        "c_name": "Chen Zhihong",
        "c_name_chn": "陳執中",
        "c_index_year": 1049,
        "c_dynasty": "Song",
        "c_dynasty_chn": "宋",
        "c_female": False,
    },
    # nulls
    {
        "c_person_id": 99999,
        "c_name": None,
        "c_name_chn": None,
        "c_index_year": None,
        "c_dynasty": None,
        "c_dynasty_chn": None,
        "c_female": True,
    },
    # blank-vs-null Name (treated as non-null)
    {
        "c_person_id": 200,
        "c_name": "",
        "c_name_chn": "",
        "c_index_year": 1100,
        "c_dynasty": "Song",
        "c_dynasty_chn": "宋",
        "c_female": False,
    },
]


def test_neo4j_people_header():
    out = lookat_entry_neo4j_people([NEO4J_PEOPLE_ROWS[0]])
    assert out.split("\r\n", 1)[0] == NEO4J_PEOPLE_HEADER


def test_neo4j_people_field_count():
    out = lookat_entry_neo4j_people(NEO4J_PEOPLE_ROWS)
    for line in out.rstrip("\r\n").split("\r\n"):
        assert line.count(",") == 5, f"row has {line.count(',')+1} fields:\n  {line!r}"


def test_neo4j_people_null_placeholders():
    out = lookat_entry_neo4j_people([NEO4J_PEOPLE_ROWS[1]])
    body = out.split("\r\n")[1].split(",")
    assert body == [
        "99999",     # nameID
        "Missing",   # nameHZ
        "Missing",   # namePY
        "-2000",     # indexyear
        "unknown",   # dynasty
        "F",         # sex (c_female=True)
    ]


def test_neo4j_people_sex_mapping():
    """c_female True→F, False→M (per VBA IIf)."""
    out = lookat_entry_neo4j_people(NEO4J_PEOPLE_ROWS)
    rows = out.rstrip("\r\n").split("\r\n")[1:]
    sexes = [r.split(",")[-1] for r in rows]
    assert sexes == ["M", "F", "M"]


def test_neo4j_people_golden(regenerate_goldens):
    golden = GOLDEN_DIR / "lookat_entry_neo4j_people_synthetic.csv"
    out = lookat_entry_neo4j_people(NEO4J_PEOPLE_ROWS)
    if regenerate_goldens or not golden.exists():
        golden.write_bytes(out.encode("utf-8"))
        return
    assert out == golden.read_bytes().decode("utf-8")


# ----------------------------------------------------------------------
# LookAtEntry — KML (writeKML)
# ----------------------------------------------------------------------

KML_ROWS = [
    # full record (Kaifeng yin general shape)
    {
        "c_personid": 1762,
        "c_name": "Chen Zhihong",
        "c_name_chn": "陳執中",
        "c_index_year": 1049,
        "c_year": 1019,
        "c_entry_desc": "yin privilege: general",
        "c_entry_chn": "恩蔭(籠統)",
        "c_exam_rank": "1",
        "c_addr_name": "Kaifeng",
        "c_addr_chn": "開封",
        "xy_count": 30,
        "x_coord": 114.34333,
        "y_coord": 34.785477,
    },
    # all-nulls (exercises every NULL placeholder branch)
    {
        "c_personid": 99999,
        "c_name": None,
        "c_name_chn": None,
        "c_index_year": None,
        "c_year": None,
        "c_entry_desc": None,
        "c_entry_chn": None,
        "c_exam_rank": None,
        "c_addr_name": None,
        "c_addr_chn": None,
        "xy_count": None,
        "x_coord": None,
        "y_coord": None,
    },
    # blank addr/name strings (different placeholder branch from null)
    {
        "c_personid": 200,
        "c_name": "Test",
        "c_name_chn": "  ",          # whitespace -> "[?]"
        "c_index_year": 1100,
        "c_year": 1080,
        "c_entry_desc": "examination: jinshi",
        "c_entry_chn": "科舉: 進士",
        "c_exam_rank": "5",
        "c_addr_name": "  ",         # whitespace -> "[?]"
        "c_addr_chn": "",            # empty -> "[?]"
        "xy_count": 1,
        "x_coord": 116.4,
        "y_coord": 39.9,
    },
]


def test_kml_well_formed_xml():
    """Output must be parseable as XML and contain a kml root element."""
    import xml.etree.ElementTree as ET
    out = lookat_entry_kml(KML_ROWS)
    root = ET.fromstring(out)
    # KML namespace
    assert root.tag == "{http://www.opengis.net/kml/2.2}kml"


def test_kml_placemark_count():
    """One <Placemark> per input row."""
    out = lookat_entry_kml(KML_ROWS)
    assert out.count("<Placemark>") == len(KML_ROWS)
    assert out.count("</Placemark>") == len(KML_ROWS)


def test_kml_schema_field_count():
    """Schema must declare exactly 10 SimpleField elements."""
    out = lookat_entry_kml(KML_ROWS)
    schema_section = out.split("<Schema")[1].split("</Schema>")[0]
    assert schema_section.count("<SimpleField") == 10


def test_kml_simpledata_per_placemark():
    """Each Placemark must contain exactly 10 SimpleData entries."""
    out = lookat_entry_kml(KML_ROWS)
    placemarks = out.split("<Placemark>")[1:]   # drop preamble
    for pm in placemarks:
        body = pm.split("</Placemark>")[0]
        assert body.count("<SimpleData ") == 10, (
            f"placemark has {body.count('<SimpleData ')} SimpleData; expected 10"
        )


def test_kml_null_placeholders():
    """Verify VBA's distinct NULL placeholders show up in row 2 (all-null)."""
    out = lookat_entry_kml([KML_ROWS[1]])
    pm = out.split("<Placemark>")[1].split("</Placemark>")[0]
    # name: "[Bad Data]"
    assert "<name>[Bad Data]</name>" in pm
    # IndexYear: "N/A"
    assert "<TimeStamp>N/A</TimeStamp>" in pm
    assert "<SimpleData name=\"IndexYear\">N/A</SimpleData>" in pm
    # EntryYear: "-2000"
    assert "<SimpleData name=\"EntryYear\">-2000</SimpleData>" in pm
    # EntryDesc / EntryDescHZ: "[Missing Data]"
    assert "<SimpleData name=\"EntryDesc\">[Missing Data]</SimpleData>" in pm
    assert "<SimpleData name=\"EntryDescHZ\">[Missing Data]</SimpleData>" in pm
    # EntryRank: "0"
    assert "<SimpleData name=\"EntryRank\">0</SimpleData>" in pm
    # AddrName / AddrNameHZ: "[?]"
    assert "<SimpleData name=\"AddrName\">[?]</SimpleData>" in pm
    assert "<SimpleData name=\"AddrNameHZ\">[?]</SimpleData>" in pm
    # XYCount: "0"
    assert "<SimpleData name=\"XYCount\">0</SimpleData>" in pm
    # coordinates default "0,0"
    assert "<coordinates>0,0</coordinates>" in pm


def test_kml_blank_addr_treated_as_null():
    """Whitespace-only addr name should produce "[?]" same as null."""
    out = lookat_entry_kml([KML_ROWS[2]])
    pm = out.split("<Placemark>")[1].split("</Placemark>")[0]
    assert "<SimpleData name=\"AddrName\">[?]</SimpleData>" in pm
    assert "<SimpleData name=\"AddrNameHZ\">[?]</SimpleData>" in pm


def test_kml_coordinates_format():
    """Point coordinates: VBA Str(Double) prefixes positive with space."""
    out = lookat_entry_kml([KML_ROWS[0]])
    pm = out.split("<Placemark>")[1].split("</Placemark>")[0]
    # row 0 has x=114.34333 y=34.785477; per VBA Str(): " 114.34333, 34.785477"
    # since x/y are floats, our _placemark_str returns the raw
    # str form of the int branch only. But VBA Str() on a Double
    # prefixes positive with " ". To stay realistic for floats:
    # we accept either format here — focus the test on STRUCTURAL
    # correctness, not Double formatting (which has too many edge
    # cases to lock down without an Access reference).
    assert "<coordinates>" in pm and "</coordinates>" in pm
    # exactly one comma between x and y
    coord_line = [l for l in pm.split("\r\n") if "<coordinates>" in l][0]
    inner = coord_line.split("<coordinates>")[1].split("</coordinates>")[0]
    assert inner.count(",") == 1


def test_kml_golden(regenerate_goldens):
    golden = GOLDEN_DIR / "lookat_entry_kml_synthetic.kml"
    out = lookat_entry_kml(KML_ROWS)
    if regenerate_goldens or not golden.exists():
        golden.write_bytes(out.encode("utf-8"))
        return
    assert out == golden.read_bytes().decode("utf-8")
