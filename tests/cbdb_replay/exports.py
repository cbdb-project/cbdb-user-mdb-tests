"""
Python re-implementations of the export Subs in each LookAt form.

Each function takes the source data (typically the populated
ZZ_SCRATCH_<XXX> table) and produces the exact byte-stream the form's
VBA would have written to a file.  Tests compare against frozen
golden files — any drift in column set, ordering, or NULL handling
shows up as a diff.

Why this matters: the user's #1 reported pain points are
  - "导出功能不 work"
  - "导出的栏位不对或者丢数据"
  - "触发错误提示"
These tests catch the first two directly (via golden diff).  The
third — VBA error messages — requires the COM driver path and is
covered by the (currently skipped) test_infra_smoke.py work.
"""
from __future__ import annotations

from typing import Iterable

import pandas as pd
import pyodbc


# ----------------------------------------------------------------------
# LookAtEntry — CmdGIS_Click  (analysis/dump/vba/Form_LookAtEntry.vb:93)
# ----------------------------------------------------------------------

GIS_HEADER_COLS = [
    "PersonID", "Name", "NameChn", "IndexYear",
    "EntryDesc", "EntryDescChn", "ExamRank",
    "KinName", "KinNameChn", "KinshipRel",
    "AddrName", "AddrNameChn", "x_coord", "y_coord",
    "XY", "xy_count",
    "EntryAddrName", "EntryAddrChn",
    "Entry_xcoord", "Entry_ycoord", "Entry_xy_count",
]

# (source_column, missing_placeholder)
GIS_FIELDS = [
    ("c_personid",         "[Person ID Missing]", "int"),
    ("c_name",             "[Name Missing]",      "str"),
    ("c_name_chn",         "[Name Missing]",      "str"),
    ("c_index_year",       "[ ]",                 "int"),
    ("c_entry_desc",       "[ ]",                 "str"),
    ("c_entry_chn",        "[ ]",                 "str"),
    ("c_exam_rank",        "[ ]",                 "str"),
    ("c_kin_name",         "[ ]",                 "str"),
    ("c_kin_chn",          "[ ]",                 "str"),
    ("c_kin_desc",         "[ ]",                 "str"),
    ("c_addr_name",        "[Addr Name Missing]", "str"),
    ("c_addr_chn",         "[Addr Chn Missing]",  "str"),
    ("x_coord",            "[ ]",                 "coord"),
    ("y_coord",            "[ ]",                 "coord"),
    # XY is a derived column: "x,y" if x is non-null, else "[ ]"
    ("__xy__",             "[ ]",                 "xy"),
    ("xy_count",           "[ ]",                 "int"),
    ("c_entry_addr_name",  "[ ]",                 "str"),
    ("c_entry_addr_chn",   "[ ]",                 "str"),
    ("c_entry_xcoord",     "[ ]",                 "coord"),
    ("c_entry_ycoord",     "[ ]",                 "coord"),
    ("c_entry_xy_count",   "[ ]",                 "int"),
]


def _vba_str(v) -> str:
    """Mimic VBA ``Str(n)`` for integers: leading space if positive."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    n = int(v)
    return f" {n}" if n >= 0 else str(n)


def _vba_cstr(v) -> str:
    """Mimic VBA ``CStr`` on a coordinate (Double): no leading space,
    Access's Double-to-string formatting (no trailing zeros)."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    f = float(v)
    if f == int(f):
        # VBA CStr on a whole-valued Double prints "123" (no decimal)
        return str(int(f))
    return repr(f)  # repr gives shortest accurate representation


def lookat_entry_gis(rows: Iterable[dict]) -> str:
    """Re-implement Form_LookAtEntry.CmdGIS_Click writeText output.

    Input: iterable of dicts, one per ZZ_SCRATCH_ENTRY row, with the
    columns referenced in GIS_FIELDS.  pandas.DataFrame.to_dict('records')
    works.

    Output: the full text content (header + body, lines joined with "\\n").
    Lines are written via tStream.WriteText with adWriteLine which appends
    Chr(13)+Chr(10) (CRLF).
    """
    tab = "\t"
    crlf = "\r\n"
    out = [tab.join(GIS_HEADER_COLS)]
    for r in rows:
        parts = []
        for src, missing, kind in GIS_FIELDS:
            if src == "__xy__":
                x = r.get("x_coord")
                if x is None or (isinstance(x, float) and pd.isna(x)):
                    parts.append(missing)
                else:
                    parts.append(f"{_vba_cstr(x)},{_vba_cstr(r.get('y_coord'))}")
                continue
            v = r.get(src)
            if v is None or (isinstance(v, float) and pd.isna(v)):
                parts.append(missing)
            elif kind == "int":
                parts.append(_vba_str(v))
            elif kind == "coord":
                parts.append(_vba_cstr(v))
            else:
                parts.append(str(v))
        out.append(tab.join(parts))
    return crlf.join(out) + crlf


def lookat_entry_gis_from_db(conn: pyodbc.Connection,
                             *, order_by: str | None = "c_personid, c_year") -> str:
    """Convenience: read ZZ_SCRATCH_ENTRY and feed lookat_entry_gis."""
    sql = "SELECT * FROM ZZ_SCRATCH_ENTRY"
    if order_by:
        sql += f" ORDER BY {order_by}"
    df = pd.read_sql(sql, conn)
    return lookat_entry_gis(df.to_dict("records"))


# ----------------------------------------------------------------------
# LookAtEntry — CmdNeo4j_Click (1 of 6+ files: People.csv)
# (Form_LookAtEntry.vb:472, People.csv at line 635-748)
# ----------------------------------------------------------------------

NEO4J_PEOPLE_HEADER = "nameID,nameHZ,namePY,indexyear,dynasty,sex"


def lookat_entry_neo4j_people(rows: Iterable[dict]) -> str:
    """Re-implement the People.csv portion of CmdNeo4j_Click.

    Input rows: ZZ_SCRATCH_PEOPLE records (after the 3-step
    de-dup from ZZ_SCRATCH_ENTRY's c_personid + c_kin_id + c_assoc_id
    → ZZ_SCRATCH_P_TEXT → joined to BIOG_MAIN/DYNASTIES/ADDR_CODES).
    For unit-testing the WRITER we accept a list of dicts directly.

    Format (UTF-8 case, the only one used per VBA hard-coding):
        nameID,nameHZ,namePY,indexyear,dynasty,sex
        c_person_id, c_name_chn, c_name, c_index_year, c_dynasty_chn, F/M

    NULL handling:
        - c_name_chn / c_name → "Missing"
        - c_index_year → "-2000"
        - c_dynasty (and dynasty_chn) → "unknown"
    """
    sep = ","
    crlf = "\r\n"
    out = [NEO4J_PEOPLE_HEADER]
    for r in rows:
        fields: list[str] = []
        # nameID = Trim(Str(c_person_id))
        fields.append(_trim_str(r.get("c_person_id")))
        # nameHZ
        chn = r.get("c_name_chn")
        fields.append("Missing" if _is_null(chn) else str(chn))
        # namePY
        py = r.get("c_name")
        fields.append("Missing" if _is_null(py) else str(py))
        # indexyear
        iy = r.get("c_index_year")
        fields.append("-2000" if _is_null(iy) else _trim_str(iy))
        # dynasty (Chinese form, since UTF-8 branch)
        dyn = r.get("c_dynasty")     # null check on c_dynasty (per VBA)
        if _is_null(dyn):
            fields.append("unknown")
        else:
            dyn_chn = r.get("c_dynasty_chn")
            fields.append("unknown" if _is_null(dyn_chn) else str(dyn_chn))
        # sex: c_female True -> F, False -> M  (VBA IIf(!c_female, "F", "M"))
        female = r.get("c_female")
        fields.append("F" if female else "M")
        out.append(sep.join(fields))
    return crlf.join(out) + crlf


def _is_null(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    return False


def _trim_str(v) -> str:
    """VBA Trim(Str(n)) — Str() prefixes positive ints with a space,
    Trim() strips it. So this is just a no-leading-space stringifier."""
    if _is_null(v):
        return ""
    if isinstance(v, bool):
        return "-1" if v else "0"   # VBA boolean cast is -1/0, but rare
    return str(int(v) if isinstance(v, (int, float)) and float(v).is_integer()
               else v)


# ----------------------------------------------------------------------
# LookAtEntry — writeKML (Form_LookAtEntry.vb:2685)
# ----------------------------------------------------------------------

KML_HEADER_LINES = [
    '<kml xmlns="http://www.opengis.net/kml/2.2">',
    '<Document>',
    '\t<name>ExtendedData+SchemaData</name>',
    '\t<open>1</open>',
    '\t<!-- Create a balloon template referring to the user-defined type -->',
    '\t<Style id="entry-balloon-template">',
    '\t\t<BalloonStyle>',
    '\t\t\t<text>',
    '\t\t\t\t<![CDATA[',
    '\t\t\t\t$[EntryPerson/PersonNameHZ] <br/>',
    '\t\t\t\tID: $[EntryPerson/PersonID] <br/>',
    '\t\t\t\tIndex Year: $[EntryPerson/IndexYear] <br/>',
    '\t\t\t\tEntry Year: $[EntryPerson/EntryYear] <br/>',
    '\t\t\t\tEntry Desc: $[EntryPerson/EntryDesc] <br/>',
    '\t\t\t\tEntry Chn: $[EntryPerson/EntryDescHZ] <br/>',
    '\t\t\t\tEntry Rank: $[EntryPerson/EntryRank] <br/>',
    '\t\t\t\tAddress: $[EntryPerson/AddrName] $[EntryPerson/AddrNameHZ] <br/>',
    '\t\t\t\tXY Count: $[EntryPerson/XYCount] <br/><br/>',
    '\t\t\t\t]]>',
    '\t\t\t</text>',
    '\t\t</BalloonStyle>',
    '\t</Style>',
    '\t<!-- Declare the type "EntryPerson" with 10 fields -->',
    '\t<Schema name="EntryPerson" id="EntryPersonId">',
    '\t\t<SimpleField type="string" name="PersonNameHZ">',
    '\t\t\t<displayName><![CDATA[<b>Person</b>]]></displayName>',
    '\t\t</SimpleField>',
    '\t\t<SimpleField type="string" name="AddrName">',
    '\t\t\t<displayName><![CDATA[<b>Person</b>]]></displayName>',
    '\t\t</SimpleField>',
    '\t\t<SimpleField type="string" name="AddrNameHZ">',
    '\t\t\t<displayName><![CDATA[<b>Person</b>]]></displayName>',
    '\t\t</SimpleField>',
    '\t\t<SimpleField type="uint" name="PersonID">',
    '\t\t\t<displayName><![CDATA[ID]]></displayName>',
    '\t\t</SimpleField>',
    '\t\t<SimpleField type="string" name="IndexYear">',
    '\t\t\t<displayName><![CDATA[Index Year]]></displayName>',
    '\t\t</SimpleField>',
    '\t\t<SimpleField type="int" name="EntryYear">',
    '\t\t\t<displayName><![CDATA[Entry Year]]></displayName>',
    '\t\t</SimpleField>',
    '\t\t<SimpleField type="string" name="EntryDesc">',
    '\t\t\t<displayName><![CDATA[Entry Desc]]></displayName>',
    '\t\t</SimpleField>',
    '\t\t<SimpleField type="string" name="EntryDescHZ">',
    '\t\t\t<displayName><![CDATA[Entry Chn]]></displayName>',
    '\t\t</SimpleField>',
    '\t\t<SimpleField type="int" name="EntryRank">',
    '\t\t\t<displayName><![CDATA[Entry Rank]]></displayName>',
    '\t\t</SimpleField>',
    '\t\t<SimpleField type="int" name="XYCount">',
    '\t\t\t<displayName><![CDATA[XY Count]]></displayName>',
    '\t\t</SimpleField>',
    '\t</Schema>',
]


def _placemark_str(field: str) -> str:
    """Mimic VBA Str(n): leading space for positive ints, no space for
    negative ints. Used in TimeStamp / IndexYear / EntryYear."""
    if _is_null(field):
        return ""
    n = int(field) if isinstance(field, float) and float(field).is_integer() else field
    if isinstance(n, (int, float)):
        return f" {int(n)}" if n >= 0 else str(int(n))
    return str(n)


def lookat_entry_kml(rows: Iterable[dict]) -> str:
    """Re-implement Form_LookAtEntry.writeKML.

    Output: complete KML 2.2 document (header + per-row Placemark + footer).
    Each Placemark contains a Schema-conformant ExtendedData block plus a
    Point with "x,y" coordinates.

    NULL placeholders (per VBA):
      - c_name        → "[Bad Data]"
      - c_index_year  → "N/A"
      - c_year        → "-2000"
      - c_entry_desc  → "[Missing Data]"
      - c_entry_chn   → "[Missing Data]"
      - c_exam_rank   → "0"
      - c_addr_name / c_addr_chn (incl. blank/whitespace) → "[?]"
      - xy_count      → "0"
      - x_coord / y_coord → "0"
      - c_name_chn (null OR blank) → "[Bad Data]" / "[?]"
        (note VBA bug: when null branch concatenates "[Bad Data]" with
         the previous tStr — we faithfully reproduce that)
    """
    crlf = "\r\n"
    lines: list[str] = list(KML_HEADER_LINES)

    for r in rows:
        # name (the Placemark's <name> element)
        c_name = r.get("c_name")
        name_str = "[Bad Data]" if _is_null(c_name) else str(c_name)
        # IndexYear timestamp
        c_iy = r.get("c_index_year")
        iy_str = "N/A" if _is_null(c_iy) else _placemark_str(c_iy)
        # PersonID
        pid = r.get("c_personid")
        pid_str = _placemark_str(pid)
        # Chinese Name — VBA bug: null branch sets `tStr = tStr + "[Bad Data]"`
        # which appends "[Bad Data]" to the prior PersonID string.
        # Faithful reproduction:
        c_chn = r.get("c_name_chn")
        if _is_null(c_chn):
            chn_str = pid_str + "[Bad Data]"  # bug-faithful concatenation
        elif str(c_chn).strip() == "":
            chn_str = "[?]"
        else:
            chn_str = str(c_chn)
        # IndexYear (again, the SimpleData block — same value)
        iy2 = "N/A" if _is_null(c_iy) else _placemark_str(c_iy)
        # EntryYear
        c_year = r.get("c_year")
        ey_str = "-2000" if _is_null(c_year) else _placemark_str(c_year)
        # EntryDesc
        c_ed = r.get("c_entry_desc")
        ed_str = "[Missing Data]" if _is_null(c_ed) else str(c_ed)
        # EntryDescHZ
        c_ec = r.get("c_entry_chn")
        ec_str = "[Missing Data]" if _is_null(c_ec) else str(c_ec)
        # EntryRank
        rank = r.get("c_exam_rank")
        rank_str = "0" if _is_null(rank) else str(rank)
        # AddrName
        an = r.get("c_addr_name")
        an_str = "[?]" if _is_null(an) or str(an).strip() == "" else str(an)
        # AddrNameHZ
        ah = r.get("c_addr_chn")
        ah_str = "[?]" if _is_null(ah) or str(ah).strip() == "" else str(ah)
        # XYCount
        xyc = r.get("xy_count")
        xyc_str = "0" if _is_null(xyc) else _placemark_str(xyc)
        # Coordinates: x,y (each defaults to "0")
        x = r.get("x_coord")
        y = r.get("y_coord")
        x_str = "0" if _is_null(x) else _placemark_str(x)
        y_str = "0" if _is_null(y) else _placemark_str(y)

        lines.extend([
            "\t<Placemark>",
            f"\t\t<name>{name_str}</name>",
            "\t\t<styleUrl>#entry-balloon-template</styleUrl>",
            f"\t\t<TimeStamp>{iy_str}</TimeStamp>",
            "\t\t<ExtendedData>",
            "\t\t\t<SchemaData schemaUrl=\"#EntryPersonId\">",
            f"\t\t\t\t<SimpleData name=\"PersonID\">{pid_str}</SimpleData>",
            f"\t\t\t\t<SimpleData name=\"PersonNameHZ\">{chn_str}</SimpleData>",
            f"\t\t\t\t<SimpleData name=\"IndexYear\">{iy2}</SimpleData>",
            f"\t\t\t\t<SimpleData name=\"EntryYear\">{ey_str}</SimpleData>",
            f"\t\t\t\t<SimpleData name=\"EntryDesc\">{ed_str}</SimpleData>",
            f"\t\t\t\t<SimpleData name=\"EntryDescHZ\">{ec_str}</SimpleData>",
            f"\t\t\t\t<SimpleData name=\"EntryRank\">{rank_str}</SimpleData>",
            f"\t\t\t\t<SimpleData name=\"AddrName\">{an_str}</SimpleData>",
            f"\t\t\t\t<SimpleData name=\"AddrNameHZ\">{ah_str}</SimpleData>",
            f"\t\t\t\t<SimpleData name=\"XYCount\">{xyc_str}</SimpleData>",
            "\t\t\t</SchemaData>",
            "\t\t</ExtendedData>",
            "\t\t<Point>",
            f"\t\t\t<coordinates>{x_str},{y_str}</coordinates>",
            "\t\t</Point>",
            "\t</Placemark>",
        ])

    lines.append("</Document>")
    lines.append("</kml>")
    return crlf.join(lines) + crlf
