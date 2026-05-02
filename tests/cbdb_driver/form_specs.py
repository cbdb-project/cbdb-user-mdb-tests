"""
Per-form metadata for the matrix test framework.

For each LookAt form we record:
  - result_table: where CmdQuery_Click writes
  - picker_table + picker_column: where the codes picker writes
  - addr_picker_table: optional, secondary picker for addresses
  - year_frame_ctl: name of the OptionGroup for year mode
  - year_mode_codes: dict mapping our 'index'/'dynasty'/etc -> int value
  - cmd_caption: button caption (Run Query / Run / Run Network ...)
  - cmd_name: COM control name
  - insert_cols: columns the VBA INSERT writes (catches schema drift)
  - source_sql_template: independent SQL to verify info-loss
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable


@dataclass
class FormSpec:
    name: str                              # form name (LookAtEntry etc.)
    result_table: str
    picker_table: str | None               # primary code picker
    picker_column: str = "c_entry_code"
    addr_picker_table: str | None = None
    year_frame_ctl: str = "FrameYears"     # or FrameFilterYears
    cmd_caption: str = "Run Query"
    cmd_name: str = "CmdQuery"
    insert_cols: set[str] = field(default_factory=set)
    person_id_col: str = "c_personid"      # some forms use c_person_id

    # CmdStoreID_Click writes:
    #   INSERT INTO ZZ_STORE_PERSON_ID (c_personid)
    #   SELECT DISTINCT <storeid_source_col> FROM <storeid_source_table>
    #   [<storeid_source_filter>]
    # then UPDATE PersonIDSource SET SourceForm = '<personid_source_label>'.
    # If storeid_source_table is None the form has no CmdStoreID button.
    storeid_source_table: str | None = None
    storeid_source_col: str = "c_personid"
    storeid_source_filter: str = ""        # e.g. " WHERE c_person_id > 0"
    personid_source_label: str | None = None  # GroupData omits the UPDATE

    # CmdRecallID_Click reads ZZ_STORE_PERSON_ID and INSERTs into this table.
    # Only 4 forms have a CmdRecallID button.
    recallid_target_table: str | None = None


# Per Form_LookAt<X>.vb INSERT INTO ZZ_SCRATCH_<Y> ( ... ):

LOOKATENTRY = FormSpec(
    name="LookAtEntry",
    result_table="ZZ_SCRATCH_ENTRY",
    picker_table="ZZ_SCRATCH_ENTRY_CODE",
    picker_column="c_entry_code",
    addr_picker_table="ZZ_SCRATCH_ADDR",
    year_frame_ctl="FrameYears",
    insert_cols={
        "c_personid", "c_name", "c_name_chn", "c_index_year",
        "c_index_year_type_code", "c_dy", "c_entry_code", "c_year",
        "c_sequence", "c_exam_rank", "c_addr_id", "c_kin_id",
        "c_kin_code", "c_assoc_id", "c_assoc_code",
        "c_parental_status_code", "c_entry_addr_id", "c_source",
        "c_inst_code", "c_inst_name_code", "c_addr_type",
    },
    storeid_source_table="ZZ_SCRATCH_ENTRY",
    storeid_source_col="c_personid",
    personid_source_label="Entry",
)

LOOKATSTATUS = FormSpec(
    name="LookAtStatus",
    result_table="ZZ_SCRATCH_STATUS",
    picker_table="ZZ_STATUS_CODE",
    picker_column="c_status_code",
    addr_picker_table="ZZ_SCRATCH_ADDR",
    year_frame_ctl="FrameFilterYears",
    insert_cols={
        "c_personid", "c_name", "c_name_chn", "c_index_year",
        "c_sex", "c_addr_id", "c_dy", "c_status_code",
        "c_status_desc", "c_status_desc_chn", "c_source",
        "c_index_year_type_code", "c_sequence",
    },
    storeid_source_table="ZZ_SCRATCH_P_STATUS",
    storeid_source_col="c_person_id",
    personid_source_label="Status",
)

LOOKATTEXTS = FormSpec(
    name="LookAtTexts",
    result_table="ZZ_SCRATCH_BIOG_TEXT_DATA",
    picker_table="ZZ_TEXT_BIBLCAT_CODES",
    picker_column="c_text_cat_code",
    addr_picker_table="ZZ_SCRATCH_ADDR",
    year_frame_ctl="FrameFilterYears",
    insert_cols={
        "c_personid", "c_name", "c_name_chn", "c_index_year",
        "c_index_year_type_code", "c_dy", "c_sex", "c_addr_id",
        "c_textid", "c_title", "c_title_chn", "c_role_id",
        "c_role_desc", "c_role_desc_chn", "c_source",
    },
    storeid_source_table="ZZ_SCRATCH_P_TEXT",
    storeid_source_col="c_person_id",
    personid_source_label="Texts",
)

LOOKATASSOCIATIONS = FormSpec(
    name="LookAtAssociations",
    result_table="ZZ_SCRATCH_ASSOC",
    picker_table="ZZ_ASSOC_CODE",
    picker_column="c_assoc_code",
    addr_picker_table="ZZ_SCRATCH_ADDR",
    year_frame_ctl="FrameFilterYears",
    person_id_col="c_person_id",
    insert_cols={
        "c_person_id", "c_name", "c_name_chn", "c_index_year",
        "c_index_year_type_code", "c_sex", "c_addr_id", "c_dy",
        "c_assoc_code", "c_kin_code", "c_kin_id", "c_assoc_id",
        "c_assoc_kin_code", "c_assoc_kin_id", "c_assoc_count",
        "c_assoc_first_year", "c_assoc_last_year", "c_source",
    },
    storeid_source_table="ZZ_SCRATCH_P_ASSOC",
    storeid_source_col="c_person_id",
    personid_source_label="Associations",
)

LOOKATOFFICE = FormSpec(
    name="LookAtOffice",
    result_table="ZZ_SCRATCH_OFFICE",
    picker_table="ZZ_OFFICE_CODE",
    picker_column="c_office_id",
    year_frame_ctl="FrameFilterYears",
    insert_cols={
        "c_posting_id", "c_personid", "c_index_year", "c_female",
        "c_person_dy", "c_office_id", "c_sequence", "c_firstyear",
        "c_lastyear", "c_appt_code", "c_assume_office_code",
        "c_inst_code",
    },
    storeid_source_table="ZZ_SCRATCH_P_OFFICE",
    storeid_source_col="c_personid",
    personid_source_label="Office",
)

LOOKATPLACE = FormSpec(
    name="LookAtPlace",
    result_table="ZZ_PLACE",
    picker_table="ZZ_SCRATCH_ADDR",   # picker is the address table itself
    picker_column="c_addr_id",
    insert_cols={
        "c_personid", "c_name", "c_name_chn", "c_index_year",
        "c_female", "c_dy", "c_addr_id", "c_addr_name",
        "c_firstyear", "c_lastyear", "c_rel_type", "c_source",
    },
    storeid_source_table="ZZ_PLACE",
    storeid_source_col="c_personid",
    personid_source_label="Place",
)

LOOKATASSOCIATIONPAIRS = FormSpec(
    name="LookAtAssociationPairs",
    result_table="ZZ_SOCIAL_NETWORK",
    picker_table=None,   # uses person-id pair via TxtID1 / TxtID2
    picker_column="",
    cmd_caption="Run Network",
    insert_cols={
        "c_person_id", "c_node_id", "c_kin_id", "c_assoc_kin_id",
        "c_assoc_claimer_id",
    },
    person_id_col="c_person_id",
    storeid_source_table="ZZ_SCRATCH_PEOPLE",
    storeid_source_col="c_person_id",
    storeid_source_filter=" WHERE c_person_id > 0",
    personid_source_label="AssocPairs",
    recallid_target_table="ZZ_SCRATCH_IMPORT_PEOPLE",
)

LOOKATKINSHIP = FormSpec(
    name="LookAtKinship",
    result_table="ZZ_SCRATCH_KIN",
    picker_table="ZZ_SCRATCH_IMPORT_PEOPLE",
    picker_column="c_person_id",
    cmd_caption="Run",
    cmd_name="CmdRun",
    insert_cols={
        "c_person_id", "c_kin_id", "c_kin_rel",
    },
    person_id_col="c_person_id",
    storeid_source_table="ZZ_SCRATCH_PEOPLE",
    storeid_source_col="c_person_id",
    personid_source_label="Kinship",
    recallid_target_table="ZZ_SCRATCH_IMPORT_PEOPLE",
)

LOOKATNETWORKS = FormSpec(
    name="LookAtNetworks",
    result_table="ZZ_SOCIAL_NETWORK",
    picker_table="ZZ_SCRATCH_IMPORT_PEOPLE",
    picker_column="c_person_id",
    cmd_caption="Run",
    cmd_name="CmdRun",
    insert_cols={
        "c_person_id", "c_node_id",
    },
    person_id_col="c_person_id",
    storeid_source_table="ZZ_SCRATCH_PEOPLE",
    storeid_source_col="c_person_id",
    personid_source_label="Networks",
    recallid_target_table="ZZ_SCRATCH_IMPORT_PEOPLE",
)

LOOKATGROUPDATA = FormSpec(
    name="LookAtGroupData",
    result_table="ZZ_SCRATCH_PEOPLE",
    picker_table="ZZ_SCRATCH_IMPORT_PEOPLE",
    picker_column="c_person_id",
    cmd_caption="Run",
    cmd_name="CmdRun",
    insert_cols={
        "c_person_id", "c_name", "c_name_chn",
    },
    person_id_col="c_person_id",
    storeid_source_table="ZZ_SCRATCH_IMPORT_PEOPLE",
    storeid_source_col="c_person_id",
    # Form_LookAtGroupData.CmdStoreID_Click does NOT update PersonIDSource.
    personid_source_label=None,
    recallid_target_table="ZZ_SCRATCH_IMPORT_PEOPLE",
)

ALL_SPECS = {
    "LookAtEntry": LOOKATENTRY,
    "LookAtStatus": LOOKATSTATUS,
    "LookAtTexts": LOOKATTEXTS,
    "LookAtAssociations": LOOKATASSOCIATIONS,
    "LookAtOffice": LOOKATOFFICE,
    "LookAtPlace": LOOKATPLACE,
    "LookAtAssociationPairs": LOOKATASSOCIATIONPAIRS,
    "LookAtKinship": LOOKATKINSHIP,
    "LookAtNetworks": LOOKATNETWORKS,
    "LookAtGroupData": LOOKATGROUPDATA,
}


# ----------------------------------------------------------------------
# CmdImport* spec table (roadmap item 13)
#
# Every LookAt form has one or more `CmdImport*_Click` handlers that:
#   1. Pop `Application.FileDialog(msoFileDialogOpen)` for a delimited
#      text file.
#   2. `DoCmd.TransferText acImportDelim, "<spec>", "TempImportList",
#      tFileName, 0` to load it into TempImportList(ImportID, ...).
#   3. `INSERT INTO InputErrorList (c_ID) SELECT TempImportList.ImportID
#      FROM <source_table> RIGHT JOIN TempImportList ON
#      <source_table>.<source_col> = TempImportList.ImportID WHERE
#      <source_table>.<source_col> IS NULL`  -- the unmatched IDs.
#   4. `INSERT INTO <target_table> (<target_col>) SELECT DISTINCT
#      TempImportList.ImportID FROM <source_table> INNER JOIN
#      TempImportList ON ...` -- the matched IDs.
#   5. (Sometimes) sets a `gUse*` global and/or enables follow-up
#      buttons / changes a TxtXxx caption to "[Imported List]".
#
# The seven distinct TransferText specs (saved in MSysIMEXSpecs) imply
# three file formats:
#   - "EntryListImport / AssocCodeListImport / OfficeListImport /
#      StatusCodeListImport / TextBiblcatListImport Specification"
#      → tab-separated, three columns (ImportID, ImportDesc, ImportDescChn)
#   - "ImportPeopleList_Space" → space-separated, one column (ImportID)
#   - "ImportPlaceList_Space"  → comma-separated, one column (ImportID)
# ----------------------------------------------------------------------


@dataclass
class ImportSpec:
    form: str                  # form name (LookAtEntry, ...)
    button: str                # button name (CmdImportEntryCodes, ...)
    target_table: str          # ZZ_SCRATCH_<X>
    target_col: str            # c_entry_code, c_person_id, c_addr_id, ...
    source_table: str          # validation source (ENTRY_CODES, BIOG_MAIN, ...)
    source_col: str            # join column on source side
    file_sep: str              # "\t" / " " / ","
    file_extra_cols: int = 0   # 0 (one-col files) or 2 (tab specs)
    # `expected_global` is documented for reference (which gUse* the
    # handler is supposed to flip).  The test does NOT assert it — see
    # cbdb_driver.vba_session._inject_autodetect for why an inject-
    # based reader had to be backed out (JET re-entrancy).  The data
    # assertion (target table + InputErrorList) is the meaningful
    # contract; the global is set in the same code path.
    expected_global: tuple[str, bool] | None = None   # ("gUseADDRID", True)


# Auto-detected by reading every `Sub CmdImport*_Click` body in
# analysis/dump/vba/Form_LookAt*.vb.  17 buttons total.
ALL_IMPORTS: list[ImportSpec] = [
    # ---- LookAtEntry ----
    ImportSpec("LookAtEntry", "CmdImportEntryCodes",
               "ZZ_SCRATCH_ENTRY_CODE", "c_entry_code",
               "ENTRY_CODES", "c_entry_code",
               file_sep="\t", file_extra_cols=2),
    ImportSpec("LookAtEntry", "CmdImportPlaces",
               "ZZ_SCRATCH_ADDR", "c_addr_id",
               "ADDR_CODES", "c_addr_id",
               file_sep=",", expected_global=("gUseADDRID", True)),
    # ---- LookAtStatus ----
    ImportSpec("LookAtStatus", "CmdImportStatusCodes",
               "ZZ_STATUS_CODE", "c_status_code",
               "STATUS_CODES", "c_status_code",
               file_sep="\t", file_extra_cols=2),
    ImportSpec("LookAtStatus", "CmdImportPlaces",
               "ZZ_SCRATCH_ADDR_LIST", "c_addr_id",
               "ADDR_CODES", "c_addr_id",
               file_sep=","),
    # ---- LookAtTexts ----
    ImportSpec("LookAtTexts", "CmdImportTextCategories",
               "ZZ_TEXT_BIBLCAT_CODES", "c_text_cat_code",
               "TEXT_BIBLCAT_CODES", "c_text_cat_code",
               file_sep="\t", file_extra_cols=2),
    ImportSpec("LookAtTexts", "CmdImportPlaces",
               "ZZ_SCRATCH_ADDR_LIST", "c_addr_id",
               "ADDR_CODES", "c_addr_id",
               file_sep=","),
    # ---- LookAtAssociations ----
    ImportSpec("LookAtAssociations", "CmdImportAssociations",
               "ZZ_ASSOC_CODE", "c_assoc_code",
               "ASSOC_CODES", "c_assoc_code",
               file_sep="\t", file_extra_cols=2),
    ImportSpec("LookAtAssociations", "CmdImportPlaces",
               "ZZ_SCRATCH_ADDR_LIST", "c_addr_id",
               "ADDR_CODES", "c_addr_id",
               file_sep=","),
    # ---- LookAtOffice ----
    ImportSpec("LookAtOffice", "CmdImportOffices",
               "ZZ_OFFICE_CODE", "c_office_id",
               "OFFICE_CODES", "c_office_id",
               file_sep="\t", file_extra_cols=2),
    ImportSpec("LookAtOffice", "CmdImportPlaceOffice",
               "ZZ_SCRATCH_ADDR_OFFICE", "c_addr_id",
               "ADDR_CODES", "c_addr_id",
               file_sep=",",
               expected_global=("gUseOfficeADDRID", True)),
    ImportSpec("LookAtOffice", "CmdImportPlacePeople",
               "ZZ_SCRATCH_ADDR_PEOPLE", "c_addr_id",
               "ADDR_CODES", "c_addr_id",
               file_sep=",",
               expected_global=("gUsePeopleADDRID", True)),
    # ---- LookAtPlace ----
    ImportSpec("LookAtPlace", "CmdImportPlaces",
               "ZZ_SCRATCH_ADDR", "c_addr_id",
               "ADDR_CODES", "c_addr_id",
               file_sep=",",
               expected_global=("gUseADDRID", True)),
    # ---- LookAtKinship ----
    ImportSpec("LookAtKinship", "CmdImport",
               "ZZ_SCRATCH_IMPORT_PEOPLE", "c_person_id",
               "BIOG_MAIN", "c_personid",
               file_sep=" "),
    # ---- LookAtGroupData ----
    ImportSpec("LookAtGroupData", "CmdImport",
               "ZZ_SCRATCH_IMPORT_PEOPLE", "c_person_id",
               "BIOG_MAIN", "c_personid",
               file_sep=" "),
    # ---- LookAtAssociationPairs ----
    ImportSpec("LookAtAssociationPairs", "CmdImportList",
               "ZZ_SCRATCH_IMPORT_PEOPLE", "c_person_id",
               "BIOG_MAIN", "c_personid",
               file_sep=" "),
    # ---- LookAtNetworks ---- (form open hangs in this driver — see
    # findings.md / matrix `_xfail_marks` — but we still record the
    # spec so the test file can mark them skip without re-deriving.)
    ImportSpec("LookAtNetworks", "CmdImportPeople",
               "ZZ_SCRATCH_IMPORT_PEOPLE", "c_person_id",
               "BIOG_MAIN", "c_personid",
               file_sep=" ",
               expected_global=("gUsePersonID", True)),
    ImportSpec("LookAtNetworks", "CmdImportPlaces",
               "ZZ_SCRATCH_ADDR_LIST", "c_addr_id",
               "ADDR_CODES", "c_addr_id",
               file_sep=",",
               expected_global=("gUseADDRID", True)),
]


# ----------------------------------------------------------------------
# CmdSave* spec table (roadmap item 14)
#
# Each `CmdSave*_Click` handler pops `Application.FileDialog(
# msoFileDialogSaveAs)`, opens an ADODB.Stream (utf-8, BOM stripped via
# `Position = 3` + binary CopyTo), runs a `SELECT ... FROM ZZ_SCRATCH_<X>
# INNER JOIN <source_codes>` and writes one tab-separated line per row.
#
# Two output formats observed in the dump:
#   - 3-column: `<id>\t<desc>\t<desc_chn>`
#     (CmdSaveEntryCodes, CmdSaveAssociations)
#   - 1-column-with-trailing-tab: `<id>\t`
#     (CmdSaveOffices, CmdSaveStatusCodes, CmdSaveTextCategories)
#
# The file is UTF-8 without BOM.  `tStream.Position = 3` skips the
# 3-byte UTF-8 BOM that ADODB.Stream prepends when Charset="utf-8";
# the CopyTo to a binary stream then writes the rest verbatim.
# ----------------------------------------------------------------------


@dataclass
class SaveSpec:
    form: str                  # form name (LookAtEntry, ...)
    button: str                # button name (CmdSaveEntryCodes, ...)
    source_table: str          # ZZ_SCRATCH_<X> the save reads from
    source_col: str            # the id column being written
    codes_table: str           # joined-against table for desc lookup
    desc_cols: tuple[str, ...] = ()  # desc columns the SQL projects
    initial_filename: str = ""       # the form's default name


ALL_SAVES: list[SaveSpec] = [
    SaveSpec("LookAtEntry", "CmdSaveEntryCodes",
             "ZZ_SCRATCH_ENTRY_CODE", "c_entry_code",
             "ENTRY_CODES",
             desc_cols=("c_entry_desc", "c_entry_desc_chn"),
             initial_filename="entry_id_list.txt"),
    SaveSpec("LookAtAssociations", "CmdSaveAssociations",
             "ZZ_ASSOC_CODE", "c_assoc_code",
             "ASSOC_CODES",
             desc_cols=("c_assoc_desc", "c_assoc_desc_chn"),
             initial_filename="assoc_code_list.txt"),
    # CmdSaveOffices SELECTs (id, c_office_chn, c_office_trans) but
    # WRITES `id\t<trans-or-empty>\t<chn>` — note the trans/chn swap
    # between SELECT and output.  The check uses output order.
    SaveSpec("LookAtOffice", "CmdSaveOffices",
             "ZZ_OFFICE_CODE", "c_office_id",
             "OFFICE_CODES",
             desc_cols=("c_office_trans", "c_office_chn"),
             initial_filename="office_id_list.txt"),
    SaveSpec("LookAtStatus", "CmdSaveStatusCodes",
             "ZZ_STATUS_CODE", "c_status_code",
             "STATUS_CODES",
             desc_cols=("c_status_desc", "c_status_desc_chn"),
             initial_filename="status_code_list.txt"),
    SaveSpec("LookAtTexts", "CmdSaveTextCategories",
             "ZZ_TEXT_BIBLCAT_CODES", "c_text_cat_code",
             "TEXT_BIBLCAT_CODES",
             desc_cols=("c_text_cat_desc", "c_text_cat_desc_chn"),
             initial_filename="text_bilblcat_list.txt"),  # note: typo 'bilblcat' is in CBDB
]
