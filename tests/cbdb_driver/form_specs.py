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
