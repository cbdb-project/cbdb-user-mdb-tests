-- QueryDef name: BM IY Rule 07W Husband XC Query
-- Source: CBDB_20260430_DATA.mdb

UPDATE (KIN_DATA AS KIN_DATA_1 INNER JOIN ((BIOG_MAIN INNER JOIN KIN_DATA ON BIOG_MAIN.c_personid = KIN_DATA.c_kin_id) INNER JOIN tmpBM_NIY ON KIN_DATA.c_personid = tmpBM_NIY.c_personid) ON (KIN_DATA_1.c_kin_id = KIN_DATA.c_personid) AND (KIN_DATA_1.c_personid = KIN_DATA.c_kin_id)) INNER JOIN ENTRY_DATA ON BIOG_MAIN.c_personid = ENTRY_DATA.c_personid SET tmpBM_NIY.c_index_year = [ENTRY_DATA].[c_year]+42, tmpBM_NIY.c_rule = "7W"
WHERE (((tmpBM_NIY.c_index_year) Is Null Or (tmpBM_NIY.c_index_year)=0) AND ((KIN_DATA.c_kin_code)=134) AND ((KIN_DATA_1.c_kin_code)<>39 And (KIN_DATA_1.c_kin_code)<>168) AND ((ENTRY_DATA.c_year)>0) AND ((ENTRY_DATA.c_entry_code)=257));

