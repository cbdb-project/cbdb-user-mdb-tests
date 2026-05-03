-- QueryDef name: BM IY Rule 04W Query
-- Source: CBDB_20260430_DATA.mdb

UPDATE (BIOG_MAIN AS BIOG_MAIN_1 INNER JOIN (BIOG_MAIN INNER JOIN KIN_DATA ON BIOG_MAIN.c_personid = KIN_DATA.c_personid) ON BIOG_MAIN_1.c_personid = KIN_DATA.c_kin_id) INNER JOIN KIN_DATA AS KIN_DATA_1 ON (BIOG_MAIN_1.c_personid = KIN_DATA_1.c_personid) AND (BIOG_MAIN.c_personid = KIN_DATA_1.c_kin_id) SET BIOG_MAIN.c_index_year = [BIOG_MAIN_1].[c_birthyear]+62, BIOG_MAIN.c_notes = 'Index year algorithmically generated: Rule 2; '+[BIOG_MAIN].[c_notes]
WHERE (((BIOG_MAIN.c_index_year) Is Null Or (BIOG_MAIN.c_index_year)=0) AND ((KIN_DATA.c_kin_code)=134) AND ((BIOG_MAIN_1.c_birthyear)>0) AND ((KIN_DATA_1.c_kin_code)<>168));

