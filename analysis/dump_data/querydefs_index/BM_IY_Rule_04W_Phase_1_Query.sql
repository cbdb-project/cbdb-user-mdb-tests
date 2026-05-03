-- QueryDef name: BM IY Rule 04W Phase 1 Query
-- Source: CBDB_20260430_DATA.mdb

INSERT INTO tmpIndexYear ( c_personid, c_index_year )
SELECT BIOG_MAIN.c_personid, BIOG_MAIN_1.c_birthyear
FROM (BIOG_MAIN AS BIOG_MAIN_1 INNER JOIN (BIOG_MAIN INNER JOIN KIN_DATA ON BIOG_MAIN.c_personid = KIN_DATA.c_personid) ON BIOG_MAIN_1.c_personid = KIN_DATA.c_kin_id) INNER JOIN KIN_DATA AS KIN_DATA_1 ON (BIOG_MAIN_1.c_personid = KIN_DATA_1.c_personid) AND (BIOG_MAIN.c_personid = KIN_DATA_1.c_kin_id)
WHERE (((BIOG_MAIN.c_index_year) Is Null Or (BIOG_MAIN.c_index_year)=0) AND ((KIN_DATA.c_kin_code)=134) AND ((BIOG_MAIN_1.c_birthyear)>0) AND ((KIN_DATA_1.c_kin_code)<>168));

