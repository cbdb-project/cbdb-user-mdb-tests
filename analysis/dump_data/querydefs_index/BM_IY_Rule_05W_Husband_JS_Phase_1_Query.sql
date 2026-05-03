-- QueryDef name: BM IY Rule 05W Husband JS Phase 1 Query
-- Source: CBDB_20260430_DATA.mdb

INSERT INTO tmpIndexYear ( c_personid, c_index_year )
SELECT BIOG_MAIN_1.c_personid, ENTRY_DATA.c_year
FROM BIOG_MAIN AS BIOG_MAIN_1 INNER JOIN ((BIOG_MAIN INNER JOIN KIN_DATA ON BIOG_MAIN.c_personid = KIN_DATA.c_kin_id) INNER JOIN ENTRY_DATA ON BIOG_MAIN.c_personid = ENTRY_DATA.c_personid) ON BIOG_MAIN_1.c_personid = KIN_DATA.c_personid
WHERE (((BIOG_MAIN_1.c_index_year) Is Null Or (BIOG_MAIN_1.c_index_year)=0) AND ((KIN_DATA.c_kin_code)=134) AND ((ENTRY_DATA.c_year)>0) AND ((ENTRY_DATA.c_entry_code)=36 Or (ENTRY_DATA.c_entry_code)=165) AND ((BIOG_MAIN.c_index_year) Is Null Or (BIOG_MAIN.c_index_year)=0));

