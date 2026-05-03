-- QueryDef name: BM IY Rule 08 Father BY Phase 1 Query
-- Source: CBDB_20260430_DATA.mdb

INSERT INTO tmpIndexYear ( c_personid, c_index_year )
SELECT BIOG_MAIN_1.c_personid, BIOG_MAIN.c_birthyear
FROM BIOG_MAIN AS BIOG_MAIN_1 INNER JOIN (BIOG_MAIN INNER JOIN KIN_DATA ON BIOG_MAIN.c_personid = KIN_DATA.c_kin_id) ON BIOG_MAIN_1.c_personid = KIN_DATA.c_personid
WHERE (((BIOG_MAIN_1.c_index_year) Is Null Or (BIOG_MAIN_1.c_index_year)=0) AND ((KIN_DATA.c_kin_code)=75) AND ((BIOG_MAIN.c_birthyear)>0));

