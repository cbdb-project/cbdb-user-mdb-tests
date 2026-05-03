-- QueryDef name: BM IY Rule 12 SIL BY Phase 1 Doublecheck Query
-- Source: CBDB_20260430_DATA.mdb

SELECT BIOG_MAIN.c_personid, BIOG_MAIN.c_index_year, BIOG_MAIN_1.c_birthyear
FROM BIOG_MAIN AS BIOG_MAIN_1 INNER JOIN (BIOG_MAIN INNER JOIN KIN_DATA ON BIOG_MAIN.c_personid = KIN_DATA.c_personid) ON BIOG_MAIN_1.c_personid = KIN_DATA.c_kin_id
WHERE (((BIOG_MAIN_1.c_birthyear)>0) AND ((KIN_DATA.c_kin_code)=181 Or (KIN_DATA.c_kin_code)=201 Or (KIN_DATA.c_kin_code)=224 Or (KIN_DATA.c_kin_code)=332));

