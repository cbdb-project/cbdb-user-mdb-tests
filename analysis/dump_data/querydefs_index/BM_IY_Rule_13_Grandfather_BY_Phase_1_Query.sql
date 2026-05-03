-- QueryDef name: BM IY Rule 13 Grandfather BY Phase 1 Query
-- Source: CBDB_20260430_DATA.mdb

INSERT INTO tmpIndexYear ( c_personid, c_index_year )
SELECT BIOG_MAIN_1.c_personid, Max(BIOG_MAIN.c_birthyear) AS MaxOfc_birthyear
FROM BIOG_MAIN AS BIOG_MAIN_1 INNER JOIN (BIOG_MAIN INNER JOIN KIN_DATA ON BIOG_MAIN.c_personid = KIN_DATA.c_kin_id) ON BIOG_MAIN_1.c_personid = KIN_DATA.c_personid
WHERE (((BIOG_MAIN_1.c_index_year) Is Null Or (BIOG_MAIN_1.c_index_year)=0) AND ((KIN_DATA.c_kin_code)=62) AND ((BIOG_MAIN.c_birthyear)>0))
GROUP BY BIOG_MAIN_1.c_personid;

