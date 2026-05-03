-- QueryDef name: BM IY Rule 14 Father IY Phase 1 Query
-- Source: CBDB_20260430_DATA.mdb

INSERT INTO tmpIndexYear ( c_index_year, c_personid )
SELECT Max(BIOG_MAIN.c_index_year) AS MaxOfc_index_year, BIOG_MAIN_1.c_personid
FROM BIOG_MAIN AS BIOG_MAIN_1 INNER JOIN (BIOG_MAIN INNER JOIN KIN_DATA ON BIOG_MAIN.c_personid = KIN_DATA.c_kin_id) ON BIOG_MAIN_1.c_personid = KIN_DATA.c_personid
WHERE (((BIOG_MAIN_1.c_index_year) Is Null Or (BIOG_MAIN_1.c_index_year)=0) And ((KIN_DATA.c_kin_code)=75) And (InStr(1,BIOG_MAIN.c_notes,"Index year algorithmically")="0"))
GROUP BY BIOG_MAIN_1.c_personid
HAVING (((Max(BIOG_MAIN.c_index_year))>0));

