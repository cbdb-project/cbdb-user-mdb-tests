-- QueryDef name: BM IY Rule 03 BY Phase 1 Query
-- Source: CBDB_20260430_DATA.mdb

INSERT INTO tmpIndexYear ( c_personid, c_index_year )
SELECT BIOG_MAIN.c_personid, BIOG_MAIN.c_birthyear
FROM BIOG_MAIN
WHERE (((BIOG_MAIN.c_index_year) Is Null Or (BIOG_MAIN.c_index_year)=0) AND ((BIOG_MAIN.c_birthyear)>0) AND ((BIOG_MAIN.c_deathyear) Is Null Or (BIOG_MAIN.c_deathyear)=0));

