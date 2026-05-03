-- QueryDef name: BM IY Rule 01A DY Phase 1 Query
-- Source: CBDB_20260430_DATA.mdb

INSERT INTO tmpIndexYear ( c_index_year, c_personid )
SELECT BIOG_MAIN.c_deathyear, BIOG_MAIN.c_personid
FROM BIOG_MAIN
WHERE (((BIOG_MAIN.c_index_year) Is Null Or (BIOG_MAIN.c_index_year)=0) And ((BIOG_MAIN.c_birthyear)>=BIOG_MAIN.c_deathyear-60) And ((BIOG_MAIN.c_deathyear)>0));

