-- QueryDef name: BM IY Rule 05 JS Phase 1 Query
-- Source: CBDB_20260430_DATA.mdb

INSERT INTO tmpIndexYear ( c_personid, c_index_year )
SELECT BIOG_MAIN.c_personid, ENTRY_DATA.c_year
FROM BIOG_MAIN INNER JOIN ENTRY_DATA ON BIOG_MAIN.c_personid = ENTRY_DATA.c_personid
WHERE (((ENTRY_DATA.c_year)>0) AND ((BIOG_MAIN.c_index_year) Is Null Or (BIOG_MAIN.c_index_year)=0) AND ((ENTRY_DATA.c_entry_code)=36 Or (ENTRY_DATA.c_entry_code)=165));

