-- QueryDef name: BM IY Rule 13 Grandfather BY Phase 2 Null Query
-- Source: CBDB_20260430_DATA.mdb

UPDATE BIOG_MAIN INNER JOIN tmpIndexYear ON BIOG_MAIN.c_personid = tmpIndexYear.c_personid SET BIOG_MAIN.c_index_year = [tmpIndexYear].[c_index_year]+119, BIOG_MAIN.c_notes = "Index year algorithmically generated: Rule 13; "
WHERE (((BIOG_MAIN.c_notes) Is Null));

