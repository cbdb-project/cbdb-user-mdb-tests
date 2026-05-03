-- QueryDef name: BM IY Rule 06 JR Phase 1 Query
-- Source: CBDB_20260430_DATA.mdb

UPDATE BIOG_MAIN INNER JOIN ENTRY_DATA ON BIOG_MAIN.c_personid = ENTRY_DATA.c_personid SET BIOG_MAIN.c_index_year = [ENTRY_DATA].[c_year]+33, BIOG_MAIN.c_notes = 'Index year algorithmically generated: Rule 6; '+[BIOG_MAIN].[c_notes]
WHERE (((BIOG_MAIN.c_index_year) Is Null Or (BIOG_MAIN.c_index_year)=0) AND ((ENTRY_DATA.c_entry_code)=39) AND ((ENTRY_DATA.c_year)>0));

