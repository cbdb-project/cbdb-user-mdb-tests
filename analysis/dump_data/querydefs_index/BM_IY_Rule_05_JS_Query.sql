-- QueryDef name: BM IY Rule 05 JS Query
-- Source: CBDB_20260430_DATA.mdb

UPDATE BIOG_MAIN INNER JOIN ENTRY_DATA ON BIOG_MAIN.c_personid = ENTRY_DATA.c_personid SET BIOG_MAIN.c_index_year = [ENTRY_DATA].[c_year]+30, BIOG_MAIN.c_notes = 'Index year algorithmically generated: Rule 3; '+[BIOG_MAIN].[c_notes]
WHERE (((BIOG_MAIN.c_index_year) Is Null Or (BIOG_MAIN.c_index_year)=0) AND ((ENTRY_DATA.c_entry_code)=36 Or (ENTRY_DATA.c_entry_code)=165) AND ((ENTRY_DATA.c_year)>0));

