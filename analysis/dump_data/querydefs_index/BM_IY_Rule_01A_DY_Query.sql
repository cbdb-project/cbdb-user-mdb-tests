-- QueryDef name: BM IY Rule 01A DY Query
-- Source: CBDB_20260430_DATA.mdb

UPDATE BIOG_MAIN SET BIOG_MAIN.c_index_year = [BIOG_MAIN].[c_deathyear], BIOG_MAIN.c_notes = 'Index year algorithmically generated: Rule 13; '+[BIOG_MAIN].[c_notes]
WHERE (((BIOG_MAIN.c_index_year) Is Null Or (BIOG_MAIN.c_index_year)=0) And ((BIOG_MAIN.c_birthyear)>=BIOG_MAIN.c_deathyear-60) And ((BIOG_MAIN.c_deathyear)>0));

