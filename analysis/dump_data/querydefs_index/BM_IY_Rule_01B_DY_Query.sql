-- QueryDef name: BM IY Rule 01B DY Query
-- Source: CBDB_20260430_DATA.mdb

UPDATE BIOG_MAIN SET BIOG_MAIN.c_index_year = [BIOG_MAIN].[c_birthyear]+59, BIOG_MAIN.c_notes = [BIOG_MAIN].[c_notes]='Index year algorithmically generated: Rule 1; '+[BIOG_MAIN].[c_notes]
WHERE (((BIOG_MAIN.c_index_year) Is Null Or (BIOG_MAIN.c_index_year)=0) And ((BIOG_MAIN.c_birthyear)<=BIOG_MAIN.c_deathyear-59) And ((BIOG_MAIN.c_deathyear)>0));

