-- QueryDef name: BM IY Rule 02 DY Query
-- Source: CBDB_20260430_DATA.mdb

UPDATE BIOG_MAIN SET BIOG_MAIN.c_index_year = [BIOG_MAIN].[c_deathyear], BIOG_MAIN.c_notes = [BIOG_MAIN].[c_notes]='Index year algorithmically generated: Rule 1; '+[BIOG_MAIN].[c_notes]
WHERE (((BIOG_MAIN.c_index_year) Is Null Or (BIOG_MAIN.c_index_year)=0) AND ((BIOG_MAIN.c_birthyear) Is Null Or (BIOG_MAIN.c_birthyear)=0) AND ((BIOG_MAIN.c_deathyear)>0));

