-- QueryDef name: BM IY Dynasty Fix Query
-- Source: CBDB_20260430_DATA.mdb

UPDATE BIOG_MAIN1 INNER JOIN BIOG_MAIN ON BIOG_MAIN1.c_personid = BIOG_MAIN.c_personid SET BIOG_MAIN.c_notes = [BIOG_MAIN].[c_notes]+[BIOG_MAIN1].[c_notes]
WHERE (((Right([BIOG_MAIN].[c_notes],36))="Dynasty calculated from index year; "));

