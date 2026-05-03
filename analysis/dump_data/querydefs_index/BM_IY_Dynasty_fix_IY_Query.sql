-- QueryDef name: BM IY Dynasty fix IY Query
-- Source: CBDB_20260430_DATA.mdb

UPDATE BIOG_MAIN INNER JOIN tmpBM_NIY_finished ON BIOG_MAIN.c_personid = tmpBM_NIY_finished.c_personid SET BIOG_MAIN.c_index_year = -1000
WHERE (((BIOG_MAIN.c_index_year)>0 And (BIOG_MAIN.c_index_year)<300) AND ((BIOG_MAIN.c_dy)>4 And (BIOG_MAIN.c_dy)<21));

