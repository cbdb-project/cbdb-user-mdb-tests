-- QueryDef name: BM IY Dynasty Query
-- Source: CBDB_20260430_DATA.mdb

UPDATE BIOG_MAIN, DYNASTIES SET BIOG_MAIN.c_dy = [DYNASTIES].[c_dy], BIOG_MAIN.c_notes = "Dynasty calculated from index year; "
WHERE (((BIOG_MAIN.c_dy) Is Null Or (BIOG_MAIN.c_dy)=0) And ((BIOG_MAIN.c_index_year)>0 And (BIOG_MAIN.c_index_year)>=DYNASTIES.c_start And (BIOG_MAIN.c_index_year)<=DYNASTIES.c_end));

