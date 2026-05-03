-- QueryDef name: BM IY Rule 07 XC Query
-- Source: CBDB_20260430_DATA.mdb

UPDATE ENTRY_DATA INNER JOIN tmpBM_NIY ON ENTRY_DATA.c_personid = tmpBM_NIY.c_personid SET tmpBM_NIY.c_index_year = [ENTRY_DATA].[c_year]+39, tmpBM_NIY.c_rule = "7"
WHERE (((tmpBM_NIY.c_index_year) Is Null Or (tmpBM_NIY.c_index_year)=0) AND ((ENTRY_DATA.c_entry_code)=257) AND ((ENTRY_DATA.c_year)>0));

