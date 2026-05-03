-- QueryDef name: BM IY Rule 08 Father BY Query
-- Source: CBDB_20260430_DATA.mdb

UPDATE BIOG_MAIN AS BIOG_MAIN_1 INNER JOIN (BIOG_MAIN INNER JOIN KIN_DATA ON BIOG_MAIN.c_personid = KIN_DATA.c_kin_id) ON BIOG_MAIN_1.c_personid = KIN_DATA.c_personid SET BIOG_MAIN_1.c_index_year = [BIOG_MAIN].[c_birthyear]+89, BIOG_MAIN_1.c_notes = 'Index year algorithmically generated: Rule 5W; '+[BIOG_MAIN_1].[c_notes]
WHERE (((BIOG_MAIN_1.c_index_year) Is Null Or (BIOG_MAIN_1.c_index_year)=0) AND ((KIN_DATA.c_kin_code)=75) AND ((BIOG_MAIN.c_birthyear)>0) AND ((BIOG_MAIN.c_index_year) Is Null Or (BIOG_MAIN.c_index_year)=0));

