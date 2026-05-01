Option Compare Database
Public gFirstTime As Integer

Private Sub Form_Current()
    Dim tRst As DAO.Recordset
    
    If gFirstTime = 0 Then
        gFirstTime = 1
    Else
        Set tRst = Forms!CBDB_Browser_2!BIOG_MAIN_2_Subform.Form.Recordset
        tRst.FindFirst "c_personid = " & Str(c_personid.Value)
        If tRst.NoMatch Then
            gPersonID = 0
            Forms!CBDB_Browser_2.Form.CmdSaveToFile.Enabled = False
        Else
            gPersonID = c_personid.Value
            Forms!CBDB_Browser_2!BIOG_MAIN_2_Subform.Form.Refresh
            Call getKinship(c_personid.Value)
            Forms!CBDB_Browser_2.Form.CmdSaveToFile.Enabled = True
        End If
    End If
End Sub

Private Sub getKinship(t_personid As Long)
    '
    ' this routine searches for the immediate kin of the current person
    '
    Dim tMaxUp As Integer, tMaxDown As Integer, tMaxCol As Integer, tMaxMarr As Integer
    Dim tTrue As Integer, tFalse As Integer, tLoopCount As Long, tErrorStr As String
    Dim tContinue As Integer, tAddrID As Long, tExitDo As Boolean, tRecCount As Long, tRecDelete As Long
    Dim tRstDummy As DAO.Recordset, tAppendQuery As QueryDef
    Dim tSeekStr As String, tLoopMax As Long, tLoopInfoStr As String, tKinQueryStr As String, tQueryStr As String
    Dim tNodeDistQueryStr As String, tPruneTmpQueryDupesStr As String, tPruneTmpQuery As String
    Dim tPruneInversesQueryStr1 As String, tPruneTmpInversesQueryStr1 As String, tPruneInversesQueryStr2 As String, tAppendQueryStr As String
    Dim tPruneTmpInversesQueryStr2 As String
    Dim tKinFirstQueryStr As String, tPruneTmpQueryDupesStr2 As String, tPruneTmpQuery2 As String
    
    tTrue = -1
    tFalse = 0
    
    tLoopMax = 10
    tMaxUp = 2
    tMaxDown = 2
    tMaxCol = 1
    tMaxMarr = 1
    '
    Dim KinQuery As DAO.QueryDef
    Dim prm As DAO.Parameter
    Dim cmdSQL As ADODB.Command, tRecDeleted As Long, strSQL As String
    '
    ' Clear the tables
    '
    Set cmdSQL = New ADODB.Command
    cmdSQL.ActiveConnection = CurrentProject.Connection
    cmdSQL.CommandType = adCmdText
    '
    '  clear the working files
    '
    cmdSQL.CommandText = "Delete * from ZZ_KIN_LIST"
    cmdSQL.Execute tRecCount
    '
    cmdSQL.CommandText = "Delete * from ZZ_KIN_LIST_TMP"
    cmdSQL.Execute tRecCount
    '
    ' now zap the ego-relative form person file
    '
    cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_KIN"
    cmdSQL.Execute tRecCount
    '
    cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_KINNET"
    cmdSQL.Execute tRecDeleted
    
    '
    '  this copies the people on the import list (which is just the selected person if one does not use a list)
    '
    tQueryStr = "INSERT INTO ZZ_KIN_LIST ( c_personid, c_kin_id, c_kinrel, c_kinrel_total, c_kinrel_total_raw, c_kinrel_total_simplified, " + _
        "c_kin_code, c_up_total, c_down_total, c_mar_total, c_col_total, c_distance, c_up, c_down, c_mar, c_col, " + _
        "c_prior_female, c_kin_female, c_kin_sex, c_female, c_sex, c_personid_root ) " + _
        "SELECT BIOG_MAIN.c_personid, BIOG_MAIN.c_personid AS c_kin_id, " + _
            "'ego' AS c_kin_rel, 'ego' AS c_kin_rel_total, 'ego' AS c_kin_rel_total_raw, 'ego' AS c_kin_rel_total_simplified, -3 AS c_kin_code, " + _
            "0 AS c_up_total, 0 AS c_down_total, 0 AS c_mar_total, 0 AS c_col_total, 0 AS c_distance, " + _
            "0 AS c_up, 0 AS c_down, 0 AS c_mar, 0 AS c_col, BIOG_MAIN.c_female AS c_prior_female, BIOG_MAIN.c_female AS c_kin_female, " + _
            "iif(BIOG_MAIN.c_female,'F','M'), BIOG_MAIN.c_female, iif(BIOG_MAIN.c_female,'F','M'), BIOG_MAIN.c_personid AS c_personid_root " + _
        "FROM BIOG_MAIN WHERE BIOG_MAIN.c_personid = " + Str(t_personid)
    
    ' the initial list of "ego" roots is now intialized in ZZ_KIN_LIST, and the personID is stored as c_personid_root:  use this to create ZZ_KIN_LIST_TMP
    ' in the first query, one begins to build out with a first layer of kinship relations
    ' as the first layer, we put the kin_rel as both the kin_rel and the kin_rel_total
    
    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecCount
    
    ' the new logic is to not test the metrics until after the reduction routine is run
    ' the reduction routine will remove the first layer of kin who do not meet the 2-2-1-1 test criterion; these will need to be added back in at the end
    
    tKinFirstQueryStr = "INSERT INTO ZZ_KIN_LIST_TMP ( c_personid, c_kin_id, c_kinrel, c_kinrel_total, c_kinrel_total_raw, c_kinrel_total_simplified, " + _
                "c_kin_code, c_up_total, c_down_total, c_mar_total, c_col_total, c_distance, c_up, c_down, c_mar, c_col, c_personid_root, " + _
                "c_prior_female, c_kin_female, c_kin_sex, c_female, c_sex, c_notes, c_source ) " + _
            "SELECT DISTINCT KIN_DATA.c_personid, KIN_DATA.c_kin_id, KINSHIP_CODES.c_kinrel, KINSHIP_CODES.c_kinrel_simplified AS c_kinrel_total, " + _
                "KINSHIP_CODES.c_kinrel AS c_kinrel_total_raw, KINSHIP_CODES.c_kinrel_simplified AS c_kinrel_total_simplified, KIN_DATA.c_kin_code, " + _
                "KINSHIP_CODES.c_upstep AS c_up_total, KINSHIP_CODES.c_dwnstep AS c_down_total, KINSHIP_CODES.c_marstep AS c_mar_total, " + _
                "KINSHIP_CODES.c_colstep AS c_col_total, 0 AS c_distance, KINSHIP_CODES.c_upstep, KINSHIP_CODES.c_dwnstep, KINSHIP_CODES.c_marstep, " + _
                "KINSHIP_CODES.c_colstep,  ZZ_KIN_LIST.c_personid_root, ZZ_KIN_LIST.c_female AS c_prior_female, BIOG_MAIN.c_female, " + _
                "IIf([BIOG_MAIN].[c_female], 'F', 'M') AS c_kin_sex, BIOG_MAIN_1.c_female, IIf([BIOG_MAIN_1].[c_female], 'F', 'M') AS c_sex, " + _
                "'Notes: ' + [BIOG_MAIN_1].[c_name_chn] + ' > ' + [BIOG_MAIN].[c_name_chn] + ' (' + [KINSHIP_CODES].[c_kinrel] + ') ' AS c_notes, " + _
                "KIN_DATA.c_source " + _
            "FROM ( ( ( ZZ_KIN_LIST INNER JOIN KIN_DATA ON ZZ_KIN_LIST.c_kin_id = KIN_DATA.c_personid ) INNER JOIN KINSHIP_CODES " + _
                "ON KIN_DATA.c_kin_code = KINSHIP_CODES.c_kincode ) INNER JOIN BIOG_MAIN ON KIN_DATA.c_kin_id = BIOG_MAIN.c_personid ) " + _
                "INNER JOIN BIOG_MAIN AS BIOG_MAIN_1 ON KIN_DATA.c_personid = BIOG_MAIN_1.c_personid"
  
  ' each subsequent layer adds the new kin_rel to the kin_rel_total and the total cumulative steps are summed
    
    tKinQueryStr = "INSERT INTO ZZ_KIN_LIST_TMP ( c_personid, c_kin_id, c_kinrel, c_kinrel_total_simplified, c_kinrel_total, c_kinrel_total_raw, c_kin_code, " + _
                "c_up_total, c_down_total, c_mar_total, c_col_total, c_distance, c_up, c_down, c_mar, c_col, c_personid_root, c_prior_female, " + _
                "c_kin_female, c_kin_sex, c_female, c_sex, c_notes, c_source ) " + _
            "SELECT DISTINCT KIN_DATA.c_personid, KIN_DATA.c_kin_id, KINSHIP_CODES.c_kinrel, [ZZ_KIN_LIST].[c_kinrel_total_simplified] + " + _
                "[KINSHIP_CODES].[c_kinrel_simplified] AS c_kinrel_total_simplified, " + _
                "[ZZ_KIN_LIST].[c_kinrel_total] + [KINSHIP_CODES].[c_kinrel_simplified] AS c_kinrel_total, " + _
                "[ZZ_KIN_LIST].[c_kinrel_total_raw] + [KINSHIP_CODES].[c_kinrel] AS c_kinrel_total_raw, " + _
                "KIN_DATA.c_kin_code, [KINSHIP_CODES].[c_upstep] + [ZZ_KIN_LIST].[c_up_total] AS c_up_total, " + _
                "[KINSHIP_CODES].[c_dwnstep] + [ZZ_KIN_LIST].[c_down_total] AS c_down_total, " + _
                "[KINSHIP_CODES].[c_marstep] + [ZZ_KIN_LIST].[c_mar_total] AS c_mar_total, " + _
                "[KINSHIP_CODES].[c_colstep] + [ZZ_KIN_LIST].[c_col_total] AS c_col_total, ZZ_KIN_LIST.c_distance, " + _
                "KINSHIP_CODES.c_upstep, KINSHIP_CODES.c_dwnstep, KINSHIP_CODES.c_marstep, KINSHIP_CODES.c_colstep, ZZ_KIN_LIST.c_personid_root, " + _
                "ZZ_KIN_LIST.c_female AS c_prior_female, BIOG_MAIN_1.c_female, IIf([BIOG_MAIN_1].[c_female], 'F', 'M') AS c_kin_sex, " + _
                "BIOG_MAIN.c_female, IIf([BIOG_MAIN].[c_female], 'F', 'M') AS Expr2, " + _
                "[ZZ_KIN_LIST].[c_notes] + ' > ' + [BIOG_MAIN_1].[c_name_chn] + ' (' + [KINSHIP_CODES].[c_kinrel] + ') ' AS c_notes, KIN_DATA.c_source " + _
            "FROM BIOG_MAIN AS BIOG_MAIN_1 INNER JOIN ( BIOG_MAIN INNER JOIN ( KINSHIP_CODES INNER JOIN ( ZZ_KIN_LIST INNER JOIN KIN_DATA " + _
                "ON ZZ_KIN_LIST.c_kin_id = KIN_DATA.c_personid ) ON KINSHIP_CODES.c_kincode = KIN_DATA.c_kin_code) " + _
                "ON BIOG_MAIN.c_personid = KIN_DATA.c_personid ) ON BIOG_MAIN_1.c_personid = KIN_DATA.c_kin_id " + _
            "WHERE ((ZZ_KIN_LIST.c_distance)="

  '  the various queries for cleaning up the results (need editing)
    '  ZZ_KIN_LIST is our collection of current results
    '  ZZ_KIN_LIST_TMP is the new material coming from the most recent query loop which looks for kin of the c_kin_id
    
    '  if the new kin (in c_kin_id) does not already show up as a relative of someone else already in the database, this is one more step distant
    tNodeDistQueryStr = "UPDATE ZZ_KIN_LIST_TMP LEFT JOIN ZZ_KIN_LIST ON ZZ_KIN_LIST_TMP.c_kin_id = ZZ_KIN_LIST.c_kin_id " + _
        "SET ZZ_KIN_LIST_TMP.c_distance = [ZZ_KIN_LIST_TMP].[c_distance]+1 " + _
        "WHERE (((ZZ_KIN_LIST.c_personid) Is Null))"
    '
    '  for insurance, explicitly delete duplicate results
    '
    tPruneTmpQuery = "UPDATE ZZ_KIN_LIST INNER JOIN ZZ_KIN_LIST_TMP ON " + _
        "(ZZ_KIN_LIST.c_kin_id = ZZ_KIN_LIST_TMP.c_kin_id) AND " + _
        "(ZZ_KIN_LIST.c_personid = ZZ_KIN_LIST_TMP.c_personid) " + _
        "SET ZZ_KIN_LIST_TMP.c_delete = 1;"

    '  delete inverse results
    tPruneTmpQuery2 = "UPDATE ZZ_KIN_LIST INNER JOIN ZZ_KIN_LIST_TMP ON " + _
        "(ZZ_KIN_LIST.c_personid = ZZ_KIN_LIST_TMP.c_kin_id) AND " + _
        "(ZZ_KIN_LIST.c_kin_id = ZZ_KIN_LIST_TMP.c_personid) " + _
        "SET ZZ_KIN_LIST_TMP.c_delete = 1;"
    
    tPruneTmpQueryDupesStr = "UPDATE ZZ_KIN_LIST_TMP AS ZZ_KIN_LIST_TMP_1 INNER JOIN " + _
        "ZZ_KIN_LIST_TMP ON (ZZ_KIN_LIST_TMP_1.c_personid = ZZ_KIN_LIST_TMP.c_personid) " + _
        "AND (ZZ_KIN_LIST_TMP_1.c_kin_id = ZZ_KIN_LIST_TMP.c_kin_id) " + _
        "AND (ZZ_KIN_LIST_TMP_1.c_kin_code = ZZ_KIN_LIST_TMP.c_kin_code) " + _
        "SET ZZ_KIN_LIST_TMP.c_delete = 1 " + _
        "WHERE (([ZZ_KIN_LIST_TMP].[c_up_total]*1000+[ZZ_KIN_LIST_TMP].[c_down_total]*100+[ZZ_KIN_LIST_TMP].[c_col_total]*10+[ZZ_KIN_LIST_TMP].[c_mar_total]>" + _
        "[ZZ_KIN_LIST_TMP_1].[c_up_total]*1000+[ZZ_KIN_LIST_TMP_1].[c_down_total]*100+[ZZ_KIN_LIST_TMP_1].[c_col_total]*10+[ZZ_KIN_LIST_TMP_1].[c_mar_total]))"
    
    '  if the data is good I should not need to do this
    tPruneTmpQueryDupesStr2 = "UPDATE ZZ_KIN_LIST_TMP AS ZZ_KIN_LIST_TMP_1 INNER JOIN " + _
        "ZZ_KIN_LIST_TMP ON (ZZ_KIN_LIST_TMP_1.c_personid = ZZ_KIN_LIST_TMP.c_personid) " + _
        "AND (ZZ_KIN_LIST_TMP_1.c_kin_id = ZZ_KIN_LIST_TMP.c_kin_id) " + _
        "AND (ZZ_KIN_LIST_TMP_1.c_kin_code = ZZ_KIN_LIST_TMP.c_kin_code) " + _
        "SET ZZ_KIN_LIST_TMP.c_delete = 1 " + _
        "WHERE (((StrComp([ZZ_KIN_LIST_TMP].[c_kinrel_total_raw], [ZZ_KIN_LIST_TMP_1].[c_kinrel_total_raw])) > 0))"
        
    tPruneInversesQueryStr1 = "UPDATE ZZ_KIN_LIST INNER JOIN (KINSHIP_CODES INNER JOIN ZZ_KIN_LIST_TMP ON KINSHIP_CODES.c_kincode = ZZ_KIN_LIST_TMP.c_kin_code) ON " + _
        "(ZZ_KIN_LIST.c_kin_id = ZZ_KIN_LIST_TMP.c_personid) AND (ZZ_KIN_LIST.c_personid = ZZ_KIN_LIST_TMP.c_kin_id) SET ZZ_KIN_LIST_TMP.c_delete = 1 " + _
        "WHERE (((ZZ_KIN_LIST.c_kin_code)=[KINSHIP_CODES].[c_kin_pair1])) OR (((ZZ_KIN_LIST.c_kin_code)=[KINSHIP_CODES].[c_kin_pair2]))"
        
    tPruneInversesQueryStr2 = "UPDATE (ZZ_KIN_LIST INNER JOIN ZZ_KIN_LIST_TMP ON (ZZ_KIN_LIST.c_personid = ZZ_KIN_LIST_TMP.c_kin_id) AND " + _
        "(ZZ_KIN_LIST.c_kin_id = ZZ_KIN_LIST_TMP.c_personid)) INNER JOIN KINSHIP_CODES ON ZZ_KIN_LIST.c_kin_code = KINSHIP_CODES.c_kincode SET ZZ_KIN_LIST_TMP.c_delete = 1 " + _
        "WHERE (((ZZ_KIN_LIST_TMP.c_kin_code)=[KINSHIP_CODES].[c_kin_pair1] Or (ZZ_KIN_LIST_TMP.c_kin_code)=[KINSHIP_CODES].[c_kin_pair2]))"

    
    tPruneTmpInversesQueryStr1 = "UPDATE KINSHIP_CODES INNER JOIN (ZZ_KIN_LIST_TMP AS ZZ_KIN_LIST_TMP_1 " + _
        "INNER JOIN ZZ_KIN_LIST_TMP ON (ZZ_KIN_LIST_TMP_1.c_personid = ZZ_KIN_LIST_TMP.c_kin_id) AND " + _
        "(ZZ_KIN_LIST_TMP_1.c_kin_id = ZZ_KIN_LIST_TMP.c_personid)) ON " + _
        "KINSHIP_CODES.c_kincode = ZZ_KIN_LIST_TMP.c_kin_code SET ZZ_KIN_LIST_TMP.c_delete = 1 " + _
        "WHERE (((ZZ_KIN_LIST_TMP.c_distance)>[ZZ_KIN_LIST_TMP_1].[c_distance]) AND " + _
        "((ZZ_KIN_LIST_TMP_1.c_kin_code)=[KINSHIP_CODES].[c_kin_pair1])) OR " + _
        "(((ZZ_KIN_LIST_TMP.c_distance)=[ZZ_KIN_LIST_TMP_1].[c_distance]) AND " + _
        "((ZZ_KIN_LIST_TMP_1.c_kin_code)=[KINSHIP_CODES].[c_kin_pair1]) AND " + _
        "((ZZ_KIN_LIST_TMP.c_personid)>[ZZ_KIN_LIST_TMP_1].[c_personid])) OR " + _
        "(((ZZ_KIN_LIST_TMP.c_distance)>[ZZ_KIN_LIST_TMP_1].[c_distance]) AND " + _
        "((ZZ_KIN_LIST_TMP_1.c_kin_code)=[KINSHIP_CODES].[c_kin_pair2])) OR " + _
        "(((ZZ_KIN_LIST_TMP.c_distance)=[ZZ_KIN_LIST_TMP_1].[c_distance]) AND " + _
        "((ZZ_KIN_LIST_TMP_1.c_kin_code)=[KINSHIP_CODES].[c_kin_pair2]) AND " + _
        "((ZZ_KIN_LIST_TMP.c_personid)>[ZZ_KIN_LIST_TMP_1].[c_personid]))"
    
    tPruneTmpInversesQueryStr2 = "UPDATE KINSHIP_CODES INNER JOIN (ZZ_KIN_LIST_TMP AS ZZ_KIN_LIST_TMP_1 " + _
        "INNER JOIN ZZ_KIN_LIST_TMP ON (ZZ_KIN_LIST_TMP_1.c_personid = ZZ_KIN_LIST_TMP.c_kin_id) AND " + _
        "(ZZ_KIN_LIST_TMP_1.c_kin_id = ZZ_KIN_LIST_TMP.c_personid)) ON " + _
        "KINSHIP_CODES.c_kincode = ZZ_KIN_LIST_TMP.c_kin_code SET ZZ_KIN_LIST_TMP.c_delete = 1 " + _
        "WHERE (((ZZ_KIN_LIST_TMP.c_distance)<[ZZ_KIN_LIST_TMP_1].[c_distance]) AND " + _
        "((ZZ_KIN_LIST_TMP_1.c_kin_code)=[KINSHIP_CODES].[c_kin_pair1])) OR " + _
        "(((ZZ_KIN_LIST_TMP.c_distance)=[ZZ_KIN_LIST_TMP_1].[c_distance]) AND " + _
        "((ZZ_KIN_LIST_TMP_1.c_kin_code)=[KINSHIP_CODES].[c_kin_pair1]) AND " + _
        "((ZZ_KIN_LIST_TMP.c_personid)<[ZZ_KIN_LIST_TMP_1].[c_personid])) OR " + _
        "(((ZZ_KIN_LIST_TMP.c_distance)<[ZZ_KIN_LIST_TMP_1].[c_distance]) AND " + _
        "((ZZ_KIN_LIST_TMP_1.c_kin_code)=[KINSHIP_CODES].[c_kin_pair2])) OR " + _
        "(((ZZ_KIN_LIST_TMP.c_distance)=[ZZ_KIN_LIST_TMP_1].[c_distance]) AND " + _
        "((ZZ_KIN_LIST_TMP_1.c_kin_code)=[KINSHIP_CODES].[c_kin_pair2]) AND " + _
        "((ZZ_KIN_LIST_TMP.c_personid)<[ZZ_KIN_LIST_TMP_1].[c_personid]))"
    
    tAppendQueryStr = "INSERT INTO ZZ_KIN_LIST ( c_personid, c_kin_id, c_kin_code, c_personid_root, c_kinrel, " + _
        "c_kinrel_total, c_kinrel_total_raw, c_kinrel_total_simplified, c_up, c_down, c_col, c_mar, c_up_total, c_down_total, c_col_total, " + _
        "c_mar_total, c_distance, c_female, c_sex, c_kin_female, c_kin_sex, c_prior_female, c_notes, c_source, c_source_text_chn, c_source_text ) " + _
        "SELECT DISTINCT ZZ_KIN_LIST_TMP.c_personid, ZZ_KIN_LIST_TMP.c_kin_id, ZZ_KIN_LIST_TMP.c_kin_code, " + _
            "ZZ_KIN_LIST_TMP.c_personid_root, ZZ_KIN_LIST_TMP.c_kinrel, ZZ_KIN_LIST_TMP.c_kinrel_total, " + _
            "ZZ_KIN_LIST_TMP.c_kinrel_total_raw, ZZ_KIN_LIST_TMP.c_kinrel_total_simplified, " + _
            "ZZ_KIN_LIST_TMP.c_up, ZZ_KIN_LIST_TMP.c_down, ZZ_KIN_LIST_TMP.c_col, ZZ_KIN_LIST_TMP.c_mar, " + _
            "ZZ_KIN_LIST_TMP.c_up_total, ZZ_KIN_LIST_TMP.c_down_total, ZZ_KIN_LIST_TMP.c_col_total, " + _
            "ZZ_KIN_LIST_TMP.c_mar_total, ZZ_KIN_LIST_TMP.c_distance, ZZ_KIN_LIST_TMP.c_female, ZZ_KIN_LIST_TMP.c_sex, " + _
            "ZZ_KIN_LIST_TMP.c_kin_female, ZZ_KIN_LIST_TMP.c_kin_sex, ZZ_KIN_LIST_TMP.c_prior_female, ZZ_KIN_LIST_TMP.c_notes, " + _
            "ZZ_KIN_LIST_TMP.c_source, ZZ_KIN_LIST_TMP.c_source_text_chn, ZZ_KIN_LIST_TMP.c_source_text " + _
        "FROM ZZ_KIN_LIST_TMP"

    tLoopCount = 1
    tExitDo = False
    
    Do While tLoopCount <= tLoopMax And tRecCount > 0
        If tLoopCount = 1 Then
            ' MsgBox "Running first query"
            cmdSQL.CommandText = tKinFirstQueryStr
        Else
            ' MsgBox "Running query"
            cmdSQL.CommandText = tKinQueryStr + Str(tLoopCount - 1) + ")"
        End If
        cmdSQL.Execute tRecCount
        
        If tRecCount > 0 Then
            '
            '  process the results for addition
            '
            '  update the distance
            '
            'MsgBox "Fixing node distance"
            cmdSQL.CommandText = tNodeDistQueryStr
            cmdSQL.Execute tRecDelete
            '
            '  then mark the duplicates and delete them
            '
            'MsgBox "Fixing dupes 1"
            cmdSQL.CommandText = tPruneTmpQuery
            cmdSQL.Execute tRecDelete
            cmdSQL.CommandText = "DELETE * FROM ZZ_KIN_LIST_TMP WHERE ZZ_KIN_LIST_TMP.c_delete = 1"
            cmdSQL.Execute tRecDelete
            '
            'MsgBox "Fixing dupes 2"
            cmdSQL.CommandText = tPruneTmpQuery2
            cmdSQL.Execute tRecDelete
            cmdSQL.CommandText = "DELETE * FROM ZZ_KIN_LIST_TMP WHERE ZZ_KIN_LIST_TMP.c_delete = 1"
            cmdSQL.Execute tRecDelete
            '
            'MsgBox "Fixing dupes 3"
            cmdSQL.CommandText = tPruneTmpQueryDupesStr
            cmdSQL.Execute tRecDelete
            cmdSQL.CommandText = "DELETE * FROM ZZ_KIN_LIST_TMP WHERE ZZ_KIN_LIST_TMP.c_delete = 1"
            cmdSQL.Execute tRecDelete
            '
            'MsgBox "Fixing dupes 4"
            cmdSQL.CommandText = tPruneTmpQueryDupesStr2
            cmdSQL.Execute tRecDelete
            cmdSQL.CommandText = "DELETE * FROM ZZ_KIN_LIST_TMP WHERE ZZ_KIN_LIST_TMP.c_delete = 1"
            cmdSQL.Execute tRecDelete
            '
            'MsgBox "Fixing inverses"
            cmdSQL.CommandText = tPruneTmpInversesQueryStr1
            cmdSQL.Execute tRecDelete
            cmdSQL.CommandText = "DELETE * FROM ZZ_KIN_LIST_TMP WHERE ZZ_KIN_LIST_TMP.c_delete = 1"
            cmdSQL.Execute tRecDelete
            '
            cmdSQL.CommandText = tPruneInversesQueryStr1
            cmdSQL.Execute tRecDelete
            cmdSQL.CommandText = "DELETE * FROM ZZ_KIN_LIST_TMP WHERE ZZ_KIN_LIST_TMP.c_delete = 1"
            cmdSQL.Execute tRecDelete
            '
            cmdSQL.CommandText = tPruneInversesQueryStr2
            cmdSQL.Execute tRecDelete
            cmdSQL.CommandText = "DELETE * FROM ZZ_KIN_LIST_TMP WHERE ZZ_KIN_LIST_TMP.c_delete = 1"
            cmdSQL.Execute tRecDelete
            '
            '  now simplify the kinship string, add to the total and reduce if possible
            '
            '
            '  Reduce kinship strings
            '
            '  first just get the string length
            
            cmdSQL.CommandText = "UPDATE ZZ_KIN_LIST_TMP SET ZZ_KIN_LIST_TMP.c_kinrel_len = Len([ZZ_KIN_LIST_TMP].[c_kinrel_total])"
            cmdSQL.Execute tRecDelete
            '
            '  then deal with len = 2
            '
            cmdSQL.CommandText = "UPDATE ZZ_KIN_LIST_TMP " + _
                "SET ZZ_KIN_LIST_TMP.c_kinrel_root_text = Left([ZZ_KIN_LIST_TMP].[c_kinrel_total],[ZZ_KIN_LIST_TMP].[c_kinrel_len]-2), " + _
                    "ZZ_KIN_LIST_TMP.c_kinrel_test_text = Right([ZZ_KIN_LIST_TMP].[c_kinrel_total],2), " + _
                    "ZZ_KIN_LIST_TMP.c_kinrel_root_text_simplified = Left([ZZ_KIN_LIST_TMP].[c_kinrel_total_simplified],[ZZ_KIN_LIST_TMP].[c_kinrel_len]-2)" + _
                "WHERE (((ZZ_KIN_LIST_TMP.c_kinrel_len)=2))"
            cmdSQL.Execute tRecDelete
            '
            '   replace where relevant
                
            cmdSQL.CommandText = "UPDATE KINREL_REDUCTION RIGHT JOIN ZZ_KIN_LIST_TMP " + _
                "ON KINREL_REDUCTION.c_kinrel_target = ZZ_KIN_LIST_TMP.c_kinrel_test_text " + _
                "SET ZZ_KIN_LIST_TMP.c_kinrel_total = [ZZ_KIN_LIST_TMP].[c_kinrel_root_text]+[KINREL_REDUCTION].[c_kinrel_replacement], " + _
                    "ZZ_KIN_LIST_TMP.c_kinrel_total_simplified = ZZ_KIN_LIST_TMP.c_kinrel_root_text_simplified + " + _
                                "'(' +  KINREL_REDUCTION.c_kinrel_target + '>' +[KINREL_REDUCTION].[c_kinrel_replacement] + ')', " + _
                    "ZZ_KIN_LIST_TMP.c_notes = [ZZ_KIN_LIST_TMP].[c_notes] + " + _
                            "'(' +  KINREL_REDUCTION.c_kinrel_target + '>' +[KINREL_REDUCTION].[c_kinrel_replacement] + ') ', " + _
                    "ZZ_KIN_LIST_TMP.c_up_total = [ZZ_KIN_LIST_TMP].[c_up_total]+[KINREL_REDUCTION].[c_up_change], " + _
                    "ZZ_KIN_LIST_TMP.c_down_total = [ZZ_KIN_LIST_TMP].[c_down_total]+[KINREL_REDUCTION].[c_down_change], " + _
                    "ZZ_KIN_LIST_TMP.c_col_total = [ZZ_KIN_LIST_TMP].[c_col_total]+[KINREL_REDUCTION].[c_col_change], " + _
                    "ZZ_KIN_LIST_TMP.c_mar_total = [ZZ_KIN_LIST_TMP].[c_mar_total]+[KINREL_REDUCTION].[c_mar_change] " + _
                "WHERE ((( KINREL_REDUCTION.c_kinrel_target) Is Not Null AND  KINREL_REDUCTION.c_required ))"
            cmdSQL.Execute tRecDelete
                
            '  then deal with len > 2
            '
            '   copy the target string and string root
            
            cmdSQL.CommandText = "UPDATE ZZ_KIN_LIST_TMP " + _
                "SET ZZ_KIN_LIST_TMP.c_kinrel_root_text = Left([ZZ_KIN_LIST_TMP].[c_kinrel_total],[ZZ_KIN_LIST_TMP].[c_kinrel_len]-2), " + _
                    "ZZ_KIN_LIST_TMP.c_kinrel_test_text = Right([ZZ_KIN_LIST_TMP].[c_kinrel_total],2), " + _
                    "ZZ_KIN_LIST_TMP.c_kinrel_root_text_simplified = Left([ZZ_KIN_LIST_TMP].[c_kinrel_total_simplified],[ZZ_KIN_LIST_TMP].[c_kinrel_len]-2)" + _
                "WHERE (((ZZ_KIN_LIST_TMP.c_kinrel_len)>2))"
            cmdSQL.Execute tRecDelete
            '
            '   replace where relevant
            
            cmdSQL.CommandText = "UPDATE KINREL_REDUCTION RIGHT JOIN ZZ_KIN_LIST_TMP " + _
                "ON  KINREL_REDUCTION.c_kinrel_target = ZZ_KIN_LIST_TMP.c_kinrel_test_text " + _
                "SET ZZ_KIN_LIST_TMP.c_kinrel_total = [ZZ_KIN_LIST_TMP].[c_kinrel_root_text]+[KINREL_REDUCTION].[c_kinrel_replacement], " + _
                    "ZZ_KIN_LIST_TMP.c_kinrel_total_simplified = ZZ_KIN_LIST_TMP.c_kinrel_root_text_simplified + " + _
                             "'(' + KINREL_REDUCTION.c_kinrel_target + '>' +[KINREL_REDUCTION].[c_kinrel_replacement] + ')', " + _
                    "ZZ_KIN_LIST_TMP.c_notes = [ZZ_KIN_LIST_TMP].[c_notes] + " + _
                            "'(' + KINREL_REDUCTION.c_kinrel_target + '>' +[KINREL_REDUCTION].[c_kinrel_replacement] + ') ', " + _
                    "ZZ_KIN_LIST_TMP.c_up_total = [ZZ_KIN_LIST_TMP].[c_up_total]+[KINREL_REDUCTION].[c_up_change], " + _
                    "ZZ_KIN_LIST_TMP.c_down_total = [ZZ_KIN_LIST_TMP].[c_down_total]+[KINREL_REDUCTION].[c_down_change], " + _
                    "ZZ_KIN_LIST_TMP.c_col_total = [ZZ_KIN_LIST_TMP].[c_col_total]+[KINREL_REDUCTION].[c_col_change], " + _
                    "ZZ_KIN_LIST_TMP.c_mar_total = [ZZ_KIN_LIST_TMP].[c_mar_total]+[KINREL_REDUCTION].[c_mar_change] " + _
                "WHERE ((( KINREL_REDUCTION.c_kinrel_target) Is Not Null AND KINREL_REDUCTION.c_required ))"
            cmdSQL.Execute tRecDelete
            '
            '  now mark the records with bad metrics
            '
            cmdSQL.CommandText = "UPDATE ZZ_KIN_LIST_TMP SET ZZ_KIN_LIST_TMP.c_delete = 1 " + _
                "WHERE (((ZZ_KIN_LIST_TMP.c_up_total)>" + Str(tMaxUp) + ")) OR " + _
                      "(((ZZ_KIN_LIST_TMP.c_down_total)>" + Str(tMaxDown) + ")) OR " + _
                      "(((ZZ_KIN_LIST_TMP.c_col_total)>" + Str(tMaxCol) + ")) OR " + _
                      "(((ZZ_KIN_LIST_TMP.c_mar_total)>" + Str(tMaxMarr) + "))"
            cmdSQL.Execute tRecDelete
            cmdSQL.CommandText = "DELETE * FROM ZZ_KIN_LIST_TMP WHERE ZZ_KIN_LIST_TMP.c_delete = 1"
            cmdSQL.Execute tRecDelete
            '
            '  one final test:  the only difference between ZZ_KIN_LIST_TEMP records is in PRIOR_FEMALE
            '
            cmdSQL.CommandText = "UPDATE ZZ_KIN_LIST_TMP AS ZZ_KIN_LIST_TMP_1 INNER JOIN ZZ_KIN_LIST_TMP " + _
                "ON (ZZ_KIN_LIST_TMP_1.c_kin_id = ZZ_KIN_LIST_TMP.c_kin_id) AND " + _
                "(ZZ_KIN_LIST_TMP_1.c_personid = ZZ_KIN_LIST_TMP.c_personid) AND " + _
                "(ZZ_KIN_LIST_TMP_1.c_kin_code = ZZ_KIN_LIST_TMP.c_kin_code) " + _
                "SET ZZ_KIN_LIST_TMP.c_delete = 1 " + _
                "WHERE (((ZZ_KIN_LIST_TMP.c_prior_female)=True) AND ((ZZ_KIN_LIST_TMP_1.c_prior_female)=False))"
            cmdSQL.Execute tRecDelete
            '
            cmdSQL.CommandText = "DELETE * FROM ZZ_KIN_LIST_TMP WHERE c_delete = 1"
            cmdSQL.Execute tRecDelete
            '
            '  it turns out that getting rid of inverse records is tougher than I would like, so we try one last time
            '
            cmdSQL.CommandText = tPruneTmpInversesQueryStr2
            cmdSQL.Execute tRecDelete
            '
            cmdSQL.CommandText = "DELETE * FROM ZZ_KIN_LIST_TMP WHERE c_delete = 1"
            cmdSQL.Execute tRecDelete
            '
            ' one last pair of clean-up routines is necessary.  One can arrive at the same results through different paths
            '
            ' first, take the shorter path
            '
            cmdSQL.CommandText = "UPDATE ZZ_KIN_LIST_TMP INNER JOIN ZZ_KIN_LIST_TMP AS ZZ_KIN_LIST_TMP_1 ON " + _
                "(ZZ_KIN_LIST_TMP.c_personid_root = ZZ_KIN_LIST_TMP_1.c_personid_root) " + _
                "AND (ZZ_KIN_LIST_TMP.c_kin_id = ZZ_KIN_LIST_TMP_1.c_kin_id) AND (ZZ_KIN_LIST_TMP.c_personid = ZZ_KIN_LIST_TMP_1.c_personid) " + _
                "AND (ZZ_KIN_LIST_TMP.c_kinrel = ZZ_KIN_LIST_TMP_1.c_kinrel) " + _
                "SET ZZ_KIN_LIST_TMP_1.c_delete = 1 " + _
                "WHERE (((Len([ZZ_KIN_LIST_TMP].[c_notes]))<Len([ZZ_KIN_LIST_TMP_1].[c_notes])))"
            cmdSQL.Execute tRecDelete
            '
            cmdSQL.CommandText = "DELETE * FROM ZZ_KIN_LIST_TMP WHERE c_delete = 1"
            cmdSQL.Execute tRecDelete
            '
            'MsgBox "Last step"
            '
            ' then take the string with the smaller value
            '
            cmdSQL.CommandText = "UPDATE ZZ_KIN_LIST_TMP INNER JOIN ZZ_KIN_LIST_TMP AS ZZ_KIN_LIST_TMP_1 ON " + _
                "(ZZ_KIN_LIST_TMP.c_personid_root = ZZ_KIN_LIST_TMP_1.c_personid_root) " + _
                "AND (ZZ_KIN_LIST_TMP.c_kin_id = ZZ_KIN_LIST_TMP_1.c_kin_id) AND (ZZ_KIN_LIST_TMP.c_personid = ZZ_KIN_LIST_TMP_1.c_personid) " + _
                "AND (ZZ_KIN_LIST_TMP.c_kinrel = ZZ_KIN_LIST_TMP_1.c_kinrel) " + _
                "SET ZZ_KIN_LIST_TMP_1.c_delete = 1 " + _
                "WHERE ('X'+[ZZ_KIN_LIST_TMP].[c_notes] > 'X'+[ZZ_KIN_LIST_TMP_1].[c_notes])"
            cmdSQL.Execute tRecDelete
            '
            cmdSQL.CommandText = "DELETE * FROM ZZ_KIN_LIST_TMP WHERE c_delete = 1"
            cmdSQL.Execute tRecDelete
            '
            ' now append the results:  if no records are added, this should stop the looping
            '
            ' MsgBox "Copying to ZZ_KIN_LIST"
            cmdSQL.CommandText = tAppendQueryStr
            cmdSQL.Execute tRecCount
            '
            '  and clear ZZ_KIN_LIST_TMP
            '
            cmdSQL.CommandText = "DELETE * FROM ZZ_KIN_LIST_TMP"
            cmdSQL.Execute tRecDelete
            '
        End If
        
        tLoopCount = tLoopCount + 1
        
        If tLoopCount > tLoopMax Then
            MsgBox "Loop limit hit."
            tExitDo = True
            Exit Do
        End If
    Loop
    '
    '  clean up the results
    '
    'MsgBox "Fixing ZZ_KIN_LIST inverses"
    cmdSQL.CommandText = tAppendQueryStr
    cmdSQL.Execute tRecDelete
    '
    '  copy to the ego-relative kinship table
    '
    '  Before copying we need to clean up the data
    '
    '  There is a bug in the algorithm that creates the occasional null value in c_kinrel_total.  To debug, for the moment plug the hole
    '
    'MsgBox "Patching NULL bug"
    cmdSQL.CommandText = "UPDATE ZZ_KIN_LIST SET ZZ_KIN_LIST.c_kinrel_total_simplified = '[Program Error]' WHERE (((ZZ_KIN_LIST.c_kinrel_total_simplified) Is Null))"
    cmdSQL.Execute tRecDelete
    '
    'MsgBox "Inserting ego-relative"
    tQueryStr = "INSERT INTO ZZ_KIN_LIST_TMP ( c_personid, c_kin_id, c_kinrel, c_kinrel_total, c_kinrel_total_simplified, c_up, c_down, c_col, c_mar, " + _
            "c_notes, c_kin_code, c_source, c_source_text, c_source_text_chn ) " + _
        "SELECT DISTINCT ZZ_KIN_LIST.c_personid_root, ZZ_KIN_LIST.c_kin_id, ZZ_KIN_LIST.c_kinrel_total_raw, ZZ_KIN_LIST.c_kinrel_total, " + _
            "ZZ_KIN_LIST.c_kinrel_total_simplified, ZZ_KIN_LIST.c_up_total, ZZ_KIN_LIST.c_down_total, ZZ_KIN_LIST.c_col_total, ZZ_KIN_LIST.c_mar_total, " + _
            "ZZ_KIN_LIST.c_notes, 0 AS c_kin_code, ZZ_KIN_LIST.c_source, ZZ_KIN_LIST.c_source_text, ZZ_KIN_LIST.c_source_text_chn " + _
        "FROM ZZ_KIN_LIST"

    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecDelete
    
    '  first just get the string length
            
    cmdSQL.CommandText = "UPDATE ZZ_KIN_LIST_TMP SET ZZ_KIN_LIST_TMP.c_kinrel_len = Len([ZZ_KIN_LIST_TMP].[c_kinrel_total_simplified])"
    cmdSQL.Execute tRecDelete
    
    '  delete the longer strings (this may solve most of the problems)
    
    cmdSQL.CommandText = "UPDATE ZZ_KIN_LIST_TMP INNER JOIN ZZ_KIN_LIST_TMP AS ZZ_KIN_LIST_TMP_1 ON ZZ_KIN_LIST_TMP.c_kin_id = ZZ_KIN_LIST_TMP_1.c_kin_id " + _
        "SET ZZ_KIN_LIST_TMP_1.c_delete = 1 " + _
        "WHERE (([ZZ_KIN_LIST_TMP_1].[c_kinrel_len]>[ZZ_KIN_LIST_TMP].[c_kinrel_len]))"
    cmdSQL.Execute tRecDelete
    
    cmdSQL.CommandText = "DELETE * FROM ZZ_KIN_LIST_TMP WHERE ZZ_KIN_LIST_TMP.c_delete = 1"
    cmdSQL.Execute tRecDelete
    '
    '  the next version uses the string-compare function because sometimes the strings are of the same length
    '
    cmdSQL.CommandText = "UPDATE ZZ_KIN_LIST_TMP AS ZZ_KIN_LIST_TMP_1 INNER JOIN " + _
        "ZZ_KIN_LIST_TMP ON (ZZ_KIN_LIST_TMP_1.c_kin_id = ZZ_KIN_LIST_TMP.c_kin_id) " + _
        "SET ZZ_KIN_LIST_TMP.c_delete = 1 " + _
        "WHERE (((StrComp([ZZ_KIN_LIST_TMP].[c_kinrel_total_simplified], [ZZ_KIN_LIST_TMP_1].[c_kinrel_total_simplified])) > 0))"
    cmdSQL.Execute tRecDelete
    
    cmdSQL.CommandText = "DELETE * FROM ZZ_KIN_LIST_TMP WHERE c_delete = 1"
    cmdSQL.Execute tRecDelete
    '
    '  the last version uses the total kinship path:  take the shortest value
    '
    cmdSQL.CommandText = "UPDATE ZZ_KIN_LIST_TMP AS ZZ_KIN_LIST_TMP_1 INNER JOIN " + _
        "ZZ_KIN_LIST_TMP ON (ZZ_KIN_LIST_TMP_1.c_kin_id = ZZ_KIN_LIST_TMP.c_kin_id) " + _
        "SET ZZ_KIN_LIST_TMP.c_delete = 1 " + _
        "WHERE (([ZZ_KIN_LIST_TMP].[c_down] + [ZZ_KIN_LIST_TMP].[c_col] + [ZZ_KIN_LIST_TMP].[c_mar] + [ZZ_KIN_LIST_TMP].[c_up]>" + _
                "[ZZ_KIN_LIST_TMP_1].[c_down]+[ZZ_KIN_LIST_TMP_1].[c_col]+[ZZ_KIN_LIST_TMP_1].[c_mar]+[ZZ_KIN_LIST_TMP_1].[c_up]))"
    cmdSQL.Execute tRecDelete
    
    cmdSQL.CommandText = "DELETE * FROM ZZ_KIN_LIST_TMP WHERE c_delete = 1"
    cmdSQL.Execute tRecDelete
    '
    ' one last clean-up:  remove results that are the same but takes a longer path to get there
    '
    cmdSQL.CommandText = "UPDATE ZZ_KIN_LIST_TMP INNER JOIN ZZ_KIN_LIST_TMP AS ZZ_KIN_LIST_TMP_1 " + _
        "ON (ZZ_KIN_LIST_TMP.c_kinrel = ZZ_KIN_LIST_TMP_1.c_kinrel) AND (ZZ_KIN_LIST_TMP.c_personid = ZZ_KIN_LIST_TMP_1.c_personid) AND " + _
           "(ZZ_KIN_LIST_TMP.c_kin_id = ZZ_KIN_LIST_TMP_1.c_kin_id) " + _
        "SET ZZ_KIN_LIST_TMP_1.c_delete = 1 " + _
        "WHERE ((Len([ZZ_KIN_LIST_TMP].[c_notes])<Len([ZZ_KIN_LIST_TMP_1].[c_notes])))"
    cmdSQL.Execute tRecDelete
    '
    cmdSQL.CommandText = "DELETE * FROM ZZ_KIN_LIST_TMP WHERE c_delete = 1"
    cmdSQL.Execute tRecDelete
    '
    'MsgBox "Last step"
    cmdSQL.CommandText = "UPDATE ZZ_KIN_LIST_TMP INNER JOIN ZZ_KIN_LIST_TMP AS ZZ_KIN_LIST_TMP_1 " + _
        "ON (ZZ_KIN_LIST_TMP.c_kinrel = ZZ_KIN_LIST_TMP_1.c_kinrel) AND (ZZ_KIN_LIST_TMP.c_personid = ZZ_KIN_LIST_TMP_1.c_personid) AND " + _
           "(ZZ_KIN_LIST_TMP.c_kin_id = ZZ_KIN_LIST_TMP_1.c_kin_id) " + _
        "SET ZZ_KIN_LIST_TMP_1.c_delete = 1 " + _
        "WHERE ('X'+[ZZ_KIN_LIST_TMP].[c_notes] > 'X'+[ZZ_KIN_LIST_TMP_1].[c_notes])"
    cmdSQL.Execute tRecDelete
    '
    cmdSQL.CommandText = "DELETE * FROM ZZ_KIN_LIST_TMP WHERE c_delete = 1"
    cmdSQL.Execute tRecDelete
    '
    '  finally, add in the source text information while copying results
    '
    tQueryStr = "INSERT INTO ZZ_SCRATCH_KIN (c_person_id, c_kin_id, c_kin_rel, c_kin_rel_total, c_kin_rel_0, c_up, c_down, c_collateral, c_marriage, c_notes, " + _
                "c_source, c_source_text, c_source_text_chn ) " + _
            "SELECT DISTINCT ZZ_KIN_LIST_TMP.c_personid, ZZ_KIN_LIST_TMP.c_kin_id, ZZ_KIN_LIST_TMP.c_kinrel, ZZ_KIN_LIST_TMP.c_kinrel_total, " + _
                "ZZ_KIN_LIST_TMP.c_kinrel_total_simplified, ZZ_KIN_LIST_TMP.c_up, ZZ_KIN_LIST_TMP.c_down, ZZ_KIN_LIST_TMP.c_col, " + _
                "ZZ_KIN_LIST_TMP.c_mar, ZZ_KIN_LIST_TMP.c_notes, ZZ_KIN_LIST_TMP.c_source, TEXT_CODES.c_title, TEXT_CODES.c_title_chn " + _
            "FROM ZZ_KIN_LIST_TMP LEFT JOIN TEXT_CODES ON ZZ_KIN_LIST_TMP.c_source = TEXT_CODES.c_textid"

     cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecDelete
    '
    ' add back in the kin not captured by the 2-2-1-1 parameters before updating the information
    '
    cmdSQL.CommandText = "INSERT INTO ZZ_SCRATCH_KIN ( c_person_id, c_kin_id, c_kin_code, c_kin_rel, c_kin_rel_0, c_kin_rel_total, c_up, c_down, c_marriage, " + _
            "c_collateral, c_source, c_pages, c_notes, c_source_text_chn, c_source_text ) " + _
        "SELECT KIN_DATA.c_personid, KIN_DATA.c_kin_id, KIN_DATA.c_kin_code, KINSHIP_CODES.c_kinrel, KINSHIP_CODES.c_kinrel AS c_kel_rel_0, KINSHIP_CODES.c_kinrel, " + _
            "KINSHIP_CODES.c_upstep, KINSHIP_CODES.c_dwnstep, KINSHIP_CODES.c_marstep, KINSHIP_CODES.c_colstep, KIN_DATA.c_source, KIN_DATA.c_pages, " + _
            "KIN_DATA.c_notes, TEXT_CODES.c_title_chn, TEXT_CODES.c_title " + _
        "FROM ( TEXT_CODES RIGHT JOIN KIN_DATA ON TEXT_CODES.c_textid = KIN_DATA.c_source ) INNER JOIN KINSHIP_CODES " + _
            "ON KIN_DATA.c_kin_code = KINSHIP_CODES.c_kincode " + _
        "WHERE ( ((KIN_DATA.c_personid) = " + Str(t_personid) + ") AND ((KINSHIP_CODES.c_upstep) > 2) ) " + _
            "OR ( ((KIN_DATA.c_personid) = " + Str(t_personid) + ") AND ((KINSHIP_CODES.c_dwnstep) > 2) ) " + _
            "OR ( ((KIN_DATA.c_personid) = " + Str(t_personid) + ") AND ((KINSHIP_CODES.c_marstep) > 1) ) " + _
            "OR ( ((KIN_DATA.c_personid) = " + Str(t_personid) + ") AND ((KINSHIP_CODES.c_colstep) > 1) )"

    cmdSQL.Execute tRecDelete
       
    tQueryStr = "UPDATE ZZ_SCRATCH_KIN INNER JOIN ( ( BIOG_MAIN LEFT JOIN ADDR_CODES ON BIOG_MAIN.c_index_addr_id = ADDR_CODES.c_addr_id ) LEFT JOIN BIOG_ADDR_CODES " + _
                    "ON BIOG_MAIN.c_index_addr_type_code = BIOG_ADDR_CODES.c_addr_type) ON ZZ_SCRATCH_KIN.c_kin_id = BIOG_MAIN.c_personid " + _
                "SET ZZ_SCRATCH_KIN.c_kin_name = [BIOG_MAIN].[c_name], ZZ_SCRATCH_KIN.c_kin_chn = [BIOG_MAIN].[c_name_chn], " + _
                    "ZZ_SCRATCH_KIN.c_kin_index_year = [BIOG_MAIN].[c_index_year], ZZ_SCRATCH_KIN.c_kin_female = [BIOG_MAIN].[c_female], " + _
                    "ZZ_SCRATCH_KIN.c_kin_sex = IIf([BIOG_MAIN].[c_female], 'F', 'M'), ZZ_SCRATCH_KIN.c_kin_addr_id = [BIOG_MAIN].[c_index_addr_id], " + _
                    "ZZ_SCRATCH_KIN.c_kin_addr_name = [ADDR_CODES].[c_name], ZZ_SCRATCH_KIN.c_kin_addr_chn = [ADDR_CODES].[c_name_chn], " + _
                    "ZZ_SCRATCH_KIN.c_kin_addr_type = [BIOG_MAIN].[c_index_addr_type_code], " + _
                    "ZZ_SCRATCH_KIN.c_kin_addr_desc = [BIOG_ADDR_CODES].[c_addr_desc], " + _
                    "ZZ_SCRATCH_KIN.c_kin_addr_desc_chn = [BIOG_ADDR_CODES].[c_addr_desc_chn], ZZ_SCRATCH_KIN.kin_x_coord = [ADDR_CODES].[x_coord], " + _
                    "ZZ_SCRATCH_KIN.kin_y_coord = [ADDR_CODES].[y_coord]"
    
    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecDelete
    '
    '  get the index year and dynasty descriptive data
    '
    cmdSQL.CommandText = "UPDATE ( DYNASTIES RIGHT JOIN ( ZZ_SCRATCH_KIN INNER JOIN BIOG_MAIN ON ZZ_SCRATCH_KIN.c_kin_id = BIOG_MAIN.c_personid ) " + _
                    "ON DYNASTIES.c_dy = BIOG_MAIN.c_dy ) INNER JOIN INDEXYEAR_TYPE_CODES " + _
                    "ON BIOG_MAIN.c_index_year_type_code = INDEXYEAR_TYPE_CODES.c_index_year_type_code " + _
                "SET ZZ_SCRATCH_KIN.c_kin_index_year_type_code = [BIOG_MAIN].[c_index_year_type_code], " + _
                    "ZZ_SCRATCH_KIN.c_kin_index_year_type_desc = [INDEXYEAR_TYPE_CODES].[c_index_year_type_desc], " + _
                    "ZZ_SCRATCH_KIN.c_kin_index_year_type_hz = [INDEXYEAR_TYPE_CODES].[c_index_year_type_hz],  " + _
                    "ZZ_SCRATCH_KIN.c_kin_dy = [BIOG_MAIN].[c_dy], " + _
                    "ZZ_SCRATCH_KIN.c_kin_dynasty = [DYNASTIES].[c_dynasty], ZZ_SCRATCH_KIN.c_kin_dynasty_chn = [DYNASTIES].[c_dynasty_chn]"
    cmdSQL.Execute tRecDelete
    
Exit_getKinship:
    '
    Exit Sub

Err_getKinship:
    MsgBox Err.Description + tErrorStr
    Resume Exit_getKinship

    Return
End Sub
