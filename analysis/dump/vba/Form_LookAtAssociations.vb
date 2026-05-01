Option Compare Database
Public gRstPeople As DAO.Recordset, gDisplayLanguage As String, gLabelsOK As Boolean
Public gImportPlaces As Boolean, gUseADDRID As Boolean, gFromStr As String, gToStr As String
Public gAssocCodeStr As String, gAssocTypeStr As String
Public gFromDynasty As Integer, gToDynasty As Integer, gUseIndexYears As Boolean, gUseDynasties As Boolean, _
        gFromDynastyBegin As Integer, gFromDynastyEnd As Integer, gToDynastyBegin As Integer, gToDynastyEnd As Integer

Private Sub ChkIndexYears_Click()
    If TxtFromYear.Enabled Then
        TxtFromYear.Enabled = False
        TxtToYear.Enabled = False
    Else
        TxtFromYear.Enabled = True
        TxtToYear.Enabled = True
    End If
End Sub

Private Sub CmdAllDynasties_Click()
    gFromDynasty = -2
    gToDynasty = -2
    TxtFromDynasty.Value = ""
    TxtFromDynastyPY.Value = "All"
    TxtToDynasty.Value = ""
    TxtToDynastyPY.Value = "All"

End Sub

Private Sub CmdFromDynasty_Click()
    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim strFromDynasty As String

    If gFromDynasty < 0 Then
        strFromDynasty = ""
    Else
        strFromDynasty = Str(gFromDynasty)
    End If
    
    stDocName = "frmPickDynasty"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , acDialog, strFromDynasty
    
    If CurrentProject.AllForms("frmPickDynasty").IsLoaded Then
        
        Forms!frmpickdynasty!frmDYNASTIES.Form!Dy_Code.SetFocus
        gFromDynasty = Forms!frmpickdynasty!frmDYNASTIES.Form!Dy_Code.Value
        
        Forms!frmpickdynasty!frmDYNASTIES.Form!c_start.SetFocus
        gFromDynastyBegin = Forms!frmpickdynasty!frmDYNASTIES.Form!c_start.Value
        
        Forms!frmpickdynasty!frmDYNASTIES.Form!c_end.SetFocus
        gFromDynastyEnd = Forms!frmpickdynasty!frmDYNASTIES.Form!c_end.Value
        '
        ' check to see if we have a problem and reject selection
        '
        If gToDynasty > -1 Then
            If gFromDynastyBegin > gToDynastyEnd Then
                MsgBox "Warning:  There is a problem with chronology:  the 'From' Dynasty begins after the 'To' Dynasty ends!", vbExclamation
                gFromDynasty = -1
                TxtFromDynasty.Value = ""
                TxtFromDynastyPY.Value = ""
            End If
        End If
        '
        '  value is OK
        '
        If gFromDynasty > -1 Then
            Forms!frmpickdynasty!frmDYNASTIES.Form!c_dynasty.SetFocus
            TxtFromDynastyPY.Value = Forms!frmpickdynasty!frmDYNASTIES.Form!c_dynasty.Value
            
            Forms!frmpickdynasty!frmDYNASTIES.Form!c_dynasty_chn.SetFocus
            TxtFromDynasty.Value = Forms!frmpickdynasty!frmDYNASTIES.Form!c_dynasty_chn.Value
        End If
        
        DoCmd.Close acForm, stDocName
        '
        ' reset ToDynasty if necessary (-2 = all dynasties)
        '
        If gToDynasty = -2 Then
            gToDynasty = -1
            TxtToDynasty.Value = ""
            TxtToDynastyPY.Value = ""
        End If
        '
    End If

End Sub

Private Sub CmdGephi_Click()
    On Error GoTo Err_CmdGephi_Click
    '
    '  This program will dump the results of the search to a .gdf file
    '
    '  for the moment I'll just describe the format of the .gdf file
    '
    '  nodedef> name, label, labelvisible, style, pinyin VARCHAR(50), nodedist INT
    '      name = str(c_person_id)
    '      label = c_name_chn
    '      style = 4 (text inside a rectangle)
    '      pinyin = c_name
    '      nodedist = c_node_dist INT
    '      indexyear = c_index_year INT
    '      sex = c_female > (F,M)
    '
    '  edgedef> node1, node2, label, labelvisible, edge_desc VARCHAR(50)
    '      node1 = str(c_person_id) for node1
    '      node2 = str(c_node_id) for node2
    '      label = c_link_chn
    '      edge_desc = c_link_desc
    '      edgetype= c_link_type (K,N)
    '
    '  the central question is whether to do distance optimizations
    '
    '  first see if there are any records to process
    '
    If ZZ_SCRATCH_ASSOC.Form.Recordset.RecordCount = 0 Then
        MsgBox "There are no records to save."
        GoTo Exit_CmdGephi_Click
    End If
    '
    If ZZ_SCRATCH_P_ASSOC.Form.Recordset.RecordCount = 0 Then
        MsgBox "There are no records to save."
        GoTo Exit_CmdGephi_Click
    End If
    '
    '  next get a file
    '
    Dim dlgSaveAs As FileDialog
    Dim tFileNum As Integer
    Dim tFileName As String, tFN As Variant
    Dim tRstNode As DAO.Recordset
    Dim tRstEdge As DAO.Recordset
    Dim tStr As String, tC As String, ti As Integer, tCodeStr As String
    'Dim tFileSystem, tGDF
    
    '  to write to a UTF-8 file, use the ADO stream object
    Dim tStream As ADODB.Stream
    Set tStream = New ADODB.Stream
    
    tPinyin = False
    If CodeFrame.Value = 1 Then
        tStream.Charset = "utf-8"
        tCodeStr = "UTF8"
    ElseIf CodeFrame.Value = 2 Then
        tStream.Charset = "big5"
        tCodeStr = "BIG5"
    ElseIf CodeFrame.Value = 3 Then
        tStream.Charset = "gb18030"
        tCodeStr = "GB18030"
    Else
        tStream.Charset = "ascii"
        tCodeStr = "ASCII"
        tPinyin = True
    End If
    
    
    Set dlgSaveAs = Application.FileDialog(msoFileDialogSaveAs)

    'Use a With...End With block to reference the FileDialog object.
    With dlgSaveAs
        .InitialFileName = "assoc_" + tCodeStr + ".gdf"
        If .Show = -1 Then
            '
            tFileName = ""
            For Each tFN In .SelectedItems
                tFileName = tFN
                If Not tFileName = "" Then
                    Exit For
                End If
            Next
            If tFileName = "" Then
                MsgBox "Bad file Name."
                GoTo Exit_CmdGephi_Click
            Else
                '  make sure the file name has a gdf extension
                If Len(tFileName) < 5 Then
                    tFileName = tFileName + ".gdf"
                ElseIf Not (LCase(Right(tFileName, 4)) = ".gdf") Then
                    tFileName = tFileName + ".gdf"
                End If
            End If
            '
            '  now process the file (second true removed to make ASCII)
            '
            'Set tFileSystem = CreateObject("Scripting.FileSystemObject")
            'Set tGDF = tFileSystem.CreateTextFile(tFileName, True, True)
            '
            tStream.Mode = adModeReadWrite
            tStream.Type = adTypeText
            tStream.Open

            ' process the two tables
            '
            Set tRstEdge = ZZ_SCRATCH_ASSOC.Form.Recordset
            Set tRstNode = ZZ_SCRATCH_P_ASSOC.Form.Recordset
            tC = Chr(44) ' the comma
            tQuote = Chr(34) 'the Quote delimiter
            '
            ' first the nodes:  define the record structure
            '   if ASCII, no pinyin field, no characters
            '
            If tCodeStr = "ASCII" Then
                tStr = "nodedef> name VARCHAR" + tC + "label VARCHAR" + tC + "labelvisible BOOLEAN" + _
                    tC + "style INT" + tC + "indexyear INT" + tC + "sex VARCHAR(1)" + _
                    tC + "addr_name VARCHAR" + tC + "latitude DOUBLE" + tC + "longitude DOUBLE"
            Else
                tStr = "nodedef> name VARCHAR" + tC + "label VARCHAR" + tC + "labelvisible BOOLEAN" + _
                    tC + "style INT" + tC + "pinyin VARCHAR(50)" + tC + "indexyear INT" + tC + "sex VARCHAR(1)" + _
                    tC + "addr_chn VARCHAR" + tC + "addr_name VARCHAR" + tC + "latitude DOUBLE" + tC + "longitude DOUBLE"
            End If
            tStream.WriteText tStr, adWriteLine
            'tGDF.WriteLine (tStr)
            '
            With tRstNode
                .MoveFirst
                Do While Not .EOF
                    '  name = the ID of the person
                    tStr = Trim(Str(!c_person_id)) + tC
                    
                    '  label
                    If tCodeStr = "ASCII" Then
                        If IsNull(!c_name) Then
                            tStr = tStr + tC
                        Else
                            tStr = tStr + !c_name + tC
                        End If
                    Else
                        If IsNull(!c_name_chn) Then
                            tStr = tStr + tC
                        Else
                            tStr = tStr + !c_name_chn + tC
                        End If
                    End If
                    '  labelvisible = true, style = 4 (text inside a rectangle)
                    tStr = tStr + "true" + tC + "4" + tC
                    
                    If Not (tCodeStr = "ASCII") Then
                        '  pinyin = c_name
                        tStr = tStr + !c_name + tC
                    End If
                    
                    '  indexyear = c_index_year INT
                    If IsNull(!c_index_year) Then
                        tStr = tStr + "-2000" + tC
                    Else
                        tStr = tStr + Trim(Str(!c_index_year)) + tC
                    End If
                    
                    '   sex = F,M
                    tStr = tStr + !c_sex + tC
                    
                    '  address name(s)
                    If tCodeStr = "ASCII" Then
                        If IsNull(!c_addr_name) Then
                            tStr = tStr + tC
                        Else
                            tStr = tStr + !c_addr_name + tC
                        End If
                    Else
                        If IsNull(!c_addr_chn) Then
                            tStr = tStr + tC
                        Else
                            tStr = tStr + !c_addr_chn + tC
                        End If
                        If IsNull(!c_addr_name) Then
                            tStr = tStr + tC
                        Else
                            tStr = tStr + !c_addr_name + tC
                        End If
                    End If
                    '
                    '   latitude = !y_coord
                    If IsNull(!y_coord) Then
                        tStr = tStr + "0.0" + tC
                    Else
                        tStr = tStr + Str(!y_coord) + tC
                    End If
                    
                    '   longitude = !x_coord
                    If IsNull(!x_coord) Then
                        tStr = tStr + "0.0"
                    Else
                        tStr = tStr + Str(!x_coord)
                    End If
                    tStream.WriteText tStr, adWriteLine
                    'tGDF.WriteLine (tStr)
                    .MoveNext
                Loop
            End With
            '
            ' now the edges:  define the record structure
            '   if ASCII, the label is the assoc_desc and there is not edge_desc
            '
            If tCodeStr = "ASCII" Then
                tStr = "edgedef> node1 VARCHAR" + tC + "node2 VARCHAR" + tC + "label VARCHAR(50)"
            Else
                tStr = "edgedef> node1 VARCHAR" + tC + "node2 VARCHAR" + tC + "label VARCHAR" + tC + "edge_desc VARCHAR(50)"
            End If
            tStream.WriteText tStr, adWriteLine
            'tGDF.WriteLine (tStr)

            With tRstEdge
                .MoveFirst
                Do While Not .EOF
                    '   node1 = str(c_person_id) for node1
                    tStr = Trim(Str(!c_person_id)) + tC
                    
                    '   node2 = str(c_assoc_id) for node2
                    tStr = tStr + Trim(Str(!c_assoc_id)) + tC
                    
                    '   label
                    If tCodeStr = "ASCII" Then
                        If IsNull(!c_assoc_desc) Then
                            tStr = tStr + tQuote + "[none]" + tQuote
                        Else
                            tStr = tStr + tQuote + Trim(Left(!c_assoc_desc + Space(50), 50)) + tQuote
                        End If
                    Else
                        If IsNull(!c_assoc_desc_chn) Then
                            tStr = tStr + tC
                        Else
                            tStr = tStr + tQuote + !c_assoc_desc_chn + tQuote + tC
                        End If
                    End If
                    
                    If Not (tCodeStr = "ASCII") Then
                        '   edge_desc = c_link_desc
                        If IsNull(!c_assoc_desc) Then
                            tStr = tStr + tQuote + "[none]" + tQuote
                        Else
                            tStr = tStr + tQuote + Trim(Left(!c_assoc_desc + Space(50), 50)) + tQuote
                        End If
                    End If
                    
                    tStream.WriteText tStr, adWriteLine
                    'tGDF.WriteLine (tStr)
                    .MoveNext
                Loop
            End With
            '
            ' now make sure all the data is copied to tStream
            tStream.Flush
            ' and write the stream to the file
            tStream.SaveToFile tFileName, adSaveCreateOverWrite
            '
            tStream.Close
            Set tStream = Nothing
            '
            'tGDF.Close
            '
            Set tRstNode = Nothing
            Set tRstEdge = Nothing
            'Set tGDF = Nothing
            'Set tFileSystem = Nothing
        Else
            'The user pressed Cancel.
        End If
    End With

    'Set the object variable to Nothing.
    Set dlgSaveAs = Nothing
    
Exit_CmdGephi_Click:
    Exit Sub

Err_CmdGephi_Click:
    MsgBox Err.Description
    Resume Exit_CmdGephi_Click
    
End Sub

Private Sub CmdImportAssociations_Click()
On Error GoTo Err_CmdImportAssociations_Click
    
    Dim stDocName As String, tRstAssociations As DAO.Recordset
    Dim stLinkCriteria As String, tRstImportAssociations As DAO.Recordset
    Dim tString As String, tOfficeID As Long, ti As Integer, tStrID As String, tQuit As Boolean
    Dim tLen As Integer, cmdSQL As ADODB.Command

    Dim dlgSaveAs As FileDialog
    Dim tFileNum As Integer
    Dim tFileName As String, tFN As Variant
    Dim tFileSystem, tList
    
    ' first see if we already have a list
    
    tQuit = False
    '
    If Not tQuit Then
        '  open the list
        
        Set dlgSaveAs = Application.FileDialog(msoFileDialogOpen)
    
        'Use a With...End With block to reference the FileDialog object.
        With dlgSaveAs
            .InitialFileName = ""
            If .Show = -1 Then
                '
                tFileName = ""
                For Each tFN In .SelectedItems
                    tFileName = tFN
                    If Not tFileName = "" Then
                        Exit For
                    End If
                Next
                If tFileName = "" Then
                    MsgBox "Bad file Name."
                    GoTo Exit_CmdImportAssociations_Click
                End If
            End If
        End With
        '
        ' Clear the address table now that we are ready to go
        '
        Set cmdSQL = New ADODB.Command
        cmdSQL.ActiveConnection = CurrentProject.Connection
        cmdSQL.CommandType = adCmdText
        '
        cmdSQL.CommandText = "Delete * from ZZ_ASSOC_CODE"
        cmdSQL.Execute tRecDeleted
        '
        cmdSQL.CommandText = "Delete * from InputErrorList"
        cmdSQL.Execute tRecDeleted
        
        cmdSQL.CommandText = "Delete * from TempImportList"
        cmdSQL.Execute tRecDeleted
        
        DoCmd.TransferText acImportDelim, "AssocCodeListImport Specification", "TempImportList", tFileName, 0
        '    TransferType=acImportDelim
        '    SpecificationName = "TempImportList" (apparently it is saved in the database itself)
        '    TableName = "TempImportList"  (probably requires that I drop the table first, but I can test)
        '    HasFieldNames = False (0)
        '
        '  copy the bad IDs
        '
        tStrSQL = "INSERT INTO InputErrorList ( c_ID ) SELECT TempImportList.ImportID " + _
            "FROM ASSOC_CODES RIGHT JOIN TempImportList ON ASSOC_CODES.c_assoc_code = TempImportList.ImportID " + _
            "WHERE (((ASSOC_CODES.c_assoc_code) Is Null))"

        cmdSQL.CommandText = tStrSQL
        cmdSQL.Execute tRecDeleted
        
        If tRecDeleted > 0 Then
            MsgBox "Some ID were not successfully imported:  please look at InputErrorList."
        End If
        '
        '  copy the good IDs
        '
        tStrSQL = "INSERT INTO ZZ_ASSOC_CODE ( c_assoc_code ) SELECT DISTINCT TempImportList.ImportID " + _
            "FROM ASSOC_CODES INNER JOIN TempImportList ON ASSOC_CODES.c_assoc_code = TempImportList.ImportID"

        cmdSQL.CommandText = tStrSQL
        cmdSQL.Execute tRecDeleted
        
        Me.TxtTypeDesc.Value = ""
        Me.TxtTypeChn.Value = ""
        If tRecDeleted > 0 Then
            Me.TxtAssocDesc.Value = "[Imported List]"
            Me.TxtAssocChn.Value = "[Imported List]"
            Me.CmdQuery.Enabled = True
            Me.CmdSaveAssociations.Enabled = True
        Else
            Me.TxtAssocDesc.Value = ""
            Me.TxtAssocChn.Value = ""
            Me.CmdQuery.Enabled = False
            Me.CmdSaveAssociations.Enabled = False
        End If
        
        Set cmdSQL = Nothing
    End If
    
Exit_CmdImportAssociations_Click:
    Exit Sub

Err_CmdImportAssociations_Click:
    MsgBox Err.Description
    Resume Exit_CmdImportAssociations_Click

End Sub

Private Sub CmdNeo4j_Click()
On Error GoTo Err_CmdNeo4j_Click
    '
    '  This program will dump the results of the search to four CSV files
    '
    '  for the moment I'll just describe the format of the CSV file
    '  Note:  Neo4j seems to treat all fields as strings, so there is no need to explicitly mark strings
    '
    '  People.CSV
    '  nameID, NameHZ, NamePY, indexyear, sex
    '      nameID = c_person_id
    '      nameHZ = c_name_chn
    '      namePY = c_name
    '      indexyear = c_index_year
    '      personDynasty = c_dynasty
    '      sex = c_female > (F,M)
    '
    '  Places.CSV
    '      placeID = c_addr_id
    '      placeHZ = c_addr_chn
    '      placePY = c_addr_name
    '      placeX  = x_coord
    '      placeY  = y_coord
    '
    '  PeoplePlaces.CSV
    '      nameID
    '      placeID
    '      personPlaceRelation
    '
    '  PeopleKinship.CSV
    '  node1_ID, node2_ID, kinshipRelation
    '      node1 = str(c_person_id) for node1
    '      node2 = str(c_node_id) for node2
    '      kinshipRelation = c_link_desc
    '
    '  first see if there are any records to process
    '
    If Me.ZZ_SCRATCH_P_ASSOC.Form.Recordset.RecordCount = 0 Then
        MsgBox "There are no records to save."
        GoTo Exit_CmdNeo4j_Click
    End If
    '
    '  allocate the file variables
    '
    Dim dlgSaveAs As FileDialog
    Dim tFileNum As Integer, tFileName As String, tFN As Variant
    '
    '  next get the People file
    '
    Dim tRstPeople As DAO.Recordset, tRstAssoc As DAO.Recordset, tRstPlace As DAO.Recordset, tRstPeoplePlace As DAO.Recordset
    Dim tStr As String, tC As String, tQueryStr As String, tRstAssocCode As DAO.Recordset
    Dim gStream As ADODB.Stream, tCodeStr As String
    '
    
    ' set up the stream to write to
    
    Set gStream = New ADODB.Stream
    If CodeFrame.Value = 1 Then
        gStream.Charset = "utf-8"
        tCodeStr = "UTF8"
    ElseIf CodeFrame.Value = 2 Then
        gStream.Charset = "big5"
        tCodeStr = "BIG5"
    ElseIf CodeFrame.Value = 3 Then
        gStream.Charset = "gb2312"
        tCodeStr = "GB2312"
    Else
        gStream.Charset = "ascii"
        tCodeStr = "ascii"
    End If

    Set dlgSaveAs = Application.FileDialog(msoFileDialogSaveAs)

        dlgSaveAs.InitialFileName = "People_" + tCodeStr + ".csv"
        If dlgSaveAs.Show = -1 Then
            '
            tFileName = ""
            For Each tFN In dlgSaveAs.SelectedItems
                tFileName = tFN
                If Not tFileName = "" Then
                    Exit For
                End If
            Next
            If tFileName = "" Then
                MsgBox "Bad file Name."
                GoTo Exit_CmdNeo4j_Click
            Else
                '  make sure the file name has a txt extension
                If Len(tFileName) < 5 Then
                    tFileName = tFileName + ".csv"
                ElseIf Not (LCase(Right(tFileName, 4)) = ".csv") Then
                    tFileName = tFileName + ".csv"
                End If
            End If
            '
            '  now process the file (second true removed to make ASCII)
            '
            'Set tFileSystem = CreateObject("Scripting.FileSystemObject")
            'Set tGDF = tFileSystem.CreateTextFile(tFileName, True, True)
            '
            '  we have a file name:  now open the stream for writing
            
            gStream.Mode = adModeReadWrite
            gStream.Type = adTypeText
            gStream.Open

            '
            '  prepare the temp tables for the people, place, peoplePlace and assoc codes
            
            Dim cmdSQL As ADODB.Command
            Set cmdSQL = New ADODB.Command
            cmdSQL.ActiveConnection = CurrentProject.Connection
            cmdSQL.CommandType = adCmdText
            '
            '  Get the people from 5 sources: c_person_id, c_assoc_id, c_kin_id, c_assoc_kin_id, and c_assoc_claimer_id
            
            cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_P_TEXT"
            cmdSQL.Execute tRecDeleted
            '
            '  copy the data for people (1)
            '
            tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id ) SELECT DISTINCT ZZ_SCRATCH_ASSOC.c_person_id " + _
                        "FROM ZZ_SCRATCH_ASSOC WHERE (((ZZ_SCRATCH_ASSOC.c_person_id)>0))"
            '
            cmdSQL.CommandText = tQueryStr
            cmdSQL.Execute tRecDeleted
            '
            '  copy the data for people (2)
            '
            tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id ) SELECT DISTINCT ZZ_SCRATCH_ASSOC.c_assoc_id " + _
                        "FROM ZZ_SCRATCH_ASSOC WHERE (((ZZ_SCRATCH_ASSOC.c_assoc_id)>0))"
            '
            cmdSQL.CommandText = tQueryStr
            cmdSQL.Execute tRecDeleted
            '
            '  copy the data for people (3)
            '
            tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id ) SELECT DISTINCT ZZ_SCRATCH_ASSOC.c_kin_id " + _
                        "FROM ZZ_SCRATCH_ASSOC WHERE (((ZZ_SCRATCH_ASSOC.c_kin_id)>0))"
            '
            cmdSQL.CommandText = tQueryStr
            cmdSQL.Execute tRecDeleted
            '
            '  copy the data for people (4)
            '
            tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id ) SELECT DISTINCT ZZ_SCRATCH_ASSOC.c_assoc_kin_id " + _
                        "FROM ZZ_SCRATCH_ASSOC WHERE (((ZZ_SCRATCH_ASSOC.c_assoc_kin_id)>0))"
            '
            cmdSQL.CommandText = tQueryStr
            cmdSQL.Execute tRecDeleted
            '
            '  copy the data for people (5)
            '
            tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id ) SELECT DISTINCT ZZ_SCRATCH_ASSOC.c_assoc_claimer_id " + _
                        "FROM ZZ_SCRATCH_ASSOC WHERE (((ZZ_SCRATCH_ASSOC.c_assoc_claimer_id)>0))"
            '
            cmdSQL.CommandText = tQueryStr
            cmdSQL.Execute tRecDeleted
            '
            ' now combine them into ZZ_SCRATCH_PEOPLE and get additional information
            '
            cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_PEOPLE"
            cmdSQL.Execute tRecDeleted
            '
            tQueryStr = "INSERT INTO ZZ_SCRATCH_PEOPLE ( c_person_id, c_name, c_name_chn, c_index_year, c_index_year_type_code, c_dy, c_addr_id, c_index_addr_type_code, c_female ) " + _
                        "SELECT DISTINCT ZZ_SCRATCH_P_TEXT.c_person_id, BIOG_MAIN.c_name, BIOG_MAIN.c_name_chn, BIOG_MAIN.c_index_year, BIOG_MAIN.c_index_year_type_code, " + _
                            "BIOG_MAIN.c_dy, BIOG_MAIN.c_index_addr_id, BIOG_MAIN.c_index_addr_type_code, BIOG_MAIN.c_female " + _
                        "FROM ZZ_SCRATCH_P_TEXT INNER JOIN BIOG_MAIN ON ZZ_SCRATCH_P_TEXT.c_person_id = BIOG_MAIN.c_personid"
            '
            cmdSQL.CommandText = tQueryStr
            cmdSQL.Execute tRecDeleted
       '
       '  now get the rest of the information
        '
        tQueryStr = "UPDATE ( ( ( ZZ_SCRATCH_PEOPLE LEFT JOIN INDEXYEAR_TYPE_CODES " + _
                        "ON ZZ_SCRATCH_PEOPLE.c_index_year_type_code = INDEXYEAR_TYPE_CODES.c_index_year_type_code )  LEFT JOIN DYNASTIES " + _
                        "ON ZZ_SCRATCH_PEOPLE.c_dy = DYNASTIES.c_dy ) LEFT JOIN ADDR_CODES ON ZZ_SCRATCH_PEOPLE.c_addr_id = ADDR_CODES.c_addr_id ) " + _
                        "LEFT JOIN BIOG_ADDR_CODES ON ZZ_SCRATCH_PEOPLE.c_addr_type = BIOG_ADDR_CODES.c_addr_type " + _
                    "SET ZZ_SCRATCH_PEOPLE.c_index_year_type_desc = [INDEXYEAR_TYPE_CODES].[c_index_year_type_desc], " + _
                        "ZZ_SCRATCH_PEOPLE.c_index_year_type_hz = [INDEXYEAR_TYPE_CODES].[c_index_year_type_hz], " + _
                            "ZZ_SCRATCH_PEOPLE.c_dynasty = [DYNASTIES].[c_dynasty],  ZZ_SCRATCH_PEOPLE.c_dynasty_chn = [DYNASTIES].[c_dynasty_chn], " + _
                            "ZZ_SCRATCH_PEOPLE.c_addr_name = [ADDR_CODES].[c_name], ZZ_SCRATCH_PEOPLE.c_addr_chn = [ADDR_CODES].[c_name_chn], " + _
                            "ZZ_SCRATCH_PEOPLE.c_addr_desc = [BIOG_ADDR_CODES].[c_addr_desc], ZZ_SCRATCH_PEOPLE.c_addr_desc_chn = [BIOG_ADDR_CODES].[c_addr_desc_chn], " + _
                            "ZZ_SCRATCH_PEOPLE.x_coord = [ADDR_CODES].[x_coord], ZZ_SCRATCH_PEOPLE.y_coord = [ADDR_CODES].[y_coord] "
            '
            cmdSQL.CommandText = tQueryStr
            cmdSQL.Execute tRecDeleted
            '
            Set tRstPeople = CurrentDb.OpenRecordset("ZZ_SCRATCH_PEOPLE", dbOpenDynaset)
            '
            ' process the four tables
            '
            tC = Chr(44) ' the comma
            '
            ' first the nodes:  define the record structure
            '
            '  if the file is strictly ASCII, the label is the pinyin, but if there are characters, then we add a pinyin field
            If tCodeStr = "ascii" Then
                tStr = "nameID" + tC + "namePY" + tC + "indexyear" + tC + "dynasty" + tC + "sex"
            Else
                tStr = "nameID" + tC + "nameHZ" + tC + "namePY" + tC + "indexyear" + tC + "dynasty" + tC + "sex"
            End If
            gStream.WriteText tStr, adWriteLine
            'tGDF.WriteLine (tStr)
            '
            With tRstPeople
                .MoveFirst
                Do While Not .EOF
                    '  the ID of the person
                    tStr = Trim(Str(!c_person_id)) + tC
                    '
                    '  name
                    '
                    If tCodeStr = "ascii" Then
                        If IsNull(!c_name) Then
                            tStr = tStr + tC
                        Else
                            tStr = tStr + !c_name + tC
                        End If
                    Else
                        If IsNull(!c_name_chn) Then
                            tStr = tStr + "Missing" + tC
                        Else
                            tStr = tStr + !c_name_chn + tC
                        End If
                        
                        If IsNull(!c_name) Then
                            tStr = tStr + "Missing" + tC
                        Else
                            tStr = tStr + !c_name + tC
                        End If
                    End If
                    '
                    '  indexyear = c_index_year INT
                    '
                    If IsNull(!c_index_year) Then
                        tStr = tStr + "-2000" + tC
                    Else
                        tStr = tStr + Trim(Str(!c_index_year)) + tC
                    End If
                    
                    '  dynasty information
                    '
                    If IsNull(!c_dynasty) Then
                        tStr = tStr + "unknown" + tC
                    Else
                        If tCodeStr = "ascii" Then
                            tStr = tStr + !c_dynasty + tC
                        Else
                            tStr = tStr + !c_dynasty_chn + tC
                        End If
                    End If
                    '
                    '   sex = c_female > (F,M)
                    tStr = tStr + IIf(!c_female, "F", "M")
                    '
                    gStream.WriteText tStr, adWriteLine
                    '
                    .MoveNext
                Loop
            End With
            ' now make sure all the data is copied to tStream
            gStream.Flush
            ' and write the stream to the file
            gStream.SaveToFile tFileName, adSaveCreateOverWrite
            '
            gStream.Close
        Else
            'The user pressed Cancel.
            GoTo Exit_CmdNeo4j_Click
        End If
        '
        '  now places:  since the association "event" is not linked to a place, the only addresses are the index addresses
        '               of the people involved, recorded in ZZ_SCRATCH_PEOPLE
        '
        '  get a file name
        '
        dlgSaveAs.InitialFileName = "Places_" + tCodeStr + ".csv"
        If dlgSaveAs.Show = -1 Then
            '
            tFileName = ""
            For Each tFN In dlgSaveAs.SelectedItems
                tFileName = tFN
                If Not tFileName = "" Then
                    Exit For
                End If
            Next
            If tFileName = "" Then
                MsgBox "Bad file Name."
                GoTo Exit_CmdNeo4j_Click
            Else
                '  make sure the file name has a txt extension
                If Len(tFileName) < 5 Then
                    tFileName = tFileName + ".csv"
                ElseIf Not (LCase(Right(tFileName, 4)) = ".csv") Then
                    tFileName = tFileName + ".csv"
                End If
            End If
            '
            gStream.Open
            '
            '  now process the file
            '
            tQueryStr = "SELECT DISTINCT ZZ_SCRATCH_PEOPLE.c_addr_id, ZZ_SCRATCH_PEOPLE.c_addr_name, ZZ_SCRATCH_PEOPLE.c_addr_chn, " + _
                            "ZZ_SCRATCH_PEOPLE.x_coord, ZZ_SCRATCH_PEOPLE.y_coord " + _
                        "FROM ZZ_SCRATCH_PEOPLE"

            Set tRstPlace = CurrentDb.OpenRecordset(tQueryStr)
            '
            If tCodeStr = "ascii" Then
                tStr = "placeID" + tC + "placePY" + tC + "placeX" + tC + "placeY"
            Else
                tStr = "placeID" + tC + "placePY" + tC + "placeHZ" + tC + "placeX" + tC + "placeY"
            End If
            gStream.WriteText tStr, adWriteLine
            With tRstPlace
                .MoveFirst
                Do While Not .EOF
                    '  the ID of the place
                    If Not IsNull(!c_addr_id) Then
                        tStr = Trim(Str(!c_addr_id)) + tC
                        '
                        '   address name
                        
                        If IsNull(!c_addr_name) Then
                            tStr = tStr + "unknown" + tC
                        Else
                            tStr = tStr + !c_addr_name + tC
                        End If
                        '
                        If Not (tCodeStr = "ascii") Then
                            If IsNull(!c_addr_chn) Then
                                tStr = tStr + "unknown" + tC
                            Else
                                tStr = tStr + !c_addr_chn + tC
                            End If
                        End If
                        
                        '   latitude = !y_coord
                        If IsNull(!y_coord) Then
                            tStr = tStr + "0.0" + tC
                        Else
                            tStr = tStr + Str(!y_coord) + tC
                        End If
                        
                        '   longitude = !x_coord
                        If IsNull(!x_coord) Then
                            tStr = tStr + "0.0" + tC
                        Else
                            tStr = tStr + Str(!x_coord) + tC
                        End If
                        '
                        gStream.WriteText tStr, adWriteLine
                    End If
                    .MoveNext
                Loop
            End With
            '
            ' now make sure all the data is copied to tStream
            gStream.Flush
            ' and write the stream to the file
            gStream.SaveToFile tFileName, adSaveCreateOverWrite
            '
            gStream.Close
        Else
            'The user pressed Cancel.
            GoTo Exit_CmdNeo4j_Click
        End If
        '
        '  now peoplePlaces
        '
        dlgSaveAs.InitialFileName = "PeoplePlaces_" + tCodeStr + ".csv"
        If dlgSaveAs.Show = -1 Then
            '
            tFileName = ""
            For Each tFN In dlgSaveAs.SelectedItems
                tFileName = tFN
                If Not tFileName = "" Then
                    Exit For
                End If
            Next
            If tFileName = "" Then
                MsgBox "Bad file Name."
                GoTo Exit_CmdNeo4j_Click
            Else
                '  make sure the file name has a txt extension
                If Len(tFileName) < 5 Then
                    tFileName = tFileName + ".csv"
                ElseIf Not (LCase(Right(tFileName, 4)) = ".csv") Then
                    tFileName = tFileName + ".csv"
                End If
            End If
            '
            gStream.Open
            '
            tQueryStr = "SELECT DISTINCT ZZ_SCRATCH_PEOPLE.c_person_id, ZZ_SCRATCH_PEOPLE.c_addr_id, ZZ_SCRATCH_PEOPLE.c_addr_type, " + _
                            "ZZ_SCRATCH_PEOPLE.c_addr_desc, ZZ_SCRATCH_PEOPLE.c_addr_desc_chn " + _
                        "FROM ZZ_SCRATCH_PEOPLE"

            Set tRstPeoplePlace = CurrentDb.OpenRecordset(tQueryStr)
            
            tStr = "nameID" + tC + "placeID" + tC + "personPlaceTrans" + tC + "personPlaceHZ"
            
            gStream.WriteText tStr, adWriteLine
            
            With tRstPeoplePlace
                .MoveFirst
                Do While Not .EOF
                    If Not IsNull(!c_addr_id) Then
                        '
                        tStr = Trim(Str(!c_person_id)) + tC
                        '
                        tStr = tStr + Trim(Str(!c_addr_id)) + tC
                        '
                        tStr = tStr + Trim(Str(!c_addr_type))
                        gStream.WriteText tStr, adWriteLine
                    End If
                    .MoveNext
                Loop
            End With
            '
            ' now make sure all the data is copied to tStream
            gStream.Flush
            ' and write the stream to the file
            gStream.SaveToFile tFileName, adSaveCreateOverWrite
            '
            gStream.Close
        Else
            'The user pressed Cancel.
            GoTo Exit_CmdNeo4j_Click
        End If
        '
        '  now the association records
        '
        dlgSaveAs.InitialFileName = "PeopleAssociations_" + tCodeStr + ".csv"
        If dlgSaveAs.Show = -1 Then
            '
            tFileName = ""
            For Each tFN In dlgSaveAs.SelectedItems
                tFileName = tFN
                If Not tFileName = "" Then
                    Exit For
                End If
            Next
            If tFileName = "" Then
                MsgBox "Bad file Name."
                GoTo Exit_CmdNeo4j_Click
            Else
                '  make sure the file name has a txt extension
                If Len(tFileName) < 5 Then
                    tFileName = tFileName + ".csv"
                ElseIf Not (LCase(Right(tFileName, 4)) = ".csv") Then
                    tFileName = tFileName + ".csv"
                End If
            End If
            '
            gStream.Open
            '
            ' now the associations:  define the record structure
            ' Because of the complexity of the primary key, this gets a bit complicated
            '
            Set tRstAssoc = CurrentDb.OpenRecordset("ZZ_SCRATCH_ASSOC", dbOpenDynaset)
            '
            tStr = "Person1_ID" + tC + "Person2_ID" + tC + "Association_Code" + tC + "Kin_ID" + tC + "Kin_Code" + tC + _
                    "AssocKin_ID" + tC + "AssocKin_Code" + tC + "LiteraryGenreCode" + tC + "OccasionCode" + tC + _
                    "TopicCode" + tC + "InstitutionCode" + tC + "TextTitle" + tC + "AssociationClaimer_ID"
            gStream.WriteText tStr, adWriteLine
            'tGDF.WriteLine (tStr)

            With tRstAssoc
                .MoveFirst
                Do While Not .EOF
                    If Not IsNull(!c_assoc_code) Then
                        tStr = Trim(Str(!c_person_id)) + tC
                        '   node1 = str(c_person_id) for node1
                        tStr = tStr + Trim(Str(!c_assoc_id)) + tC
                        '   node2 = str(c_node_id) for node2
                        tStr = tStr + Trim(Str(!c_assoc_code)) + tC
                        
                        '   kin ID
                        If IsNull(!c_kin_id) Then
                            tStr = tStr + "0" + tC
                        Else
                            tStr = tStr + Str(!c_kin_id) + tC
                        End If
                        
                        '   kin code
                        If IsNull(!c_kin_code) Then
                            tStr = tStr + "0" + tC
                        Else
                            tStr = tStr + Str(!c_kin_code) + tC
                        End If
                        
                        '   assoc kin ID
                        If IsNull(!c_assoc_kin_id) Then
                            tStr = tStr + "0" + tC
                        Else
                            tStr = tStr + Str(!c_assoc_kin_id) + tC
                        End If
                        
                        '   assoc kin code
                        If IsNull(!c_assoc_kin_code) Then
                            tStr = tStr + "0" + tC
                        Else
                            tStr = tStr + Str(!c_assoc_kin_code) + tC
                        End If
                        
                        '   literary genre code
                        If IsNull(!c_litgenre_code) Then
                            tStr = tStr + "0" + tC
                        Else
                            tStr = tStr + Str(!c_litgenre_code) + tC
                        End If
                        
                        '   occasion code
                        If IsNull(!c_occasion_code) Then
                            tStr = tStr + "0" + tC
                        Else
                            tStr = tStr + Str(!c_occasion_code) + tC
                        End If
                        
                        '   topic code
                        If IsNull(!c_topic_code) Then
                            tStr = tStr + "0" + tC
                        Else
                            tStr = tStr + Str(!c_topic_code) + tC
                        End If
                        
                        '   institution code
                        If IsNull(!c_inst_code) Then
                            tStr = tStr + "0" + tC
                        Else
                            tStr = tStr + Str(!c_inst_code) + tC
                        End If
                        
                        '   text title
                        If IsNull(!c_text_title) Then
                            tStr = tStr + "N/A" + tC
                        Else
                            tStr = tStr + !c_text_title + tC
                        End If
                        
                        '   association claimer ID
                        If IsNull(!c_assoc_claimer_id) Then
                            tStr = tStr + "0" + tC
                        Else
                            tStr = tStr + Str(!c_assoc_claimer_id)
                        End If
                        
                        gStream.WriteText tStr, adWriteLine
                    End If
                    .MoveNext
                Loop
            End With
            '
            ' now make sure all the data is copied to tStream
            gStream.Flush
            ' and write the stream to the file
            gStream.SaveToFile tFileName, adSaveCreateOverWrite
            '
            gStream.Close
        Else
            'The user pressed Cancel.
        End If
        '
        '  now the association codes
        '
        dlgSaveAs.InitialFileName = "AssociationCodes_" + tCodeStr + ".csv"
        If dlgSaveAs.Show = -1 Then
            '
            tFileName = ""
            For Each tFN In dlgSaveAs.SelectedItems
                tFileName = tFN
                If Not tFileName = "" Then
                    Exit For
                End If
            Next
            If tFileName = "" Then
                MsgBox "Bad file Name."
                GoTo Exit_CmdNeo4j_Click
            Else
                '  make sure the file name has a txt extension
                If Len(tFileName) < 5 Then
                    tFileName = tFileName + ".csv"
                ElseIf Not (LCase(Right(tFileName, 4)) = ".csv") Then
                    tFileName = tFileName + ".csv"
                End If
            End If
            '
            gStream.Open
            '
            tQueryStr = "SELECT DISTINCT ZZ_SCRATCH_ASSOC.c_assoc_code, ZZ_SCRATCH_ASSOC.c_assoc_type, ZZ_SCRATCH_ASSOC.c_assoc_desc, " + _
                            "ZZ_SCRATCH_ASSOC.c_assoc_desc_chn " + _
                        "FROM ZZ_SCRATCH_ASSOC"

            Set tRstAssocCode = CurrentDb.OpenRecordset(tQueryStr, dbOpenDynaset)
            '
            If tCodeStr = "ascii" Then
                tStr = "AssociationCode" + tC + "AssociationTypeID" + tC + "AssociationTrans"
            Else
                tStr = "AssociationCode" + tC + "AssociationTypeID" + tC + "AssociationTrans" + tC + "AssociationHZ"
            End If
            
            gStream.WriteText tStr, adWriteLine
            
            With tRstAssocCode
                .MoveFirst
                Do While Not .EOF
                    If Not IsNull(!c_assoc_code) Then
                        '
                        tStr = Trim(Str(!c_assoc_code)) + tC
                        '
                        tStr = tStr + Trim(!c_assoc_type) + tC
                        '
                        tStr = tStr + Trim(!c_assoc_desc)
                        '
                        If Not (tCodeStr = "ascii") Then
                            tStr = tStr + tC + Trim(!c_assoc_desc_chn)
                        End If
                        gStream.WriteText tStr, adWriteLine
                    End If
                    .MoveNext
                Loop
            End With
            '
            ' now make sure all the data is copied to tStream
            gStream.Flush
            ' and write the stream to the file
            gStream.SaveToFile tFileName, adSaveCreateOverWrite
            '
            gStream.Close
        Else
            'The user pressed Cancel.
            GoTo Exit_CmdNeo4j_Click
        End If
        '
        '  there are codes that MAY require additional tables: c_kin_code, c_litgenrte_code, c_occasion_code, c_topic_code, c_inst_code
        '
        '  test for kin codes
        '
        tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id ) " + _
                    "SELECT DISTINCT ZZ_SCRATCH_ASSOC.c_person_id " + _
                    "FROM ZZ_SCRATCH_ASSOC " + _
                    "WHERE (((ZZ_SCRATCH_ASSOC.c_kin_code)>0))"
        '
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecDeleted
        '
        ' debug
        '
        MsgBox "Kinship code records = " + Trim(Str(tRecDeleted))
        '
        If tRecDeleted > 0 Then
            dlgSaveAs.InitialFileName = "KinshipCodes_" + tCodeStr + ".csv"
            If dlgSaveAs.Show = -1 Then
                '
                tFileName = ""
                For Each tFN In dlgSaveAs.SelectedItems
                    tFileName = tFN
                    If Not tFileName = "" Then
                        Exit For
                    End If
                Next
                If tFileName = "" Then
                    MsgBox "Bad file Name."
                    GoTo Exit_CmdNeo4j_Click
                Else
                    '  make sure the file name has a txt extension
                    If Len(tFileName) < 5 Then
                        tFileName = tFileName + ".csv"
                    ElseIf Not (LCase(Right(tFileName, 4)) = ".csv") Then
                        tFileName = tFileName + ".csv"
                    End If
                End If
                '
                gStream.Open
                '
                ' there is an additional complication for kinship codes because there are two sources: c_kin_code and c_assoc_kin_code
                '
                cmdSQL.CommandText = "DELETE * FROM ZZ_KIN_LIST_TMP"
                cmdSQL.Execute tRecDeleted
                
                tQueryStr = "INSERT INTO ZZ_KIN_LIST_TMP ( c_kin_code, c_kinrel, c_kinrel_total ) " + _
                            "SELECT DISTINCT ZZ_SCRATCH_ASSOC.c_kin_code, ZZ_SCRATCH_ASSOC.c_kin_desc, ZZ_SCRATCH_ASSOC.c_kin_desc_chn " + _
                            "FROM ZZ_SCRATCH_ASSOC " + _
                            "WHERE (ZZ_SCRATCH_ASSOC.c_kin_code > 0)"
                cmdSQL.CommandText = tQueryStr
                cmdSQL.Execute tRecDeleted

                tQueryStr = "INSERT INTO ZZ_KIN_LIST_TMP ( c_kin_code, c_kinrel, c_kinrel_total ) " + _
                            "SELECT DISTINCT ZZ_SCRATCH_ASSOC.c_assoc_kin_code, ZZ_SCRATCH_ASSOC.c_assoc_kin_desc, ZZ_SCRATCH_ASSOC.c_assoc_kin_desc_chn " + _
                            "FROM ZZ_SCRATCH_ASSOC " + _
                            "WHERE (ZZ_SCRATCH_ASSOC.c_assoc_kin_code > 0)"
                cmdSQL.CommandText = tQueryStr
                cmdSQL.Execute tRecDeleted

                '
                tQueryStr = "SELECT DISTINCT ZZ_KIN_LIST_TMP.c_kin_code, ZZ_KIN_LIST_TMP.c_kinrel, ZZ_KIN_LIST_TMP.c_kinrel_total " + _
                            "FROM ZZ_KIN_LIST_TMP"
    
                Set tRstAssocCode = CurrentDb.OpenRecordset(tQueryStr)
                
                If tCodeStr = "ascii" Then
                    tStr = "KinshipCode" + tC + "KinshipTrans"
                Else
                    tStr = "KinshipCode" + tC + "KinshipTrans" + tC + "KinshipHZ"
                End If
                
                gStream.WriteText tStr, adWriteLine
                
                With tRstAssocCode
                    .MoveFirst
                    Do While Not .EOF
                        If Not IsNull(!c_kin_code) Then
                            '
                            tStr = Trim(Str(!c_kin_code)) + tC
                            '
                            tStr = tStr + Trim(!c_kinrel)
                            '
                            If Not (tCodeStr = "ascii") Then
                                tStr = tStr + tC + Trim(!c_kinrel_total)
                            End If
                            gStream.WriteText tStr, adWriteLine
                        End If
                        .MoveNext
                    Loop
                End With
                '
                ' now make sure all the data is copied to tStream
                gStream.Flush
                ' and write the stream to the file
                gStream.SaveToFile tFileName, adSaveCreateOverWrite
                '
                gStream.Close
            Else
                'The user pressed Cancel.
                GoTo Exit_CmdNeo4j_Click
            End If
        End If
        '
        '  test for literary genre codes
        '
        tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id ) " + _
                    "SELECT DISTINCT ZZ_SCRATCH_ASSOC.c_person_id " + _
                    "FROM ZZ_SCRATCH_ASSOC " + _
                    "WHERE (((ZZ_SCRATCH_ASSOC.c_litgenre_code)>0))"
        '
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecDeleted
        '
        ' debug
        '
        MsgBox "Literary genre code records = " + Trim(Str(tRecDeleted))
        '
        If tRecDeleted > 0 Then
            dlgSaveAs.InitialFileName = "LiteraryGenreCodes_" + tCodeStr + ".csv"
            If dlgSaveAs.Show = -1 Then
                '
                tFileName = ""
                For Each tFN In dlgSaveAs.SelectedItems
                    tFileName = tFN
                    If Not tFileName = "" Then
                        Exit For
                    End If
                Next
                If tFileName = "" Then
                    MsgBox "Bad file Name."
                    GoTo Exit_CmdNeo4j_Click
                Else
                    '  make sure the file name has a txt extension
                    If Len(tFileName) < 5 Then
                        tFileName = tFileName + ".csv"
                    ElseIf Not (LCase(Right(tFileName, 4)) = ".csv") Then
                        tFileName = tFileName + ".csv"
                    End If
                End If
                '
                gStream.Open
                '
                tQueryStr = "SELECT DISTINCT ZZ_SCRATCH_ASSOC.c_litgenre_code, ZZ_SCRATCH_ASSOC.c_litgenre_desc, ZZ_SCRATCH_ASSOC.c_litgenre_desc_chn " + _
                            "FROM ZZ_SCRATCH_ASSOC " + _
                            "WHERE (((ZZ_SCRATCH_ASSOC.c_litgenre_code)>0))"
    
                Set tRstAssocCode = CurrentDb.OpenRecordset(tQueryStr)
                
                If tCodeStr = "ascii" Then
                    tStr = "LitGenreCode" + tC + "LitGenreTrans"
                Else
                    tStr = "LitGenreCode" + tC + "LitGenreTrans" + tC + "LitGenreHZ"
                End If
                
                gStream.WriteText tStr, adWriteLine
                
                With tRstAssocCode
                    .MoveFirst
                    Do While Not .EOF
                        If Not IsNull(!c_litgenre_code) Then
                            '
                            tStr = Trim(Str(!c_litgenre_code)) + tC
                            '
                            tStr = tStr + Trim(!c_litgenre_desc)
                            '
                            If Not (tCodeStr = "ascii") Then
                                tStr = tStr + tC + Trim(!c_litgenre_desc_chn)
                            End If
                            gStream.WriteText tStr, adWriteLine
                        End If
                        .MoveNext
                    Loop
                End With
                '
                ' now make sure all the data is copied to tStream
                gStream.Flush
                ' and write the stream to the file
                gStream.SaveToFile tFileName, adSaveCreateOverWrite
                '
                gStream.Close
            Else
                'The user pressed Cancel.
                GoTo Exit_CmdNeo4j_Click
            End If
        End If
        '
        '  test for institution codes
        '
        tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id ) " + _
                    "SELECT DISTINCT ZZ_SCRATCH_ASSOC.c_person_id " + _
                    "FROM ZZ_SCRATCH_ASSOC " + _
                    "WHERE (((ZZ_SCRATCH_ASSOC.c_inst_code)>0))"
        '
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecDeleted
        '
        ' debug
        '
        MsgBox "Institution code records = " + Trim(Str(tRecDeleted))
        '
        If tRecDeleted > 0 Then
            dlgSaveAs.InitialFileName = "InstitutionCodes_" + tCodeStr + ".csv"
            If dlgSaveAs.Show = -1 Then
                '
                tFileName = ""
                For Each tFN In dlgSaveAs.SelectedItems
                    tFileName = tFN
                    If Not tFileName = "" Then
                        Exit For
                    End If
                Next
                If tFileName = "" Then
                    MsgBox "Bad file Name."
                    GoTo Exit_CmdNeo4j_Click
                Else
                    '  make sure the file name has a txt extension
                    If Len(tFileName) < 5 Then
                        tFileName = tFileName + ".csv"
                    ElseIf Not (LCase(Right(tFileName, 4)) = ".csv") Then
                        tFileName = tFileName + ".csv"
                    End If
                End If
                '
                gStream.Open
                '
                tQueryStr = "SELECT DISTINCT ZZ_SCRATCH_ASSOC.c_inst_code, ZZ_SCRATCH_ASSOC.c_inst_name_py, ZZ_SCRATCH_ASSOC.c_inst_name_hz " + _
                            "FROM ZZ_SCRATCH_ASSOC " + _
                            "WHERE (((ZZ_SCRATCH_ASSOC.c_inst_code)>0))"
    
                Set tRstAssocCode = CurrentDb.OpenRecordset(tQueryStr)
                
                If tCodeStr = "ascii" Then
                    tStr = "InstitutionCode" + tC + "InstitutionNamePY"
                Else
                    tStr = "InstitutionCode" + tC + "InstitutionNamePY" + tC + "InstitutionNameHZ"
                End If
                
                gStream.WriteText tStr, adWriteLine
                
                With tRstAssocCode
                    .MoveFirst
                    Do While Not .EOF
                        If Not IsNull(!c_inst_code) Then
                            '
                            tStr = Trim(Str(!c_inst_code)) + tC
                            '
                            tStr = tStr + Trim(!c_inst_name_py)
                            '
                            If Not (tCodeStr = "ascii") Then
                                tStr = tStr + tC + Trim(!c_inst_name_hz)
                            End If
                            gStream.WriteText tStr, adWriteLine
                        End If
                        .MoveNext
                    Loop
                End With
                '
                ' now make sure all the data is copied to tStream
                gStream.Flush
                ' and write the stream to the file
                gStream.SaveToFile tFileName, adSaveCreateOverWrite
                '
                gStream.Close
            Else
                'The user pressed Cancel.
                GoTo Exit_CmdNeo4j_Click
            End If
        End If
        '
        '  test for occasion codes
        '
        tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id ) " + _
                    "SELECT DISTINCT ZZ_SCRATCH_ASSOC.c_person_id " + _
                    "FROM ZZ_SCRATCH_ASSOC " + _
                    "WHERE (((ZZ_SCRATCH_ASSOC.c_occasion_code)>0))"
        '
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecDeleted
        '
        ' debug
        '
        MsgBox "Occasion code records = " + Trim(Str(tRecDeleted))
        '
        If tRecDeleted > 0 Then
            dlgSaveAs.InitialFileName = "OccasionCodes_" + tCodeStr + ".csv"
            If dlgSaveAs.Show = -1 Then
                '
                tFileName = ""
                For Each tFN In dlgSaveAs.SelectedItems
                    tFileName = tFN
                    If Not tFileName = "" Then
                        Exit For
                    End If
                Next
                If tFileName = "" Then
                    MsgBox "Bad file Name."
                    GoTo Exit_CmdNeo4j_Click
                Else
                    '  make sure the file name has a txt extension
                    If Len(tFileName) < 5 Then
                        tFileName = tFileName + ".csv"
                    ElseIf Not (LCase(Right(tFileName, 4)) = ".csv") Then
                        tFileName = tFileName + ".csv"
                    End If
                End If
                '
                gStream.Open
                '
                tQueryStr = "SELECT DISTINCT ZZ_SCRATCH_ASSOC.c_occasion_code, ZZ_SCRATCH_ASSOC.c_occasion_desc, ZZ_SCRATCH_ASSOC.c_occasion_desc_chn " + _
                            "FROM ZZ_SCRATCH_ASSOC " + _
                            "WHERE (((ZZ_SCRATCH_ASSOC.c_occasion_code)>0))"
    
                Set tRstAssocCode = CurrentDb.OpenRecordset(tQueryStr)
                
                If tCodeStr = "ascii" Then
                    tStr = "OccasionCode" + tC + "OccasionTrans"
                Else
                    tStr = "OccasionCode" + tC + "OccasionTrans" + tC + "OccasionHZ"
                End If
                
                gStream.WriteText tStr, adWriteLine
                
                With tRstAssocCode
                    .MoveFirst
                    Do While Not .EOF
                        If Not IsNull(!c_occasion_code) Then
                            '
                            tStr = Trim(Str(!c_occasion_code)) + tC
                            '
                            tStr = tStr + Trim(!c_occasion_desc)
                            '
                            If Not (tCodeStr = "ascii") Then
                                tStr = tStr + tC + Trim(!c_occasion_desc_chn)
                            End If
                            gStream.WriteText tStr, adWriteLine
                        End If
                        .MoveNext
                    Loop
                End With
                '
                ' now make sure all the data is copied to tStream
                gStream.Flush
                ' and write the stream to the file
                gStream.SaveToFile tFileName, adSaveCreateOverWrite
                '
                gStream.Close
            Else
                'The user pressed Cancel.
                GoTo Exit_CmdNeo4j_Click
            End If
        End If
        '
        '  test for topic codes
        '
        tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id ) " + _
                    "SELECT DISTINCT ZZ_SCRATCH_ASSOC.c_person_id " + _
                    "FROM ZZ_SCRATCH_ASSOC " + _
                    "WHERE (((ZZ_SCRATCH_ASSOC.c_topic_code)>0))"
        '
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecDeleted
        '
        ' debug
        '
        MsgBox "Topic code records = " + Trim(Str(tRecDeleted))
        '
        If tRecDeleted > 0 Then
            dlgSaveAs.InitialFileName = "TopicCodes_" + tCodeStr + ".csv"
            If dlgSaveAs.Show = -1 Then
                '
                tFileName = ""
                For Each tFN In dlgSaveAs.SelectedItems
                    tFileName = tFN
                    If Not tFileName = "" Then
                        Exit For
                    End If
                Next
                If tFileName = "" Then
                    MsgBox "Bad file Name."
                    GoTo Exit_CmdNeo4j_Click
                Else
                    '  make sure the file name has a txt extension
                    If Len(tFileName) < 5 Then
                        tFileName = tFileName + ".csv"
                    ElseIf Not (LCase(Right(tFileName, 4)) = ".csv") Then
                        tFileName = tFileName + ".csv"
                    End If
                End If
                '
                gStream.Open
                '
                '  I need to use a join to TOPIC_CODES
                '
                tQueryStr = "SELECT DISTINCT ZZ_SCRATCH_ASSOC.c_topic_code, SCHOLARLYTOPIC_CODES.c_topic_desc, SCHOLARLYTOPIC_CODES.c_topic_desc_chn " + _
                            "FROM ZZ_SCRATCH_ASSOC INNER JOIN SCHOLARLYTOPIC_CODES ON ZZ_SCRATCH_ASSOC.c_topic_code = SCHOLARLYTOPIC_CODES.c_topic_code " + _
                            "WHERE (((ZZ_SCRATCH_ASSOC.c_topic_code)>0))"
    
                Set tRstAssocCode = CurrentDb.OpenRecordset(tQueryStr)
                
                If tCodeStr = "ascii" Then
                    tStr = "TopicCode" + tC + "TopicTrans"
                Else
                    tStr = "TopicCode" + tC + "TopicTrans" + tC + "TopicHZ"
                End If
                
                gStream.WriteText tStr, adWriteLine
                
                With tRstAssocCode
                    .MoveFirst
                    Do While Not .EOF
                        If Not IsNull(!c_topic_code) Then
                            '
                            tStr = Trim(Str(!c_topic_code)) + tC
                            '
                            tStr = tStr + Trim(!c_topic_desc)
                            '
                            If Not (tCodeStr = "ascii") Then
                                tStr = tStr + tC + Trim(!c_topic_desc_chn)
                            End If
                            gStream.WriteText tStr, adWriteLine
                        End If
                        .MoveNext
                    Loop
                End With
                '
                ' now make sure all the data is copied to tStream
                gStream.Flush
                ' and write the stream to the file
                gStream.SaveToFile tFileName, adSaveCreateOverWrite
                '
                gStream.Close
            Else
                'The user pressed Cancel.
                GoTo Exit_CmdNeo4j_Click
            End If
        End If
    MsgBox "Finished saving to Neo4j"

    'Set the object variable to Nothing.
    Set dlgSaveAs = Nothing
    
Exit_CmdNeo4j_Click:
    Exit Sub

Err_CmdNeo4j_Click:
    MsgBox Err.Description
    Resume Exit_CmdNeo4j_Click

End Sub

Private Sub CmdPickAssoc_Click()
    On Error GoTo Err_CmdPickAssoc_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim strAssoc As String

    TxtAssocCode.Visible = True
    TxtAssocCode.SetFocus
    strAssoc = TxtAssocCode.TEXT
    
    stDocName = "frmPickAssoc_multi"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , acDialog, strAssoc
    
    If CurrentProject.AllForms("frmPickAssoc_multi").IsLoaded Then
        Dim intAssoc As Integer
        Dim strAssoc_DESC As String
            
        Forms!frmPickAssoc_multi.Form!TxtAssocID.Visible = True
        Forms!frmPickAssoc_multi.Form!TxtAssocID.SetFocus
        intAssoc = Forms!frmPickAssoc_multi.Form!TxtAssocID.Value
        Forms!frmPickAssoc_multi.Form!subTreeView.SetFocus
        Forms!frmPickAssoc_multi.Form!TxtAssocID.Visible = False
        TxtAssocCode.Value = intAssoc
        gAssocCodeStr = Trim(Str(intAssoc))
        
        If TxtAssocCode.Value = -1 Or TxtAssocCode.Value = -2 Then
            If TxtAssocCode.Value = -1 Then
                TxtAssocDesc.Value = "[[All]]"
                TxtAssocChn.Value = "[[All]]"
            Else
                TxtAssocDesc.Value = "[[Multi]]"
                TxtAssocChn.Value = "[[" + ChrW(22810) + ChrW(36984) + "]]"
            End If
            
            Forms!frmPickAssoc_multi.Form!TxtTypeID.Visible = True
            Forms!frmPickAssoc_multi.Form!TxtTypeID.SetFocus
            strAssoc_DESC = Forms!frmPickAssoc_multi.Form!TxtTypeID.Value
            Forms!frmPickAssoc_multi.Form!subTreeView.SetFocus
            Forms!frmPickAssoc_multi.Form!TxtTypeID.Visible = False
            TxtTypeCode.Value = strAssoc_DESC
            gAssocTypeStr = Trim(strAssoc_DESC)
            
            If TxtTypeCode.Value = "000" Then
                TxtTypeDesc.Value = "[ALL]"
                TxtTypeChn.Value = "[ALL]"
            Else
                Forms!frmPickAssoc_multi.Form!TxtTypeDesc.Visible = True
                Forms!frmPickAssoc_multi.Form!TxtTypeDesc.SetFocus
                strAssoc_DESC = Forms!frmPickAssoc_multi.Form!TxtTypeDesc.Value
                Forms!frmPickAssoc_multi.Form!subTreeView.SetFocus
                Forms!frmPickAssoc_multi.Form!TxtTypeDesc.Visible = False
                TxtTypeDesc.Value = strAssoc_DESC
                    
                Forms!frmPickAssoc_multi.Form!TxtTypeDescChn.Visible = True
                Forms!frmPickAssoc_multi.Form!TxtTypeDescChn.SetFocus
                strAssoc_DESC = Forms!frmPickAssoc_multi.Form!TxtTypeDescChn.Value
                Forms!frmPickAssoc_multi.Form!subTreeView.SetFocus
                Forms!frmPickAssoc_multi.Form!TxtTypeDescChn.Visible = False
                TxtTypeChn.Value = strAssoc_DESC
            End If
        Else
            Forms!frmPickAssoc_multi.Form!TxtAssocDesc.Visible = True
            Forms!frmPickAssoc_multi.Form!TxtAssocDesc.SetFocus
            strAssoc_DESC = Forms!frmPickAssoc_multi.Form!TxtAssocDesc.Value
            Forms!frmPickAssoc_multi.Form!subTreeView.SetFocus
            Forms!frmPickAssoc_multi.Form!TxtAssocDesc.Visible = False
            TxtAssocDesc.Value = strAssoc_DESC
            
            Forms!frmPickAssoc_multi.Form!TxtAssocDescChn.Visible = True
            Forms!frmPickAssoc_multi.Form!TxtAssocDescChn.SetFocus
            strAssoc_DESC = Forms!frmPickAssoc_multi.Form!TxtAssocDescChn.Value
            Forms!frmPickAssoc_multi.Form!subTreeView.SetFocus
            Forms!frmPickAssoc_multi.Form!TxtAssocDescChn.Visible = False
            TxtAssocChn.Value = strAssoc_DESC
            
            TxtTypeCode.Value = ""
            TxtTypeDesc.Value = "N/A"
            TxtTypeChn.Value = "N/A"
        End If
                
                
        DoCmd.Close acForm, stDocName
        '
        CmdQuery.Enabled = True
        CmdSaveAssociations.Enabled = True
    Else
        CmdQuery.Enabled = False
        CmdSaveAssociations.Enabled = False
    End If
            
    CmdPickAssoc.SetFocus
    TxtAssocCode.Visible = False
        
Exit_CmdPickAssoc_Click:
    Exit Sub

Err_CmdPickAssoc_Click:
    MsgBox Err.Description
    Resume Exit_CmdPickAssoc_Click
    
End Sub

Private Sub CmdQuery_Click()
    On Error GoTo Err_Run_Query

    Dim rst As DAO.Recordset, tContinue As Integer
    Dim tRstAssoc As DAO.Recordset, tRstAddrList As DAO.Recordset, tRstDummy As DAO.Recordset
    Dim tQueryInsertStr As String, tQuerySelectStr As String, tQueryFromStr As String, tQueryWhereStr As String
    Dim tQueryStr As String, tRecDrop As Long, tStrWhereSQL As String
    Dim tUseAddr As Boolean
    
    Dim cmdSQL As ADODB.Command, tRecCount As Long
    
    Set cmdSQL = New ADODB.Command
    '
    cmdSQL.ActiveConnection = CurrentProject.Connection
    cmdSQL.CommandType = adCmdText
    '
    '  to clear the table, close and then delete records
    '
    Set tRstAssoc = ZZ_SCRATCH_ASSOC.Form.Recordset
    Set tRstDummy = CurrentDb.OpenRecordset("Z_SCRATCH_DUMMY_AC", dbOpenDynaset)
    Set ZZ_SCRATCH_ASSOC.Form.Recordset = tRstDummy
    tRstAssoc.Close
        '
    cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_ASSOC"
    cmdSQL.Execute tRecCount
    '
    '  now the people table
    '
    Set gRstPeople = ZZ_SCRATCH_P_ASSOC.Form.Recordset
    Set tRstDummy = CurrentDb.OpenRecordset("Z_SCRATCH_DUMMY_AP", dbOpenDynaset)
    Set ZZ_SCRATCH_P_ASSOC.Form.Recordset = tRstDummy
    gRstPeople.Close
    '
    cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_P_ASSOC"
    cmdSQL.Execute tRecCount
    
    '
    ' now see if address IDs will be used.  If so, zap the scratch file and repopulate
    '
    ' MsgBox "About to process address"
    If gUseADDRID Then
        '
        '  the strategy here is to fill the scratch file with all the relevant addresses from ZZZ_BELONGS_TO
        '
        cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_ADDR"
        cmdSQL.Execute tRecCount
        '
        If ChkSubUnits.Value Then
            tQueryStr = "INSERT INTO ZZ_SCRATCH_ADDR ( c_addr_id ) " + _
                "SELECT DISTINCT ZZZ_BELONGS_TO.c_addr_id " + _
                "FROM ZZ_SCRATCH_ADDR_LIST INNER JOIN ZZZ_BELONGS_TO ON " + _
                "ZZ_SCRATCH_ADDR_LIST.c_addr_id = ZZZ_BELONGS_TO.c_belongs_to"
        Else
            tQueryStr = "INSERT INTO ZZ_SCRATCH_ADDR ( c_addr_id ) " + _
                "SELECT DISTINCT c_addr_id " + _
                "FROM ZZ_SCRATCH_ADDR_LIST"
        End If
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecCount
        '
        '  see if we need to use the historical XY search
        '
        If ChkXYRef.Value Then
            '
            '  the strategy here is to dump the IDs to ZZ_ADDRESSES then copy to ZZ_SCRATCH_ADDR_LIST
            '  (I borrow ZZ_ADDRESSES from the Pick Addresses form in order to keep the initial selection
            '   of addresses for the query intact.)
            '
            '  zap the list
            '
            tQueryStr = "DELETE * FROM ZZ_ADDRESSES"
            cmdSQL.CommandText = tQueryStr
            cmdSQL.Execute tRecDeleted
            '
            '  run the query
            '
            '  FrameXY.Value = 2 :: Narrow, FrameXY.Value = 1 :: Broad
            '
            If FrameXY.Value = 2 Then
                tStrWhereSQL = "WHERE (((ADDR_CODES.x_coord)>=([ADDR_CODES_1].[x_coord]-0.03) And " + _
                    "(ADDR_CODES.x_coord)<=([ADDR_CODES_1].[x_coord]+0.03)) AND " + _
                    "((ADDR_CODES.y_coord)>=([ADDR_CODES_1].[y_coord]-0.03) And " + _
                    "(ADDR_CODES.y_coord)<=([ADDR_CODES_1].[y_coord]+0.03)))"
            Else
                tStrWhereSQL = "WHERE (((ADDR_CODES.x_coord)>=([ADDR_CODES_1].[x_coord]-0.06) And " + _
                    "(ADDR_CODES.x_coord)<=([ADDR_CODES_1].[x_coord]+0.06)) AND " + _
                    "((ADDR_CODES.y_coord)>=([ADDR_CODES_1].[y_coord]-0.06) And " + _
                    "(ADDR_CODES.y_coord)<=([ADDR_CODES_1].[y_coord]+0.06)))"
            End If
            '
            tQueryStr = "INSERT INTO ZZ_ADDRESSES ( c_addr_id )SELECT DISTINCT ADDR_CODES.c_addr_id " + _
                "FROM ADDR_CODES, ZZ_SCRATCH_ADDR INNER JOIN ADDR_CODES AS ADDR_CODES_1 ON " + _
                "ZZ_SCRATCH_ADDR.c_addr_id = ADDR_CODES_1.c_addr_id " + tStrWhereSQL
                
            cmdSQL.CommandText = tQueryStr
            cmdSQL.Execute tRecDeleted
            '
            ' now get the address IDs from the initial list that have no xy coordinates
            '
            tQueryStr = "INSERT INTO ZZ_ADDRESSES ( c_addr_id ) SELECT ZZ_SCRATCH_ADDR.c_addr_id " + _
                "FROM ZZ_SCRATCH_ADDR INNER JOIN ADDR_CODES ON " + _
                "ZZ_SCRATCH_ADDR.c_addr_id = ADDR_CODES.c_addr_id " + _
                "WHERE (((ADDR_CODES.x_coord) Is Null)) OR (((ADDR_CODES.y_coord) Is Null))"
            cmdSQL.CommandText = tQueryStr
            cmdSQL.Execute tRecDeleted
            '
            '  zap ZZ_SCRATCH_ADDR
            '
            tQueryStr = "DELETE * FROM ZZ_SCRATCH_ADDR"
            cmdSQL.CommandText = tQueryStr
            cmdSQL.Execute tRecDeleted
            '
            '  copy the list
            '
            tQueryStr = "INSERT INTO ZZ_SCRATCH_ADDR ( c_addr_id )SELECT DISTINCT ZZ_ADDRESSES.c_addr_id " + _
                "FROM ZZ_ADDRESSES"
            cmdSQL.CommandText = tQueryStr
            cmdSQL.Execute tRecDeleted
            '
            '  zap the temporary list
            '
            tQueryStr = "DELETE * FROM ZZ_ADDRESSES"
            cmdSQL.CommandText = tQueryStr
            cmdSQL.Execute tRecDeleted
            '
        End If
        
        tUseAddr = True
    Else
        tUseAddr = False
    End If

    ' next build the appropriate query string
    
    tQueryInsertStr = "INSERT INTO ZZ_SCRATCH_ASSOC (c_person_id, c_name, c_name_chn, c_index_year, c_index_year_type_code, c_sex, c_addr_id, c_dy, " + _
                "c_assoc_code, c_kin_code, c_kin_id, c_assoc_id, c_assoc_kin_code, c_assoc_kin_id, c_assoc_count, c_assoc_first_year, " + _
                "c_assoc_last_year, c_source, c_assoc_place_addr_id, c_litgenre_code, c_occasion_code, " + _
                "c_topic_code, c_inst_code, c_inst_name_code, c_text_title, c_assoc_claimer_id, c_distance ) "
                    

    tQuerySelectStr = "SELECT ASSOC_DATA.c_personid, BIOG_MAIN.c_name, BIOG_MAIN.c_name_chn, BIOG_MAIN.c_index_year, BIOG_MAIN.c_index_year_type_code, " + _
                "iif(BIOG_MAIN.c_female,'F','M'), BIOG_MAIN.c_index_addr_id, BIOG_MAIN.c_dy, ASSOC_DATA.c_assoc_code, ASSOC_DATA.c_kin_code, " + _
                "ASSOC_DATA.c_kin_id, ASSOC_DATA.c_assoc_id, ASSOC_DATA.c_assoc_kin_code, ASSOC_DATA.c_assoc_kin_id, ASSOC_DATA.c_assoc_count, " + _
                "ASSOC_DATA.c_assoc_first_year, ASSOC_DATA.c_assoc_last_year, ASSOC_DATA.c_source, ASSOC_DATA.c_addr_id, ASSOC_DATA.c_litgenre_code, " + _
                "ASSOC_DATA.c_occasion_code, ASSOC_DATA.c_topic_code, ASSOC_DATA.c_inst_code, ASSOC_DATA.c_inst_name_code, ASSOC_DATA.c_text_title, " + _
                "ASSOC_DATA.c_assoc_claimer_id , 1 AS c_distance "

    
    '  set the from tables and the dynasties table, if needed
    
    '  With the introduction of multi-select, I now join the ZZ_ASSOC_CODE to the NONKIN to get the associations

    If tUseAddr Then
        If gUseDynasties And gToDynasty > -2 Then
            tQueryFromStr = "FROM DYNASTIES INNER JOIN ( (  ( ZZ_ASSOC_CODE INNER JOIN ( ASSOC_DATA INNER JOIN BIOG_MAIN ON ASSOC_DATA.c_personid = BIOG_MAIN.c_personid ) " + _
                            "ON ZZ_ASSOC_CODE.c_assoc_code = ASSOC_DATA.c_assoc_code ) INNER JOIN ASSOC_CODE_TYPE_REL " + _
                            "ON ASSOC_DATA.c_assoc_code = ASSOC_CODE_TYPE_REL.c_assoc_code ) INNER JOIN ZZ_SCRATCH_ADDR " + _
                            "ON BIOG_MAIN.c_index_addr_id = ZZ_SCRATCH_ADDR.c_addr_id ) ON DYNASTIES.c_dy = BIOG_MAIN.c_dy "

        Else
            tQueryFromStr = "FROM ( ( ZZ_ASSOC_CODE INNER JOIN ( ASSOC_DATA INNER JOIN BIOG_MAIN ON ASSOC_DATA.c_personid = BIOG_MAIN.c_personid ) " + _
                            "ON ZZ_ASSOC_CODE.c_assoc_code = ASSOC_DATA.c_assoc_code ) INNER JOIN ASSOC_CODE_TYPE_REL " + _
                            "ON ASSOC_DATA.c_assoc_code = ASSOC_CODE_TYPE_REL.c_assoc_code ) INNER JOIN ZZ_SCRATCH_ADDR " + _
                            "ON BIOG_MAIN.c_index_addr_id = ZZ_SCRATCH_ADDR.c_addr_id "

        End If
    Else
        If gUseDynasties And gToDynasty > -2 Then

            tQueryFromStr = "FROM DYNASTIES INNER JOIN ( ( ZZ_ASSOC_CODE INNER JOIN ( ASSOC_DATA INNER JOIN BIOG_MAIN ON ASSOC_DATA.c_personid = BIOG_MAIN.c_personid ) " + _
                            "ON ZZ_ASSOC_CODE.c_assoc_code = ASSOC_DATA.c_assoc_code ) INNER JOIN ASSOC_CODE_TYPE_REL " + _
                            "ON ASSOC_DATA.c_assoc_code = ASSOC_CODE_TYPE_REL.c_assoc_code ) ON DYNASTIES.c_dy = BIOG_MAIN.c_dy "
        Else
            tQueryFromStr = "FROM ( ZZ_ASSOC_CODE INNER JOIN ( ASSOC_DATA INNER JOIN BIOG_MAIN ON ASSOC_DATA.c_personid = BIOG_MAIN.c_personid ) " + _
                            "ON ZZ_ASSOC_CODE.c_assoc_code = ASSOC_DATA.c_assoc_code )  INNER JOIN ASSOC_CODE_TYPE_REL " + _
                            "ON ASSOC_DATA.c_assoc_code = ASSOC_CODE_TYPE_REL.c_assoc_code "
        End If
    End If
    '  set the where conditions
    '
    '  Start with the index years
    '
    tQueryWhereStr = ""
    
    If gUseIndexYears Then
        '
        '  four possibilities
        '
        If gFromStr = "" And gToStr = "" Then
            tQueryWhereStr = ""
        ElseIf gFromStr = "" Then
            tQueryWhereStr = "WHERE (((BIOG_MAIN.c_index_year)<=" + gToStr + ") "
        ElseIf gToStr = "" Then
            tQueryWhereStr = "WHERE (((BIOG_MAIN.c_index_year)>=" + gFromStr + ") "
        Else
            tQueryWhereStr = "WHERE (((BIOG_MAIN.c_index_year)<=" + gToStr + ") AND ((BIOG_MAIN.c_index_year)>=" + gFromStr + ") "
        End If
    ElseIf gUseDynasties Then
        '
        '  five possibilities (all, just from, just to, both from and to, and a cluelessly unset parameter)
        '
        If gFromDynasty = -2 Then
            tQueryWhereStr = "Where (((BIOG_MAIN.c_dy) > 0 ) "
        ElseIf gFromDynasty = -1 And gToDynasty > 0 Then
            tQueryWhereStr = "WHERE (((DYNASTIES.c_start)<" + Str(gToDynastyEnd) + ") "
        ElseIf gFromDynasty > 0 And gToDynasty = -1 Then
            tQueryWhereStr = "WHERE (((DYNASTIES.c_end)>" + Str(gFromDynastyBegin) + ") "
        ElseIf gFromDynasty = gToDynasty And gFromDynasty > 0 Then
            tQueryWhereStr = "WHERE (((DYNASTIES.c_dy)=" + Str(gFromDynasty) + ") "
        ElseIf gFromDynasty > 0 And gToDynasty > 0 Then
            tQueryWhereStr = "WHERE (((DYNASTIES.c_end)>" + Str(gFromDynastyBegin) + ") AND " + _
                "((DYNASTIES.c_start)<" + Str(gToDynastyEnd) + ") "
        Else
            tQueryWhereStr = ""
        End If
    End If

    If Not (tQueryWhereStr = "") Then
        tQueryWhereStr = tQueryWhereStr + ")"
    End If

    cmdSQL.CommandText = tQueryInsertStr + tQuerySelectStr + tQueryFromStr + tQueryWhereStr
    cmdSQL.Execute tRecCount
    '
    '  Because the query is complex enough as is, I add some xy information and get information from BIOG_MAIN in three separate steps
    '
    If tRecCount > 0 Then
        cmdSQL.CommandText = "UPDATE ZZ_SCRATCH_ASSOC INNER JOIN ADDR_CODES ON ZZ_SCRATCH_ASSOC.c_assoc_place_addr_id = ADDR_CODES.c_addr_id " + _
                                        "SET ZZ_SCRATCH_ASSOC.c_assoc_place_addr_xcoord = [ADDR_CODES].[x_coord], " + _
                                            "ZZ_SCRATCH_ASSOC.c_assoc_place_addr_ycoord = [ADDR_CODES].[y_coord], " + _
                                            "ZZ_SCRATCH_ASSOC.c_assoc_place_addr_name = [ADDR_CODES].[c_name], " + _
                                            "ZZ_SCRATCH_ASSOC.c_assoc_place_addr_chn = [ADDR_CODES].[c_name_chn]"
        cmdSQL.Execute tRecCount

    ' fill in the outer join information for the association
        
    tQueryUpdateStr = "UPDATE ( ( ( SCHOLARLYTOPIC_CODES RIGHT JOIN ( OCCASION_CODES  RIGHT JOIN ( ADDR_CODES RIGHT JOIN ( INDEXYEAR_TYPE_CODES " + _
                                    "RIGHT JOIN ZZ_SCRATCH_ASSOC ON INDEXYEAR_TYPE_CODES.c_index_year_type_code = ZZ_SCRATCH_ASSOC.c_index_year_type_code ) " + _
                        "ON ADDR_CODES.c_addr_id = ZZ_SCRATCH_ASSOC.c_addr_id ) " + _
                        "ON OCCASION_CODES.c_occasion_code = ZZ_SCRATCH_ASSOC.c_occasion_code ) " + _
                        "ON SCHOLARLYTOPIC_CODES.c_topic_code = ZZ_SCRATCH_ASSOC.c_topic_code ) LEFT JOIN LITERARYGENRE_CODES " + _
                        "ON ZZ_SCRATCH_ASSOC.c_litgenre_code = LITERARYGENRE_CODES.c_lit_genre_code ) LEFT JOIN SOCIAL_INSTITUTION_NAME_CODES " + _
                        "ON ZZ_SCRATCH_ASSOC.c_inst_name_code = SOCIAL_INSTITUTION_NAME_CODES.c_inst_name_code ) " + _
                        "LEFT JOIN DYNASTIES ON ZZ_SCRATCH_ASSOC.c_dy = DYNASTIES.c_dy "

    tQuerySetStr = "SET ZZ_SCRATCH_ASSOC.c_dynasty = [DYNASTIES].[c_dynasty], ZZ_SCRATCH_ASSOC.c_dynasty_chn = [DYNASTIES].[c_dynasty_chn], " + _
                        "ZZ_SCRATCH_ASSOC.c_index_year_type_desc = [INDEXYEAR_TYPE_CODES].[c_index_year_type_desc], " + _
                        "ZZ_SCRATCH_ASSOC.c_index_year_type_hz = [INDEXYEAR_TYPE_CODES].[c_index_year_type_hz], " + _
                        "ZZ_SCRATCH_ASSOC.c_addr_name = [ADDR_CODES].[c_name], ZZ_SCRATCH_ASSOC.c_addr_chn = [ADDR_CODES].[c_name_chn], " + _
                        "ZZ_SCRATCH_ASSOC.x_coord = [ADDR_CODES].[x_coord], ZZ_SCRATCH_ASSOC.y_coord = [ADDR_CODES].[y_coord], " + _
                        "ZZ_SCRATCH_ASSOC.c_litgenre_desc = [LITERARYGENRE_CODES].[c_lit_genre_desc], " + _
                        "ZZ_SCRATCH_ASSOC.c_litgenre_desc_chn = [LITERARYGENRE_CODES].[c_lit_genre_desc_chn], " + _
                        "ZZ_SCRATCH_ASSOC.c_occasion_desc = [OCCASION_CODES].[c_occasion_desc], " + _
                        "ZZ_SCRATCH_ASSOC.c_occasion_desc_chn = [OCCASION_CODES].[c_occasion_desc_chn], " + _
                        "ZZ_SCRATCH_ASSOC.c_topic_desc = [SCHOLARLYTOPIC_CODES].[c_topic_desc], " + _
                        "ZZ_SCRATCH_ASSOC.c_topic_desc_chn = [SCHOLARLYTOPIC_CODES].[c_topic_desc_chn], " + _
                        "ZZ_SCRATCH_ASSOC.c_inst_name_py = [SOCIAL_INSTITUTION_NAME_CODES].[c_inst_name_py], " + _
                        "ZZ_SCRATCH_ASSOC.c_inst_name_hz = [SOCIAL_INSTITUTION_NAME_CODES].[c_inst_name_hz]"

        cmdSQL.CommandText = tQueryUpdateStr + tQuerySetStr
        cmdSQL.Execute tRecCount

    ' now get the basic data for the associate (inner join), the kin, the associate's kin, and the claimer (outer joins)

    tQueryUpdateStr = "UPDATE ( ( ( ( ( ZZ_SCRATCH_ASSOC LEFT JOIN BIOG_MAIN ON ZZ_SCRATCH_ASSOC.c_assoc_id = BIOG_MAIN.c_personid ) INNER JOIN ASSOC_CODES " + _
                        "ON ZZ_SCRATCH_ASSOC.c_assoc_code = ASSOC_CODES.c_assoc_code ) LEFT JOIN ASSOC_CODE_TYPE_REL " + _
                        "ON ZZ_SCRATCH_ASSOC.c_assoc_code = ASSOC_CODE_TYPE_REL.c_assoc_code ) LEFT JOIN BIOG_MAIN AS BIOG_MAIN_1 " + _
                        "ON ZZ_SCRATCH_ASSOC.c_kin_id = BIOG_MAIN_1.c_personid ) LEFT JOIN BIOG_MAIN AS BIOG_MAIN_2 " + _
                        "ON ZZ_SCRATCH_ASSOC.c_assoc_kin_id = BIOG_MAIN_2.c_personid ) LEFT JOIN BIOG_MAIN AS BIOG_MAIN_3 " + _
                        "ON ZZ_SCRATCH_ASSOC.c_assoc_claimer_id = BIOG_MAIN_3.c_personid "

    tQuerySetStr = "SET ZZ_SCRATCH_ASSOC.c_assoc_name = [BIOG_MAIN].[c_name], ZZ_SCRATCH_ASSOC.c_assoc_chn = [BIOG_MAIN].[c_name_chn], " + _
                        "ZZ_SCRATCH_ASSOC.c_assoc_index_year = [BIOG_MAIN].[c_index_year], " + _
                        "ZZ_SCRATCH_ASSOC.c_assoc_index_year_type_code = [BIOG_MAIN].[c_index_year_type_code], " + _
                        "ZZ_SCRATCH_ASSOC.c_assoc_dy = [BIOG_MAIN].[c_dy], ZZ_SCRATCH_ASSOC.c_assoc_sex = IIf([BIOG_MAIN].[c_female], 'F', 'M'), " + _
                        "ZZ_SCRATCH_ASSOC.c_assoc_type = [ASSOC_CODE_TYPE_REL].[c_assoc_type_code], " + _
                        "ZZ_SCRATCH_ASSOC.c_assoc_desc = [ASSOC_CODES].[c_assoc_desc], " + _
                        "ZZ_SCRATCH_ASSOC.c_assoc_desc_chn = [ASSOC_CODES].[c_assoc_desc_chn], ZZ_SCRATCH_ASSOC.c_kin_name = [BIOG_MAIN_1].[c_name], " + _
                        "ZZ_SCRATCH_ASSOC.c_kin_chn = [BIOG_MAIN_1].[c_name_chn], ZZ_SCRATCH_ASSOC.c_assoc_kin_name = [BIOG_MAIN_2].[c_name], " + _
                        "ZZ_SCRATCH_ASSOC.c_assoc_kin_chn = [BIOG_MAIN_2].[c_name_chn], " + _
                        "ZZ_SCRATCH_ASSOC.c_assoc_claimer_name = [BIOG_MAIN_3].[c_name], " + _
                        "ZZ_SCRATCH_ASSOC.c_assoc_claimer_name_chn = [BIOG_MAIN_3].[c_name_chn], " + _
                        "ZZ_SCRATCH_ASSOC.c_assoc_claimer_index_year = [BIOG_MAIN_3].[c_index_year]"

        cmdSQL.CommandText = tQueryUpdateStr + tQuerySetStr
        cmdSQL.Execute tRecCount

    ' now fill in the outer join information for the associate

    tQueryUpdateStr = "UPDATE ( ( ( ( ( ZZ_SCRATCH_ASSOC LEFT JOIN INDEXYEAR_TYPE_CODES " + _
            "ON ZZ_SCRATCH_ASSOC.c_assoc_index_year_type_code = INDEXYEAR_TYPE_CODES.c_index_year_type_code ) LEFT JOIN DYNASTIES " + _
            "ON ZZ_SCRATCH_ASSOC.c_assoc_dy = DYNASTIES.c_dy ) LEFT JOIN ADDR_CODES ON ZZ_SCRATCH_ASSOC.c_assoc_addr_id = ADDR_CODES.c_addr_id ) " + _
            "LEFT JOIN KINSHIP_CODES ON ZZ_SCRATCH_ASSOC.c_kin_code = KINSHIP_CODES.c_kincode ) LEFT JOIN KINSHIP_CODES AS KINSHIP_CODES_1 " + _
            "ON ZZ_SCRATCH_ASSOC.c_assoc_kin_code = KINSHIP_CODES_1.c_kincode ) LEFT JOIN TEXT_CODES ON ZZ_SCRATCH_ASSOC.c_source = TEXT_CODES.c_textid "
    tQuerySetStr = "SET ZZ_SCRATCH_ASSOC.c_assoc_index_year_type_desc = [INDEXYEAR_TYPE_CODES].[c_index_year_type_desc], " + _
                        "ZZ_SCRATCH_ASSOC.c_assoc_index_year_type_hz = [INDEXYEAR_TYPE_CODES].[c_index_year_type_hz], " + _
                        "ZZ_SCRATCH_ASSOC.c_assoc_dynasty = [DYNASTIES].[c_dynasty], " + _
                        "ZZ_SCRATCH_ASSOC.c_assoc_dynasty_chn = [DYNASTIES].[c_dynasty_chn], " + _
                        "ZZ_SCRATCH_ASSOC.c_assoc_addr_name = [ADDR_CODES].[c_name], ZZ_SCRATCH_ASSOC.c_assoc_addr_chn = [ADDR_CODES].[c_name_chn], " + _
                        "ZZ_SCRATCH_ASSOC.assoc_xcoord = [ADDR_CODES].[x_coord], ZZ_SCRATCH_ASSOC.assoc_ycoord = [ADDR_CODES].[y_coord], " + _
                        "ZZ_SCRATCH_ASSOC.c_kin_desc = [KINSHIP_CODES].[c_kinrel], ZZ_SCRATCH_ASSOC.c_kin_desc_chn = [KINSHIP_CODES].[c_kinrel_chn], " + _
                        "ZZ_SCRATCH_ASSOC.c_assoc_kin_desc = [KINSHIP_CODES_1].[c_kinrel], " + _
                        "ZZ_SCRATCH_ASSOC.c_assoc_kin_desc_chn = [KINSHIP_CODES_1].[c_kinrel_chn], " + _
                        "ZZ_SCRATCH_ASSOC.c_source_text = [TEXT_CODES].[c_title], ZZ_SCRATCH_ASSOC.c_source_text_chn = [TEXT_CODES].[c_title_chn]"

        cmdSQL.CommandText = tQueryUpdateStr + tQuerySetStr
        cmdSQL.Execute tRecCount

    End If
    '
    '  the next step is to clean up the data (remove duplicates) and make the table of people from the associations
    '
    If tRecCount = 0 Then
        CmdPajek.Enabled = False
        CmdUCINet.Enabled = False
        CmdGephi.Enabled = False
    Else
        CmdPajek.Enabled = True
        CmdUCINet.Enabled = True
        CmdGephi.Enabled = True
        '
        '  remove duplicated associations
        '
        '  (1) mark the passive members of the pair
        '
        tQueryStr = "UPDATE (ZZ_SCRATCH_ASSOC AS ZZ_SCRATCH_ASSOC_1 INNER JOIN ZZ_SCRATCH_ASSOC ON " + _
            "(ZZ_SCRATCH_ASSOC.c_person_id = ZZ_SCRATCH_ASSOC_1.c_assoc_id) AND " + _
            "(ZZ_SCRATCH_ASSOC_1.c_person_id = ZZ_SCRATCH_ASSOC.c_assoc_id)) " + _
            "INNER JOIN ASSOC_CODES ON (ZZ_SCRATCH_ASSOC_1.c_assoc_code = ASSOC_CODES.c_assoc_pair) AND " + _
            "(ZZ_SCRATCH_ASSOC.c_assoc_code = ASSOC_CODES.c_assoc_code) " + _
            "SET ZZ_SCRATCH_ASSOC.c_delete = 1 " + _
            "WHERE (((ASSOC_CODES.c_assoc_role_type)='P'))"
        '
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecCount
        '
        '  (2) mark the higher person ID from mutual relations
        '
        tQueryStr = "UPDATE (ZZ_SCRATCH_ASSOC AS ZZ_SCRATCH_ASSOC_1 INNER JOIN ZZ_SCRATCH_ASSOC ON " + _
            "(ZZ_SCRATCH_ASSOC.c_person_id = ZZ_SCRATCH_ASSOC_1.c_assoc_id) AND " + _
            "(ZZ_SCRATCH_ASSOC_1.c_person_id = ZZ_SCRATCH_ASSOC.c_assoc_id)) INNER JOIN ASSOC_CODES ON " + _
            "(ZZ_SCRATCH_ASSOC_1.c_assoc_code = ASSOC_CODES.c_assoc_pair) AND " + _
            "(ZZ_SCRATCH_ASSOC.c_assoc_code = ASSOC_CODES.c_assoc_code) " + _
            "SET ZZ_SCRATCH_ASSOC.c_delete = 1 " + _
            "WHERE (((ASSOC_CODES.c_assoc_role_type)='M') AND ((ZZ_SCRATCH_ASSOC.c_person_id)>[ZZ_SCRATCH_ASSOC_1].[c_person_id]))"

        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecCount
        '
        '  now delete
        '
        tQueryStr = "DELETE * FROM ZZ_SCRATCH_ASSOC WHERE c_delete = 1"
        '
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecCount
        '
        '  now get the people:  dump the person first, and then the associate, into a temporary table, then copy
        '
        tQueryStr = "DELETE * FROM ZZ_SCRATCH_IMPORT_PEOPLE"
        '
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecCount
        '
        tQueryStr = "INSERT INTO ZZ_SCRATCH_IMPORT_PEOPLE ( c_person_id, c_name, c_name_chn, c_index_year, c_dy, c_dynasty, c_dynasty_chn, " + _
                "c_sex, c_addr_id, c_addr_name, c_addr_chn, x_coord, y_coord ) " + _
            "SELECT DISTINCT ZZ_SCRATCH_ASSOC.c_person_id, ZZ_SCRATCH_ASSOC.c_name, ZZ_SCRATCH_ASSOC.c_name_chn, " + _
                "ZZ_SCRATCH_ASSOC.c_index_year, ZZ_SCRATCH_ASSOC.c_dy, ZZ_SCRATCH_ASSOC.c_dynasty, ZZ_SCRATCH_ASSOC.c_dynasty_chn, " + _
                "ZZ_SCRATCH_ASSOC.c_sex, ZZ_SCRATCH_ASSOC.c_addr_id, ZZ_SCRATCH_ASSOC.c_addr_name, " + _
                "ZZ_SCRATCH_ASSOC.c_addr_chn, ZZ_SCRATCH_ASSOC.x_coord, ZZ_SCRATCH_ASSOC.y_coord " + _
            "FROM ZZ_SCRATCH_ASSOC"
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecCount
        '
        '  now copy all the assoc IDs that do not also appear as person IDs
        '
        tQueryStr = "INSERT INTO ZZ_SCRATCH_IMPORT_PEOPLE ( c_person_id, c_name, c_name_chn, c_index_year, c_dy, c_dynasty, c_dynasty_chn, " + _
                "c_sex, c_addr_id, c_addr_name, c_addr_chn, x_coord, y_coord ) " + _
            "SELECT DISTINCT ZZ_SCRATCH_ASSOC.c_assoc_id, ZZ_SCRATCH_ASSOC.c_assoc_name, ZZ_SCRATCH_ASSOC.c_assoc_chn, " + _
                "ZZ_SCRATCH_ASSOC.c_assoc_index_year, ZZ_SCRATCH_ASSOC.c_assoc_dy, ZZ_SCRATCH_ASSOC.c_assoc_dynasty, ZZ_SCRATCH_ASSOC.c_assoc_dynasty_chn, " + _
                "ZZ_SCRATCH_ASSOC.c_assoc_sex, ZZ_SCRATCH_ASSOC.c_assoc_addr_id, " + _
                "ZZ_SCRATCH_ASSOC.c_assoc_addr_name, ZZ_SCRATCH_ASSOC.c_assoc_addr_chn, ZZ_SCRATCH_ASSOC.assoc_xcoord, ZZ_SCRATCH_ASSOC.assoc_ycoord " + _
            "FROM ZZ_SCRATCH_ASSOC LEFT JOIN ZZ_SCRATCH_ASSOC AS ZZ_SCRATCH_ASSOC_1 ON ZZ_SCRATCH_ASSOC.c_assoc_id = ZZ_SCRATCH_ASSOC_1.c_person_id " + _
            "WHERE (((ZZ_SCRATCH_ASSOC_1.c_assoc_code) Is Null))"
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecCount
            
        tQueryStr = "INSERT INTO ZZ_SCRATCH_P_ASSOC ( c_person_id, c_name, c_name_chn, c_index_year, c_dy, c_dynasty, c_dynasty_chn, c_sex, c_addr_id, " + _
             "c_addr_name, c_addr_chn, x_coord, y_coord, c_addr_type, c_index_year_type_code ) " + _
             "SELECT DISTINCT ZZ_SCRATCH_IMPORT_PEOPLE.c_person_id, ZZ_SCRATCH_IMPORT_PEOPLE.c_name, ZZ_SCRATCH_IMPORT_PEOPLE.c_name_chn, " + _
             "ZZ_SCRATCH_IMPORT_PEOPLE.c_index_year, ZZ_SCRATCH_IMPORT_PEOPLE.c_dy, ZZ_SCRATCH_IMPORT_PEOPLE.c_dynasty, " + _
             "ZZ_SCRATCH_IMPORT_PEOPLE.c_dynasty_chn, ZZ_SCRATCH_IMPORT_PEOPLE.c_sex, ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_id, " + _
             "ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_name, ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_chn, ZZ_SCRATCH_IMPORT_PEOPLE.x_coord, " + _
             "ZZ_SCRATCH_IMPORT_PEOPLE.y_coord, BIOG_MAIN.c_index_addr_type_code, BIOG_MAIN.c_index_year_type_code " + _
             "FROM BIOG_MAIN INNER JOIN ZZ_SCRATCH_IMPORT_PEOPLE ON BIOG_MAIN.c_personid = ZZ_SCRATCH_IMPORT_PEOPLE.c_person_id"
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecCount
        '
        '  now get the new index year type and index address type information
        '
        cmdSQL.CommandText = "UPDATE ( ZZ_SCRATCH_P_ASSOC LEFT JOIN INDEXYEAR_TYPE_CODES " + _
            "ON ZZ_SCRATCH_P_ASSOC.c_index_year_type_code = INDEXYEAR_TYPE_CODES.c_index_year_type_code ) LEFT JOIN BIOG_ADDR_CODES " + _
            "ON ZZ_SCRATCH_P_ASSOC.c_addr_type = BIOG_ADDR_CODES.c_addr_type " + _
        "SET ZZ_SCRATCH_P_ASSOC.c_index_year_type_desc = [INDEXYEAR_TYPE_CODES].[c_index_year_type_desc], " + _
            "ZZ_SCRATCH_P_ASSOC.c_index_year_type_hz = [INDEXYEAR_TYPE_CODES].[c_index_year_type_hz], " + _
            "ZZ_SCRATCH_P_ASSOC.c_addr_desc = [BIOG_ADDR_CODES].[c_addr_desc], ZZ_SCRATCH_P_ASSOC.c_addr_desc_chn = [BIOG_ADDR_CODES].[c_addr_desc_chn]"
        cmdSQL.Execute tRecCount

        '
        '  the final step is to calculate the xy_count
        '
        If tRecCount > 0 Then
            '
            cmdSQL.CommandText = "Delete * from tmpXY"
            cmdSQL.Execute tRecDeleted
            '
            tQueryStr = "INSERT INTO tmpXY ( x_coord, y_coord, CountOfx_coord, CountOfy_coord ) " + _
                "SELECT ZZ_SCRATCH_P_ASSOC.x_coord, ZZ_SCRATCH_P_ASSOC.y_coord, Count(ZZ_SCRATCH_P_ASSOC.x_coord) " + _
                "AS CountOfx_coord, Count(ZZ_SCRATCH_P_ASSOC.y_coord) AS CountOfy_coord " + _
                "FROM ZZ_SCRATCH_P_ASSOC " + _
                "GROUP BY ZZ_SCRATCH_P_ASSOC.x_coord, ZZ_SCRATCH_P_ASSOC.y_coord"
            '
            cmdSQL.CommandText = tQueryStr
            cmdSQL.Execute tRecCount
            '
            tQueryStr = "UPDATE tmpXY INNER JOIN ZZ_SCRATCH_P_ASSOC ON (tmpXY.y_coord = " + _
                "ZZ_SCRATCH_P_ASSOC.y_coord) AND (tmpXY.x_coord = ZZ_SCRATCH_P_ASSOC.x_coord) " + _
                "SET ZZ_SCRATCH_P_ASSOC.xy_count = [tmpXY].[CountOfx_coord]"
        
            cmdSQL.CommandText = tQueryStr
            cmdSQL.Execute tRecCount
            CmdGIS.Enabled = True
            CmdStoreID.Enabled = True
            CmdNeo4j.Enabled = True
        Else
            CmdNeo4j.Enabled = False
            CmdGIS.Enabled = False
            CmdStoreID.Enabled = False
        End If
    End If

Exit_Run_Query:
    '
    '  now reopen the tables
    '
    Set tRstAssoc = CurrentDb.OpenRecordset("ZZ_SCRATCH_ASSOC", dbOpenDynaset)
    Set ZZ_SCRATCH_ASSOC.Form.Recordset = tRstAssoc
    '
    Set gRstPeople = CurrentDb.OpenRecordset("ZZ_SCRATCH_P_ASSOC", dbOpenDynaset)
    Set ZZ_SCRATCH_P_ASSOC.Form.Recordset = gRstPeople
    '
    '  close everything
    '
    Set rst = Nothing
    Set AssocQuery = Nothing
    Set AddressQuery = Nothing
    Set tRstDummy = Nothing
    Set cmdSQL = Nothing
    Exit Sub

Err_Run_Query:
    MsgBox Err.Description
    Resume Exit_Run_Query
    
End Sub

Private Sub CmdGIS_Click()
On Error GoTo Err_CmdGIS_Click
    '
    '  If it is a KML file, call the routine and exit
    '
    If ChkKML.Value Then
        Call writeKML
        Exit Sub
    End If
    
    Dim dlgSaveAs As FileDialog
    Dim tFileNum As Integer
    Dim tFileName As String, tFN As Variant, tFemale As String
    Dim tRstNode As DAO.Recordset
    Dim tStr As String, tC As String, ti As Integer, tPinyin As Boolean
    Dim tFileSystem, tGDF
    '
    '  This program will dump the results to a .gis file
    '
    If ZZ_SCRATCH_P_ASSOC.Form.Recordset.RecordCount = 0 Then
        MsgBox "There are no records to save."
        GoTo Exit_CmdGIS_Click
    End If
    '
    Dim tStream As ADODB.Stream
    Set tStream = New ADODB.Stream
    
    tPinyin = False
    If GISFrame.Value = 1 Then
        tStream.Charset = "utf-8"
        tCodeStr = "UTF8"
    ElseIf GISFrame.Value = 2 Then
        tStream.Charset = "ascii"
        tCodeStr = "ASCII"
        tPinyin = True
    Else
        tStream.Charset = "gb18030"
        tCodeStr = "GB18030"
    End If
    tStream.Mode = adModeReadWrite
    tStream.Type = adTypeText
    tStream.Open
    '
    '  next get a file
    '
    
    Set dlgSaveAs = Application.FileDialog(msoFileDialogSaveAs)

    dlgSaveAs.InitialFileName = "network_gis_" + tCodeStr + ".tab"
    If dlgSaveAs.Show = -1 Then
        '
        tFileName = ""
        For Each tFN In dlgSaveAs.SelectedItems
            tFileName = tFN
            If Not tFileName = "" Then
                Exit For
            End If
        Next
        If tFileName = "" Then
            MsgBox "Bad file Name."
            GoTo Exit_CmdGIS_Click
        Else
            '  make sure the file name has a txt extension
            If Len(tFileName) < 5 Then
                tFileName = tFileName + ".tab"
            ElseIf Not (LCase(Right(tFileName, 4)) = ".tab") Then
                tFileName = tFileName + ".tab"
            End If
        End If
        '
        '  write the file
        '
        'Name,NameChn,Female,IndexYear,AddrName,AddrChn,X,Y,xy_count,NodeDist
        '
        ' process the table
        '
        Set tRstNode = ZZ_SCRATCH_P_ASSOC.Form.Recordset
        tC = Chr(9) ' the tab
        '
        With tRstNode
            '
            ' write the header
            '
            If tPinyin Then
                tStr = "Name" + tC + "Female" + tC + "IndexYear" + tC + _
                    "AddrName" + tC + "X" + tC + "Y" + tC + "xy_count"
            Else
                tStr = "Name" + tC + "NameChn" + tC + "Female" + tC + "IndexYear" + tC + _
                    "AddrName" + tC + "AddrChn" + tC + "X" + tC + "Y" + tC + "xy_count"
            End If
            tStream.WriteText tStr, adWriteLine
            .MoveFirst
            Do While Not .EOF
                ' must guard against NULLs
                '
                If Trim(!c_name) = "" Then
                    tStr = "[?]" + tC
                Else
                    tStr = !c_name + tC
                End If
                
                If Not tPinyin Then
                    If Trim(!c_name_chn) = "" Then
                        tStr = tStr + "[?]" + tC
                    Else
                        tStr = tStr + !c_name_chn + tC
                    End If
                End If
                
                If IsNull(!c_sex) Then
                    tStr = tStr + "[?]" + tC
                Else
                    tStr = tStr + !c_sex + tC
                End If
                
                If IsNull(!c_index_year) Then
                    tStr = tStr + "-2000" + tC
                Else
                    tStr = tStr + Str(!c_index_year) + tC
                End If
                
                ' here guard against blanks as well
                
                If IsNull(!c_addr_name) Then
                    tStr = tStr + "[?]" + tC
                ElseIf Trim(!c_addr_name) = "" Then
                    tStr = tStr + "[?]" + tC
                Else
                    tStr = tStr + !c_addr_name + tC
                End If
                
                If Not tPinyin Then
                    If IsNull(!c_addr_chn) Then
                        tStr = tStr + "[?]" + tC
                    ElseIf Trim(!c_addr_chn) = "" Then
                        tStr = tStr + "[?]" + tC
                    Else
                        tStr = tStr + !c_addr_chn + tC
                    End If
                End If
                
                If IsNull(!x_coord) Then
                    tStr = tStr + "0" + tC
                Else
                    tStr = tStr + Str(!x_coord) + tC
                End If
                
                If IsNull(!y_coord) Then
                    tStr = tStr + "0" + tC
                Else
                    tStr = tStr + Str(!y_coord) + tC
                End If
                
                If IsNull(!xy_count) Then
                    tStr = tStr + "0"
                Else
                    tStr = tStr + Str(!xy_count)
                End If
                
                tStream.WriteText tStr, adWriteLine
                .MoveNext
            Loop
        End With
        ' now make sure all the data is copied to tStream
        tStream.Flush
        ' and write the stream to the file
        tStream.SaveToFile tFileName, adSaveCreateOverWrite
        '
    Else
        'The user pressed Cancel.
    End If

    Set tRstNode = Nothing
            
    tStream.Close
    Set tStream = Nothing
    'Set the object variable to Nothing.
    Set dlgSaveAs = Nothing
    
Exit_CmdGIS_Click:
    Exit Sub

Err_CmdGIS_Click:
    MsgBox Err.Description
    Resume Exit_CmdGIS_Click
    
End Sub


Private Sub CmdSaveAssociations_Click()
On Error GoTo Err_CmdSaveAssociations_Click
    '
    '  This program will store the current list of office IDs to a .txt file
    '
    Dim tStream As ADODB.Stream, tStreamNoBOM As ADODB.Stream
    Set tStream = New ADODB.Stream
        
    tStream.Charset = "utf-8"
    tStream.Mode = adModeReadWrite
    tStream.Type = adTypeText
    tStream.Open
    '
    Set tStreamNoBOM = New ADODB.Stream
    tStreamNoBOM.Type = adTypeBinary
    tStreamNoBOM.Open
    '
    '  next get a file
    '
    Dim dlgSaveAs As FileDialog
    Dim tFileNum As Integer
    Dim tFileName As String, tFN As Variant, tFemale As String
    Dim tRstIDs As DAO.Recordset
    Dim tStr As String, tTab As String, ti As Integer
    Dim tFileSystem, tGDF
        
    Set dlgSaveAs = Application.FileDialog(msoFileDialogSaveAs)
    
    dlgSaveAs.InitialFileName = "assoc_code_list.txt"
    If dlgSaveAs.Show = -1 Then
        '
        tFileName = ""
        For Each tFN In dlgSaveAs.SelectedItems
            tFileName = tFN
            If Not tFileName = "" Then
                Exit For
            End If
        Next
        If tFileName = "" Then
            MsgBox "Bad file Name."
            GoTo Exit_CmdSaveAssociations_Click
        Else
            '  make sure the file name has a txt extension
            If Len(tFileName) < 5 Then
                tFileName = tFileName + ".txt"
            ElseIf Not (LCase(Right(tFileName, 4)) = ".txt") Then
                tFileName = tFileName + ".txt"
            End If
        End If
        '
        '  write the file
        '
        ' process the table
        '
        tStr = "SELECT ZZ_ASSOC_CODE.c_assoc_code, ASSOC_CODES.c_assoc_desc, ASSOC_CODES.c_assoc_desc_chn " + _
            "FROM ZZ_ASSOC_CODE INNER JOIN ASSOC_CODES ON ZZ_ASSOC_CODE.c_assoc_code = ASSOC_CODES.c_assoc_code"

        Set tRstIDs = CurrentDb.OpenRecordset(tStr, dbOpenDynaset)
        '
        tTab = Chr(9)
        
        With tRstIDs
            
            .MoveFirst
            ' MsgBox "writing file"
            Do While Not .EOF
                '
                tStr = Str(!c_assoc_code) + tTab + !c_assoc_desc + tTab + !c_assoc_desc_chn
                tStream.WriteText tStr, adWriteLine
                .MoveNext
            Loop
        End With
        ' now make sure all the data is copied to tStream
        tStream.Flush
        tStream.Position = 3
        ' and write the stream to the file
        tStream.CopyTo tStreamNoBOM
        tStreamNoBOM.SaveToFile tFileName, adSaveCreateOverWrite
        '
    Else
        'The user pressed Cancel.
    End If
    
    Set tRstIDs = Nothing
                
    tStream.Close
    Set tStream = Nothing
    tStreamNoBOM.Close
    Set tStreamNoBOM = Nothing
    'Set the object variable to Nothing.
    Set dlgSaveAs = Nothing
    
Exit_CmdSaveAssociations_Click:
    Exit Sub

Err_CmdSaveAssociations_Click:
    MsgBox Err.Description
    Resume Exit_CmdSaveAssociations_Click

End Sub

Private Sub CmdToDynasty_Click()
    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim strToDynasty As String

    If gToDynasty = -1 Then
        strToDynasty = ""
    Else
        strToDynasty = Str(gToDynasty)
    End If
    
    stDocName = "frmPickDynasty"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , acDialog, strFromDynasty
    
    If CurrentProject.AllForms("frmPickDynasty").IsLoaded Then
        
        Forms!frmpickdynasty!frmDYNASTIES.Form!Dy_Code.SetFocus
        gToDynasty = Forms!frmpickdynasty!frmDYNASTIES.Form!Dy_Code.Value
        
        Forms!frmpickdynasty!frmDYNASTIES.Form!c_start.SetFocus
        gToDynastyBegin = Forms!frmpickdynasty!frmDYNASTIES.Form!c_start.Value
        
        Forms!frmpickdynasty!frmDYNASTIES.Form!c_end.SetFocus
        gToDynastyEnd = Forms!frmpickdynasty!frmDYNASTIES.Form!c_end.Value
        '
        ' check to see if we have a problem and reject selection if needed
        '
        If gFromDynasty > -1 Then
            If gFromDynastyBegin > gToDynastyEnd Then
                MsgBox "Warning:  There is a problem with chronology:  the 'From' Dynasty begins after the 'To' Dynasty ends!", vbExclamation
                gToDynasty = -1
                TxtToDynasty.Value = ""
                TxtToDynastyPY.Value = ""
            End If
        End If
        '
        '  value is OK
        '
        If gToDynasty > -1 Then
            Forms!frmpickdynasty!frmDYNASTIES.Form!c_dynasty.SetFocus
            TxtToDynastyPY.Value = Forms!frmpickdynasty!frmDYNASTIES.Form!c_dynasty.Value
            
            Forms!frmpickdynasty!frmDYNASTIES.Form!c_dynasty_chn.SetFocus
            TxtToDynasty.Value = Forms!frmpickdynasty!frmDYNASTIES.Form!c_dynasty_chn.Value
        End If
        
        DoCmd.Close acForm, stDocName
        '
        ' reset FromDynasty if necessary (-2 = all dynasties)
        '
        If gFromDynasty = -2 Then
            gFromDynasty = -1
            TxtFromDynasty.Value = ""
            TxtFromDynastyPY.Value = ""
        End If
        '
    End If

End Sub

Private Sub CmdUCINet_Click()
On Error GoTo Err_CmdUCINet_Click
    '
    '  This program will dump the results of the search to a .vna file
    '
    '  for the moment I'll just describe the format of the .vna file
    '
    '  *node data
    '  ID index_year sex x_coord y_coord nodedist
    '      ID = str(c_person_id)
    '      indexyear = c_index_year INT
    '      nodedist = c_node_dist INT
    '      sex = c_female > (F,M)
    '  *node properties
    '  ID color shape size shortlabel active
    '      color = red (1), orange (2), yellow (3), green (4), blue (5)
    '      shortlabel = c_name
    '      shape = 2
    '      active = TRUE
    '
    '  *tie data
    '  from to edgetype nodedist
    '      from = str(c_person_id)
    '      to = str(c_node_id)
    '      edgetype= c_link_type (K,N)
    '
    '  *tie properties
    '  from to color size active
    '      from = str(c_person_id)
    '      to = str(c_node_id)
    '      color = red (255), orange (26367), yellow (65535), green (32768), blue (16711680)
    '      size = 1-5 (the weight)
    '
    '  the central question is whether to do distance optimizations
    '
    '  first see if there are any records to process
    '
    If ZZ_SCRATCH_ASSOC.Form.Recordset.RecordCount = 0 Then
        MsgBox "There are no records to save."
        GoTo Exit_CmdUCINet_Click
    End If
    '
    If ZZ_SCRATCH_P_ASSOC.Form.Recordset.RecordCount = 0 Then
        MsgBox "There are no records to save."
        GoTo Exit_CmdUCINet_Click
    End If
    '
    '  next get a file
    '
    Dim dlgSaveAs As FileDialog
    Dim tFileNum As Integer
    Dim tFileName As String, tFN As Variant
    Dim tRstNode As DAO.Recordset, tRstAssocType As DAO.Recordset
    Dim tRstEdge As DAO.Recordset
    Dim tStr As String, tC As String, ti As Integer, tSearchStr As String
    Dim tColor(20) As String, tQuote As String
    Dim tFileSystem, tVNA
    
    Set dlgSaveAs = Application.FileDialog(msoFileDialogSaveAs)

    ' open the assoc type look-up table
    Set tRstAssocType = CurrentDb.OpenRecordset("ASSOC_CODE_TYPE_REL", dbOpenDynaset)

    'Use a With...End With block to reference the FileDialog object.
    With dlgSaveAs
        .InitialFileName = "network.vna"
        If .Show = -1 Then
            '
            tFileName = ""
            For Each tFN In .SelectedItems
                tFileName = tFN
                If Not tFileName = "" Then
                    Exit For
                End If
            Next
            If tFileName = "" Then
                MsgBox "Bad file Name."
                GoTo Exit_CmdUCINet_Click
            Else
                '  make sure the file name has a vna extension
                If Len(tFileName) < 5 Then
                    tFileName = tFileName + ".vna"
                ElseIf Not (LCase(Right(tFileName, 4)) = ".vna") Then
                    tFileName = tFileName + ".vna"
                End If
            End If
            '
            '  now process the file (second true removed to make ASCII)
            '
            Set tFileSystem = CreateObject("Scripting.FileSystemObject")
            Set tVNA = tFileSystem.CreateTextFile(tFileName, True)

            ' process the two tables
            '
            Set tRstEdge = CurrentDb.OpenRecordset("ZZ_SCRATCH_ASSOC", dbOpenDynaset)
            Set tRstNode = CurrentDb.OpenRecordset("ZZ_SCRATCH_P_ASSOC", dbOpenDynaset)
            tQuote = Chr(34) ' the quotation mark
            '
            ' first the nodes:  define the node data structure
            tVNA.WriteLine ("*node data")
            tVNA.WriteLine ("ID index_year sex x_coord y_coord")
            '
            With tRstNode
                .MoveFirst
                Do While Not .EOF
                    '  name = the ID of the person
                    tStr = Trim(Str(!c_person_id)) + " "
                    '
                    '  indexyear = c_index_year INT
                    If IsNull(!c_index_year) Then
                        tStr = tStr + "0 "
                    Else
                        tStr = tStr + Trim(Str(!c_index_year)) + " "
                    End If
                    '
                    '   sex = c_female > (F,M)
                    If !c_sex = "F" Then
                        tStr = tStr + tQuote + "F" + tQuote + " "
                    Else
                        tStr = tStr + tQuote + "M" + tQuote + " "
                    End If
                    '
                    '   x_coord
                    If IsNull(!x_coord) Then
                        tStr = tStr + "0 "
                    Else
                        tStr = tStr + Trim(Str(!x_coord)) + " "
                    End If
                    '
                    '   y_coord
                    If IsNull(!y_coord) Then
                        tStr = tStr + "0 "
                    Else
                        tStr = tStr + Trim(Str(!y_coord)) + " "
                    End If
                    '
                    tVNA.WriteLine (tStr)
                    .MoveNext
                Loop
            End With
            '
            ' now the node properties
            '
            ' Note:  ACTIVE removed as a property (MAF 2018/07/22)
            '
            tVNA.WriteLine ("*node properties")
            tVNA.WriteLine ("ID shape size shortlabel")
            '
            With tRstNode
                .MoveFirst
                Do While Not .EOF
                    '  ID = the ID of the person
                    tStr = Trim(Str(!c_person_id)) + " "
                    '
                    '  shape = 2? / size = 1?
                    tStr = tStr + "2 1 "
                    '
                    '  shortlabel (+ Active = TRUE removed)
                    If IsNull(!c_name) Then
                        tStr = tStr + "[Missing]"
                    Else
                        tStr = tStr + tQuote + !c_name + tQuote
                    End If
                    tVNA.WriteLine (tStr)
                    .MoveNext
                Loop
            End With
            '
            ' now the edges:  define the record structure
            '
            tStr = "from to " + tQuote + "EdgeWeight" + tQuote + " " + tQuote + "edgedesc" + tQuote
            tVNA.WriteLine ("*tie data")
            tVNA.WriteLine (tStr)
            '
            '  For the moment, I am not combining parallel edges
            '
            With tRstEdge
                .MoveFirst
                Do While Not .EOF
                    '
                    '   From = str(c_person_id) for node1
                    tStr = Trim(Str(!c_person_id)) + " "
                    '
                    '   to = str(c_assoc_id) for node2
                    tStr = tStr + Trim(Str(!c_assoc_id)) + " 1 "
                    '
                    '   edgedesc
                    '
                    tStr = tStr + tQuote + Trim(!c_assoc_desc) + tQuote
                    '
                    tVNA.WriteLine (tStr)
                    .MoveNext
                Loop
            End With
            '
            ' now the edges properties
            '
            'tVNA.WriteLine ("*tie properties")
            'tVNA.WriteLine ("from to color size active")

            'With tRstEdge
                '.MoveFirst
                'Do While Not .EOF
                    '
                    '   from = str(c_person_id) for node1
                    'tStr = Trim(Str(!c_person_id)) + " "
                    '
                    '   to = str(c_node_id) for node2
                    'tStr = tStr + Trim(Str(!c_node_id)) + " 1 "
                    '
                    '   color = black (1), blue (2), green (3), yellow (4), orange (5)
                    'tStr = tStr + tColor(!c_edge_dist)
                    '
                    '   size = 1?  active = TRUE
                    'tStr = tStr + "1 TRUE"
                    '
                    'tVNA.WriteLine (tStr)
                    '.MoveNext
                'Loop
            'End With
            '
            tVNA.Close
            '
            Set tRstNode = Nothing
            Set tRstEdge = Nothing
            Set tVNA = Nothing
            Set tFileSystem = Nothing
            Set tRstAssocType = Nothing
        Else
            'The user pressed Cancel.
        End If
    End With

    'Set the object variable to Nothing.
    Set dlgSaveAs = Nothing
    
Exit_CmdUCINet_Click:
    Exit Sub

Err_CmdUCINet_Click:
    MsgBox Err.Description
    Resume Exit_CmdUCINet_Click
    
End Sub

Private Sub Form_Open(Cancel As Integer)
    Dim cmdSQL As ADODB.Command, tRecDeleted As Variant
    Dim tRstAssocCode As DAO.Recordset, tRstDummy As DAO.Recordset
    
    Set cmdSQL = New ADODB.Command
    cmdSQL.ActiveConnection = CurrentProject.Connection
    cmdSQL.CommandType = adCmdText
    '
    '  to clear the tables, briefly close and then delete records
    '
    Set tRstAssocCode = ZZ_SCRATCH_ASSOC.Form.Recordset
    Set tRstDummy = CurrentDb.OpenRecordset("Z_SCRATCH_DUMMY_AC", dbOpenDynaset)
    Set ZZ_SCRATCH_ASSOC.Form.Recordset = tRstDummy
    tRstAssocCode.Close
    '
    cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_ASSOC"
    cmdSQL.Execute tRecDeleted
    '
    '  now reopen
    '
    Set tRstAssocCode = CurrentDb.OpenRecordset("ZZ_SCRATCH_ASSOC", dbOpenDynaset)
    Set ZZ_SCRATCH_ASSOC.Form.Recordset = tRstAssocCode
    '
    Set tRstAssocCode = ZZ_SCRATCH_P_ASSOC.Form.Recordset
    Set tRstDummy = CurrentDb.OpenRecordset("Z_SCRATCH_DUMMY_AP", dbOpenDynaset)
    Set ZZ_SCRATCH_P_ASSOC.Form.Recordset = tRstDummy
    tRstAssocCode.Close
    '
    cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_P_ASSOC"
    cmdSQL.Execute tRecDeleted
    '
    '  now reopen
    '
    Set tRstAssocCode = CurrentDb.OpenRecordset("ZZ_SCRATCH_P_ASSOC", dbOpenDynaset)
    Set ZZ_SCRATCH_P_ASSOC.Form.Recordset = tRstAssocCode
    
    '  first determine the language
    gLCID = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    If gLCID = 2052 Or gLCID = 3076 Then      ' 2052 = PRC, 3076 = Hong Kong
        gDisplayLanguage = "S"
    ElseIf gLCID = 4100 Or gLCID = 1028 Then  ' 4100 = Singapore, 1028 = Taiwan
        gDisplayLanguage = "T"
        Call changeDisplayLanguage
    Else
        gDisplayLanguage = "E"
        Call changeDisplayLanguage
    End If
    
    '  set the index year and dynasty default values
    
    gFromStr = "-200"
    gToStr = "1911"
    TxtFromYear.Value = -200
    TxtToYear.Value = 1911
    gFromDynasty = -1
    gToDynasty = -1
    gUseIndexYears = False
    gUseDynasties = False
End Sub

Private Sub CmdPajek_Click()
On Error GoTo Err_CmdPajek_Click
    '
    '  This program will dump the results of the search to a .net file
    '
    '  for the moment I'll just describe the format of the .gdf file
    '
    '  *Vertices NUM
    '  ID label "box" ic [color] bc [color]
    '      ID = str(c_person_id)
    '      label = c_name_chn
    '      color = red (1), orange (2), yellow (3), green (4), blue (5)
    '
    '  *Edges
    '  node1 node2 1 l "label"
    '      node1 = str(c_person_id) for node1
    '      node2 = str(c_node_id) for node2
    '      color = red (1), orange (2), yellow (3), green (4), blue (5)
    '      label = c_link_desc
    '
    '
    '  first see if there are any records to process
    '
    If ZZ_SCRATCH_ASSOC.Form.Recordset.RecordCount = 0 Then
        MsgBox "There are no records to save."
        GoTo Exit_CmdPajek_Click
    End If
    '
    If ZZ_SCRATCH_P_ASSOC.Form.Recordset.RecordCount = 0 Then
        MsgBox "There are no records to save."
        GoTo Exit_CmdPajek_Click
    End If
    '
    '  next get a file
    '
    Dim dlgSaveAs As FileDialog
    Dim tFileNum As Integer
    Dim tFileName As String, tFN As Variant
    Dim tRstNode As DAO.Recordset, tRstNodeList As DAO.Recordset
    Dim tRstEdge As DAO.Recordset, tRstEdgeList As DAO.Recordset
    Dim tStr As String, tC As String, ti As Integer, tQuote As String, tFindStr As String, tPinyin As Boolean
    Dim tColor(20) As String, tStrNode1 As String, tStrNode2 As String, tCodeStr As String, tRecDeleted As Long
    
    '  to write to a UTF-8 file, use the ADO stream object
    Dim tStream As ADODB.Stream
    Set tStream = New ADODB.Stream
    
    tPinyin = False
    If CodeFrame.Value = 1 Then
        tStream.Charset = "utf-8"
        tCodeStr = "UTF8.net"
    ElseIf CodeFrame.Value = 2 Then
        tStream.Charset = "big5"
        tCodeStr = "BIG5.net"
    ElseIf CodeFrame.Value = 3 Then
        tStream.Charset = "gb18030"
        tCodeStr = "GB18030.net"
    Else
        tStream.Charset = "ascii"
        tCodeStr = "ascii.net"
        tPinyin = True
    End If
    tStream.Mode = adModeReadWrite
    tStream.Type = adTypeText
    tStream.Open
    
    
    Set dlgSaveAs = Application.FileDialog(msoFileDialogSaveAs)


    'Use a With...End With block to reference the FileDialog object.
    With dlgSaveAs
        .InitialFileName = "network_" + tCodeStr
        If .Show = -1 Then
            '
            tFileName = ""
            For Each tFN In .SelectedItems
                tFileName = tFN
                If Not tFileName = "" Then
                    Exit For
                End If
            Next
            If tFileName = "" Then
                MsgBox "Bad file Name."
                GoTo Exit_CmdPajek_Click
            Else
                '  make sure the file name has a net extension
                If Len(tFileName) < 5 Then
                    tFileName = tFileName + ".net"
                ElseIf Not (LCase(Right(tFileName, 4)) = ".net") Then
                    tFileName = tFileName + ".net"
                End If
            End If
            '
            '  zap and open the scratch file
            '
            Dim cmdSQL As ADODB.Command
            Set cmdSQL = New ADODB.Command
            cmdSQL.ActiveConnection = CurrentProject.Connection
            cmdSQL.CommandType = adCmdText
            '
            cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_PAJEK"
            cmdSQL.Execute tRecDeleted
            '
            Set tRstNodeList = CurrentDb.OpenRecordset("ZZ_SCRATCH_PAJEK", dbOpenTable)
            tRstNodeList.Index = "c_ID"
            '
            cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_PAJEK_EDGE"
            cmdSQL.Execute tRecDeleted
            '
            ' set the Quote delimiter
            '
            tQuote = Chr(34)
            '
            ' define the colors for the nodes
            '
            tColor(1) = "Black"
            tColor(2) = "Blue"
            tColor(3) = "Green"
            tColor(4) = "Yellow"
            tColor(5) = "Orange"
            For ti = 6 To 20
                tColor(ti) = "Red"
            Next
            '
            ' process the two tables
            '
            Set tRstEdge = ZZ_SCRATCH_ASSOC.Form.Recordset
            Set tRstNode = ZZ_SCRATCH_P_ASSOC.Form.Recordset
            tC = Chr(44) ' the comma
            '
            ' first the nodes:  define the record structure
            '
            tRstNode.MoveFirst
            tStr = "*Vertices " + Trim(Str(tRstNode.RecordCount))
            tStream.WriteText tStr, adWriteLine
            '
            ti = 1
            With tRstNode
                .MoveFirst
                Do While Not .EOF
                    tStream.WriteText Trim(Str(ti)) + " "
                    '
                    If IsNull(!c_name_chn) Then
                        If !c_name = "" Or Left(!c_name, 12) = "**BAD DATA**" Then
                            tStream.WriteText Chr(34)
                            tStream.WriteText "Error-" + Trim(Str(!c_person_id))
                            tStream.WriteText Chr(34)
                            tStream.WriteText " box ", adWriteLine
                        Else
                            tStream.WriteText Chr(34)
                            tStream.WriteText !c_name
                            tStream.WriteText Chr(34)
                            tStream.WriteText " box ", adWriteLine
                        End If
                    Else
                        If !c_name_chn = "" Then
                            If !c_name = "" Or Left(!c_name, 12) = "**BAD DATA**" Then
                                tStream.WriteText Chr(34)
                                tStream.WriteText "Error-" + Trim(Str(!c_person_id))
                                tStream.WriteText Chr(34)
                                tStream.WriteText " box ", adWriteLine
                            Else
                                tStream.WriteText Chr(34)
                                tStream.WriteText !c_name
                                tStream.WriteText Chr(34)
                                tStream.WriteText " box ", adWriteLine
                            End If
                        Else
                            tStream.WriteText Chr(34)
                            If tPinyin Then
                                tStream.WriteText !c_name
                            Else
                                tStream.WriteText !c_name_chn
                            End If
                            If ChkIDs.Value Then
                                tStream.WriteText " (" + Trim(Str(!c_person_id)) + ")"
                            End If
                            tStream.WriteText Chr(34)
                            tStream.WriteText " box ", adWriteLine
                        End If
                    End If
                    '
                    '  add the node to the list
                    '
                    tRstNodeList.AddNew
                    tRstNodeList!c_v_num = Str(ti)
                    tRstNodeList!c_ID = !c_person_id
                    tRstNodeList.Update
                    '
                    .MoveNext
                    ti = ti + 1
                Loop
            End With
            tRstNodeList.Close
            '
            ' now the edges:  define the record structure
            '
            tStream.WriteText "*Edges", adWriteLine
            '
            '  first aggregate the data to a temporary table (use the edge weight to count the number of records)
            '
            cmdSQL.CommandText = "SELECT ZZ_SCRATCH_ASSOC.c_person_id, ZZ_SCRATCH_ASSOC.c_assoc_id, " + _
                "Count(ZZ_SCRATCH_ASSOC.c_assoc_code) AS CountOfc_assoc_code, " + _
                "Sum(ZZ_SCRATCH_ASSOC.c_assoc_count) AS SumOfc_assoc_count INTO tmp_pajek_edge " + _
                "FROM ZZ_SCRATCH_ASSOC GROUP BY ZZ_SCRATCH_ASSOC.c_person_id, ZZ_SCRATCH_ASSOC.c_assoc_id"
            cmdSQL.Execute tRecDeleted
            '
            '  now join to the node IDs and copy to the edge table
            '
            cmdSQL.CommandText = "INSERT INTO ZZ_SCRATCH_PAJEK_EDGE ( c_node_1, c_node_2, c_edge_weight, c_edge_count, " + _
                "c_node_1_str, c_node_2_str ) " + _
                "SELECT Val([ZZ_SCRATCH_PAJEK].[c_v_num]) AS c_node_1, Val(ZZ_SCRATCH_PAJEK_1.c_v_num) AS c_node_2, " + _
                "tmp_pajek_edge.CountOfc_assoc_code, tmp_pajek_edge.SumOfc_assoc_count, [ZZ_SCRATCH_PAJEK].[c_v_num], " + _
                "[ZZ_SCRATCH_PAJEK_1].[c_v_num] " + _
                "FROM ZZ_SCRATCH_PAJEK AS ZZ_SCRATCH_PAJEK_1 INNER JOIN (ZZ_SCRATCH_PAJEK INNER JOIN " + _
                "tmp_pajek_edge ON ZZ_SCRATCH_PAJEK.c_ID = tmp_pajek_edge.c_person_id) " + _
                "ON ZZ_SCRATCH_PAJEK_1.c_ID = tmp_pajek_edge.c_assoc_id"
            cmdSQL.Execute tRecDeleted
            '
            cmdSQL.CommandText = "DROP TABLE tmp_pajek_edge"
            cmdSQL.Execute tRecDeleted
            '
            '  now fill in the edge description.
            '
            If tPinyin Then
                tQueryStr = "UPDATE ((ZZ_SCRATCH_PAJEK_EDGE INNER JOIN ZZ_SCRATCH_PAJEK " + _
                    "ON ZZ_SCRATCH_PAJEK_EDGE.c_node_1_str = ZZ_SCRATCH_PAJEK.c_v_num) INNER JOIN " + _
                    "ZZ_SCRATCH_PAJEK AS ZZ_SCRATCH_PAJEK_1 " + _
                    "ON ZZ_SCRATCH_PAJEK_EDGE.c_node_2_str = ZZ_SCRATCH_PAJEK_1.c_v_num) INNER JOIN " + _
                    "ZZ_SCRATCH_ASSOC ON (ZZ_SCRATCH_ASSOC.c_assoc_id = ZZ_SCRATCH_PAJEK_1.c_ID) " + _
                    "AND (ZZ_SCRATCH_PAJEK.c_ID = ZZ_SCRATCH_ASSOC.c_person_id) " + _
                    "SET ZZ_SCRATCH_PAJEK_EDGE.c_edge_desc = [ZZ_SCRATCH_ASSOC].[c_assoc_desc] " + _
                    "WHERE (((ZZ_SCRATCH_PAJEK_EDGE.c_edge_weight)=1))"
            Else
                tQueryStr = "UPDATE ((ZZ_SCRATCH_PAJEK_EDGE INNER JOIN ZZ_SCRATCH_PAJEK " + _
                    "ON ZZ_SCRATCH_PAJEK_EDGE.c_node_1_str = ZZ_SCRATCH_PAJEK.c_v_num) INNER JOIN " + _
                    "ZZ_SCRATCH_PAJEK AS ZZ_SCRATCH_PAJEK_1 " + _
                    "ON ZZ_SCRATCH_PAJEK_EDGE.c_node_2_str = ZZ_SCRATCH_PAJEK_1.c_v_num) INNER JOIN " + _
                    "ZZ_SCRATCH_ASSOC ON (ZZ_SCRATCH_ASSOC.c_assoc_id = ZZ_SCRATCH_PAJEK_1.c_ID) " + _
                    "AND (ZZ_SCRATCH_PAJEK.c_ID = ZZ_SCRATCH_ASSOC.c_person_id) " + _
                    "SET ZZ_SCRATCH_PAJEK_EDGE.c_edge_desc = [ZZ_SCRATCH_ASSOC].[c_assoc_desc] + ' ' + " + _
                    "[ZZ_SCRATCH_ASSOC].[c_assoc_desc_chn] WHERE (((ZZ_SCRATCH_PAJEK_EDGE.c_edge_weight)=1))"
            End If
            
            cmdSQL.CommandText = tQueryStr
            cmdSQL.Execute tRecDeleted
            
            tQueryStr = "UPDATE ZZ_SCRATCH_PAJEK_EDGE SET ZZ_SCRATCH_PAJEK_EDGE.c_edge_desc = " + _
                "'Parallel Edges merged' WHERE (((ZZ_SCRATCH_PAJEK_EDGE.c_edge_weight)>1))"

            cmdSQL.CommandText = tQueryStr
            cmdSQL.Execute tRecDeleted
            '
            '  open the table
            '
            Set tRstEdgeList = CurrentDb.OpenRecordset("ZZ_SCRATCH_PAJEK_EDGE", dbOpenDynaset)

            With tRstEdgeList
                .MoveFirst
                Do While Not .EOF
                    tStr = !c_node_1_str + " " + !c_node_2_str
                    '
                    ' now get the weight
                    '
                    If !c_edge_count < 6 Then
                        tStr = tStr + " " + Trim(Str(!c_edge_count)) + " "
                    Else
                        tStr = tStr + " 5 "
                    End If
                    '
                    ' now get the label
                    '
                    tStr = tStr + "l " + tQuote
                    If !c_edge_weight = 1 Then
                        tStr = tStr + !c_edge_desc + tQuote + " "
                    Else
                        tStr = tStr + Trim(Str(!c_edge_count)) + " relations merged" + tQuote + " "
                        '
                    End If
                            
                    tStream.WriteText tStr, adWriteLine
                    '
                    .MoveNext
                Loop
            End With
            '
            ' now make sure all the data is copied to tStream
            tStream.Flush
            ' and write the stream to the file
            tStream.SaveToFile tFileName, adSaveCreateOverWrite
            '
            tStream.Close
            Set tStream = Nothing
            '
            Set tRstNode = Nothing
            Set tRstEdge = Nothing
            Set tRstNodeList = Nothing
            Set tRstEdgeList = Nothing
        Else
            'The user pressed Cancel.
        End If
    End With

    'Set the object variable to Nothing.
    Set dlgSaveAs = Nothing
    
Exit_CmdPajek_Click:
    Exit Sub

Err_CmdPajek_Click:
    MsgBox Err.Description
    Resume Exit_CmdPajek_Click
    
End Sub
Private Sub guess_write()
    '
    '  This program will dump the results of the search to a .gdf file
    '
    '  for the moment I'll just describe the format of the .gdf file
    '
    '  nodedef> name, color, label, labelvisible, style, pinyin VARCHAR(50), nodedist INT
    '      name = str(c_person_id)
    '      color = red (1), orange (2), yellow (3), green (4), blue (5)
    '      label = c_name_chn
    '      style = 4 (text inside a rectangle)
    '      pinyin = c_name
    '      nodedist = c_node_dist INT
    '      indexyear = c_index_year INT
    '      sex = c_female > (F,M)
    '
    '  edgedef> node1, node2, color, label, labelvisible, edge_desc VARCHAR(50)
    '      node1 = str(c_person_id) for node1
    '      node2 = str(c_node_id) for node2
    '      color = red (1), orange (2), yellow (3), green (4), blue (5)
    '      label = c_link_chn
    '      edge_desc = c_link_desc
    '
    '  the central question is whether to do distance optimizations
    '
    '  first see if there are any records to process
    '
    If ZZ_SCRATCH_ASSOC.Form.Recordset.RecordCount = 0 Then
        MsgBox "There are no records to save."
        GoTo Exit_CmdGUESS_Click
    End If
    '
    '  next get a file
    '
    Dim dlgSaveAs As FileDialog
    Dim tFileNum As Integer
    Dim tFileName As String, tFN As Variant
    Dim tRstNode As DAO.Recordset
    Dim tRstEdge As DAO.Recordset
    Dim tStr As String, tC As String, ti As Integer
    Dim tColor(50) As String, tMetricSum As Integer
    Dim tFileSystem, tGDF
    
    Set dlgSaveAs = Application.FileDialog(msoFileDialogSaveAs)


    'Use a With...End With block to reference the FileDialog object.
    With dlgSaveAs
        .InitialFileName = "default.gdf"
        If .Show = -1 Then
            '
            tFileName = ""
            For Each tFN In .SelectedItems
                tFileName = tFN
                If Not tFileName = "" Then
                    Exit For
                End If
            Next
            If tFileName = "" Then
                MsgBox "Bad file Name."
                GoTo Exit_CmdGUESS_Click
            End If
            '
            '  now process the file (second true removed to make ASCII)
            '
            Set tFileSystem = CreateObject("Scripting.FileSystemObject")
            Set tGDF = tFileSystem.CreateTextFile(tFileName, True, True)

            ' define the colors for the nodes
            '
            tColor(1) = "white"
            tColor(2) = "blue"
            tColor(3) = "green"
            tColor(4) = "yellow"
            tColor(5) = "orange"
            For ti = 6 To 50
                tColor(ti) = "red"
            Next
            '
            ' process the two tables
            '
            Set tRstEdge = ZZ_SCRATCH_ASSOC.Form.Recordset
            Set tRstNode = ZZ_SCRATCH_P_ASSOC.Form.Recordset
            tC = Chr(44) ' the comma
            '
            ' first the nodes:  define the record structure
            tStr = "nodedef> name" + tC + "color" + tC + "label" + tC + "labelvisible"
            tStr = tStr + tC + "style" + tC + "pinyin VARCHAR(50)"
            tStr = tStr + tC + "indexyear INT" + tC + "sex VARCHAR(1)"
            tGDF.WriteLine (tStr)
            '
            With tRstNode
                .MoveFirst
                Do While Not .EOF
                    tStr = Trim(Str(!c_person_id)) + tC
                    '  name = the ID of the person
                    '
                    tStr = tStr + tColor(1) + tC
                    '
                    If IsNull(!c_name_chn) Then
                        tStr = tStr + tC
                    Else
                        tStr = tStr + !c_name_chn + tC
                    End If
                    '  label
                    tStr = tStr + "true" + tC + "4" + tC
                    '  labelvisible = true, style = 4 (text inside a rectangle)
                    If IsNull(!c_name) Then
                        tStr = tStr + tC
                    Else
                        tStr = tStr + !c_name + tC
                    End If
                    '  pinyin = c_name
                    '
                    If IsNull(!c_index_year) Then
                        tStr = tStr + "-2000" + tC
                    Else
                        tStr = tStr + Trim(Str(!c_index_year)) + tC
                    End If
                    '  indexyear = c_index_year INT
                    If Not IsNull(!c_sex) Then
                        tStr = tStr + !c_sex
                    End If
                    tGDF.WriteLine (tStr)
                    .MoveNext
                Loop
            End With
            '
            ' now the edges:  define the record structure
            tStr = "edgedef> node1" + tC + "node2" + tC + "color" + tC + "label"
            tGDF.WriteLine (tStr)

            With tRstEdge
                .MoveFirst
                Do While Not .EOF
                    tStr = Trim(Str(!c_person_id)) + tC
                    '   node1 = str(c_person_id) for node1
                    tStr = tStr + Trim(Str(!c_assoc_id)) + tC
                    '   node2 = str(c_node_id) for node2
                    '
                    tStr = tStr + tColor(1) + tC
                    '  color = white (1), blue (2), green (3), yellow (4), orange (5)
                    '
                    If IsNull(!c_assoc_desc) Then
                        tStr = tStr + tC
                    Else
                        tStr = tStr + !c_assoc_desc
                    End If
                    '   label = the association
                    '
                    tGDF.WriteLine (tStr)
                    .MoveNext
                Loop
            End With
            '
            tGDF.Close
            '
            Set tRstNode = Nothing
            Set tRstEdge = Nothing
            Set tGDF = Nothing
            Set tFileSystem = Nothing
        Else
            'The user pressed Cancel.
        End If
    End With

    'Set the object variable to Nothing.
    Set dlgSaveAs = Nothing
    
Exit_CmdGUESS_Click:
    Exit Sub

Err_CmdGUESS_Click:
    MsgBox Err.Description
    Resume Exit_CmdGUESS_Click
    

End Sub

Private Sub CmdFanti_Click()
On Error GoTo Err_CmdFanti_Click

    If gDisplayLanguage = "T" Then
        gDisplayLanguage = "E"
    Else
        gDisplayLanguage = "T"
    End If

    Call changeDisplayLanguage

Exit_CmdFanti_Click:
    Exit Sub

Err_CmdFanti_Click:
    MsgBox Err.Description
    Resume Exit_CmdFanti_Click
    
End Sub
Private Sub CmdJianti_Click()
On Error GoTo Err_CmdJianti_Click

    If gDisplayLanguage = "S" Then
        gDisplayLanguage = "E"
    Else
        gDisplayLanguage = "S"
    End If

    Call changeDisplayLanguage

Exit_CmdJianti_Click:
    Exit Sub

Err_CmdJianti_Click:
    MsgBox Err.Description
    Resume Exit_CmdJianti_Click
    
End Sub

Public Sub changeDisplayLanguage()
    Dim tLabelLanguage(3, 37) As String, tLang As Integer
    
    Dim tRstLabelList As DAO.Recordset, ti As Integer
    
    Set tRstLabelList = CurrentDb.OpenRecordset("FormLabels", dbOpenTable)
    
    tRstLabelList.Index = "label"
    
    gLabelsOK = False
    With tRstLabelList
        .MoveFirst
        ti = 1
        
        Do While ti < 37 And Not .EOF
            If !c_form = "LAA" Then
                gLabelsOK = True
                If ti <> !c_label_id Then
                    MsgBox "Uh oh:  mismatched label table"
                    gLabelsOK = False
                    Exit Do
                End If
                tLabelLanguage(1, ti) = !c_english
                tLabelLanguage(2, ti) = !c_fanti
                tLabelLanguage(3, ti) = !c_jianti
                ti = ti + 1
            End If
            .MoveNext
        Loop
    End With
    ' tRstLabelList.Close
    Set tRstLabelList = Nothing
    
    If gLabelsOK Then
        If gDisplayLanguage = "E" Then
            tLang = 1
        ElseIf gDisplayLanguage = "T" Then
            tLang = 2
        Else
            tLang = 3
        End If
        '
        '  now comes the basic routine
        '
        Me.LblFrom.Caption = tLabelLanguage(tLang, 1)
        Me.LblTo.Caption = tLabelLanguage(tLang, 2)
        Me.LblType.Caption = tLabelLanguage(tLang, 3)
        Me.CmdPickAssoc.Caption = tLabelLanguage(tLang, 4)
        Me.CmdQuery.Caption = tLabelLanguage(tLang, 5)
        Me.CmdGIS.Caption = tLabelLanguage(tLang, 6)
        Me.CmdPajek.Caption = tLabelLanguage(tLang, 7)
        Me.CmdFanti.Caption = tLabelLanguage(tLang, 8)
        Me.CmdJianti.Caption = tLabelLanguage(tLang, 9)
        Me.PageAssoc.Caption = tLabelLanguage(tLang, 10)
        Me.PagePeople.Caption = tLabelLanguage(tLang, 11)
        'Me.LblChkIndexYears.Caption = tLabelLanguage(tLang, 12)
        Me.CmdSelectPlace.Caption = tLabelLanguage(tLang, 13)
        Me.CmdImportPlaces.Caption = tLabelLanguage(tLang, 14)
        Me.CmdAllPlaces.Caption = tLabelLanguage(tLang, 15)
        Me.LblIDs.Caption = tLabelLanguage(tLang, 16)
        Me.LblDisplay.Caption = tLabelLanguage(tLang, 17)
        Me.CmdHelp.Caption = tLabelLanguage(tLang, 18)
        Me.LblXYRef.Caption = tLabelLanguage(tLang, 19)
        Me.LblNarrow.Caption = tLabelLanguage(tLang, 20)
        Me.LblBroad.Caption = tLabelLanguage(tLang, 21)
        Me.CmdStoreID.Caption = tLabelLanguage(tLang, 22)
        Me.CmdUCINet.Caption = tLabelLanguage(tLang, 23)
        Me.LblChkSubUnits.Caption = tLabelLanguage(tLang, 24)
        Me.CmdGephi.Caption = tLabelLanguage(tLang, 25)
        Me.LblDynasties.Caption = tLabelLanguage(tLang, 26)
        Me.CmdFromDynasty.Caption = tLabelLanguage(tLang, 27)
        Me.CmdToDynasty.Caption = tLabelLanguage(tLang, 28)
        Me.CmdAllDynasties.Caption = tLabelLanguage(tLang, 29)
        Me.LblIndexYears.Caption = tLabelLanguage(tLang, 30)
        Me.LblOptNoDates.Caption = tLabelLanguage(tLang, 31)
        Me.LblOptIndexYears.Caption = tLabelLanguage(tLang, 32)
        Me.LblOptDynasties.Caption = tLabelLanguage(tLang, 33)
        Me.CmdNeo4j.Caption = tLabelLanguage(tLang, 34)
        Me.CmdImportAssociations.Caption = tLabelLanguage(tLang, 35)
        Me.CmdSaveAssociations.Caption = tLabelLanguage(tLang, 36)
    End If
    
End Sub
Private Sub CmdSelectPlace_Click()
On Error GoTo Err_CmdSelectPlace_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim strADDR As String
    Dim cmdSQL As ADODB.Command
                
    Set cmdSQL = New ADODB.Command
    cmdSQL.ActiveConnection = CurrentProject.Connection
    cmdSQL.CommandType = adCmdText

    TxtAddrID.Visible = True
    TxtAddrID.SetFocus
    strADDR = TxtAddrID.TEXT

    stDocName = "frmPickAddresses_multi"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , acDialog, strADDR
    
    If CurrentProject.AllForms("frmPickAddresses_multi").IsLoaded Then
        Dim tAddrID As Long, tRstAddr As DAO.Recordset
        Dim strADDR_CHN As String, strADDR_PY As String
                           
        gUseADDRID = True
        CmdAllPlaces.Enabled = True
        ChkXYRef.Enabled = True
        ChkSubUnits.Enabled = True
        FrameXY.Enabled = True
        
        Forms!frmPickAddresses_multi.Form!TxtAddrFilter.Visible = True
        Forms!frmPickAddresses_multi.Form!TxtAddrFilter.SetFocus
        If Forms!frmPickAddresses_multi.Form!TxtAddrFilter.Value Then
            '
            TxtAddrID.Value = 0
            strADDR_PY = Forms!frmPickAddresses_multi.Form!TxtFilterPY
            strADDR_CHN = Forms!frmPickAddresses_multi.Form!TxtFilterChn
            
            If strADDR_CHN = "" Then
                TxtPlaceChn.Value = "[[Filter]]"
                TxtPlace.Value = "[[" + strADDR_PY + "]]"
            Else
                TxtPlaceChn.Value = "[[" + strADDR_CHN + "]]"
                TxtPlace.Value = "[[Filter]]"
            End If
        Else
            Forms!frmPickAddresses_multi.Form!TxtSelectCount.Visible = True
            Forms!frmPickAddresses_multi.Form!TxtSelectCount.SetFocus
            If Forms!frmPickAddresses_multi.Form!TxtSelectCount.Value > 1 Then
                TxtPlaceChn.Value = "[[" + ChrW(22810) + ChrW(36984) + "]]"
                TxtPlace.Value = "[[Multi-Select]]"
                TxtAddrID.Value = 0
            Else
                '  only one record in ZZ_ADDRESSES: get its field values
                '
                Set tRstAddr = CurrentDb.OpenRecordset("ZZ_ADDRESSES", dbOpenDynaset)
                tRstAddr.MoveFirst
                'MsgBox "Checking zz_addresses:  no records"
                TxtAddrID.Value = tRstAddr!c_addr_id
                TxtPlaceChn.Value = tRstAddr!c_name_chn
                TxtPlace.Value = tRstAddr!c_name
                tRstAddr.Close
                Set tRstAddr = Nothing
           End If
        End If
        '
        ' now copy the records
        '
        cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_ADDR_LIST"
        cmdSQL.Execute tRecDeleted
            
        cmdSQL.CommandText = "INSERT INTO ZZ_SCRATCH_ADDR_LIST ( c_addr_id ) SELECT DISTINCT " + _
            "ZZ_ADDRESSES.c_addr_id FROM ZZ_ADDRESSES"
        cmdSQL.Execute tRecDeleted
    End If
            
    DoCmd.Close acForm, stDocName
    CmdSelectPlace.SetFocus
    TxtAddrID.Visible = False

Exit_CmdSelectPlace_Click:
    Exit Sub

Err_CmdSelectPlace_Click:
    MsgBox Err.Description
    Resume Exit_CmdSelectPlace_Click
    
End Sub
Private Sub CmdAllPlaces_Click()
On Error GoTo Err_CmdAllPlaces_Click

        TxtAddrID.Value = -1
                
        TxtPlaceChn.Value = ""
        TxtPlace.Value = ""
        gUseADDRID = False
        ChkXYRef.Enabled = False
        ChkSubUnits.Enabled = False
        FrameXY.Enabled = False
     
Exit_CmdAllPlaces_Click:
    Exit Sub

Err_CmdAllPlaces_Click:
    MsgBox Err.Description
    Resume Exit_CmdAllPlaces_Click
  
End Sub
Private Sub CmdImportPlaces_Click()
    On Error GoTo Err_CmdImportPlaces_Click
    
    Dim stDocName As String, tRstAddresses As DAO.Recordset, tRstImportPlaces As DAO.Recordset
    Dim stLinkCriteria As String
    Dim tString As String, tAddrID As Long, ti As Integer, tStrID As String, tLen As Integer, tQuit As Boolean

    Dim dlgSaveAs As FileDialog
    Dim tFileNum As Integer
    Dim tFileName As String, tFN As Variant
    Dim tFileSystem, tList
    
    tQuit = False
    '
    If Not tQuit Then
        '  open the list
        
        Set dlgSaveAs = Application.FileDialog(msoFileDialogOpen)
    
        'Use a With...End With block to reference the FileDialog object.
        With dlgSaveAs
            .InitialFileName = ""
            If .Show = -1 Then
                '
                tFileName = ""
                For Each tFN In .SelectedItems
                    tFileName = tFN
                    If Not tFileName = "" Then
                        Exit For
                    End If
                Next
                If tFileName = "" Then
                    MsgBox "Bad file Name."
                    GoTo Exit_CmdImportPlaces_Click
                End If
            End If
        End With
        '
        ' Clear the address table now that we are ready to go
        '
        Set cmdSQL = New ADODB.Command
        cmdSQL.ActiveConnection = CurrentProject.Connection
        cmdSQL.CommandType = adCmdText
        '
        cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_ADDR_LIST"
        cmdSQL.Execute tRecDeleted
        
        cmdSQL.CommandText = "Delete * from InputErrorList"
        cmdSQL.Execute tRecDeleted
        
        cmdSQL.CommandText = "Delete * from TempImportList"
        cmdSQL.Execute tRecDeleted
        
        DoCmd.TransferText acImportDelim, "ImportPlaceList_Space", "TempImportList", tFileName, 0
        '    TransferType=acImportDelim
        '    SpecificationName = "TempImportList" (apparently it is saved in the database itself)
        '    TableName = "TempImportList"  (probably requires that I drop the table first, but I can test)
        '    HasFieldNames = False (0)
        '
        '  copy the bad IDs
        '
        tStrSQL = "INSERT INTO InputErrorList ( c_ID ) SELECT TempImportList.ImportID " + _
            "FROM ADDR_CODES RIGHT JOIN TempImportList ON ADDR_CODES.c_addr_id = TempImportList.ImportID " + _
            "WHERE (((ADDR_CODES.c_addr_id) Is Null))"

        cmdSQL.CommandText = tStrSQL
        cmdSQL.Execute tRecDeleted
        
        If tRecDeleted > 0 Then
            MsgBox "Some ID were not successfully imported:  please look at InputErrorList."
        End If
        '
        '  copy the good IDs
        '
        tStrSQL = "INSERT INTO ZZ_SCRATCH_ADDR_LIST ( c_addr_id ) SELECT DISTINCT TempImportList.ImportID " + _
            "FROM ADDR_CODES INNER JOIN TempImportList ON ADDR_CODES.c_addr_id = TempImportList.ImportID"

        cmdSQL.CommandText = tStrSQL
        cmdSQL.Execute tRecDeleted
        
        If tRecDeleted > 0 Then
            Me.TxtPlace.Value = "[Imported List]"
            Me.TxtPlaceChn.Value = "[Imported List]"
            gUseADDRID = False
            ChkXYRef.Enabled = True
            ChkSubUnits.Enabled = True
            FrameXY.Enabled = True
        End If
        
        Set cmdSQL = Nothing
        Set tFileSystem = Nothing
    End If
        
Exit_CmdImportPlaces_Click:
    Exit Sub

Err_CmdImportPlaces_Click:
    MsgBox Err.Description
    Resume Exit_CmdImportPlaces_Click
        
End Sub



Private Sub FrameFilterYears_Click()
    '
    '  the simplest approach is to turn it all off and then turn on the appropriate objects
    
    ' disable all
    Me.CmdFromDynasty.Enabled = False
    Me.CmdToDynasty.Enabled = False
    Me.CmdAllDynasties.Enabled = False
    Me.TxtFromDynasty.Enabled = False
    Me.TxtFromDynastyPY.Enabled = False
    Me.TxtToDynasty.Enabled = False
    Me.TxtToDynastyPY.Enabled = False
    Me.TxtFromDynasty.Locked = False
    Me.TxtFromDynastyPY.Locked = False
    Me.TxtToDynasty.Locked = False
    Me.TxtToDynastyPY.Locked = False
        
    Me.TxtFromYear.Enabled = False
    Me.TxtToYear.Enabled = False
    
    gUseIndexYears = False
    gUseDynasties = False
        
    If FrameFilterYears.Value = 2 Then
        
        ' enable index years
        Me.TxtFromYear.Enabled = True
        Me.TxtToYear.Enabled = True
        gUseIndexYears = True
    
    ElseIf FrameFilterYears.Value = 3 Then
        
        '  enable dynasties
        Me.CmdFromDynasty.Enabled = True
        Me.CmdToDynasty.Enabled = True
        Me.CmdAllDynasties.Enabled = True
        Me.TxtFromDynasty.Enabled = True
        Me.TxtFromDynastyPY.Enabled = True
        Me.TxtToDynasty.Enabled = True
        Me.TxtToDynastyPY.Enabled = True
        Me.TxtFromDynasty.Locked = True
        Me.TxtFromDynastyPY.Locked = True
        Me.TxtToDynasty.Locked = True
        Me.TxtToDynastyPY.Locked = True
        gUseDynasties = True
    
    End If

End Sub

Private Sub TxtFromYear_LostFocus()
    gFromStr = Trim(TxtFromYear.TEXT)
End Sub

Private Sub TxtToYear_LostFocus()
    gToStr = Trim(TxtToYear.TEXT)
End Sub

Private Sub CmdHelp_Click()
    Dim tStrPDF As String
    
    tStrPDF = Application.CurrentProject.Path + "\HelpFiles\HelpFile_LookAtAssociations.pdf"
    
    'MsgBox tStrPDF
    
    Application.FollowHyperlink tStrPDF, , True
    
End Sub

Private Sub writeKML()
On Error GoTo Err_writeKML
    '
    '  This program will dump the results to a .gis file
    '
    If ZZ_SCRATCH_P_ASSOC.Form.Recordset.RecordCount = 0 Then
        MsgBox "There are no records to save."
        GoTo Exit_writeKML
    End If
    '
    Dim tStream As ADODB.Stream
    Set tStream = New ADODB.Stream
    
    tPinyin = False
    If GISFrame.Value = 1 Then
        tStream.Charset = "utf-8"
        tCodeStr = "UTF8"
    ElseIf GISFrame.Value = 2 Then
        tStream.Charset = "ascii"
        tCodeStr = "ASCII"
        tPinyin = True
    Else
        tStream.Charset = "gb18030"
        tCodeStr = "GB18030"
    End If
    tStream.Mode = adModeReadWrite
    tStream.Type = adTypeText
    tStream.Open
    '
    '  next get a file
    '
    
    Set dlgSaveAs = Application.FileDialog(msoFileDialogSaveAs)

    dlgSaveAs.InitialFileName = "network_gis_" + tCodeStr + ".kml"
    If dlgSaveAs.Show = -1 Then
        '
        tFileName = ""
        For Each tFN In dlgSaveAs.SelectedItems
            tFileName = tFN
            If Not tFileName = "" Then
                Exit For
            End If
        Next
        If tFileName = "" Then
            MsgBox "Bad file Name."
            GoTo Exit_writeKML
        Else
            '  make sure the file name has a txt extension
            If Len(tFileName) < 5 Then
                tFileName = tFileName + ".kml"
            ElseIf Not (LCase(Right(tFileName, 4)) = ".kml") Then
                tFileName = tFileName + ".kml"
            End If
        End If
        '
        '  write the file
        '
        'Name,NameChn,Female,IndexYear,AddrName,AddrChn,X,Y,xy_count
        '
        ' process the table
        '
        Set tRstNode = ZZ_SCRATCH_P_ASSOC.Form.Recordset
        tC = Chr(9) ' the tab
        tDQ = Chr(34) ' the double quotation mark
        '
        ' write the header
        '
        tStream.WriteText "<kml xmlns=" + tDQ + "http://www.opengis.net/kml/2.2" + tDQ + ">", adWriteLine
        tStream.WriteText "<Document>", adWriteLine
        tStream.WriteText tC + "<name>ExtendedData+SchemaData</name>", adWriteLine
        tStream.WriteText tC + "<open>1</open>", adWriteLine '"
        tStream.WriteText tC + "<!-- Create a balloon template referring to the user-defined type -->", adWriteLine
        tStream.WriteText tC + "<Style id=" + tDQ + "assoc-balloon-template" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + "<BalloonStyle>", adWriteLine
        tStream.WriteText tC + tC + tC + "<text>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "<![CDATA[", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "ID: $[AssocGIS/PersonID] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "Name Chn: $[AssocGIS/NameHZ] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "Index Year: $[AssocGIS/IndexYear] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "Sex: $[AssocGIS/Sex] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "Address: $[AssocGIS/AddrName] $[AssocGIS/AddrHZ] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "XY Count: $[AssocGIS/XYCount] <br/><br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "]]>", adWriteLine
        tStream.WriteText tC + tC + tC + "</text>", adWriteLine
        tStream.WriteText tC + tC + "</BalloonStyle>", adWriteLine
        tStream.WriteText tC + "</Style>", adWriteLine
        tStream.WriteText tC + "<!-- Declare the type " + tDQ + "AssocGIS" + tDQ + " with 6 fields -->", adWriteLine
        tStream.WriteText tC + "<Schema name=" + tDQ + "AssocGIS" + tDQ + " id=" + tDQ + "AssocGISId" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "uint" + tDQ + " name=" + tDQ + "PersonID" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Person ID]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "NameHZ" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Name Chn]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "Sex" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Sex]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "AddrName" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Address]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "AddrHZ" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Address Chn]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "IndexYear" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Index Year]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "int" + tDQ + " name=" + tDQ + "XYCount" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[XY Count]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        tStream.WriteText tC + "</Schema>", adWriteLine
        
        With tRstNode
            '
            .MoveFirst
            Do While Not .EOF
                ' must guard against NULLs, even where there should not be any
                '
                '  write the point header
                '
                tStream.WriteText tC + "<Placemark>", adWriteLine
                
                If IsNull(!c_name) Then
                    tStr = "[Bad Data] "
                Else
                    tStr = !c_name
                End If
                tStream.WriteText tC + tC + "<name>" + tStr + "</name>", adWriteLine
                
                tStream.WriteText tC + tC + "<styleUrl>#assoc-balloon-template</styleUrl>", adWriteLine
                '
                '  First Year as time stamp
                '
                If IsNull(!c_index_year) Then
                    tStr = "N/A"
                Else
                    tStr = Str(!c_index_year)
                End If
                tStream.WriteText tC + tC + "<TimeStamp>" + tStr + "</TimeStamp>", adWriteLine
                '
                tStream.WriteText tC + tC + "<ExtendedData>", adWriteLine
                tStream.WriteText tC + tC + tC + "<SchemaData schemaUrl=" + tDQ + "#AssocGISId" + tDQ + ">", adWriteLine
                '
                '  person ID
                '
                tStr = Str(!c_person_id)
                tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "PersonID" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                '
                '  Person Name Chn
                '
                If IsNull(!c_name_chn) Then
                    tStr = tStr + "[Bad Data]"
                Else
                    If Trim(!c_name_chn) = "" Then
                        tStr = "[?]"
                    Else
                        tStr = !c_name_chn
                    End If
                End If
                tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "NameHZ" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                '
                '  Index Year
                '
                If IsNull(!c_index_year) Then
                    tStr = "N/A"
                Else
                    tStr = Str(!c_index_year)
                End If
                tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "IndexYear" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                '
                If IsNull(!c_sex) Then
                    tStr = "[?]"
                Else
                    tStr = !c_sex
                End If
                tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "Sex" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                '
                '  Address Name
                '
                If IsNull(!c_addr_name) Then
                    tStr = "[?]"
                ElseIf Trim(!c_addr_name) = "" Then
                    tStr = "[?]"
                Else
                    tStr = !c_addr_name
                End If
                tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "AddrName" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                '
                '  Address Name Chinese
                '
                If IsNull(!c_addr_chn) Then
                    tStr = "[?]"
                ElseIf Trim(!c_addr_chn) = "" Then
                    tStr = "[?]"
                Else
                    tStr = !c_addr_chn
                End If
                tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "AddrHZ" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                '
                '  XY Count
                '
                If IsNull(!xy_count) Then
                    tStr = "0"
                Else
                    tStr = Str(!xy_count)
                End If
                tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "XYCount" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                '
                tStream.WriteText tC + tC + tC + "</SchemaData>", adWriteLine
                tStream.WriteText tC + tC + "</ExtendedData>", adWriteLine
                tStream.WriteText tC + tC + "<Point>", adWriteLine
                '
                '  coordinates
                If IsNull(!x_coord) Then
                    tStr = "0"
                Else
                    tStr = Str(!x_coord)
                End If
                
                If IsNull(!y_coord) Then
                    tStr = tStr + ",0"
                Else
                    tStr = tStr + "," + Str(!y_coord)
                End If
                tStream.WriteText tC + tC + tC + "<coordinates>" + tStr + "</coordinates>", adWriteLine
                '
                '  footer
                '
                tStream.WriteText tC + tC + "</Point>", adWriteLine
                tStream.WriteText tC + "</Placemark>", adWriteLine
                .MoveNext
            Loop
        End With
        '
        '  footer
        '
        tStream.WriteText "</Document>", adWriteLine
        tStream.WriteText "</kml>", adWriteLine
    Else
        'The user pressed Cancel.
    End If

    ' now make sure all the data is copied to tStream
    tStream.Flush
    ' and write the stream to the file
    tStream.SaveToFile tFileName, adSaveCreateOverWrite
    '
    Set tRstNode = Nothing
            
    tStream.Close
    Set tStream = Nothing
    'Set the object variable to Nothing.
    Set dlgSaveAs = Nothing

Exit_writeKML:
    Exit Sub

Err_writeKML:
    MsgBox Err.Description
    Resume Exit_writeKML

End Sub
Private Sub CmdStoreID_Click()
    Dim cmdSQL As ADODB.Command, tRecCount As Variant
    
    Set cmdSQL = New ADODB.Command
    cmdSQL.ActiveConnection = CurrentProject.Connection
    cmdSQL.CommandType = adCmdText
        '
    If DCount("*", "ZZ_STORE_PERSON_ID") > 0 Then
        ' Display message.
        If MsgBox("Do you wish to replace the current stored values?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            Exit Sub
        Else
            cmdSQL.CommandText = "Delete * from ZZ_STORE_PERSON_ID"
            cmdSQL.Execute tRecCount
        End If
    End If

    tStrQuery = "INSERT INTO ZZ_STORE_PERSON_ID ( c_personid ) SELECT DISTINCT ZZ_SCRATCH_P_ASSOC.c_person_id FROM ZZ_SCRATCH_P_ASSOC"
    
    cmdSQL.CommandText = tStrQuery
    cmdSQL.Execute tRecCount
    MsgBox "Person IDs successfully stored.  Click on 'Recall Person IDs' to reuse these IDs in other forms."
    '
    '  update storage source
    cmdSQL.CommandText = "UPDATE PersonIDSource SET SourceForm ='Associations' WHERE PersonIDSource.LineNum =1"
    cmdSQL.Execute tRecCount

End Sub


