Option Compare Database
Dim gRstPeople As DAO.Recordset, gDisplayLanguage As String, gLabelsOK As Boolean
Public gUseADDRID As Boolean, gUseIndexYears As Boolean, gUseEntryYears As Boolean, gUseDynasties As Boolean
Public gFromDynasty As Integer, gToDynasty As Integer
Public gFromDynastyBegin As Integer, gFromDynastyEnd As Integer, gToDynastyBegin As Integer, gToDynastyEnd As Integer


Private Sub ChkUseYears_Click()
    If ChkUseYears.Value = True Then
        ' ChkUseYears.Value = False
        TxtFromYear.Enabled = True
        TxtToYear.Enabled = True
        FrameYears.Enabled = True
    Else
        ' ChkUseYears.Value = False
        TxtFromYear.Enabled = False
        TxtToYear.Enabled = False
        FrameYears.Enabled = False
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

Private Sub CmdGIS_Click()
On Error GoTo Err_CmdGIS_Click
    '
    '  If it is a KML file, call the routine and exit
    '
    If ChkKML.Value Then
        Call writeKML
        Exit Sub
    End If
    '
    '  This program will dump the results to a .gis file
    '
    If Entry_Address_Query.Form.Recordset.RecordCount = 0 Then
        MsgBox "There are no records to save."
        GoTo Exit_CmdGIS_Click
    End If
    '
    ' I leave this code here as the frame if I need to find a way to define the coding of the export file
    
    Dim tStream As ADODB.Stream
    Set tStream = New ADODB.Stream
        
    If GISFrame.Value = 1 Then
        tStream.Charset = "utf-8"
        tCodeStr = "UTF8"
    Else
        tStream.Charset = "gb18030"
        tCodeStr = "GB18030"
    End If
    '
    '  next get a file
    '
    Dim dlgSaveAs As FileDialog
    Dim tFileNum As Integer
    Dim tFileName As String, tFN As Variant, tC As String
    Dim tRstGIS As DAO.Recordset
    Dim tStr As String, tStrFileType As String
        
    Set dlgSaveAs = Application.FileDialog(msoFileDialogSaveAs)
    '
    ' .tab for default
    '
    tStrFileType = ".tab"
    dlgSaveAs.InitialFileName = "entry_gis_" + tCodeStr + tStrFileType
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
                tFileName = tFileName + tStrFileType
            ElseIf Not (LCase(Right(tFileName, 4)) = tStrFileType) Then
                tFileName = tFileName + tStrFileType
            End If
        End If
        '
        'tStrQuery = "SELECT ZZ_SCRATCH_ENTRY.c_personid, ZZ_SCRATCH_ENTRY.c_name, ZZ_SCRATCH_ENTRY.c_name_chn, " + _
            "ZZ_SCRATCH_ENTRY.c_index_year, ZZ_SCRATCH_ENTRY.c_entry_desc, ZZ_SCRATCH_ENTRY.c_entry_chn, " + _
            "ZZ_SCRATCH_ENTRY.c_exam_rank, ZZ_SCRATCH_ENTRY.c_kin_name, ZZ_SCRATCH_ENTRY.c_kin_chn, ZZ_SCRATCH_ENTRY.c_kin_desc, " + _
            "ZZ_SCRATCH_ENTRY.c_addr_name, ZZ_SCRATCH_ENTRY.c_addr_chn, " + _
            "str(ZZ_SCRATCH_ENTRY.x_coord) as X, str(ZZ_SCRATCH_ENTRY.y_coord) as Y, " _
            "str(ZZ_SCRATCH_ENTRY.x_coord) + ',' + str(ZZ_SCRATCH_ENTRY.y_coord) AS XY, " + _
            "ZZ_SCRATCH_ENTRY.xy_count, ZZ_SCRATCH_ENTRY.c_entry_addr_name, ZZ_SCRATCH_ENTRY.c_entry_addr_chn, " + _
            "ZZ_SCRATCH_ENTRY.c_entry_xcoord, ZZ_SCRATCH_ENTRY.c_entry_ycoord, ZZ_SCRATCH_ENTRY.c_entry_xy_count " + _
            "FROM ZZ_SCRATCH_ENTRY " _
            "WHERE ZZ_SCRATCH_ENTRY.c_addr_id > 0"
        'DoCmd.TransferText acExportDelim, , "ENTRY_GIS_QUERY", tFileName, True
        
        tC = Chr(9)  ' tab
        '
        
        '  we have a file name:  now open the stream for writing
        '
        tStream.Mode = adModeReadWrite
        tStream.Type = adTypeText
        
        tStream.Open
        '
        ' process the table
        '
        Set tRstGIS = CurrentDb.OpenRecordset("ZZ_SCRATCH_ENTRY", dbOpenDynaset)
        '
        ' write the header
        '
        tStr = "PersonID" + tC + "Name" + tC + "NameChn" + tC + "IndexYear" + tC + "EntryDesc" + tC + _
            "EntryDescChn" + tC + "ExamRank" + tC + "KinName" + tC + "KinNameChn" + tC + "KinshipRel" + tC + _
            "AddrName" + tC + "AddrNameChn" + tC + "x_coord" + tC + "y_coord" + tC + "XY" + tC + "xy_count" + tC + _
            "EntryAddrName" + tC + "EntryAddrChn" + tC + "Entry_xcoord" + tC + "Entry_ycoord" + tC + "Entry_xy_count"
        tStream.WriteText tStr, adWriteLine
        '
        With tRstGIS
            .MoveFirst
            Do While Not .EOF
                tStr = ""
                '
                If IsNull(!c_personid) Then
                    tStr = "[Person ID Missing]"
                Else
                    tStr = Str(!c_personid)
                End If
                '
                If IsNull(!c_name) Then
                    tStr = tStr + tC + "[Name Missing]"
                Else
                    tStr = tStr + tC + !c_name
                End If
                '
                If IsNull(!c_name_chn) Then
                    tStr = tStr + tC + "[Name Missing]"
                Else
                    tStr = tStr + tC + !c_name_chn
                End If
                '
                If IsNull(!c_index_year) Then
                    tStr = tStr + tC + "[ ]"
                Else
                    tStr = tStr + tC + Str(!c_index_year)
                End If
                '
                If IsNull(!c_entry_desc) Then
                    tStr = tStr + tC + "[ ]"
                Else
                    tStr = tStr + tC + !c_entry_desc
                End If
                '
                If IsNull(!c_entry_chn) Then
                    tStr = tStr + tC + "[ ]"
                Else
                    tStr = tStr + tC + !c_entry_chn
                End If
                '
                If IsNull(!c_exam_rank) Then
                    tStr = tStr + tC + "[ ]"
                Else
                    tStr = tStr + tC + !c_exam_rank
                End If
                '
                If IsNull(!c_kin_name) Then
                    tStr = tStr + tC + "[ ]"
                Else
                    tStr = tStr + tC + !c_kin_name
                End If
                '
                If IsNull(!c_kin_chn) Then
                    tStr = tStr + tC + "[ ]"
                Else
                    tStr = tStr + tC + !c_kin_chn
                End If
                '
                If IsNull(!c_kin_desc) Then
                    tStr = tStr + tC + "[ ]"
                Else
                    tStr = tStr + tC + !c_kin_desc
                End If
                '
                If IsNull(!c_addr_name) Then
                    tStr = tStr + tC + "[Addr Name Missing]"
                Else
                    tStr = tStr + tC + !c_addr_name
                End If
                '
                If IsNull(!c_addr_chn) Then
                    tStr = tStr + tC + "[Addr Chn Missing]"
                Else
                    tStr = tStr + tC + !c_addr_chn
                End If
                '
                If IsNull(!x_coord) Then
                    tStr = tStr + tC + "[ ]"
                Else
                    tStr = tStr + tC + CStr(!x_coord)
                End If
                '
                If IsNull(!y_coord) Then
                    tStr = tStr + tC + "[ ]"
                Else
                    tStr = tStr + tC + CStr(!y_coord)
                End If
                '
                If IsNull(!x_coord) Then
                    tStr = tStr + tC + "[ ]"
                Else
                    tStr = tStr + tC + CStr(!x_coord) + "," + CStr(!y_coord)
                End If
                '
                If IsNull(!xy_count) Then
                    tStr = tStr + tC + "[ ]"
                Else
                    tStr = tStr + tC + Str(!xy_count)
                End If
                '
                If IsNull(!c_entry_addr_name) Then
                    tStr = tStr + tC + "[ ]"
                Else
                    tStr = tStr + tC + !c_entry_addr_name
                End If
                '
                If IsNull(!c_entry_addr_chn) Then
                    tStr = tStr + tC + "[ ]"
                Else
                    tStr = tStr + tC + !c_entry_addr_chn
                End If
                '
                If IsNull(!c_entry_xcoord) Then
                    tStr = tStr + tC + "[ ]"
                Else
                    tStr = tStr + tC + CStr(!c_entry_xcoord)
                End If
                '
                If IsNull(!c_entry_ycoord) Then
                    tStr = tStr + tC + "[ ]"
                Else
                    tStr = tStr + tC + CStr(!c_entry_ycoord)
                End If
                '
                If IsNull(!c_entry_xy_count) Then
                    tStr = tStr + tC + "[ ]"
                Else
                    tStr = tStr + tC + Str(!c_entry_xy_count)
                End If
                '
                If Not (tStr = "") Then
                    tStream.WriteText tStr, adWriteLine
                End If
                .MoveNext
            Loop
        End With
        ' now make sure all the data is copied to tStream
        tStream.Flush
        ' and write the stream to the file
        tStream.SaveToFile tFileName, adSaveCreateOverWrite
        '
        tStream.Close
    Else
        'The user pressed Cancel.
        
    End If
    '
    Set tStream = Nothing
    'Set the object variable to Nothing.
    Set dlgSaveAs = Nothing
    
Exit_CmdGIS_Click:
    Exit Sub

Err_CmdGIS_Click:
    MsgBox Err.Description
    Resume Exit_CmdGIS_Click
End Sub

Private Sub CmdHelp_Click()
    Dim tStrPDF As String
    
    tStrPDF = Application.CurrentProject.Path + "\HelpFiles\HelpFile_LookAtEntry.pdf"
    
    'MsgBox tStrPDF
    
    Application.FollowHyperlink tStrPDF, , True
    
End Sub

Private Sub CmdImportEntryCodes_Click()
On Error GoTo Err_CmdImportEntryCodess_Click
    
    Dim stDocName As String, tRstEntryCodes As DAO.Recordset
    Dim stLinkCriteria As String, tRstImportEntryCodess As DAO.Recordset
    Dim tString As String, tEntryCode As Long, ti As Integer, tStrID As String, tQuit As Boolean
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
                    GoTo Exit_CmdImportEntryCodess_Click
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
        cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_ENTRY_CODE"
        cmdSQL.Execute tRecDeleted
        '
        cmdSQL.CommandText = "Delete * from InputErrorList"
        cmdSQL.Execute tRecDeleted
        
        cmdSQL.CommandText = "Delete * from TempImportList"
        cmdSQL.Execute tRecDeleted
        
        DoCmd.TransferText acImportDelim, "EntryListImport Specification", "TempImportList", tFileName, 0
        '    TransferType=acImportDelim
        '    SpecificationName = "TempImportList" (apparently it is saved in the database itself)
        '    TableName = "TempImportList"  (probably requires that I drop the table first, but I can test)
        '    HasFieldNames = False (0)
        '
        '  copy the bad IDs
        '
        tStrSQL = "INSERT INTO InputErrorList ( c_ID ) SELECT TempImportList.ImportID " + _
            "FROM ENTRY_CODES RIGHT JOIN TempImportList ON ENTRY_CODES.c_entry_code = TempImportList.ImportID " + _
            "WHERE (((ENTRY_CODES.c_entry_code) Is Null))"

        cmdSQL.CommandText = tStrSQL
        cmdSQL.Execute tRecDeleted
        
        If tRecDeleted > 0 Then
            MsgBox "Some ID were not successfully imported:  please look at InputErrorList."
        End If
        '
        '  copy the good IDs
        '
        tStrSQL = "INSERT INTO ZZ_SCRATCH_ENTRY_CODE ( c_entry_code ) SELECT DISTINCT TempImportList.ImportID " + _
            "FROM ENTRY_CODES INNER JOIN TempImportList ON ENTRY_CODES.c_entry_code = TempImportList.ImportID"

        cmdSQL.CommandText = tStrSQL
        cmdSQL.Execute tRecDeleted
        
        Me.TxtTypeDesc.Value = ""
        Me.TxtTypeChn.Value = ""
        If tRecDeleted > 0 Then
            Me.TxtEntryDesc.Value = "[Imported List]"
            Me.TxtEntryChn.Value = "[Imported List]"
            Me.CmdQuery.Enabled = True
            Me.CmdSaveEntryCodes.Enabled = True
        Else
            Me.TxtEntryDesc.Value = ""
            Me.TxtEntryChn.Value = ""
            Me.CmdQuery.Enabled = False
            Me.CmdSaveEntryCodes.Enabled = False
        End If
        
        Set cmdSQL = Nothing
    End If
    
Exit_CmdImportEntryCodess_Click:
    Exit Sub

Err_CmdImportEntryCodess_Click:
    MsgBox Err.Description
    Resume Exit_CmdImportEntryCodess_Click

End Sub

Private Sub CmdNeo4j_Click()
On Error GoTo Err_CmdNeo4j_Click
    '
    '  This program will dump the results of the search to five CSV files
    '
    '  for the moment I'll just describe the format of the CSV file
    '  Note:  Neo4j seems to treat all fields as strings, so there is no need to explicitly mark strings
    '
    '  1. People.CSV
    '      nameID = c_person_id
    '      nameHZ = c_name_chn
    '      namePY = c_name
    '      indexyear = c_index_year
    '      personDynasty = c_dynasty
    '      sex = c_female > (F,M)
    '
    '  2. Places.CSV
    '      placeID = c_addr_id
    '      placeHZ = c_addr_chn
    '      placePY = c_addr_name
    '      placeX  = x_coord
    '      placeY  = y_coord
    '
    '  3. PeoplePlaces.CSV
    '      nameID
    '      placeID
    '      personPlaceRelation
    '
    '  4. PeopleEntry.CSV
    '      nameID = str(c_person_id)
    '      entryID = str(c_node_id)
    '      entryPlaceID
    '      kinID
    '      kinRelID
    '      AssocPersonID
    '      AssocRelID
    '      SocialInstID
    '      SocialInstNameID
    '      EntryYear
    '      EntryDynasty
    '
    '  5. PeoplePlaceCodes
    '
    '  6. EntryCodes.CSV
    '      entryID = str(c_entry_id)
    '      entryDesc = c_entry_desc
    '
    '  7. KinCodes.CSV
    '      kinCode
    '      kinDesc
    '
    '  8. AssocCodes.CSV
    '
    '  9. Institution codes
    '
    '  first see if there are any records to process
    '
    If Entry_Address_Query.Form.Recordset.RecordCount = 0 Then
        MsgBox "There are no records to save."
        GoTo Exit_CmdNeo4j_Click
    End If
    '
    ' warn the user that a lot of files will be created
    '
    MsgBox "Neo4j requires that from 6 to 9 files be created."
    '
    '  allocate the file variables
    '
    Dim dlgSaveAs As FileDialog
    Dim tFileNum As Integer, tFileName As String, tFN As Variant
    '
    '  next get the People file
    '
    Dim tRstPeople As DAO.Recordset, tRstEntry As DAO.Recordset, tRstEntryCodes As DAO.Recordset, tRstPlace As DAO.Recordset
    Dim tRstPeopleEntry As DAO.Recordset, tRstPeoplePlace As DAO.Recordset, tStr As String, tC As String, ti As Integer, tUseList As Boolean
    Dim tQueryStr As String, tPersonID As Long
    Dim gStream As ADODB.Stream, tCodeStr As String
    '
    ' the optional recordsets
    '
    Dim tRstKinCodes As DAO.Recordset, tRstAssocCodes As DAO.Recordset, tRstInstitutions As DAO.Recordset
    '
    'Dim tFileSystem, tGDF
    
    ' set up the stream to write to
    
    Set gStream = New ADODB.Stream
    '
    ' for the moment, set the character set to UTF-8
    
    gStream.Charset = "utf-8"
    tCodeStr = "UTF8"
    'If CodeFrame.Value = 1 Then
    '    gStream.Charset = "utf-8"
    '    tCodeStr = "UTF8"
    'ElseIf CodeFrame.Value = 2 Then
    '    gStream.Charset = "big5"
    '    tCodeStr = "BIG5"
    'ElseIf CodeFrame.Value = 3 Then
    '    gStream.Charset = "gb2312"
    '    tCodeStr = "GB2312"
    'Else
    '    gStream.Charset = "ascii"
    '    tCodeStr = "ascii"
    'End If
    '
    tC = Chr(44) ' the comma
    '
    '  prepare the temp tables for the people, place, peoplePlace and entry data
            
    Dim cmdSQL As ADODB.Command
    Set cmdSQL = New ADODB.Command
    cmdSQL.ActiveConnection = CurrentProject.Connection
    cmdSQL.CommandType = adCmdText
    '
    ' start with people:  there will be three sources for people.
    '   the people who entered
    '   the kin who might have had a role in entry
    '   the associates who might have played a role
    ' the strategy is to just dump all such IDs to a scratch table and append (distinct) to a table for export
    ' ZZ_SCRATCH_P_TEXT is a convenient table for collecting IDs (no primary key)
            
    cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_P_TEXT"
    cmdSQL.Execute tRecDeleted
    '
    '  get the people IDs
    '
    tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id ) SELECT DISTINCT ZZ_SCRATCH_ENTRY.c_personid FROM ZZ_SCRATCH_ENTRY"
    '
    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecDeleted
    '
    tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id ) SELECT DISTINCT ZZ_SCRATCH_ENTRY.c_kin_id FROM ZZ_SCRATCH_ENTRY " + _
                "WHERE (((ZZ_SCRATCH_ENTRY.c_kin_id)>0))"
    '
    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecDeleted
    '
    tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id ) SELECT DISTINCT ZZ_SCRATCH_ENTRY.c_assoc_id FROM ZZ_SCRATCH_ENTRY " + _
                "WHERE (((ZZ_SCRATCH_ENTRY.c_assoc_id)>0));"
    '
    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecDeleted
    '
    '  now clear ZZ_SCRATCH_PEOPLE and copy the records
    '
    cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_PEOPLE"
    cmdSQL.Execute tRecDeleted
    '
    tQueryStr = "INSERT INTO ZZ_SCRATCH_PEOPLE ( c_person_id, c_name, c_name_chn, c_index_year, c_dynasty, c_dynasty_chn, c_female, c_addr_id, c_addr_name, c_addr_chn, c_addr_type, c_addr_desc, c_addr_desc_chn, " + _
            "x_coord, y_coord ) " + _
        "SELECT DISTINCT ZZ_SCRATCH_P_TEXT.c_person_id, BIOG_MAIN.c_name, BIOG_MAIN.c_name_chn, BIOG_MAIN.c_index_year, DYNASTIES.c_dynasty, DYNASTIES.c_dynasty_chn, BIOG_MAIN.c_female, " + _
            "BIOG_MAIN.c_index_addr_id, ADDR_CODES.c_name, ADDR_CODES.c_name_chn, BIOG_MAIN.c_index_addr_type_code, BIOG_ADDR_CODES.c_addr_desc, BIOG_ADDR_CODES.c_addr_desc_chn, " + _
            "ADDR_CODES.x_coord, ADDR_CODES.y_coord " + _
        "FROM ( ( ( ZZ_SCRATCH_P_TEXT INNER JOIN BIOG_MAIN ON ZZ_SCRATCH_P_TEXT.c_person_id = BIOG_MAIN.c_personid ) LEFT JOIN DYNASTIES ON BIOG_MAIN.c_dy = DYNASTIES.c_dy ) " + _
            "LEFT JOIN BIOG_ADDR_CODES ON BIOG_MAIN.c_index_addr_type_code = BIOG_ADDR_CODES.c_addr_type ) LEFT JOIN ADDR_CODES ON BIOG_MAIN.c_index_addr_id = ADDR_CODES.c_addr_id"
    '
    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecDeleted
    '
    Set tRstPeopleEntry = CurrentDb.OpenRecordset("ZZ_SCRATCH_ENTRY", dbOpenDynaset)
    Set tRstPeople = CurrentDb.OpenRecordset("ZZ_SCRATCH_PEOPLE", dbOpenDynaset)
    
    ' Open the People file

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
        '  we have a file name:  now open the stream for writing
            
        gStream.Mode = adModeReadWrite
        gStream.Type = adTypeText
        gStream.Open

        tRstPeople.MoveLast
        '
        ' process the four tables
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
    ' now the PeopleEntry file
    '
    dlgSaveAs.InitialFileName = "PeopleEntry_" + tCodeStr + ".csv"
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
        gStream.Mode = adModeReadWrite
        gStream.Type = adTypeText
        gStream.Open

        tStr = "NameID" + tC + "EntryCode" + tC + "EntryPlaceID" + tC + "KinID" + tC + "KinRelCode" + tC + _
                "AssocPersonID" + tC + "AssocRelCode" + tC + "SocialInstID" + tC + "EntryYear" + tC + "EntryDynasty"
        gStream.WriteText tStr, adWriteLine
        '
        With tRstPeopleEntry
            .MoveFirst
            Do While Not .EOF
                '  the ID of the person
                tStr = Trim(Str(!c_personid)) + tC
                '
                '  entry code
                '
                If IsNull(!c_entry_code) Then
                    tStr = tStr + "0" + tC
                Else
                    tStr = tStr + Trim(Str(!c_entry_code)) + tC
                End If
                '
                '  entry addr id
                '
                If IsNull(!c_entry_addr_id) Then
                    tStr = tStr + "0" + tC
                Else
                    tStr = tStr + Trim(Str(!c_entry_addr_id)) + tC
                End If
                '
                '  kin ID
                '
                If IsNull(!c_kin_id) Then
                    tStr = tStr + "0" + tC
                Else
                    tStr = tStr + Trim(Str(!c_kin_id)) + tC
                End If
                '
                '  kin rel ID
                '
                If IsNull(!c_kin_code) Then
                    tStr = tStr + "0" + tC
                Else
                    tStr = tStr + Trim(Str(!c_kin_code)) + tC
                End If
                '
                '  assoc ID
                '
                If IsNull(!c_assoc_id) Then
                    tStr = tStr + "0" + tC
                Else
                    tStr = tStr + Trim(Str(!c_assoc_id)) + tC
                End If
                '
                '  assoc desc
                '
                If IsNull(!c_assoc_code) Then
                    tStr = tStr + "N/A" + tC
                Else
                    tStr = tStr + Trim(Str(!c_assoc_code)) + tC
                End If
                '
                '  social inst ID
                '
                If IsNull(!c_inst_code) Then
                    tStr = tStr + "0" + tC
                Else
                    tStr = tStr + Right("000000" + Trim(Str(!c_inst_code)), 6) + Right("000000" + Trim(Str(!c_inst_name_code)), 6) + tC
                End If
                '
                '  entry year
                '
                If IsNull(!c_year) Then
                    tStr = tStr + "0" + tC
                Else
                    tStr = tStr + Trim(Str(!c_year)) + tC
                End If
                '
                '  dynasty
                '
                If IsNull(!c_dy) Then
                    tStr = tStr + "0"
                Else
                    tStr = tStr + Trim(Str(!c_dy))
                End If
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
    '  now places
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
        '  there are three sources of places: the list of people, the entry locations, and the list of institutions
        '  since ZZ_SCRATCH_P_TEXT has the required fields, just reuse it before copying to ZZ_ADDRESSES
        '
            
        cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_P_TEXT"
        cmdSQL.Execute tRecDeleted
        '
        '  get the people IDs
        '
        tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_addr_id, c_addr_name, c_addr_chn, x_coord, y_coord ) " + _
                    "SELECT DISTINCT ZZ_SCRATCH_PEOPLE.c_addr_id, ZZ_SCRATCH_PEOPLE.c_addr_name, ZZ_SCRATCH_PEOPLE.c_addr_chn, " + _
                        "ZZ_SCRATCH_PEOPLE.x_coord, ZZ_SCRATCH_PEOPLE.y_coord " + _
                    "FROM ZZ_SCRATCH_PEOPLE"
        '
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecDeleted
        '
        tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_addr_id, c_addr_name, c_addr_chn, x_coord, y_coord ) " + _
                    "SELECT DISTINCT ZZ_SCRATCH_ENTRY.c_entry_addr_id, ZZ_SCRATCH_ENTRY.c_entry_addr_name, ZZ_SCRATCH_ENTRY.c_entry_addr_chn, " + _
                        "ZZ_SCRATCH_ENTRY.c_entry_xcoord, ZZ_SCRATCH_ENTRY.c_entry_ycoord " + _
                    "FROM ZZ_SCRATCH_ENTRY " + _
                    "WHERE (((ZZ_SCRATCH_ENTRY.c_entry_addr_id)>0))"
        '
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecDeleted
        '
        tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_addr_id, c_addr_name, c_addr_chn, x_coord, y_coord ) " + _
                    "SELECT DISTINCT SOCIAL_INSTITUTION_ADDR.c_inst_addr_id, ADDR_CODES.c_name, ADDR_CODES.c_name_chn, ADDR_CODES.x_coord, ADDR_CODES.y_coord " + _
                    "FROM ADDR_CODES INNER JOIN (ZZ_SCRATCH_ENTRY INNER JOIN SOCIAL_INSTITUTION_ADDR " + _
                        "ON (ZZ_SCRATCH_ENTRY.c_inst_name_code = SOCIAL_INSTITUTION_ADDR.c_inst_name_code) " + _
                        "AND (ZZ_SCRATCH_ENTRY.c_inst_code = SOCIAL_INSTITUTION_ADDR.c_inst_code)) " + _
                        "ON (ADDR_CODES.c_addr_id = SOCIAL_INSTITUTION_ADDR.c_inst_addr_id) AND (ADDR_CODES.c_addr_id = SOCIAL_INSTITUTION_ADDR.c_inst_addr_id) " + _
                    "WHERE (((ZZ_SCRATCH_ENTRY.c_inst_code)>0))"
        '
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecDeleted
        '
        ' now copy the results
        cmdSQL.CommandText = "Delete * from ZZ_ADDRESSES"
        cmdSQL.Execute tRecDeleted
        '
        tQueryStr = "INSERT INTO ZZ_ADDRESSES ( c_addr_id, c_name, c_name_chn, x_coord, y_coord ) " + _
                    "SELECT DISTINCT ZZ_SCRATCH_P_TEXT.c_addr_id, ZZ_SCRATCH_P_TEXT.c_addr_name, ZZ_SCRATCH_P_TEXT.c_addr_chn, " + _
                        "ZZ_SCRATCH_P_TEXT.x_coord, ZZ_SCRATCH_P_TEXT.y_coord " + _
                    "FROM ZZ_SCRATCH_P_TEXT WHERE (((ZZ_SCRATCH_P_TEXT.c_addr_id)>0))"
        '
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecDeleted
        
        Set tRstPlace = CurrentDb.OpenRecordset("ZZ_ADDRESSES", dbOpenDynaset)
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
                        
                    If IsNull(!c_name) Then
                        tStr = tStr + "unknown" + tC
                    Else
                        tStr = tStr + !c_name + tC
                    End If
                    '
                    If Not (tCodeStr = "ascii") Then
                        If IsNull(!c_name_chn) Then
                            tStr = tStr + "unknown" + tC
                        Else
                            tStr = tStr + !c_name_chn + tC
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
                        tStr = tStr + "0.0"
                    Else
                        tStr = tStr + Str(!x_coord)
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
    '  now peoplePlaces:  use ZZ_SCRATCH_PEOPLE
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
        tQueryStr = "SELECT DISTINCT ZZ_SCRATCH_PEOPLE.c_person_id, ZZ_SCRATCH_PEOPLE.c_addr_id, ZZ_SCRATCH_PEOPLE.c_addr_type " + _
                    "FROM ZZ_SCRATCH_PEOPLE WHERE (((ZZ_SCRATCH_PEOPLE.c_addr_id) > 0))"

        Set tRstPeoplePlace = CurrentDb.OpenRecordset(tQueryStr)
            
        tStr = "nameID" + tC + "placeID" + tC + "personPlaceCode"
            
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
    '  now peoplePlaceCode:  use ZZ_SCRATCH_PEOPLE
    '
    dlgSaveAs.InitialFileName = "PeoplePlacesCodes_" + tCodeStr + ".csv"
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
        tQueryStr = "SELECT DISTINCT ZZ_SCRATCH_PEOPLE.c_addr_type, ZZ_SCRATCH_PEOPLE.c_addr_desc, ZZ_SCRATCH_PEOPLE.c_addr_desc_chn " + _
                    "FROM ZZ_SCRATCH_PEOPLE WHERE (((ZZ_SCRATCH_PEOPLE.c_addr_type) > 0))"

        Set tRstPeoplePlace = CurrentDb.OpenRecordset(tQueryStr)
            
        If tCodeStr = "ascii" Then
            tStr = "personPlaceCode" + tC + "personPlaceTrans"
        Else
            tStr = "personPlaceCode" + tC + "personPlaceTrans" + tC + "personPlaceHZ"
        End If
            
        gStream.WriteText tStr, adWriteLine
            
        With tRstPeoplePlace
            .MoveFirst
            Do While Not .EOF
                If Not IsNull(!c_addr_type) Then
                    '
                    tStr = Trim(Str(!c_addr_type)) + tC
                    '
                    tStr = tStr + !c_addr_desc
                    '
                    If Not (tCodeStr = "ascii") Then
                        tStr = tStr + tC + !c_addr_desc_chn
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
    ' finally, get entry codes, kinship codes, association codes, and institution codes, if there are any
    '
    ' now the EntryCode file
    '
    dlgSaveAs.InitialFileName = "EntryCode_" + tCodeStr + ".csv"
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
        gStream.Mode = adModeReadWrite
        gStream.Type = adTypeText
        gStream.Open

        If tCodeStr = "ascii" Then
            tStr = "EntryCode" + tC + "EntryDesc"
        Else
            tStr = "EntryCode" + tC + "EntryDesc" + tC + "EntryDescHZ"
        End If
        gStream.WriteText tStr, adWriteLine
        '
        ' get the codes
        '
        tQueryStr = "SELECT DISTINCT ZZ_SCRATCH_ENTRY.c_entry_code, ZZ_SCRATCH_ENTRY.c_entry_desc, ZZ_SCRATCH_ENTRY.c_entry_chn FROM ZZ_SCRATCH_ENTRY"
        Set tRstEntryCode = CurrentDb.OpenRecordset(tQueryStr)
        With tRstEntryCode
            .MoveFirst
            Do While Not .EOF
                '
                tStr = Trim(Str(!c_entry_code)) + tC
                '
                '  entry desc
                '
                If IsNull(!c_entry_desc) Then
                    tStr = tStr + "Missing"
                Else
                    tStr = tStr + Trim(!c_entry_desc)
                End If
                '
                '  kin ID
                '
                If Not (tCodeStr = "ascii") Then
                    tStr = tStr + tC + Trim(!c_entry_chn)
                End If
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
    cmdSQL.CommandText = "DELETE * FROM ZZ_SCRATCH_P_TEXT"
    cmdSQL.Execute tRecDeleted
    '
    tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id ) " + _
                "SELECT DISTINCT ZZ_SCRATCH_ENTRY.c_personid " + _
                "FROM ZZ_SCRATCH_ENTRY " + _
                "WHERE (((ZZ_SCRATCH_ENTRY.c_kin_code)>0))"
    '
    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecDeleted
    '
    tQueryStr = "SELECT DISTINCT ZZ_SCRATCH_ENTRY.c_kin_code, ZZ_SCRATCH_ENTRY.c_kin_desc FROM ZZ_SCRATCH_ENTRY WHERE (((ZZ_SCRATCH_ENTRY.c_kin_code)>0))"
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
            Set tRstKinCodes = CurrentDb.OpenRecordset(tQueryStr)
            '
            tStr = "KinCode" + tC + "KinDesc"
            gStream.WriteText tStr, adWriteLine
            'tGDF.WriteLine (tStr)

            With tRstKinCodes
                .MoveFirst
                Do While Not .EOF
                    If Not IsNull(!c_kin_code) Then
                        tStr = Trim(Str(!c_kin_code)) + tC
                        '
                        tStr = tStr + Trim(!c_kin_desc)
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
            'tGDF.Close
            '
        Else
            'The user pressed Cancel.
        End If
    End If
    '
    cmdSQL.CommandText = "DELETE * FROM ZZ_SCRATCH_P_TEXT"
    cmdSQL.Execute tRecDeleted
    '
    tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id ) " + _
                "SELECT DISTINCT ZZ_SCRATCH_ENTRY.c_personid " + _
                "FROM ZZ_SCRATCH_ENTRY " + _
                "WHERE (((ZZ_SCRATCH_ENTRY.c_assoc_code)>0))"
    '
    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecDeleted
    '
    tQueryStr = "SELECT DISTINCT ZZ_SCRATCH_ENTRY.c_assoc_code, ZZ_SCRATCH_ENTRY.c_assoc_desc, ZZ_SCRATCH_ENTRY.c_assoc_desc_chn " + _
                "FROM ZZ_SCRATCH_ENTRY WHERE (((ZZ_SCRATCH_ENTRY.c_assoc_code)>0))"
                
    If tRecDeleted > 0 Then
        dlgSaveAs.InitialFileName = "AssocCodes_" + tCodeStr + ".csv"
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
            Set tRstAssocCodes = CurrentDb.OpenRecordset(tQueryStr)
            '
            If tCodeStr = "ascii" Then
                tStr = "AssocCode" + tC + "AssocDesc"
            Else
                tStr = "AssocCode" + tC + "AssocDesc" + tC + "AssocDescHZ"
            End If
            gStream.WriteText tStr, adWriteLine
            'tGDF.WriteLine (tStr)

            With tRstAssocCodes
                .MoveFirst
                Do While Not .EOF
                    If Not IsNull(!c_assoc_code) Then
                        tStr = Trim(Str(!c_assoc_code)) + tC
                        '
                        tStr = tStr + Trim(!c_assoc_desc)
                        '
                        If Not (tCodeStr = "ascii") Then
                            tStr = tStr + tC + !c_assoc_desc_chn
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
    End If
    '
    ' the final selection is for social institutions
    '
    cmdSQL.CommandText = "DELETE * FROM ZZ_SCRATCH_P_TEXT"
    cmdSQL.Execute tRecDeleted
    '
    tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id ) " + _
                "SELECT DISTINCT ZZ_SCRATCH_ENTRY.c_personid " + _
                "FROM ZZ_SCRATCH_ENTRY " + _
                "WHERE (((ZZ_SCRATCH_ENTRY.c_inst_code)>0))"
    '
    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecDeleted
    '
    tQueryStr = "SELECT DISTINCT ZZ_SCRATCH_ENTRY.c_inst_code, ZZ_SCRATCH_ENTRY.c_inst_name_code, ZZ_SCRATCH_ENTRY.c_inst_name_hz, " + _
                    "ZZ_SCRATCH_ENTRY.c_inst_name_py " + _
                "FROM ZZ_SCRATCH_ENTRY WHERE (((ZZ_SCRATCH_ENTRY.c_inst_code)>0))"
                
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
            Set tRstInstitutions = CurrentDb.OpenRecordset(tQueryStr)
            '
            If tCodeStr = "ascii" Then
                tStr = "InstitutionCode" + tC + "InstitutionNamePY"
            Else
                tStr = "InstitutionCode" + tC + "InstitutionNamePY" + tC + "InstitutionNameHZ"
            End If
            gStream.WriteText tStr, adWriteLine
            'tGDF.WriteLine (tStr)

            With tRstAssocCodes
                .MoveFirst
                Do While Not .EOF
                    If Not IsNull(!c_inst_code) Then
                        tStr = Right("000000" + Trim(Str(!c_inst_code)), 6) + Right("000000" + Trim(Str(!c_inst_name_code)), 6) + tC
                        '
                        If IsNull(!c_inst_name_py) Then
                            tStr = tStr + "NameMissing"
                        Else
                            tStr = tStr + Trim(!c_inst_name_py)
                        End If
                        '
                        If Not (tCodeStr = "ascii") Then
                            If IsNull(!c_inst_name_hz) Then
                                tStr = tStr + tC + "NameMissing"
                            Else
                                tStr = tStr + tC + !c_inst_name_hz
                            End If
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
    End If

    cmdSQL.CommandText = "DELETE * FROM ZZ_SCRATCH_P_TEXT"
    cmdSQL.Execute tRecDeleted
    
    MsgBox "Finished saving to Neo4j"

    'Set the object variable to Nothing.
    Set dlgSaveAs = Nothing
    
Exit_CmdNeo4j_Click:
    Exit Sub

Err_CmdNeo4j_Click:
    MsgBox Err.Description
    Resume Exit_CmdNeo4j_Click
    
End Sub

Private Sub CmdQuery_Click()
    On Error GoTo Err_CmdQuery_Click

    Dim rst As DAO.Recordset
    Dim EntryQuery As DAO.QueryDef, AddressQuery As DAO.QueryDef
    Dim tRstDummy As DAO.Recordset, tRstAddress As DAO.Recordset
    Dim prm As DAO.Parameter, tRstBiogMain As DAO.Recordset
    Dim tRstKinCodes As DAO.Recordset, tRstADDRID As DAO.Recordset
    Dim tRstAddrList As DAO.Recordset, tQstr As String, tQt As String
    Dim cmdSQL As ADODB.Command, tStrYears As String, tStrFromYear As String, tStrToYear As String, tStrFromAddr As String, tStrFrom As String
    
    tQt = Chr(34)
    
    Set cmdSQL = New ADODB.Command
    '
    '  to clear the table, briefly close and then delete records
    '
    Set gRstPeople = Entry_Address_Query.Form.Recordset
    Set tRstDummy = CurrentDb.OpenRecordset("Z_SCRATCH_DUMMY_EC", dbOpenDynaset)
    Set Entry_Address_Query.Form.Recordset = tRstDummy
    gRstPeople.Close
    '
    cmdSQL.ActiveConnection = CurrentProject.Connection
    cmdSQL.CommandType = adCmdText
        '
    cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_ENTRY"
    cmdSQL.Execute tRecDeleted
    '
    ' now see if address IDs will be used.  If so, zap the scratch file and repopulate
    ' by looking for address that belong to the address
    
    'MsgBox "About to process address"
    If gUseADDRID Then
        '
        '  ZZ_SCRATCH_ADDR has at least one record
        '
        cmdSQL.CommandText = "DELETE * FROM ZZ_SCRATCH_ADDR_LIST"
        cmdSQL.Execute tRecDeleted

        '
        If ChkSubUnits.Value Then
            tQueryStr = "INSERT INTO ZZ_SCRATCH_ADDR_LIST ( c_addr_id ) " + _
                "SELECT DISTINCT ZZZ_BELONGS_TO.c_addr_id " + _
                "FROM ZZ_SCRATCH_ADDR INNER JOIN ZZZ_BELONGS_TO ON " + _
                "ZZ_SCRATCH_ADDR.c_addr_id = ZZZ_BELONGS_TO.c_belongs_to"
        Else
            tQueryStr = "INSERT INTO ZZ_SCRATCH_ADDR_LIST ( c_addr_id ) SELECT DISTINCT c_addr_id FROM ZZ_SCRATCH_ADDR"
        End If
        
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecDeleted

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
            tQueryStr = "INSERT INTO ZZ_ADDRESSES ( c_addr_id )SELECT DISTINCT ADDR_CODES.c_addr_id " + _
                "FROM ADDR_CODES, ZZ_SCRATCH_ADDR_LIST INNER JOIN ADDR_CODES AS ADDR_CODES_1 ON " + _
                "ZZ_SCRATCH_ADDR_LIST.c_addr_id = ADDR_CODES_1.c_addr_id " + _
                "WHERE (((ADDR_CODES.x_coord)>=([ADDR_CODES_1].[x_coord]-0.03) And " + _
                "(ADDR_CODES.x_coord)<=([ADDR_CODES_1].[x_coord]+0.03)) AND " + _
                "((ADDR_CODES.y_coord)>=([ADDR_CODES_1].[y_coord]-0.03) And " + _
                "(ADDR_CODES.y_coord)<=([ADDR_CODES_1].[y_coord]+0.03)))"
                
            cmdSQL.CommandText = tQueryStr
            cmdSQL.Execute tRecDeleted
            '
            ' now get the address IDs from the initial list that have no xy coordinates
            '
            tQueryStr = "INSERT INTO ZZ_ADDRESSES ( c_addr_id ) SELECT ZZ_SCRATCH_ADDR_LIST.c_addr_id " + _
                "FROM ZZ_SCRATCH_ADDR_LIST INNER JOIN ADDR_CODES ON " + _
                "ZZ_SCRATCH_ADDR_LIST.c_addr_id = ADDR_CODES.c_addr_id " + _
                "WHERE (((ADDR_CODES.x_coord) Is Null)) OR (((ADDR_CODES.y_coord) Is Null))"
            cmdSQL.CommandText = tQueryStr
            cmdSQL.Execute tRecDeleted
            '
            '  zap ZZ_SCRATCH_ADDR
            '
            tQueryStr = "DELETE * FROM ZZ_SCRATCH_ADDR_LIST"
            cmdSQL.CommandText = tQueryStr
            cmdSQL.Execute tRecDeleted
            '
            '  copy the list
            '
            tQueryStr = "INSERT INTO ZZ_SCRATCH_ADDR_LIST ( c_addr_id )SELECT DISTINCT ZZ_ADDRESSES.c_addr_id " + _
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
    End If
    'MsgBox "Finished processing address"
    
    tStrFromYear = Str(TxtFromYear.Value)
    tStrToYear = Str(TxtToYear.Value)
    
    gUseEntryYears = False
    gUseIndexYears = False
    gUseDynasties = False
    If FrameYears.Value = 1 Then    ' Entry Years
        If Not (tStrFromYear = "") And tStrToYear = "" Then
            tStrYears = "(ENTRY_DATA.c_year >=" + Str(TxtFromYear.Value) + ")"
        ElseIf tStrFromYear = "" And Not (tStrToYear = "") Then
            tStrYears = "(ENTRY_DATA.c_year <=" + Str(TxtToYear.Value) + ")"
        ElseIf Not (tStrFromYear = "") And Not (tStrToYear = "") Then
            tStrYears = "(ENTRY_DATA.c_year >=" + Str(TxtFromYear.Value) + " And ENTRY_DATA.c_year <=" + Str(TxtToYear.Value) + ")"
        Else
            tStrYears = ""
        End If
        If Not (tStrYears = "") Then
            gUseEntryYears = True
        End If
    ElseIf FrameYears.Value = 2 Then    ' Index Years
        If Not (tStrFromYear = "") And tStrToYear = "" Then
            tStrYears = "(BIOG_MAIN.c_index_year >=" + Str(TxtFromYear.Value) + ")"
        ElseIf tStrFromYear = "" And Not (tStrToYear = "") Then
            tStrYears = "(BIOG_MAIN.c_index_year <=" + Str(TxtToYear.Value) + ")"
        ElseIf Not (tStrFromYear = "") And Not (tStrToYear = "") Then
            tStrYears = "(BIOG_MAIN.c_index_year >=" + Str(TxtFromYear.Value) + " And BIOG_MAIN.c_index_year <=" + Str(TxtToYear.Value) + ")"
        Else
            tStrYears = ""
        End If
        If Not (tStrYears = "") Then
            gUseIndexYears = True
        End If
    ElseIf FrameYears.Value = 3 Then    ' Dynasties
        If gFromDynasty = -2 Then
            tStrYears = "((BIOG_MAIN.c_dy) > 0 ) "
        ElseIf gFromDynasty = -1 And gToDynasty > 0 Then
            tStrYears = "((DYNASTIES.c_start)<" + Str(gToDynastyEnd) + ") "
        ElseIf gFromDynasty > 0 And gToDynasty = -1 Then
            tStrYears = "((DYNASTIES.c_end)>" + Str(gFromDynastyBegin) + ") "
        ElseIf gFromDynasty = gToDynasty And gFromDynasty > 0 Then
            tStrYears = "((DYNASTIES.c_dy)=" + Str(gFromDynasty) + ") "
        ElseIf gFromDynasty > 0 And gToDynasty > 0 Then
            tStrYears = "((DYNASTIES.c_end)>" + Str(gFromDynastyBegin) + ") AND ((DYNASTIES.c_start)<" + Str(gToDynastyEnd) + ") "
        Else
            ' no constraint have been set, so just ignore
            
            tStrYears = ""
        End If
        If Not (tStrYears = "") Then
            gUseDynasties = True
        End If
    Else
        tStrYears = ""
    End If
    
    tQstr = "INSERT INTO ZZ_SCRATCH_ENTRY ( c_personid, c_name, c_name_chn, c_index_year, c_index_year_type_code, c_dy, c_entry_code, c_year, c_sequence, " + _
            "c_exam_rank, c_addr_id, c_kin_id, c_kin_code, c_assoc_id, c_assoc_code, " + _
            "c_parental_status_code, c_entry_addr_id, c_source, c_inst_code, c_inst_name_code, c_addr_type ) " + _
        "SELECT ENTRY_DATA.c_personid, BIOG_MAIN.c_name, BIOG_MAIN.c_name_chn, BIOG_MAIN.c_index_year, BIOG_MAIN.c_index_year_type_code, BIOG_MAIN.c_dy, " + _
            "ENTRY_DATA.c_entry_code, ENTRY_DATA.c_year, ENTRY_DATA.c_sequence, ENTRY_DATA.c_exam_rank, BIOG_MAIN.c_index_addr_id, ENTRY_DATA.c_kin_id, " + _
            "ENTRY_DATA.c_kin_code, ENTRY_DATA.c_assoc_id, ENTRY_DATA.c_assoc_code, ENTRY_DATA.c_parental_status_code, " + _
            "ENTRY_DATA.c_entry_addr_id, ENTRY_DATA.c_source, ENTRY_DATA.c_inst_code, ENTRY_DATA.c_inst_name_code, BIOG_MAIN.c_index_addr_type_code "
 
    ' the FROM statement gets complicated because of the nesting of the inner joins for address, entry code, and dynasty
    ' This code, for the sake of clarity, simply sets out all the options
    
    If gUseADDRID Then
        '
        '  use person address = 1, use entry address = 2
        '
        If FrameAddress.Value = 1 Then '                        join is with BIOG_MAIN.c_index_addr_id
        
        If TxtEntryDesc.Value = "[All]" And TxtTypeCode.Value = "" Then ' No entry codes
            
            If gUseDynasties Then
  
                ' join both address and dynasty but no entry code
  
                tStrFrom = " FROM DYNASTIES RIGHT JOIN ( ( BIOG_MAIN INNER JOIN ENTRY_DATA ON BIOG_MAIN.c_personid = ENTRY_DATA.c_personid ) INNER JOIN ZZ_SCRATCH_ADDR_LIST " + _
                    "ON BIOG_MAIN.c_index_addr_id = ZZ_SCRATCH_ADDR_LIST.c_addr_id ) ON DYNASTIES.c_dy = BIOG_MAIN.c_dy "
            Else
  
                ' join just address
  
                tStrFrom = " FROM ( BIOG_MAIN INNER JOIN ENTRY_DATA ON BIOG_MAIN.c_personid = ENTRY_DATA.c_personid ) INNER JOIN ZZ_SCRATCH_ADDR_LIST ON BIOG_MAIN.c_index_addr_id = ZZ_SCRATCH_ADDR_LIST.c_addr_id "
            End If
            '
        Else    ' entry code(s) are specified as well as address
            '
            ' the table ZZ_SCRATCH_ENTRY_CODE contains either one selected code or all the codes for a particular selected TYPE
     
            If gUseDynasties Then   '  this joins all three: address, entry, dynasty
  
                tStrFrom = " FROM DYNASTIES RIGHT JOIN ( ( ZZ_SCRATCH_ENTRY_CODE INNER JOIN ( BIOG_MAIN INNER JOIN ENTRY_DATA ON BIOG_MAIN.c_personid = ENTRY_DATA.c_personid ) " + _
                    "ON ZZ_SCRATCH_ENTRY_CODE.c_entry_code = ENTRY_DATA.c_entry_code ) INNER JOIN ZZ_SCRATCH_ADDR_LIST ON BIOG_MAIN.c_index_addr_id = ZZ_SCRATCH_ADDR_LIST.c_addr_id ) " + _
                    "ON DYNASTIES.c_dy = BIOG_MAIN.c_dy "
  
            Else    ' this joins just address and entry
  
                tStrFrom = "FROM ( ZZ_SCRATCH_ENTRY_CODE INNER JOIN ( BIOG_MAIN INNER JOIN ENTRY_DATA ON BIOG_MAIN.c_personid = ENTRY_DATA.c_personid ) " + _
                    "ON ZZ_SCRATCH_ENTRY_CODE.c_entry_code = ENTRY_DATA.c_entry_code ) INNER JOIN ZZ_SCRATCH_ADDR_LIST ON BIOG_MAIN.c_index_addr_id = ZZ_SCRATCH_ADDR_LIST.c_addr_id "
            End If
  
        End If
    Else '                   address join is with ENTRY_DATA.c_entry_addr_id

        If TxtEntryDesc.Value = "[All]" And TxtTypeCode.Value = "" Then ' No entry codes
            
            If gUseDynasties Then
  
                ' join both address and dynasty but no entry code
  
                tStrFrom = " FROM DYNASTIES RIGHT JOIN ( ( BIOG_MAIN INNER JOIN ENTRY_DATA ON BIOG_MAIN.c_personid = ENTRY_DATA.c_personid ) INNER JOIN ZZ_SCRATCH_ADDR_LIST " + _
                    "ON ENTRY_DATA.c_entry_addr_id = ZZ_SCRATCH_ADDR_LIST.c_addr_id ) ON DYNASTIES.c_dy = BIOG_MAIN.c_dy "
            Else
  
                ' join just address
  
                tStrFrom = " FROM ( BIOG_MAIN INNER JOIN ENTRY_DATA ON BIOG_MAIN.c_personid = ENTRY_DATA.c_personid ) INNER JOIN ZZ_SCRATCH_ADDR_LIST ON ENTRY_DATA.c_entry_addr_id = ZZ_SCRATCH_ADDR_LIST.c_addr_id "
            End If
            '
        Else    ' entry code(s) are specified as well as address
            '
            ' the table ZZ_SCRATCH_ENTRY_CODE contains either one selected code or all the codes for a particular selected TYPE
     
            If gUseDynasties Then   '  this joins all three: address, entry, dynasty
  
                tStrFrom = " FROM ZZ_SCRATCH_ENTRY_CODE INNER JOIN ( DYNASTIES RIGHT JOIN ( ( BIOG_MAIN INNER JOIN ENTRY_DATA ON BIOG_MAIN.c_personid = ENTRY_DATA.c_personid ) " + _
                    "INNER JOIN ZZ_SCRATCH_ADDR_LIST ON ENTRY_DATA.c_entry_addr_id = ZZ_SCRATCH_ADDR_LIST.c_addr_id ) ON DYNASTIES.c_dy = BIOG_MAIN.c_dy ) " + _
                    "ON ZZ_SCRATCH_ENTRY_CODE.c_entry_code = ENTRY_DATA.c_entry_code "
  
            Else    ' this joins just address and entry
  
                tStrFrom = "FROM ZZ_SCRATCH_ENTRY_CODE INNER JOIN ( ( BIOG_MAIN INNER JOIN ENTRY_DATA ON BIOG_MAIN.c_personid = ENTRY_DATA.c_personid ) INNER JOIN ZZ_SCRATCH_ADDR_LIST " + _
                    "ON ENTRY_DATA.c_entry_addr_id = ZZ_SCRATCH_ADDR_LIST.c_addr_id ) ON ZZ_SCRATCH_ENTRY_CODE.c_entry_code = ENTRY_DATA.c_entry_code "
            End If
  
        End If
    End If
        '
    Else    ' No addresses
        
        If TxtEntryDesc.Value = "[All]" And TxtTypeCode.Value = "" Then ' This is unconstrained and a bad idea unless there is a dynasty constrain
        
            ' all entry codes are OK:  selection is just by place
            If gUseDynasties Then
                tStrFrom = " FROM DYNASTIES RIGHT JOIN ( BIOG_MAIN INNER JOIN ENTRY_DATA ON BIOG_MAIN.c_personid = ENTRY_DATA.c_personid ) ON DYNASTIES.c_dy = BIOG_MAIN.c_dy "
            Else
                tStrFrom = "FROM BIOG_MAIN INNER JOIN ENTRY_DATA ON BIOG_MAIN.c_personid = ENTRY_DATA.c_personid "
            End If
        Else
            '
            If gUseDynasties Then  ' join dynasty and entry codes
            
                tStrFrom = " FROM ZZ_SCRATCH_ENTRY_CODE INNER JOIN ( DYNASTIES RIGHT JOIN ( BIOG_MAIN INNER JOIN ENTRY_DATA ON BIOG_MAIN.c_personid = ENTRY_DATA.c_personid ) " + _
            "ON DYNASTIES.c_dy = BIOG_MAIN.c_dy ) ON ZZ_SCRATCH_ENTRY_CODE.c_entry_code = ENTRY_DATA.c_entry_code "
            
            Else ' join just entry codes
            
                tStrFrom = " FROM ZZ_SCRATCH_ENTRY_CODE INNER JOIN ( BIOG_MAIN INNER JOIN ENTRY_DATA ON BIOG_MAIN.c_personid = ENTRY_DATA.c_personid ) ON ZZ_SCRATCH_ENTRY_CODE.c_entry_code = ENTRY_DATA.c_entry_code "
            End If
        End If
        '
    End If
    '
    tQstr = tQstr + tStrFrom
            
    ' add the years constraint, if needed
            
    If gUseEntryYears Or gUseIndexYears Or gUseDynasties Then
        '
        ' one last paranoid check
        If Not (tStrYears = "") Then
            tQstr = tQstr + " WHERE (" + tStrYears + ")"
        End If
    End If
    '
    ' MsgBox tQstr
    '
    '  run the query
    '
    cmdSQL.CommandText = tQstr
    cmdSQL.Execute tRecDeleted
    '
    If tRecDeleted > 0 Then
        '
        ' fill in the fields (to keep the query from getting too messy, I do this in several queries
        '
        tQstr = "UPDATE ( ( ( ( ( ( ZZ_SCRATCH_ENTRY LEFT JOIN INDEXYEAR_TYPE_CODES ON ZZ_SCRATCH_ENTRY.c_index_year_type_code = INDEXYEAR_TYPE_CODES.c_index_year_type_code ) " + _
                "INNER JOIN ENTRY_CODES ON ZZ_SCRATCH_ENTRY.c_entry_code = ENTRY_CODES.c_entry_code ) LEFT JOIN ASSOC_CODES ON ZZ_SCRATCH_ENTRY.c_assoc_code = ASSOC_CODES.c_assoc_code ) " + _
                "LEFT JOIN BIOG_MAIN ON ZZ_SCRATCH_ENTRY.c_kin_id = BIOG_MAIN.c_personid ) LEFT JOIN BIOG_MAIN AS BIOG_MAIN_1 ON ZZ_SCRATCH_ENTRY.c_assoc_id = BIOG_MAIN_1.c_personid ) " + _
                "LEFT JOIN KINSHIP_CODES ON ZZ_SCRATCH_ENTRY.c_kin_code = KINSHIP_CODES.c_kincode ) LEFT JOIN SOCIAL_INSTITUTION_NAME_CODES " + _
                "ON ZZ_SCRATCH_ENTRY.c_inst_name_code = SOCIAL_INSTITUTION_NAME_CODES.c_inst_name_code " + _
            "SET ZZ_SCRATCH_ENTRY.c_index_year_type_desc = [INDEXYEAR_TYPE_CODES].[c_index_year_type_desc], ZZ_SCRATCH_ENTRY.c_index_year_type_hz = [INDEXYEAR_TYPE_CODES].[c_index_year_type_hz], " + _
                "ZZ_SCRATCH_ENTRY.c_entry_desc = [ENTRY_CODES].[c_entry_desc], ZZ_SCRATCH_ENTRY.c_entry_chn = [ENTRY_CODES].[c_entry_desc_chn], ZZ_SCRATCH_ENTRY.c_kin_desc = [KINSHIP_CODES].[c_kinrel], " + _
                "ZZ_SCRATCH_ENTRY.c_kin_name = [BIOG_MAIN].[c_name], ZZ_SCRATCH_ENTRY.c_kin_chn = [BIOG_MAIN].[c_name_chn], " + _
                "ZZ_SCRATCH_ENTRY.c_assoc_desc = [ASSOC_CODES].[c_assoc_desc], ZZ_SCRATCH_ENTRY.c_assoc_desc_chn = [ASSOC_CODES].[c_assoc_desc_chn], " + _
                "ZZ_SCRATCH_ENTRY.c_assoc_name = [BIOG_MAIN_1].[c_name], ZZ_SCRATCH_ENTRY.c_assoc_name_chn = [BIOG_MAIN_1].[c_name_chn]"
        cmdSQL.CommandText = tQstr
        cmdSQL.Execute tRecDeleted

        tQstr = "UPDATE ( ( ( ( ( ZZ_SCRATCH_ENTRY LEFT JOIN ADDR_CODES ON ZZ_SCRATCH_ENTRY.c_addr_id = ADDR_CODES.c_addr_id ) LEFT JOIN PARENTAL_STATUS_CODES " + _
                "ON ZZ_SCRATCH_ENTRY.c_parental_status_code = PARENTAL_STATUS_CODES.c_parental_status_code ) LEFT JOIN ADDR_CODES AS ADDR_CODES_1 ON ZZ_SCRATCH_ENTRY.c_entry_addr_id = ADDR_CODES_1.c_addr_id ) " + _
                "LEFT JOIN DYNASTIES ON ZZ_SCRATCH_ENTRY.c_dy = DYNASTIES.c_dy ) LEFT JOIN TEXT_CODES ON ZZ_SCRATCH_ENTRY.c_source = TEXT_CODES.c_textid ) " + _
                "LEFT JOIN BIOG_ADDR_CODES ON ZZ_SCRATCH_ENTRY.c_addr_type = BIOG_ADDR_CODES.c_addr_type " + _
            "SET ZZ_SCRATCH_ENTRY.c_addr_name = [ADDR_CODES].[c_name], ZZ_SCRATCH_ENTRY.c_addr_chn = [ADDR_CODES].[c_name_chn], " + _
                "ZZ_SCRATCH_ENTRY.x_coord = [ADDR_CODES].[x_coord], ZZ_SCRATCH_ENTRY.y_coord = [ADDR_CODES].[y_coord], " + _
                "ZZ_SCRATCH_ENTRY.c_addr_desc = [BIOG_ADDR_CODES].[c_addr_desc], ZZ_SCRATCH_ENTRY.c_addr_desc_chn = [BIOG_ADDR_CODES].[c_addr_desc_chn], " + _
                "ZZ_SCRATCH_ENTRY.c_parental_status_desc = [PARENTAL_STATUS_CODES].[c_parental_status_desc], ZZ_SCRATCH_ENTRY.c_parental_status_desc_chn = [PARENTAL_STATUS_CODES].[c_parental_status_desc_chn], " + _
                "ZZ_SCRATCH_ENTRY.c_entry_addr_name = [ADDR_CODES_1].[c_name], ZZ_SCRATCH_ENTRY.c_entry_addr_chn = [ADDR_CODES_1].[c_name_chn], " + _
                "ZZ_SCRATCH_ENTRY.c_entry_xcoord = [ADDR_CODES_1].[x_coord], ZZ_SCRATCH_ENTRY.c_entry_ycoord = [ADDR_CODES_1].[y_coord], " + _
                "ZZ_SCRATCH_ENTRY.c_dynasty_chn = [DYNASTIES].[c_dynasty_chn], ZZ_SCRATCH_ENTRY.c_dynasty = [DYNASTIES].[c_dynasty], " + _
                "ZZ_SCRATCH_ENTRY.c_source_text = [TEXT_CODES].[c_title], ZZ_SCRATCH_ENTRY.c_source_text_chn = [TEXT_CODES].[c_title_chn]"
        cmdSQL.CommandText = tQstr
        cmdSQL.Execute tRecDeleted

        'calculate_xy_count
        '
        ' use three SQL calls
        '
        cmdSQL.CommandText = "Delete * from tmpXY"
        cmdSQL.Execute tRecDeleted
        '
        tQstr = "INSERT INTO tmpXY ( x_coord, y_coord, CountOfx_coord, CountOfy_coord ) " + _
            "SELECT ZZ_SCRATCH_ENTRY.x_coord, ZZ_SCRATCH_ENTRY.y_coord, Count(ZZ_SCRATCH_ENTRY.x_coord) AS CountOfx_coord, Count(ZZ_SCRATCH_ENTRY.y_coord) AS CountOfy_coord " + _
            "FROM ZZ_SCRATCH_ENTRY " + _
            "GROUP BY ZZ_SCRATCH_ENTRY.x_coord, ZZ_SCRATCH_ENTRY.y_coord;"
        '
        cmdSQL.CommandText = tQstr
        cmdSQL.Execute tRecDeleted
        '
        tQstr = "UPDATE tmpXY INNER JOIN ZZ_SCRATCH_ENTRY ON (tmpXY.y_coord = ZZ_SCRATCH_ENTRY.y_coord) AND (tmpXY.x_coord = ZZ_SCRATCH_ENTRY.x_coord) " + _
                "SET ZZ_SCRATCH_ENTRY.xy_count = [tmpXY].[CountOfx_coord];"

        cmdSQL.CommandText = tQstr
        cmdSQL.Execute tRecDeleted
        '
    End If
 ' now reopen
    '
    Set gRstPeople = CurrentDb.OpenRecordset("ZZ_SCRATCH_ENTRY", dbOpenDynaset)
    
    gRstPeople.MoveFirst
    If gRstPeople.RecordCount > 0 Then
        '
        CmdGIS.Enabled = True
        CmdStoreID.Enabled = True
        CmdNeo4j.Enabled = True
    Else
        CmdGIS.Enabled = False
        CmdStoreID.Enabled = False
        CmdNeo4j.Enabled = False
    End If
    Set Entry_Address_Query.Form.Recordset = gRstPeople

Exit_CmdQuery_Click:
    '
    '  close everything
    '
    Set rst = Nothing
    Set tRstKinCodes = Nothing
    Set tRstAddr = Nothing
    Set tRstBiogMain = Nothing
    Set EntryQuery = Nothing
    Set AddressQuery = Nothing
    Set tRstDummy = Nothing
    Set cmdSQL = Nothing
    Exit Sub

Err_CmdQuery_Click:
    MsgBox Err.Description
    Resume Exit_CmdQuery_Click
    
End Sub
Private Sub calculate_xy_count()
    Dim tX As Double, tY As Double, tXY As Integer, tBM As Variant, tWrite As Integer
    '
    '  the strategy is to first throw a bookmark at the first new value
    '  then count the number, then go back to the bookmark and update each record
    '
    With gRstPeople
        .Index = "xy"
        .MoveFirst
        
        tX = -1#
        tY = -1#
        tXY = 0
        tWrite = 0
        tBM = .Bookmark
        
        Do While Not .EOF
            If tX <> !x_coord Or tY <> !y_coord Then
                If tWrite = 1 Then
                    ' go back to the first record with the value
                    .Bookmark = tBM
                    Do While tX = !x_coord And tY = !y_coord
                        .Edit
                        !xy_count = tXY
                        .Update
                        .MoveNext
                    Loop
                Else
                    tWrite = 1
                End If
                '  reset
                tXY = 0
                tBM = .Bookmark
                tX = !x_coord
                tY = !y_coord
            End If
            '  increment the count and move to the next
            tXY = tXY + 1
            .MoveNext
        Loop
        '
        '  the last xy value still needs to be written
        '
        .Bookmark = tBM
        Do While Not .EOF
            .Edit
            !xy_count = tXY
            .Update
            .MoveNext
        Loop
        '
        ' now repeat the process with the entry xy
        '
        .Index = "entry_xy"
        .MoveFirst
        
        tX = -1#
        tY = -1#
        tXY = 0
        tWrite = 0
        tBM = .Bookmark
        
        Do While Not .EOF
            If tX <> !c_entry_xcoord Or tY <> !c_entry_ycoord Then
                If tWrite = 1 Then
                    ' go back to the first record with the value
                    .Bookmark = tBM
                    Do While tX = !c_entry_xcoord And tY = !c_entry_ycoord
                        .Edit
                        !c_entry_xy_count = tXY
                        .Update
                        .MoveNext
                    Loop
                Else
                    tWrite = 1
                End If
                '  reset
                tXY = 0
                tBM = .Bookmark
                tX = !c_entry_xcoord
                tY = !c_entry_ycoord
            End If
            '  increment the count and move to the next
            tXY = tXY + 1
            .MoveNext
        Loop
        '
        '  the last xy value still needs to be written
        '
        .Bookmark = tBM
        Do While Not .EOF
            .Edit
            !c_entry_xy_count = tXY
            .Update
            .MoveNext
        Loop
        .Index = "IndexYear"
    End With
End Sub
Private Sub CmdPickEntry_Click()
On Error GoTo Err_CmdPickEntry_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim strENTRY As String

    Dim cmdSQL As ADODB.Command, tStrID As String
    Dim varItm As Variant, tCount As Integer
            
    Set cmdSQL = New ADODB.Command
    cmdSQL.ActiveConnection = CurrentProject.Connection
    cmdSQL.CommandType = adCmdText
            
    TxtEntryCode.Visible = True
    TxtEntryCode.SetFocus
    strENTRY = TxtEntryCode.TEXT
    
    stDocName = "frmPickEntry_multi"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , acDialog, strENTRY
    
    If CurrentProject.AllForms("frmPickEntry_multi").IsLoaded Then
        Dim intENTRY As Integer
        Dim strENTRY_DESC As String
        
        'MsgBox "Getting Entry Code value"
        Forms!frmPickEntry_multi.Form!TxtEntryCode.Visible = True
        Forms!frmPickEntry_multi.Form!TxtEntryCode.SetFocus
        intENTRY = Forms!frmPickEntry_multi.Form!TxtEntryCode.Value
        
        Forms!frmPickEntry_multi.Form!subTreeView.SetFocus
        Forms!frmPickEntry_multi.Form!TxtEntryCode.Visible = False
        TxtEntryCode.Value = intENTRY
        
        ' zap the temporary table
            
        cmdSQL.CommandText = "Delete * from zz_scratch_entry_code"
        cmdSQL.Execute tdeleted
        
        'MsgBox "Processing values"
        If TxtEntryCode.Value < 0 Then
            If TxtEntryCode.Value = -1 Then
                TxtEntryDesc.Value = "[[All]]"
                TxtEntryChn.Value = "[[All]]"
            Else
                TxtEntryDesc.Value = "[[Multi-Select]]"
                TxtEntryChn.Value = "[[" + ChrW(22810) + ChrW(36984) + "]]"
            End If
            
            'MsgBox "Getting TxtTypeID"
            Forms!frmPickEntry_multi.Form!TxtTypeID.Visible = True
            Forms!frmPickEntry_multi.Form!TxtTypeID.SetFocus
            strENTRY_DESC = Forms!frmPickEntry_multi.Form!TxtTypeID.Value
            Forms!frmPickEntry_multi.Form!subTreeView.SetFocus
            Forms!frmPickEntry_multi.Form!TxtTypeID.Visible = False
            TxtTypeCode.Value = strENTRY_DESC
            
            If TxtTypeCode.Value = "" Then
                TxtTypeDesc.Value = "[ALL]"
                '
                ' multi-select from the root
                '
                If TxtEntryCode.Value = -2 Then
                    cmdSQL.CommandText = "INSERT INTO ZZ_SCRATCH_ENTRY_CODE ( c_entry_code ) SELECT DISTINCT c_entry_code FROM ZZ_ENTRY_CODE"
                    cmdSQL.Execute tdeleted
                End If
            Else
                'MsgBox "Getting TxtTypeDesc"
                Forms!frmPickEntry_multi.Form!TxtTypeDesc.Visible = True
                Forms!frmPickEntry_multi.Form!TxtTypeDesc.SetFocus
                strENTRY_DESC = Forms!frmPickEntry_multi.Form!TxtTypeDesc.Value
                Forms!frmPickEntry_multi.Form!subTreeView.SetFocus
                Forms!frmPickEntry_multi.Form!TxtTypeDesc.Visible = False
                TxtTypeDesc.Value = strENTRY_DESC
                    
                'MsgBox "Getting TxtTypeChn"
                Forms!frmPickEntry_multi.Form!TxtTypeChn.Visible = True
                Forms!frmPickEntry_multi.Form!TxtTypeChn.SetFocus
                strENTRY_DESC = Forms!frmPickEntry_multi.Form!TxtTypeChn.Value
                Forms!frmPickEntry_multi.Form!subTreeView.SetFocus
                Forms!frmPickEntry_multi.Form!TxtTypeChn.Visible = False
                TxtTypeChn.Value = strENTRY_DESC
                
                cmdSQL.CommandText = "INSERT INTO ZZ_SCRATCH_ENTRY_CODE ( c_entry_code ) SELECT DISTINCT c_entry_code FROM ZZ_ENTRY_CODE"
                cmdSQL.Execute tdeleted
                
            End If
        Else
            ' extract the entry code
            '
            cmdSQL.CommandText = "INSERT INTO ZZ_SCRATCH_ENTRY_CODE ( c_entry_code ) SELECT " + Str(intENTRY) + " as c_entry_code"
            cmdSQL.Execute tdeleted
            
            Forms!frmPickEntry_multi.Form!TxtEntryDesc.Visible = True
            Forms!frmPickEntry_multi.Form!TxtEntryDesc.SetFocus
            strENTRY_DESC = Forms!frmPickEntry_multi.Form!TxtEntryDesc.Value
            Forms!frmPickEntry_multi.Form!subTreeView.SetFocus
            Forms!frmPickEntry_multi.Form!TxtEntryDesc.Visible = False
            TxtEntryDesc.Value = strENTRY_DESC
                
            'MsgBox "Getting TxtEntryDescChn"
            Forms!frmPickEntry_multi.Form!TxtEntryChn.Visible = True
            Forms!frmPickEntry_multi.Form!TxtEntryChn.SetFocus
            strENTRY_DESC = Forms!frmPickEntry_multi.Form!TxtEntryChn.Value
            Forms!frmPickEntry_multi.Form!subTreeView.SetFocus
            Forms!frmPickEntry_multi.Form!TxtEntryChn.Visible = False
            TxtEntryChn.Value = strENTRY_DESC
            
            'MsgBox "Getting TxtTypeDesc"
            Forms!frmPickEntry_multi.Form!TxtTypeDesc.Visible = True
            Forms!frmPickEntry_multi.Form!TxtTypeDesc.SetFocus
            strENTRY_DESC = Forms!frmPickEntry_multi.Form!TxtTypeDesc.Value
            Forms!frmPickEntry_multi.Form!subTreeView.SetFocus
            Forms!frmPickEntry_multi.Form!TxtTypeDesc.Visible = False
            TxtTypeDesc.Value = strENTRY_DESC
                        
            'MsgBox "Getting TxtTypeChn"
            Forms!frmPickEntry_multi.Form!TxtTypeChn.Visible = True
            Forms!frmPickEntry_multi.Form!TxtTypeChn.SetFocus
            strENTRY_DESC = Forms!frmPickEntry_multi.Form!TxtTypeChn.Value
            Forms!frmPickEntry_multi.Form!subTreeView.SetFocus
            Forms!frmPickEntry_multi.Form!TxtTypeChn.Visible = False
            TxtTypeChn.Value = strENTRY_DESC
            TxtTypeCode.Value = ""
        End If
        
        ' now enable the search
        
        CmdQuery.Enabled = True
        CmdSaveEntryCodes.Enabled = True
                
        DoCmd.Close acForm, stDocName
    End If
            
    CmdPickEntry.SetFocus
    TxtEntryCode.Visible = False
        
Exit_CmdPickEntry_Click:
    Exit Sub

Err_CmdPickEntry_Click:
    MsgBox Err.Description
    Resume Exit_CmdPickEntry_Click
    
End Sub


Private Sub CmdSaveEntryCodes_Click()
On Error GoTo Err_CmdSaveEntryCodes_Click
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
    
    dlgSaveAs.InitialFileName = "entry_id_list.txt"
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
            GoTo Exit_CmdSaveEntryCodes_Click
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
        tStr = "SELECT ZZ_SCRATCH_ENTRY_CODE.c_entry_code, ENTRY_CODES.c_entry_desc, ENTRY_CODES.c_entry_desc_chn " + _
            "FROM ZZ_SCRATCH_ENTRY_CODE INNER JOIN ENTRY_CODES ON ZZ_SCRATCH_ENTRY_CODE.c_entry_code = ENTRY_CODES.c_entry_code"

        Set tRstIDs = CurrentDb.OpenRecordset(tStr, dbOpenDynaset)
        '
        tTab = Chr(9)
        
        With tRstIDs
            
            .MoveFirst
            ' MsgBox "writing file"
            Do While Not .EOF
                '
                tStr = Str(!c_entry_code) + tTab + !c_entry_desc + tTab + !c_entry_desc_chn
                tStream.WriteText tStr, adWriteLine
                .MoveNext
            Loop
        End With
        ' now make sure all the data is copied to tStream
        tStream.Flush
        ' and write the stream to the file
        tStream.Position = 3
        MsgBox "Copying to stream"
        tStream.CopyTo tStreamNoBOM
        MsgBox "Writing to file"
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
    
Exit_CmdSaveEntryCodes_Click:
    Exit Sub

Err_CmdSaveEntryCodes_Click:
    MsgBox Err.Description
    Resume Exit_CmdSaveEntryCodes_Click

End Sub

Private Sub CmdStoreID_Click()
    Dim cmdSQL As ADODB.Command, tStrQuery As String
    
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

    tStrQuery = "INSERT INTO ZZ_STORE_PERSON_ID ( c_personid ) SELECT DISTINCT ZZ_SCRATCH_ENTRY.c_personid FROM ZZ_SCRATCH_ENTRY"
    
    cmdSQL.CommandText = tStrQuery
    cmdSQL.Execute tRecCount
    MsgBox "Person IDs successfully stored.  Click on 'Recall Person IDs' to reuse these IDs in other forms."
    '
    '  update storage source
    cmdSQL.CommandText = "UPDATE PersonIDSource SET SourceForm ='Entry' WHERE PersonIDSource.LineNum =1"
    cmdSQL.Execute tRecCount

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

Private Sub Form_Open(Cancel As Integer)
    Dim cmdSQL As ADODB.Command
    Dim tRstEntryCode As DAO.Recordset, tRstDummy As DAO.Recordset
    
    Set cmdSQL = New ADODB.Command
    '
    '  to clear the table, briefly close and then delete records
    '
    Set tRstEntryCode = Entry_Address_Query.Form.Recordset
    Set tRstDummy = CurrentDb.OpenRecordset("Z_SCRATCH_DUMMY_EC", dbOpenDynaset)
    Set Entry_Address_Query.Form.Recordset = tRstDummy
    tRstEntryCode.Close
    '
    cmdSQL.ActiveConnection = CurrentProject.Connection
    cmdSQL.CommandType = adCmdText
        '
    cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_ENTRY"
    cmdSQL.Execute tRecDeleted
    '
    '  now reopen
    '
    Set tRstEntryCode = CurrentDb.OpenRecordset("ZZ_SCRATCH_ENTRY", dbOpenDynaset)
    Set Entry_Address_Query.Form.Recordset = tRstEntryCode
    
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
    
    gFromDynasty = -1
    gToDynasty = -1
    gUseIndexYears = False
    gUseDynasties = False
End Sub
Private Sub CmdExit_Click()
On Error GoTo Err_CmdExit_Click


    DoCmd.Close

Exit_CmdExit_Click:
    Exit Sub

Err_CmdExit_Click:
    MsgBox Err.Description
    Resume Exit_CmdExit_Click
    
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
    Dim tLabelLanguage(3, 34) As String, tLang As Integer
    
    Dim tRstLabelList As DAO.Recordset, ti As Integer
    
    Set tRstLabelList = CurrentDb.OpenRecordset("FormLabels", dbOpenTable)
    
    tRstLabelList.Index = "label"
    
    gLabelsOK = False
    With tRstLabelList
        .MoveFirst
        ti = 1
        
        Do While ti < 34 And Not .EOF
            If !c_form = "LAE" Then
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
        Me.CmdPickEntry.Caption = tLabelLanguage(tLang, 4)
        Me.CmdQuery.Caption = tLabelLanguage(tLang, 5)
        Me.CmdGIS.Caption = tLabelLanguage(tLang, 6)
        Me.CmdFanti.Caption = tLabelLanguage(tLang, 8)
        Me.CmdJianti.Caption = tLabelLanguage(tLang, 9)
        Me.CmdExit.Caption = tLabelLanguage(tLang, 10)
        Me.CmdSelectPlace.Caption = tLabelLanguage(tLang, 11)
        Me.CmdImportPlaces.Caption = tLabelLanguage(tLang, 12)
        Me.CmdAllPlaces.Caption = tLabelLanguage(tLang, 13)
        Me.LblIndexYears.Caption = tLabelLanguage(tLang, 14)
        ' Me.LblUseYears.Caption = tLabelLanguage(tLang, 15)
        Me.LblDisplay.Caption = tLabelLanguage(tLang, 16)
        Me.LblExamYears.Caption = tLabelLanguage(tLang, 17)
        Me.LblUsePersonAddr.Caption = tLabelLanguage(tLang, 18)
        Me.LblUseEntryAddr.Caption = tLabelLanguage(tLang, 19)
        Me.CmdStoreID.Caption = tLabelLanguage(tLang, 20)
        Me.CmdHelp.Caption = tLabelLanguage(tLang, 21)
        Me.Label37.Caption = tLabelLanguage(tLang, 22)
        Me.LblChkSubUnits.Caption = tLabelLanguage(tLang, 23)
        Me.LblDynasties.Caption = tLabelLanguage(tLang, 24)
        Me.CmdFromDynasty.Caption = tLabelLanguage(tLang, 25)
        Me.CmdToDynasty.Caption = tLabelLanguage(tLang, 26)
        Me.CmdAllDynasties.Caption = tLabelLanguage(tLang, 27)
        Me.LblYears.Caption = tLabelLanguage(tLang, 28)
        Me.LblOptNoDates.Caption = tLabelLanguage(tLang, 29)
        Me.LblOptDynasties.Caption = tLabelLanguage(tLang, 30)
        Me.CmdNeo4j.Caption = tLabelLanguage(tLang, 31)
        Me.CmdImportEntryCodes.Caption = tLabelLanguage(tLang, 32)
        Me.CmdSaveEntryCodes.Caption = tLabelLanguage(tLang, 33)
    End If
    
End Sub
Private Sub CmdAllPlaces_Click()
On Error GoTo Err_CmdAllPlaces_Click

        TxtAddrID.Value = -1
                
        TxtPlaceChn.Value = ""
        TxtPlace.Value = ""
        gUseADDRID = False
        ChkXYRef.Enabled = False
        ChkSubUnits.Enabled = False
     
Exit_CmdAllPlaces_Click:
    Exit Sub

Err_CmdAllPlaces_Click:
    MsgBox Err.Description
    Resume Exit_CmdAllPlaces_Click
  
End Sub
Private Sub CmdImportPlaces_Click()
    On Error GoTo Err_CmdImportPlaces_Click
    
    Dim stDocName As String, tRstAddresses As DAO.Recordset
    Dim stLinkCriteria As String, tRstImportPlaces As DAO.Recordset
    Dim tString As String, tAddrID As Long, ti As Integer, tStrID As String
    Dim tLen As Integer, cmdSQL As ADODB.Command

    Dim dlgSaveAs As FileDialog
    Dim tFileNum As Integer
    Dim tFileName As String, tFN As Variant
    Dim tFileSystem, tList
    
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
            '
        End If
    End With
    '
    ' Clear the address table now that we are ready to go
    '
    Set cmdSQL = New ADODB.Command
    cmdSQL.ActiveConnection = CurrentProject.Connection
    cmdSQL.CommandType = adCmdText
    '
    cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_ADDR"
    cmdSQL.Execute tRecDeleted
    '
       
    cmdSQL.CommandText = "Delete * from InputErrorList"
    cmdSQL.Execute tRecDeleted
        
    cmdSQL.CommandText = "Delete * from TempImportList"
    cmdSQL.Execute tRecDeleted
        
    DoCmd.TransferText acImportDelim, "ImportPlaceList_Space", "TempImportList", tFileName, 0
    '    TransferType=acImportDelim
    '    SpecificationName = "ImportPlaceList_Space" (apparently it is saved in the database itself)
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
    tStrSQL = "INSERT INTO ZZ_SCRATCH_ADDR ( c_addr_id ) SELECT DISTINCT TempImportList.ImportID " + _
        "FROM ADDR_CODES INNER JOIN TempImportList ON ADDR_CODES.c_addr_id = TempImportList.ImportID"

    cmdSQL.CommandText = tStrSQL
    cmdSQL.Execute tRecDeleted
        
    If tRecDeleted > 0 Then
        Me.TxtPlace.Value = "[Imported List]"
        Me.TxtPlaceChn.Value = "[Imported List]"
        gUseADDRID = True
        ChkXYRef.Enabled = True
        ChkSubUnits.Enabled = True
    End If
        
    Set cmdSQL = Nothing
    Set tFileSystem = Nothing
    
Exit_CmdImportPlaces_Click:
    Exit Sub

Err_CmdImportPlaces_Click:
    MsgBox Err.Description
    Resume Exit_CmdImportPlaces_Click
        
End Sub
Private Sub CmdSelectPlace_Click()
On Error GoTo Err_CmdSelectPlace_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim strADDR As String

    TxtAddrID.Visible = True
    TxtAddrID.SetFocus
    strADDR = TxtAddrID.TEXT

    stDocName = "frmPickAddresses_multi"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , acDialog, strADDR
    
    If CurrentProject.AllForms("frmPickAddresses_multi").IsLoaded Then
    
        '  if the user selected a group of addresses, ZZ_ADDRESSES will have records
            
        Dim tAddrID As Long, tRstAddr As DAO.Recordset
        Dim cmdSQL As ADODB.Command
                
        Set cmdSQL = New ADODB.Command
        cmdSQL.ActiveConnection = CurrentProject.Connection
        cmdSQL.CommandType = adCmdText
        
        gUseADDRID = True
        CmdAllPlaces.Enabled = True
        ChkXYRef.Enabled = True
        ChkSubUnits.Enabled = True
        
        'MsgBox "Checking zz_addresses"
        
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
        cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_ADDR"
        cmdSQL.Execute tRecDeleted
            
        cmdSQL.CommandText = "INSERT INTO ZZ_SCRATCH_ADDR ( c_addr_id ) SELECT DISTINCT " + _
            "ZZ_ADDRESSES.c_addr_id FROM ZZ_ADDRESSES"
        cmdSQL.Execute tRecDeleted
        
        DoCmd.Close acForm, "frmPickAddresses_multi"
        
    End If
    CmdSelectPlace.SetFocus
    TxtAddrID.Visible = False

Exit_CmdSelectPlace_Click:
    Exit Sub

Err_CmdSelectPlace_Click:
    MsgBox Err.Description
    Resume Exit_CmdSelectPlace_Click
    
End Sub

Private Sub writeKML()
'<kml xmlns="http://www.opengis.net/kml/2.2">
'<Document>
'   <name>ExtendedData+SchemaData</name>
'   <open>1</open>
'   <!-- Create a balloon template referring to the user-defined type -->
'   <Style id="assoc-balloon-template">
'       <BalloonStyle>
'           <text>
'              <![CDATA[
'              $[AssocPerson/PersonNameHZ] <br/>
'              ID: $[AssocPerson/PersonID] <br/>
'              Index Year: $[AssocPerson/IndexYear] <br/>
'              Address: $[AssocPerson/AddrName] $[AssocPerson/AddrNameHZ] <br/>
'              XY Count: $[AssocPerson/XYCount] <br/><br/>
'               ]]>
'           </text>
'       </BalloonStyle>
'   </Style>
'   <!-- Declare the type "AssocPerson" with 6 fields -->
'   <Schema name="AssocPerson" id="AssocPersonId">
'       <SimpleField type="string" name="PersonNameHZ">
'           <displayName><![CDATA[<b>Person</b>]]></displayName>
'       </SimpleField>
'       <SimpleField type="string" name="AddrName">
'           <displayName><![CDATA[<b>Person</b>]]></displayName>
'       </SimpleField>
'       <SimpleField type="string" name="AddrNameHZ">
'           <displayName><![CDATA[<b>Person</b>]]></displayName>
'       </SimpleField>
'       <SimpleField type="uint" name="PersonID">
'           <displayName><![CDATA[ID]]></displayName>
'       </SimpleField>
'       <SimpleField type="int" name="IndexYear">
'           <displayName><![CDATA[Index Year]]></displayName>
'       </SimpleField>
'       <SimpleField type="int" name="XYCount">
'           <displayName><![CDATA[XY Count]]></displayName>
'       </SimpleField>
'   </Schema>
'   <!-- Instantiate some Placemarks extended with AssocPerson fields -->
'   <Placemark>
'       <name>Easy trail</name>
'       <styleUrl>#assoc-balloon-template</styleUrl>
'       <ExtendedData>
'           <SchemaData schemaUrl="#AssocPersonId">
'               <SimpleData name="PersonID">3.14159</SimpleData>
'               <SimpleData name="PersonNameHZ">Pi in the sky</SimpleData>
'               <SimpleData name="IndexYear">10</SimpleData>
'               <SimpleData name="AddrName">Pi in the sky</SimpleData>
'               <SimpleData name="AddrNameHZ">Pi in the sky</SimpleData>
'               <SimpleData name="XYCount">10</SimpleData>
'           </SchemaData>
'       </ExtendedData>
'       <Point>
'           <coordinates>-122.000,37.002</coordinates>
'       </Point>
'   </Placemark>
'   <Placemark>
'       <name>Difficult trail</name>
'       <styleUrl>#assoc-balloon-template</styleUrl>
'       <ExtendedData>
'           <SchemaData schemaUrl="#AssocPersonId">
'               <SimpleData name="TrailHeadName">Mount Everest</SimpleData>
'               <SimpleData name="TrailLength">347.45</SimpleData>
'               <SimpleData name="ElevationGain">10000</SimpleData>
'           </SchemaData>
'       </ExtendedData>
'       <Point>
'           <coordinates>-121.998,37.0078</coordinates>
'       </Point>
'   </Placemark>
'</Document>
'</kml>

    Dim tStrKML As String
    '
    '  This program will dump the results to a .gis file
    '
    If Entry_Address_Query.Form.Recordset.RecordCount = 0 Then
        MsgBox "There are no records to save."
        GoTo Exit_writeKML
    End If
    '
    Dim tStream As ADODB.Stream
    Set tStream = New ADODB.Stream
    
    If GISFrame.Value = 1 Then
        tStream.Charset = "utf-8"
        tCodeStr = "UTF8"
    Else
        tStream.Charset = "gb2312"
        tCodeStr = "GB2312"
    End If
    tStream.Mode = adModeReadWrite
    tStream.Type = adTypeText
    tStream.Open
    '
    '  next get a file
    '
    Dim dlgSaveAs As FileDialog
    Dim tFileNum As Integer
    Dim tFileName As String, tFN As Variant, tFemale As String
    Dim tRstNode As DAO.Recordset
    Dim tStr As String, tC As String, tDQ As String, ti As Integer
    Dim tFileSystem, tGDF
    
    Set dlgSaveAs = Application.FileDialog(msoFileDialogSaveAs)

    dlgSaveAs.InitialFileName = "entry_gis_" + tCodeStr + ".kml"
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
        'Name,NameChn,Female,IndexYear,AddrName,AddrChn,X,Y,xy_count,NodeDist
        'Name,NameChn,IndexYear,EntryDesc,EntryChn,EntryYear,
        'AddrName,AddrChn,X,Y,xy_count
        '
        ' process the table
        '
        Set tRstNode = Entry_Address_Query.Form.Recordset
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
        tStream.WriteText tC + "<Style id=" + tDQ + "entry-balloon-template" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + "<BalloonStyle>", adWriteLine
        tStream.WriteText tC + tC + tC + "<text>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "<![CDATA[", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "$[EntryPerson/PersonNameHZ] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "ID: $[EntryPerson/PersonID] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "Index Year: $[EntryPerson/IndexYear] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "Entry Year: $[EntryPerson/EntryYear] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "Entry Desc: $[EntryPerson/EntryDesc] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "Entry Chn: $[EntryPerson/EntryDescHZ] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "Entry Rank: $[EntryPerson/EntryRank] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "Address: $[EntryPerson/AddrName] $[EntryPerson/AddrNameHZ] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "XY Count: $[EntryPerson/XYCount] <br/><br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "]]>", adWriteLine
        tStream.WriteText tC + tC + tC + "</text>", adWriteLine
        tStream.WriteText tC + tC + "</BalloonStyle>", adWriteLine
        tStream.WriteText tC + "</Style>", adWriteLine
        tStream.WriteText tC + "<!-- Declare the type " + tDQ + "EntryPerson" + tDQ + " with 10 fields -->", adWriteLine
        tStream.WriteText tC + "<Schema name=" + tDQ + "EntryPerson" + tDQ + " id=" + tDQ + "EntryPersonId" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "PersonNameHZ" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[<b>Person</b>]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "AddrName" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[<b>Person</b>]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "AddrNameHZ" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[<b>Person</b>]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "uint" + tDQ + " name=" + tDQ + "PersonID" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[ID]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "IndexYear" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Index Year]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "int" + tDQ + " name=" + tDQ + "EntryYear" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Entry Year]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "EntryDesc" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Entry Desc]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "EntryDescHZ" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Entry Chn]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "int" + tDQ + " name=" + tDQ + "EntryRank" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Entry Rank]]></displayName>", adWriteLine
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
                    tStr = "[Bad Data]"
                Else
                    tStr = !c_name
                End If
                tStream.WriteText tC + tC + "<name>" + tStr + "</name>", adWriteLine
                
                tStream.WriteText tC + tC + "<styleUrl>#entry-balloon-template</styleUrl>", adWriteLine
                '
                '  Index Year as time stamp
                '
                If IsNull(!c_index_year) Then
                    tStr = "N/A"
                Else
                    tStr = Str(!c_index_year)
                End If
                tStream.WriteText tC + tC + "<TimeStamp>" + tStr + "</TimeStamp>", adWriteLine
                '
                tStream.WriteText tC + tC + "<ExtendedData>", adWriteLine
                tStream.WriteText tC + tC + tC + "<SchemaData schemaUrl=" + tDQ + "#EntryPersonId" + tDQ + ">", adWriteLine
                '
                '  person ID
                '
                tStr = Str(!c_personid)
                tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "PersonID" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                '
                '  Chinese Name
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
                tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "PersonNameHZ" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
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
                '  Entry Year
                '
                If IsNull(!c_year) Then
                    tStr = "-2000"
                Else
                    tStr = Str(!c_year)
                End If
                tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "EntryYear" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                '
                '  Entry Desc
                '
                If IsNull(!c_entry_desc) Then
                    tStr = "[Missing Data]"
                Else
                    tStr = !c_entry_desc
                End If
                tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "EntryDesc" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                '
                '  Entry Chn
                '
                If IsNull(!c_entry_chn) Then
                    tStr = "[Missing Data]"
                Else
                    tStr = !c_entry_chn
                End If
                tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "EntryDescHZ" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                '
                '  Entry Rank
                '
                If IsNull(!c_exam_rank) Then
                    tStr = "0"
                Else
                    tStr = !c_exam_rank
                End If
                tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "EntryRank" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
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
                tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "AddrNameHZ" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
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

Private Sub FrameYears_Click()
    
    ' Turn off usage
    gUseEntryYears = False
    gUseIndexYears = False
    gUseDynasties = False
    
    ' Turn off Dynasty text boxes
    
    Me.TxtFromDynasty.Enabled = False
    Me.TxtFromDynastyPY.Enabled = False
    Me.TxtToDynasty.Enabled = False
    Me.TxtToDynastyPY.Enabled = False
    Me.TxtFromDynasty.Locked = False
    Me.TxtFromDynastyPY.Locked = falsee
    Me.TxtToDynasty.Locked = False
    Me.TxtToDynastyPY.Locked = False
    
    If FrameYears.Value = 1 Or FrameYears.Value = 2 Then
        ' entry years or index years
        Me.CmdFromDynasty.Enabled = False
        Me.CmdToDynasty.Enabled = False
        Me.CmdAllDynasties.Enabled = False
        
        Me.TxtFromYear.Enabled = True
        Me.TxtToYear.Enabled = True
        If FrameYears.Value = 1 Then
            gUseEntryYears = True
        Else
            gUseIndexYears = True
        End If
    ElseIf FrameYears.Value = 3 Then
        '  enable dynasties
        Me.CmdFromDynasty.Enabled = True
        Me.CmdToDynasty.Enabled = True
        Me.CmdAllDynasties.Enabled = True
        Me.TxtFromDynasty.Locked = True
        Me.TxtFromDynastyPY.Locked = True
        Me.TxtToDynasty.Locked = True
        Me.TxtToDynastyPY.Locked = True
        ' diaable index years
        Me.TxtFromYear.Enabled = False
        Me.TxtToYear.Enabled = False
        
        gUseDynasties = True
    Else
        '  disable all
        Me.CmdFromDynasty.Enabled = False
        Me.CmdToDynasty.Enabled = False
        Me.CmdAllDynasties.Enabled = False
        '
        Me.TxtFromYear.Enabled = False
        Me.TxtToYear.Enabled = False
    End If

End Sub

