Option Compare Database
Public gRstPeople As DAO.Recordset, gDisplayLanguage As String, gLabelsOK As Boolean
Public gImportPlaces As Boolean, gUseADDRID As Boolean, gFromStr As String, gToStr As String
Public gStatusCodeStr As String, gStatusTypeStr As String, gFromDynasty As Integer, gToDynasty As Integer
Public gFromDynastyBegin As Integer, gFromDynastyEnd As Integer, gToDynastyBegin As Integer, gToDynastyEnd As Integer




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

Private Sub CmdImportStatusCodes_Click()
On Error GoTo Err_CmdImportStatusCodes_Click
    
    Dim stDocName As String, tRstStatus As DAO.Recordset
    Dim stLinkCriteria As String, tRstImportStatusCodes As DAO.Recordset
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
                    GoTo Exit_CmdImportStatusCodes_Click
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
        cmdSQL.CommandText = "Delete * from ZZ_STATUS_CODE"
        cmdSQL.Execute tRecDeleted
        '
        cmdSQL.CommandText = "Delete * from InputErrorList"
        cmdSQL.Execute tRecDeleted
        
        cmdSQL.CommandText = "Delete * from TempImportList"
        cmdSQL.Execute tRecDeleted
        
        DoCmd.TransferText acImportDelim, "StatusCodeListImport Specification", "TempImportList", tFileName, 0
        '    TransferType=acImportDelim
        '    SpecificationName = "TempImportList" (apparently it is saved in the database itself)
        '    TableName = "TempImportList"  (probably requires that I drop the table first, but I can test)
        '    HasFieldNames = False (0)
        '
        '  copy the bad IDs
        '
        tStrSQL = "INSERT INTO InputErrorList ( c_ID ) SELECT TempImportList.ImportID " + _
            "FROM STATUS_CODES RIGHT JOIN TempImportList ON STATUS_CODES.c_status_code = TempImportList.ImportID " + _
            "WHERE (((STATUS_CODES.c_status_code) Is Null))"

        cmdSQL.CommandText = tStrSQL
        cmdSQL.Execute tRecDeleted
        
        If tRecDeleted > 0 Then
            MsgBox "Some ID were not successfully imported:  please look at InputErrorList."
        End If
        '
        '  copy the good IDs
        '
        tStrSQL = "INSERT INTO ZZ_STATUS_CODE ( c_status_code ) SELECT DISTINCT TempImportList.ImportID " + _
            "FROM STATUS_CODES INNER JOIN TempImportList ON STATUS_CODES.c_status_code = TempImportList.ImportID"

        cmdSQL.CommandText = tStrSQL
        cmdSQL.Execute tRecDeleted
        
        Me.TxtTypeDesc.Value = ""
        Me.TxtTypeChn.Value = ""
        If tRecDeleted > 0 Then
            Me.TxtStatusDesc.Value = "[Imported List]"
            Me.TxtStatusChn.Value = "[Imported List]"
            Me.CmdQuery.Enabled = True
            Me.CmdSaveStatusCodes.Enabled = True
        Else
            Me.TxtStatusDesc.Value = ""
            Me.TxtStatusChn.Value = ""
            Me.CmdQuery.Enabled = False
            Me.CmdSaveStatusCodes.Enabled = False
        End If
        
        Set cmdSQL = Nothing
    End If
    
Exit_CmdImportStatusCodes_Click:
    Exit Sub

Err_CmdImportStatusCodes_Click:
    MsgBox Err.Description
    Resume Exit_CmdImportStatusCodes_Click

End Sub

Private Sub CmdNeo4j_Click()
On Error GoTo Err_CmdNeo4j_Click
    '
    '  This program will dump the results of the search to five CSV files
    '
    ' warn the user that a lot of files will be created
    '
    'MsgBox "Neo4j requires that from 6 to 9 files be created."
    '
    '  allocate the file variables
    '
    Dim dlgSaveAs As FileDialog
    Dim tFileNum As Integer, tFileName As String, tFN As Variant
    '
    '  next get the People file
    '
    Dim tRstPeople As DAO.Recordset, tRstStatus As DAO.Recordset, tRstStatusCodes As DAO.Recordset, tRstPlace As DAO.Recordset
    Dim tRstPeopleStatus As DAO.Recordset, tRstPeoplePlace As DAO.Recordset, tStr As String, tC As String
    Dim tQueryStr As String, tPersonID As Long
    Dim gStream As ADODB.Stream, tCodeStr As String
    '
    ' set up the stream to write to
    
    Set gStream = New ADODB.Stream
    '
    If GISFrame.Value = 1 Then
        gStream.Charset = "utf-8"
        tCodeStr = "UTF8"
    ElseIf GISFrame.Value = 2 Then
        gStream.Charset = "big5"
        tCodeStr = "BIG5"
    ElseIf GISFrame.Value = 3 Then
        gStream.Charset = "gb2312"
        tCodeStr = "GB2312"
    Else
        gStream.Charset = "ascii"
        tCodeStr = "ascii"
    End If
    '
    tC = Chr(44) ' the comma
    '
    '  prepare the temp tables for the people, place, peoplePlace and entry data
            
    Dim cmdSQL As ADODB.Command
    Set cmdSQL = New ADODB.Command
    cmdSQL.ActiveConnection = CurrentProject.Connection
    cmdSQL.CommandType = adCmdText
    '
    Set tRstPeopleStatus = CurrentDb.OpenRecordset("ZZ_SCRATCH_STATUS", dbOpenDynaset)
    Set tRstPeople = CurrentDb.OpenRecordset("ZZ_SCRATCH_P_STATUS", dbOpenDynaset)
    
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
            tStr = "NameID" + tC + "NamePY" + tC + "IndexYear" + tC + "Dynasty" + tC + "Sex"
        Else
            tStr = "NameID" + tC + "NameHZ" + tC + "NamePY" + tC + "IndexYear" + tC + "Dynasty" + tC + "Sex"
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
                If IsNull(!c_sex) Then
                    tStr = tStr + "Missing"
                Else
                    tStr = tStr + !c_sex
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
    ' now the PeopleEntry file
    '
    dlgSaveAs.InitialFileName = "PeopleStatus_" + tCodeStr + ".csv"
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

        tStr = "NameID" + tC + "StatusCode" + tC + "FirstYear" + tC + "LastYear"
        gStream.WriteText tStr, adWriteLine
        '
        With tRstPeopleStatus
            .MoveFirst
            Do While Not .EOF
                '  the ID of the person
                tStr = Trim(Str(!c_personid)) + tC
                '
                '  entry code
                '
                If IsNull(!c_status_code) Then
                    tStr = tStr + "0" + tC
                Else
                    tStr = tStr + Trim(Str(!c_status_code)) + tC
                End If
                '
                '  first year
                '
                If IsNull(!c_firstyear) Then
                    tStr = tStr + "0" + tC
                Else
                    tStr = tStr + Trim(Str(!c_firstyear)) + tC
                End If
                '
                '  last year
                '
                If IsNull(!c_lastyear) Then
                    tStr = tStr + "0"
                Else
                    tStr = tStr + Trim(Str(!c_lastyear))
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
        tQueryStr = "SELECT DISTINCT ZZ_SCRATCH_P_STATUS.c_addr_id, ZZ_SCRATCH_P_STATUS.c_addr_name, ZZ_SCRATCH_P_STATUS.c_addr_chn, " + _
                        "ZZ_SCRATCH_P_STATUS.x_coord, ZZ_SCRATCH_P_STATUS.y_coord " + _
                    "FROM ZZ_SCRATCH_P_STATUS " + _
                    "WHERE (ZZ_SCRATCH_P_STATUS.c_addr_id > 0)"
        '
        Set tRstPlace = CurrentDb.OpenRecordset(tQueryStr, dbOpenDynaset)
        '
        If tCodeStr = "ascii" Then
            tStr = "PlaceID" + tC + "PlacePY" + tC + "PlaceX" + tC + "PlaceY"
        Else
            tStr = "PlaceID" + tC + "PlacePY" + tC + "PlaceHZ" + tC + "PlaceX" + tC + "PlaceY"
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
                        
                    If IsNull(!x_coord) Then
                        tStr = tStr + "0.0" + tC
                    Else
                        tStr = tStr + Str(!x_coord) + tC
                    End If
                        
                    If IsNull(!y_coord) Then
                        tStr = tStr + "0.0"
                    Else
                        tStr = tStr + Str(!y_coord)
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
        tQueryStr = "SELECT DISTINCT ZZ_SCRATCH_P_STATUS.c_person_id, BIOG_MAIN.c_index_addr_id, " + _
                        "BIOG_MAIN.c_index_addr_type_code " + _
                    "FROM ZZ_SCRATCH_P_STATUS INNER JOIN BIOG_MAIN ON ZZ_SCRATCH_P_STATUS.c_person_id = BIOG_MAIN.c_personid " + _
                    "WHERE (BIOG_MAIN.c_index_addr_type_code > 0)"

        Set tRstPeoplePlace = CurrentDb.OpenRecordset(tQueryStr)
        
        tStr = "NameID" + tC + "PlaceID" + tC + "PersonPlaceCode"
            
        gStream.WriteText tStr, adWriteLine
            
        With tRstPeoplePlace
            .MoveFirst
            Do While Not .EOF
                If Not IsNull(!c_index_addr_id) Then
                    '
                    tStr = Trim(Str(!c_person_id)) + tC
                        '
                    tStr = tStr + Trim(Str(!c_index_addr_id)) + tC
                    '
                    tStr = tStr + Trim(Str(!c_index_addr_type_code))
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
        tQueryStr = "SELECT BIOG_MAIN.c_index_addr_type_code, BIOG_ADDR_CODES.c_addr_desc, BIOG_ADDR_CODES.c_addr_desc_chn " + _
            "FROM ZZ_SCRATCH_P_STATUS INNER JOIN (BIOG_MAIN INNER JOIN BIOG_ADDR_CODES " + _
            "ON BIOG_MAIN.c_index_addr_type_code = BIOG_ADDR_CODES.c_addr_type) " + _
            "ON ZZ_SCRATCH_P_STATUS.c_person_id = BIOG_MAIN.c_personid " + _
                    "WHERE (BIOG_MAIN.c_index_addr_type_code > 0)"

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
                If Not IsNull(!c_index_addr_type_code) Then
                    '
                    tStr = Trim(Str(!c_index_addr_type_code)) + tC
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
    ' finally, get status codes
    '
    dlgSaveAs.InitialFileName = "StatusCode_" + tCodeStr + ".csv"
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
            tStr = "StatusCode" + tC + "StatusDesc"
        Else
            tStr = "StatusCode" + tC + "StatusDesc" + tC + "StatusDescHZ"
        End If
        gStream.WriteText tStr, adWriteLine
        '
        ' get the codes
        '
        tQueryStr = "SELECT DISTINCT ZZ_SCRATCH_STATUS.c_status_code, ZZ_SCRATCH_STATUS.c_status_desc, ZZ_SCRATCH_STATUS.c_status_desc_chn " + _
                    "FROM ZZ_SCRATCH_STATUS " + _
                    "WHERE (ZZ_SCRATCH_STATUS.c_status_code > 0)"
                    
        Set tRstStatusCodes = CurrentDb.OpenRecordset(tQueryStr)
        With tRstStatusCodes
            .MoveFirst
            Do While Not .EOF
                '
                tStr = Trim(Str(!c_status_code)) + tC
                '
                '  entry desc
                '
                If IsNull(!c_status_desc) Then
                    tStr = tStr + "Missing"
                Else
                    tStr = tStr + Trim(!c_status_desc)
                End If
                '
                '  kin ID
                '
                If Not (tCodeStr = "ascii") Then
                    tStr = tStr + tC + Trim(!c_status_desc_chn)
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
    
    MsgBox "Finished saving to Neo4j"
    '
    'Set the object variable to Nothing.
    Set dlgSaveAs = Nothing
    
Exit_CmdNeo4j_Click:
    Exit Sub

Err_CmdNeo4j_Click:
    MsgBox Err.Description
    Resume Exit_CmdNeo4j_Click

End Sub

Private Sub CmdPickStatus_Click()
    On Error GoTo Err_CmdPickStatus_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim strStatus As String

    TxtStatusCode.Visible = True
    TxtStatusCode.SetFocus
    strStatus = TxtStatusCode.TEXT
    
    stDocName = "frmPickStatus_multi"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , acDialog, strStatus
    
    If CurrentProject.AllForms("frmPickStatus_multi").IsLoaded Then
        Dim intStatus As Integer
        Dim strStatus_DESC As String
            
        Forms!frmPickStatus_multi.Form!TxtStatusID.Visible = True
        Forms!frmPickStatus_multi.Form!TxtStatusID.SetFocus
        intStatus = Forms!frmPickStatus_multi.Form!TxtStatusID.Value
        Forms!frmPickStatus_multi.Form!subTreeView.SetFocus
        Forms!frmPickStatus_multi.Form!TxtStatusID.Visible = False
        TxtStatusCode.Value = intStatus
        gStatusCodeStr = Trim(Str(intStatus))
        
        If TxtStatusCode.Value < 0 Then
            If TxtStatusCode.Value = -1 Then
                TxtStatusDesc.Value = "[[All]]"
                TxtStatusChn.Value = "[[All]]"
            Else
                TxtStatusDesc.Value = "[[Multi-Select]]"
                TxtStatusChn.Value = "[[" + ChrW(22810) + ChrW(36984) + "]]"
            End If
            
            Forms!frmPickStatus_multi.Form!TxtTypeID.Visible = True
            Forms!frmPickStatus_multi.Form!TxtTypeID.SetFocus
            strStatus_DESC = Forms!frmPickStatus_multi.Form!TxtTypeID.Value
            Forms!frmPickStatus_multi.Form!subTreeView.SetFocus
            Forms!frmPickStatus_multi.Form!TxtTypeID.Visible = False
            TxtTypeCode.Value = strStatus_DESC
            gStatusTypeStr = Trim(strStatus_DESC)
            
            If TxtTypeCode.Value = "000" Then
                TxtTypeDesc.Value = "[ALL]"
                TxtTypeChn.Value = "[ALL]"
            Else
                Forms!frmPickStatus_multi.Form!TxtTypeDesc.Visible = True
                Forms!frmPickStatus_multi.Form!TxtTypeDesc.SetFocus
                
                strStatus_DESC = Forms!frmPickStatus_multi.Form!TxtTypeDesc.Value
                
                Forms!frmPickStatus_multi.Form!subTreeView.SetFocus
                Forms!frmPickStatus_multi.Form!TxtTypeDesc.Visible = False
                
                TxtTypeDesc.Value = strStatus_DESC
                    
                Forms!frmPickStatus_multi.Form!TxtTypeDescChn.Visible = True
                Forms!frmPickStatus_multi.Form!TxtTypeDescChn.SetFocus
                
                strStatus_DESC = Forms!frmPickStatus_multi.Form!TxtTypeDescChn.Value
                
                Forms!frmPickStatus_multi.Form!subTreeView.SetFocus
                Forms!frmPickStatus_multi.Form!TxtTypeDescChn.Visible = False
                
                TxtTypeChn.Value = strStatus_DESC
            End If
        Else
            Forms!frmPickStatus_multi.Form!TxtStatusDesc.Visible = True
            Forms!frmPickStatus_multi.Form!TxtStatusDesc.SetFocus
            strStatus_DESC = Forms!frmPickStatus_multi.Form!TxtStatusDesc.Value
            Forms!frmPickStatus_multi.Form!subTreeView.SetFocus
            Forms!frmPickStatus_multi.Form!TxtStatusDesc.Visible = False
            TxtStatusDesc.Value = strStatus_DESC
            
            Forms!frmPickStatus_multi.Form!TxtStatusDescChn.Visible = True
            Forms!frmPickStatus_multi.Form!TxtStatusDescChn.SetFocus
            strStatus_DESC = Forms!frmPickStatus_multi.Form!TxtStatusDescChn.Value
            Forms!frmPickStatus_multi.Form!subTreeView.SetFocus
            Forms!frmPickStatus_multi.Form!TxtStatusDescChn.Visible = False
            TxtStatusChn.Value = strStatus_DESC
            
            TxtTypeCode.Value = ""
            TxtTypeDesc.Value = "N/A"
            TxtTypeChn.Value = "N/A"
        End If
                
        DoCmd.Close acForm, stDocName
        '
        CmdQuery.Enabled = True
        Me.CmdSaveStatusCodes.Enabled = True
    Else
        CmdQuery.Enabled = False
        Me.CmdSaveStatusCodes.Enabled = False
    End If
            
    CmdPickStatus.SetFocus
    TxtStatusCode.Visible = False
        
Exit_CmdPickStatus_Click:
    Exit Sub

Err_CmdPickStatus_Click:
    MsgBox Err.Description
    Resume Exit_CmdPickStatus_Click
    
End Sub

Private Sub CmdQuery_Click()
    On Error GoTo Err_Run_Query

    Dim rst As DAO.Recordset, tContinue As Integer
    Dim tRstStatus As DAO.Recordset, tRstAddrList As DAO.Recordset, tRstDummy As DAO.Recordset
    Dim tQueryInsertStr As String, tQuerySelectStr As String, tQueryFromStr As String, tQueryWhereStr As String
    Dim tQueryStr As String, tRecDrop As Long, tStrWhereSQL As String
    Dim tUseAddr As Boolean, tUseIndexYears As Boolean, tUseDynasties As Boolean
    
    Dim cmdSQL As ADODB.Command, tRecCount As Long, tRecIndexYearCount As Long
    
    Set cmdSQL = New ADODB.Command
    '
    cmdSQL.ActiveConnection = CurrentProject.Connection
    cmdSQL.CommandType = adCmdText
    '
    '  to clear the table, close and then delete records
    '
    Set tRstStatus = ZZ_SCRATCH_STATUS.Form.Recordset
    Set tRstDummy = CurrentDb.OpenRecordset("Z_SCRATCH_DUMMY_SC", dbOpenDynaset)
    Set ZZ_SCRATCH_STATUS.Form.Recordset = tRstDummy
    tRstStatus.Close
        '
    cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_STATUS"
    cmdSQL.Execute tRecCount
    '
    '  now the people table
    '
    Set gRstPeople = ZZ_SCRATCH_P_STATUS.Form.Recordset
    Set tRstDummy = CurrentDb.OpenRecordset("Z_SCRATCH_DUMMY_SP", dbOpenDynaset)
    Set ZZ_SCRATCH_P_STATUS.Form.Recordset = tRstDummy
    gRstPeople.Close
    '
    cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_P_STATUS"
    cmdSQL.Execute tRecCount
    '
    ' see whether index years or dynasties will be used
    '
    If Me.FrameFilterYears.Value = 1 Then
        tUseIndexYears = False
        tUseDynasties = False
    ElseIf Me.FrameFilterYears.Value = 2 Then
        tUseIndexYears = True
        tUseDynasties = False
    Else
        tUseIndexYears = False
        tUseDynasties = True
    End If
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
    
    ' next build the appropriate query string. Since using dynasties requires an outer join, the string construction needs to be split.
    ' one in which c_dy must have a values, and another in which it does not.
    

    '  set the from tables with two possible joins:  ZZ_SCRATCH_ADDR for addresses and DYNASTIES for dynasties
    '  the actual options are a bit more complicated, but the DYNASTIES and STATUS_CODE_TYPE_REL files are small enough that the joins should not be too costly to leave as is
    
    tQueryInsertStr = "INSERT INTO ZZ_SCRATCH_STATUS ( c_personid, c_name, c_name_chn, c_index_year, c_sex, c_addr_id, c_dy, c_status_code, c_status_desc, " + _
        "c_status_desc_chn, c_source, c_index_year_type_code, c_sequence ) "
    tQuerySelectStr = "SELECT BIOG_MAIN.c_personid, BIOG_MAIN.c_name, BIOG_MAIN.c_name_chn, BIOG_MAIN.c_index_year, IIf([BIOG_MAIN.c_female],'F','M') AS c_sex, " + _
                "BIOG_MAIN.c_index_addr_id, BIOG_MAIN.c_dy, STATUS_DATA.c_status_code, STATUS_CODES.c_status_desc, STATUS_CODES.c_status_desc_chn, " + _
                "STATUS_DATA.c_source, BIOG_MAIN.c_index_year_type_code, STATUS_DATA.c_sequence "

    If tUseDynasties Then
        If tUseAddr Then
            tQueryFromStr = "FROM ( STATUS_CODES INNER JOIN ( ZZ_STATUS_CODE INNER JOIN STATUS_DATA ON ZZ_STATUS_CODE.c_status_code = STATUS_DATA.c_status_code ) " + _
                "ON STATUS_CODES.c_status_code = STATUS_DATA.c_status_code ) INNER JOIN ( DYNASTIES RIGHT JOIN ( BIOG_MAIN INNER JOIN ZZ_SCRATCH_ADDR " + _
                "ON BIOG_MAIN.c_index_addr_id = ZZ_SCRATCH_ADDR.c_addr_id ) ON DYNASTIES.c_dy = BIOG_MAIN.c_dy ) ON STATUS_DATA.c_personid = BIOG_MAIN.c_personid "
        Else
            tQueryFromStr = "FROM STATUS_CODES INNER JOIN ( ( ZZ_STATUS_CODE INNER JOIN STATUS_DATA ON ZZ_STATUS_CODE.c_status_code = STATUS_DATA.c_status_code ) " + _
                "INNER JOIN ( DYNASTIES RIGHT JOIN BIOG_MAIN ON DYNASTIES.c_dy = BIOG_MAIN.c_dy ) ON STATUS_DATA.c_personid = BIOG_MAIN.c_personid ) " + _
                "ON STATUS_CODES.c_status_code = STATUS_DATA.c_status_code "
        End If
    Else
        If tUseAddr Then
            tQueryFromStr = "FROM STATUS_CODES INNER JOIN ( ( ZZ_STATUS_CODE INNER JOIN STATUS_DATA ON ZZ_STATUS_CODE.c_status_code = STATUS_DATA.c_status_code ) " + _
                "INNER JOIN ( BIOG_MAIN INNER JOIN ZZ_SCRATCH_ADDR ON BIOG_MAIN.c_index_addr_id = ZZ_SCRATCH_ADDR.c_addr_id ) " + _
                "ON STATUS_DATA.c_personid = BIOG_MAIN.c_personid ) ON STATUS_CODES.c_status_code = STATUS_DATA.c_status_code "
        Else
            tQueryFromStr = "FROM STATUS_CODES INNER JOIN ( ( ZZ_STATUS_CODE INNER JOIN STATUS_DATA ON ZZ_STATUS_CODE.c_status_code = STATUS_DATA.c_status_code ) " + _
                "INNER JOIN BIOG_MAIN ON STATUS_DATA.c_personid = BIOG_MAIN.c_personid ) ON STATUS_CODES.c_status_code = STATUS_DATA.c_status_code "
        End If
    End If
        
    '  set the where conditions
    '
    '  Start with the index years
    '
    tQueryWhereStr = ""
    If tUseIndexYears Then
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
            tQueryWhereStr = "WHERE (((BIOG_MAIN.c_index_year)<=" + gToStr + ") AND " + _
                "((BIOG_MAIN.c_index_year)>=" + gFromStr + ") "
        End If
    ElseIf tUseDynasties Then
        '
        '  five possibilities (all, just from, just to, both from and to, and a cluelessly unset parameter)
        '
        If gFromDynasty = -2 Then
            tQueryWhereStr = "Where (((DYNASTIES.c_dy) > 0 ) "
        ElseIf gFromDynasty = -1 And gToDynasty > 0 Then
            tQueryWhereStr = "WHERE (((DYNASTIES.c_start)<" + Str(gToDynastyEnd) + ") "
        ElseIf gFromDynasty > 0 And gToDynasty = -1 Then
            tQueryWhereStr = "WHERE (((DYNASTIES.c_end)>" + Str(gFromDynastyBegin) + ") "
        ElseIf gFromDynasty = gToDynasty And gFromDynasty > 0 Then
            tQueryWhereStr = "WHERE (((DYNASTIES.c_dy) = " + Str(gFromDynasty) + ") "
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

    '  This section has been made irrelevant since all the selected codes, from 1, to multi, to all, are in ZZ_STATUS_CODE
    '  Because I no longer rely on the status type codes, STATUS_CODE_TYPE_REL disappears from the FROM statement
    
    
    ' run the query
    
    'MsgBox tQueryInsertStr + tQuerySelectStr + tQueryFromStr + tQueryWhereStr
    
    cmdSQL.CommandText = tQueryInsertStr + tQuerySelectStr + tQueryFromStr + tQueryWhereStr
    cmdSQL.Execute tRecCount
    '
    ' now update the index year information
    '
    cmdSQL.CommandText = "UPDATE (((ZZ_SCRATCH_STATUS LEFT JOIN INDEXYEAR_TYPE_CODES " + _
        "ON ZZ_SCRATCH_STATUS.c_index_year_type_code = INDEXYEAR_TYPE_CODES.c_index_year_type_code) LEFT JOIN DYNASTIES " + _
        "ON ZZ_SCRATCH_STATUS.c_dy = DYNASTIES.c_dy) LEFT JOIN ADDR_CODES ON ZZ_SCRATCH_STATUS.c_addr_id = ADDR_CODES.c_addr_id) LEFT JOIN TEXT_CODES " + _
        "ON ZZ_SCRATCH_STATUS.c_source = TEXT_CODES.c_textid " + _
    "SET ZZ_SCRATCH_STATUS.c_title_chn = [TEXT_CODES].[c_title_chn], ZZ_SCRATCH_STATUS.c_title = [TEXT_CODES].[c_title], " + _
        "ZZ_SCRATCH_STATUS.c_index_year_type_desc = [INDEXYEAR_TYPE_CODES].[c_index_year_type_desc], " + _
        "ZZ_SCRATCH_STATUS.c_index_year_type_hz = [INDEXYEAR_TYPE_CODES].[c_index_year_type_hz], ZZ_SCRATCH_STATUS.c_dynasty = [DYNASTIES].[c_dynasty], " + _
        "ZZ_SCRATCH_STATUS.c_dynasty_chn = [DYNASTIES].[c_dynasty_chn], ZZ_SCRATCH_STATUS.c_addr_name = [ADDR_CODES].[c_name], " + _
        "ZZ_SCRATCH_STATUS.c_addr_chn = [ADDR_CODES].[c_name_chn], ZZ_SCRATCH_STATUS.x_coord = [ADDR_CODES].[x_coord], " + _
        "ZZ_SCRATCH_STATUS.y_coord = [ADDR_CODES].[y_coord]"
    cmdSQL.Execute tRecIndexYearCount

    '
    '  the next step is to make the table of people from the status records
    '
    If tRecCount > 0 Then
        '
        tQueryStr = "INSERT INTO ZZ_SCRATCH_P_STATUS ( c_person_id, c_name, c_name_chn, c_index_year, c_index_year_type_code, c_index_year_type_desc, " + _
            "c_index_year_type_hz, c_dy, c_dynasty, c_dynasty_chn, c_sex, c_addr_id, c_addr_name, c_addr_chn, x_coord, y_coord ) " + _
            "SELECT DISTINCT ZZ_SCRATCH_STATUS.c_personid, ZZ_SCRATCH_STATUS.c_name, ZZ_SCRATCH_STATUS.c_name_chn, ZZ_SCRATCH_STATUS.c_index_year, " + _
            "ZZ_SCRATCH_STATUS.c_index_year_type_code, ZZ_SCRATCH_STATUS.c_index_year_type_desc, ZZ_SCRATCH_STATUS.c_index_year_type_hz, " + _
            "ZZ_SCRATCH_STATUS.c_dy, ZZ_SCRATCH_STATUS.c_dynasty, ZZ_SCRATCH_STATUS.c_dynasty_chn, " + _
            "ZZ_SCRATCH_STATUS.c_sex, ZZ_SCRATCH_STATUS.c_addr_id, ZZ_SCRATCH_STATUS.c_addr_name, " + _
            "ZZ_SCRATCH_STATUS.c_addr_chn, ZZ_SCRATCH_STATUS.x_coord, ZZ_SCRATCH_STATUS.y_coord " + _
            "FROM ZZ_SCRATCH_STATUS"
        cmdSQL.CommandText = tQueryStr
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
                "SELECT ZZ_SCRATCH_P_STATUS.x_coord, ZZ_SCRATCH_P_STATUS.y_coord, Count(ZZ_SCRATCH_P_STATUS.x_coord) " + _
                "AS CountOfx_coord, Count(ZZ_SCRATCH_P_STATUS.y_coord) AS CountOfy_coord " + _
                "FROM ZZ_SCRATCH_P_STATUS " + _
                "GROUP BY ZZ_SCRATCH_P_STATUS.x_coord, ZZ_SCRATCH_P_STATUS.y_coord"
            '
            cmdSQL.CommandText = tQueryStr
            cmdSQL.Execute tRecCount
            '
            tQueryStr = "UPDATE tmpXY INNER JOIN ZZ_SCRATCH_P_STATUS ON (tmpXY.y_coord = " + _
                "ZZ_SCRATCH_P_STATUS.y_coord) AND (tmpXY.x_coord = ZZ_SCRATCH_P_STATUS.x_coord) " + _
                "SET ZZ_SCRATCH_P_STATUS.xy_count = [tmpXY].[CountOfx_coord]"
        
            cmdSQL.CommandText = tQueryStr
            cmdSQL.Execute tRecCount
            CmdGIS.Enabled = True
            CmdNeo4j.Enabled = True
            CmdStoreID.Enabled = True
        Else
            CmdGIS.Enabled = False
            CmdNeo4j.Enabled = False
            CmdStoreID.Enabled = False
        End If
    End If

Exit_Run_Query:
    '
    '  now reopen the tables
    '
    Set tRstStatus = CurrentDb.OpenRecordset("ZZ_SCRATCH_STATUS", dbOpenDynaset)
    Set ZZ_SCRATCH_STATUS.Form.Recordset = tRstStatus
    '
    Set gRstPeople = CurrentDb.OpenRecordset("ZZ_SCRATCH_P_STATUS", dbOpenDynaset)
    Set ZZ_SCRATCH_P_STATUS.Form.Recordset = gRstPeople
    '
    '  close everything
    '
    Set rst = Nothing
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
    If ZZ_SCRATCH_P_STATUS.Form.Recordset.RecordCount = 0 Then
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
        Set tRstNode = ZZ_SCRATCH_P_STATUS.Form.Recordset
        tC = Chr(9) ' the tab
        '
        With tRstNode
            '
            ' write the header
            '
            If tPinyin Then
                tStr = "Name" + tC + "Sex" + tC + "IndexYear" + tC + _
                    "AddrName" + tC + "X" + tC + "Y" + tC + "xy_count"
            Else
                tStr = "Name" + tC + "NameChn" + tC + "Sex" + tC + "IndexYear" + tC + _
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



Private Sub CmdSaveStatusCodes_Click()
On Error GoTo Err_CmdSaveStatusCodes_Click
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
    
    dlgSaveAs.InitialFileName = "status_code_list.txt"
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
            GoTo Exit_CmdSaveStatusCodes_Click
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
        tStr = "SELECT ZZ_STATUS_CODE.c_status_code, STATUS_CODES.c_status_desc, STATUS_CODES.c_status_desc_chn " + _
            "FROM ZZ_STATUS_CODE INNER JOIN STATUS_CODES ON ZZ_STATUS_CODE.c_status_code = STATUS_CODES.c_status_code"
        Set tRstIDs = CurrentDb.OpenRecordset(tStr, dbOpenDynaset)
        '
        tTab = Chr(9)
        With tRstIDs
            
            .MoveFirst
            ' MsgBox "writing file"
            Do While Not .EOF
                tStr = Str(!c_status_code) + tTab
                If IsNull(!c_status_desc) Then
                    tStr = tStr + "" + tTab
                Else
                    tStr = tStr + !c_status_desc + tTab
                End If
                tStr = tStr + !c_status_desc_chn
                '
                tStream.WriteText tStr, adWriteLine
                .MoveNext
            Loop
        End With
        ' now make sure all the data is copied to tStream
        tStream.Flush
        tStream.Position = 3
        tStream.CopyTo tStreamNoBOM
        ' and write the stream to the file
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
    
Exit_CmdSaveStatusCodes_Click:
    Exit Sub

Err_CmdSaveStatusCodes_Click:
    MsgBox Err.Description
    Resume Exit_CmdSaveStatusCodes_Click

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
    Dim cmdSQL As ADODB.Command, tRecDeleted As Variant
    Dim tRstStatusCode As DAO.Recordset, tRstDummy As DAO.Recordset
    
    Set cmdSQL = New ADODB.Command
    cmdSQL.ActiveConnection = CurrentProject.Connection
    cmdSQL.CommandType = adCmdText
    '
    '  to clear the tables, briefly close and then delete records
    '
    Set tRstStatusCode = ZZ_SCRATCH_STATUS.Form.Recordset
    Set tRstDummy = CurrentDb.OpenRecordset("Z_SCRATCH_DUMMY_SC", dbOpenDynaset)
    Set ZZ_SCRATCH_STATUS.Form.Recordset = tRstDummy
    tRstStatusCode.Close
    '
    cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_STATUS"
    cmdSQL.Execute tRecDeleted
    '
    '  now reopen
    '
    Set tRstStatusCode = CurrentDb.OpenRecordset("ZZ_SCRATCH_STATUS", dbOpenDynaset)
    Set ZZ_SCRATCH_STATUS.Form.Recordset = tRstStatusCode
    '
    Set tRstStatusCode = ZZ_SCRATCH_P_STATUS.Form.Recordset
    Set tRstDummy = CurrentDb.OpenRecordset("Z_SCRATCH_DUMMY_SP", dbOpenDynaset)
    Set ZZ_SCRATCH_P_STATUS.Form.Recordset = tRstDummy
    tRstStatusCode.Close
    '
    cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_P_STATUS"
    cmdSQL.Execute tRecDeleted
    '
    '  now reopen
    '
    Set tRstStatusCode = CurrentDb.OpenRecordset("ZZ_SCRATCH_P_STATUS", dbOpenDynaset)
    Set ZZ_SCRATCH_P_STATUS.Form.Recordset = tRstStatusCode
    
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
    
    '  set the index year default values
    
    gFromDynasty = -1
    gToDynasty = -1
    gFromStr = "-200"
    gToStr = "1911"
    TxtFromYear.Value = -200
    TxtToYear.Value = 1911
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
    Dim tLabelLanguage(3, 32) As String, tLang As Integer
    
    Dim tRstLabelList As DAO.Recordset, ti As Integer
    
    Set tRstLabelList = CurrentDb.OpenRecordset("FormLabels", dbOpenTable)
    
    tRstLabelList.Index = "label"
    
    gLabelsOK = False
    With tRstLabelList
        .MoveFirst
        ti = 1
        
        Do While ti < 32 And Not .EOF
            If !c_form = "LAS" Then
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
        Me.CmdPickStatus.Caption = tLabelLanguage(tLang, 4)
        Me.CmdQuery.Caption = tLabelLanguage(tLang, 5)
        Me.CmdGIS.Caption = tLabelLanguage(tLang, 6)
        Me.CmdFanti.Caption = tLabelLanguage(tLang, 7)
        Me.CmdJianti.Caption = tLabelLanguage(tLang, 8)
        Me.PageStatus.Caption = tLabelLanguage(tLang, 9)
        Me.PagePeople.Caption = tLabelLanguage(tLang, 10)
        Me.CmdSelectPlace.Caption = tLabelLanguage(tLang, 11)
        Me.CmdImportPlaces.Caption = tLabelLanguage(tLang, 12)
        Me.CmdAllPlaces.Caption = tLabelLanguage(tLang, 13)
        Me.LblDisplay.Caption = tLabelLanguage(tLang, 14)
        Me.CmdHelp.Caption = tLabelLanguage(tLang, 15)
        Me.LblXYRef.Caption = tLabelLanguage(tLang, 16)
        Me.LblNarrow.Caption = tLabelLanguage(tLang, 17)
        Me.LblBroad.Caption = tLabelLanguage(tLang, 18)
        Me.CmdStoreID.Caption = tLabelLanguage(tLang, 19)
        Me.LblChkSubUnits.Caption = tLabelLanguage(tLang, 20)
        Me.LblDynasties.Caption = tLabelLanguage(tLang, 21)
        Me.CmdFromDynasty.Caption = tLabelLanguage(tLang, 22)
        Me.CmdToDynasty.Caption = tLabelLanguage(tLang, 23)
        Me.CmdAllDynasties.Caption = tLabelLanguage(tLang, 24)
        Me.LblIndexYears.Caption = tLabelLanguage(tLang, 25)
        Me.LblNoYears.Caption = tLabelLanguage(tLang, 26)
        Me.LblUseIndexYears.Caption = tLabelLanguage(tLang, 27)
        Me.LblUseDynasty.Caption = tLabelLanguage(tLang, 28)
        
        Me.CmdNeo4j.Caption = tLabelLanguage(tLang, 29)
        Me.CmdImportStatusCodes.Caption = tLabelLanguage(tLang, 30)
        Me.CmdSaveStatusCodes.Caption = tLabelLanguage(tLang, 31)
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
    
    If FrameFilterYears.Value = 2 Then
        ' enable index years
        Me.TxtFromYear.Enabled = True
        Me.TxtToYear.Enabled = True
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
    
    tStrPDF = Application.CurrentProject.Path + "\HelpFiles\HelpFile_LookAtStatusiations.pdf"
    
    'MsgBox tStrPDF
    
    Application.FollowHyperlink tStrPDF, , True
    
End Sub

Private Sub writeKML()
On Error GoTo Err_writeKML
    '
    '  This program will dump the results to a .gis file
    '
    If ZZ_SCRATCH_P_STATUS.Form.Recordset.RecordCount = 0 Then
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
        Set tRstNode = ZZ_SCRATCH_P_STATUS.Form.Recordset
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
        tStream.WriteText tC + "<Style id=" + tDQ + "status-balloon-template" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + "<BalloonStyle>", adWriteLine
        tStream.WriteText tC + tC + tC + "<text>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "<![CDATA[", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "ID: $[StatusGIS/PersonID] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "Name Chn: $[StatusGIS/NameHZ] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "Index Year: $[StatusGIS/IndexYear] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "Sex: $[StatusGIS/Sex] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "Address: $[StatusGIS/AddrName] $[StatusGIS/AddrHZ] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "XY Count: $[StatusGIS/XYCount] <br/><br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "]]>", adWriteLine
        tStream.WriteText tC + tC + tC + "</text>", adWriteLine
        tStream.WriteText tC + tC + "</BalloonStyle>", adWriteLine
        tStream.WriteText tC + "</Style>", adWriteLine
        tStream.WriteText tC + "<!-- Declare the type " + tDQ + "StatusGIS" + tDQ + " with 6 fields -->", adWriteLine
        tStream.WriteText tC + "<Schema name=" + tDQ + "StatusGIS" + tDQ + " id=" + tDQ + "StatusGISId" + tDQ + ">", adWriteLine
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
                
                tStream.WriteText tC + tC + "<styleUrl>#status-balloon-template</styleUrl>", adWriteLine
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
                tStream.WriteText tC + tC + tC + "<SchemaData schemaUrl=" + tDQ + "#StatusGISId" + tDQ + ">", adWriteLine
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

    tStrQuery = "INSERT INTO ZZ_STORE_PERSON_ID ( c_personid ) SELECT DISTINCT ZZ_SCRATCH_P_STATUS.c_person_id FROM ZZ_SCRATCH_P_STATUS"
    
    cmdSQL.CommandText = tStrQuery
    cmdSQL.Execute tRecCount
    MsgBox "Person IDs successfully stored.  Click on 'Recall Person IDs' to reuse these IDs in other forms."
    '
    '  update storage source
    cmdSQL.CommandText = "UPDATE PersonIDSource SET SourceForm ='Status' WHERE PersonIDSource.LineNum =1"
    cmdSQL.Execute tRecCount

End Sub


