Option Compare Database
' g stands for global
Public gDisplayLanguage As String, gLCID As Integer, gStream As ADODB.Stream, gRstImportPeople As DAO.Recordset
Public gStatusRecCount As Long, gOfficeRecCount As Long, gEntryRecCount As Long, gTextRecCount As Long, gPlaceRecCount As Long
Public gChkCount As Integer, gChkGisCount As Integer, gPeopleCount As Long
 

Private Sub ChkAddr_Click()
    ' the value is AFTER the click
    If ChkAddr.Value Then
        gChkCount = gChkCount + 1
        If gPeopleCount > 0 Then
            CmdRun.Enabled = True
        End If
        FrameQueryAddress.Enabled = True
    Else
        gChkCount = gChkCount - 1
        If gChkCount = 0 Then
            CmdRun.Enabled = False
        End If
        FrameQueryAddress.Enabled = False
    End If
    'MsgBox Str(gChkCount)

End Sub

Private Sub ChkEntry_Click()
    ' the value is AFTER the click
    If ChkEntry.Value Then
        gChkCount = gChkCount + 1
        '  just as an initial check that I'm doing this right
        If gPeopleCount > 0 Then
            CmdRun.Enabled = True
        End If
    Else
        gChkCount = gChkCount - 1
        If gChkCount = 0 Then
            CmdRun.Enabled = False
        End If
    End If
    'MsgBox Str(gChkCount)

End Sub

Private Sub ChkOffice_Click()
    ' the value is AFTER the click
    If ChkOffice.Value Then
        gChkCount = gChkCount + 1
        '  just as an initial check that I'm doing this right
        If gPeopleCount > 0 Then
            CmdRun.Enabled = True
        End If
    Else
        gChkCount = gChkCount - 1
        If gChkCount = 0 Then
            CmdRun.Enabled = False
        End If
    End If
    'MsgBox Str(gChkCount)

End Sub

Private Sub ChkStatus_Click()
    ' the value is AFTER the click
    If ChkStatus.Value Then
        gChkCount = gChkCount + 1
        '  just as an initial check that I'm doing this right
        If gPeopleCount > 0 Then
            CmdRun.Enabled = True
        End If
    Else
        gChkCount = gChkCount - 1
        If gChkCount = 0 Then
            CmdRun.Enabled = False
        End If
    End If
    'MsgBox Str(gChkCount)
End Sub

Private Sub ChkText_Click()
    ' the value is AFTER the click
    If ChkText.Value Then
        gChkCount = gChkCount + 1
        If gPeopleCount > 0 Then
            CmdRun.Enabled = True
        End If
    Else
        gChkCount = gChkCount - 1
        If gChkCount = 0 Then
            CmdRun.Enabled = False
        End If
    End If
    'MsgBox Str(gChkCount)
End Sub

Private Sub CmdClose_Click()
On Error GoTo Err_CmdClose_Click

    DoCmd.Close

Exit_CmdClose_Click:
    Exit Sub

Err_CmdClose_Click:
    MsgBox Err.Description
    Resume Exit_CmdClose_Click
    
End Sub
Private Sub CmdGIS_Click()
On Error GoTo Err_CmdGIS_Click
    '
    '  If it is a KML file, call the routine and exit
    '
    If ChkGisStatus.Value Then
        Call WriteGIS_Status
    End If
    '
    If Me.ChkGisOffice Then
        Call WriteGIS_OfficeOffice
    End If
    If Me.ChkGisOfficePeople Then
        Call WriteGIS_OfficePeople
    End If
    '
    If ChkGisEntry.Value Then
        Call WriteGIS_Entry
    End If
    '
    If ChkGisText.Value Then
        Call WriteGIS_Text
    End If
    '
    If ChkGisAddr.Value Then
        Call WriteGIS_Addr
    End If
    '
Exit_CmdGIS_Click:
    Exit Sub

Err_CmdGIS_Click:
    MsgBox Err.Description
    Resume Exit_CmdGIS_Click
    
End Sub


Private Sub CmdImport_Click()
On Error GoTo Err_CmdImport_Click
    Dim cmdSQL As ADODB.Command
    Dim tStrQuestion As String, tQuit As Boolean

    Dim dlgSaveAs As FileDialog
    Dim tFileNum As Integer
    Dim tFileName As String, tFN As Variant
    Dim tFileSystem, tList
    
    Set cmdSQL = New ADODB.Command
    cmdSQL.ActiveConnection = CurrentProject.Connection
    cmdSQL.CommandType = adCmdText
    
    tQuit = False
    '
    If Not tQuit Then
        '
        '  open the list
        
        Set dlgSaveAs = Application.FileDialog(msoFileDialogOpen)
    
        'Use a With...End With block to reference the FileDialog object.
        
        tFileName = ""
        
        With dlgSaveAs
            .InitialFileName = ""
            If .Show = -1 Then
                '
                For Each tFN In .SelectedItems
                    tFileName = tFN
                    If Not tFileName = "" Then
                        Exit For
                    End If
                Next
                If tFileName = "" Then
                    MsgBox "Bad file Name."
                    GoTo Exit_CmdImport_Click
                End If
            End If
        End With
        '
        ' Clear the people table now that we are ready to go
        '
        If Not (tFileName = "") Then
            '
            cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_IMPORT_PEOPLE"
            cmdSQL.Execute tRecDeleted
            '
            cmdSQL.CommandText = "Delete * from InputErrorList"
            cmdSQL.Execute tRecDeleted
            
            cmdSQL.CommandText = "Delete * from TempImportList"
            cmdSQL.Execute tRecDeleted
            
            DoCmd.TransferText acImportDelim, "ImportPeopleList_Space", "TempImportList", tFileName, 0
            '    TransferType=acImportDelim
            '    SpecificationName = "TempImportList" (apparently it is saved in the database itself)
            '    TableName = "TempImportList"  (probably requires that I drop the table first, but I can test)
            '    HasFieldNames = False (0)
            '
            '  copy the bad IDs
            '
            tStrSQL = "INSERT INTO InputErrorList ( c_ID ) SELECT TempImportList.ImportID " + _
                "FROM BIOG_MAIN RIGHT JOIN TempImportList ON BIOG_MAIN.c_personid = TempImportList.ImportID " + _
                "WHERE (((BIOG_MAIN.c_personid) Is Null) AND (TempImportList.ImportID is Not Null))"
    
            cmdSQL.CommandText = tStrSQL
            cmdSQL.Execute tRecDeleted
            
            If tRecDeleted > 0 Then
                MsgBox "Some ID were not successfully imported:  please look at InputErrorList."
            End If
            '
            '  copy the good IDs
            '
            tStrSQL = "INSERT INTO ZZ_SCRATCH_IMPORT_PEOPLE ( c_person_id ) SELECT DISTINCT TempImportList.ImportID " + _
                "FROM BIOG_MAIN INNER JOIN TempImportList ON BIOG_MAIN.c_personid = TempImportList.ImportID"
    
            cmdSQL.CommandText = tStrSQL
            cmdSQL.Execute gPeopleCount
            
            If gPeopleCount = 0 Then
                CmdRun.Enabled = False
                CmdStoreID.Enabled = False
            Else
                If gChkCount > 0 Then
                    CmdRun.Enabled = True
                    CmdStoreID.Enabled = True
                End If
            End If
            
            Set cmdSQL = Nothing
            Set tFileSystem = Nothing
        End If
            
    End If
    
Exit_CmdImport_Click:
    Exit Sub

Err_CmdImport_Click:
    MsgBox Err.Description
    Resume Exit_CmdImport_Click

End Sub

Private Sub CmdNeo4j_Click()
On Error GoTo Err_CmdNeo4j_Click
    '
    '  The challenge here is that we have 5 tables of results:
    '       ZZ_SCRATCH_STATUS
    '       ZZ_SCRATCH_OFFICE
    '       ZZ_SCRATCH_ENTRY
    '       ZZ_SCRATCH_BIOG_TEXT_DATA
    '       ZZ_SCRATCH_BIOG_ADDR_DATA
    '
    '  We need to check all of these for people (entry has 3 IDs) and address IDs, along with their specific data
    '
    Dim dlgSaveAs As FileDialog
    Dim tFileNum As Integer, tFileName As String, tFN As Variant
    '
    Dim tRstPeople As DAO.Recordset, tRstPlace As DAO.Recordset, tRstPeoplePlace As DAO.Recordset
    Dim tRstEntryCodes As DAO.Recordset, tRstPeopleEntry As DAO.Recordset, tRstInstitutionCodes As DAO.Recordset
    Dim tRstOfficeCodes As DAO.Recordset, tRstPeopleOffice As DAO.Recordset
    Dim tRstTextCodes As DAO.Recordset, tRstPeopleText As DAO.Recordset
    Dim tRstPeopleStatus As DAO.Recordset, tRstStatusCodes As DAO.Recordset, tStr As String, tC As String
    Dim tQueryStr As String, tPersonID As Long, tCount As Long
    '
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
    '  get the People file
    '
    '  prepare the temp tables for the people, place, peoplePlace and entry data
            
    Dim cmdSQL As ADODB.Command, tRecDeleted As Long
    Set cmdSQL = New ADODB.Command
    cmdSQL.ActiveConnection = CurrentProject.Connection
    cmdSQL.CommandType = adCmdText
    
    ' collect the person IDs
    
    cmdSQL.CommandText = "DELETE * FROM ZZ_SCRATCH_P_TEXT"
    cmdSQL.Execute tRecDeleted
    '
    ' collect from entry
    '
    tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id ) SELECT DISTINCT ZZ_SCRATCH_ENTRY.c_personid FROM ZZ_SCRATCH_ENTRY"
    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecDeleted
    '
    tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id ) SELECT DISTINCT ZZ_SCRATCH_ENTRY.c_kin_id " + _
                "FROM ZZ_SCRATCH_ENTRY WHERE (ZZ_SCRATCH_ENTRY.c_kin_id > 0)"
    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecDeleted
    '
    tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id ) SELECT DISTINCT ZZ_SCRATCH_ENTRY.c_assoc_id " + _
                "FROM ZZ_SCRATCH_ENTRY WHERE (ZZ_SCRATCH_ENTRY.c_assoc_id > 0)"
    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecDeleted
    '
    ' collect from office
    '
    tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id ) SELECT DISTINCT ZZ_SCRATCH_OFFICE.c_personid FROM ZZ_SCRATCH_OFFICE"
    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecDeleted
    '
    ' collect from status
    '
    tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id ) SELECT DISTINCT ZZ_SCRATCH_STATUS.c_personid FROM ZZ_SCRATCH_STATUS"
    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecDeleted
    '
    ' collect from texts
    '
    tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id ) SELECT DISTINCT ZZ_SCRATCH_BIOG_TEXT_DATA.c_personid " + _
                "FROM ZZ_SCRATCH_BIOG_TEXT_DATA"
    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecDeleted
    '
    ' collect from addresses
    '
    tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id ) SELECT DISTINCT ZZ_SCRATCH_BIOG_ADDR_DATA.c_personid " + _
                "FROM ZZ_SCRATCH_BIOG_ADDR_DATA"
    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecDeleted
    '
    '  now get the information on people, which will be used in various queries
    '
    tQueryStr = "UPDATE ZZ_SCRATCH_P_TEXT INNER JOIN BIOG_MAIN ON ZZ_SCRATCH_P_TEXT.c_person_id = BIOG_MAIN.c_personid " + _
                "SET ZZ_SCRATCH_P_TEXT.c_name = [BIOG_MAIN].[c_name],  ZZ_SCRATCH_P_TEXT.c_name_chn = [BIOG_MAIN].[c_name_chn], " + _
                        "ZZ_SCRATCH_P_TEXT.c_sex = iif([BIOG_MAIN].[c_female],'F','M'), ZZ_SCRATCH_P_TEXT.c_index_year = [BIOG_MAIN].[c_index_year], " + _
                        "ZZ_SCRATCH_P_TEXT.c_index_year_type_code = [BIOG_MAIN].[c_index_year_type_code],     ZZ_SCRATCH_P_TEXT.c_dy = [BIOG_MAIN].[c_dy], " + _
                        "ZZ_SCRATCH_P_TEXT.c_addr_id = [BIOG_MAIN].[c_index_addr_id]"
    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecDeleted
    '
    ' now fill in the outer join information
    '
    tQueryStr = "UPDATE ( ( ZZ_SCRATCH_P_TEXT LEFT JOIN INDEXYEAR_TYPE_CODES ON ZZ_SCRATCH_P_TEXT.c_index_year_type_code = INDEXYEAR_TYPE_CODES.c_index_year_type_code ) " + _
                    "LEFT JOIN DYNASTIES ON ZZ_SCRATCH_P_TEXT.c_dy = DYNASTIES.c_dy ) LEFT JOIN ADDR_CODES " + _
                    "ON ZZ_SCRATCH_P_TEXT.c_addr_id = ADDR_CODES.c_addr_id " + _
                "SET ZZ_SCRATCH_P_TEXT.c_index_year_type_desc = [INDEXYEAR_TYPE_CODES].[c_index_year_type_desc], " + _
                    "ZZ_SCRATCH_P_TEXT.c_index_year_type_hz = [INDEXYEAR_TYPE_CODES].[c_index_year_type_hz], " + _
                    "ZZ_SCRATCH_P_TEXT.c_dynasty = [DYNASTIES].[c_dynasty], ZZ_SCRATCH_P_TEXT.c_dynasty_chn = [DYNASTIES].[c_dynasty_chn], " + _
                    "ZZ_SCRATCH_P_TEXT.c_addr_name = [ADDR_CODES].[c_name], ZZ_SCRATCH_P_TEXT.c_addr_chn = [ADDR_CODES].[c_name_chn], " + _
                    "ZZ_SCRATCH_P_TEXT.x_coord = [ADDR_CODES].[x_coord], ZZ_SCRATCH_P_TEXT.y_coord = [ADDR_CODES].[y_coord]"
    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecDeleted
    '
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
        
        Set tRstPeople = CurrentDb.OpenRecordset("ZZ_SCRATCH_P_TEXT")
        '
        ' process the four tables
        '
        ' first the nodes:  define the record structure
        '
        '  if the file is strictly ASCII, the label is the pinyin, but if there are characters, then we add a pinyin field
        If tCodeStr = "ascii" Then
            tStr = "NameID" + tC + "NamePY" + tC + "IndexYear" + tC + "IndexYearTypeCode" + tC + "IndexYearTypeDesc" + tC + _
                   "Dynasty" + tC + "Sex"
        Else
            tStr = "NameID" + tC + "NameHZ" + tC + "NamePY" + tC + "IndexYear" + tC + "IndexYearTypeCode" + tC + "IndexYearTypeDesc" + tC + _
                    "IndexYearTypeDescHZ" + tC + "Dynasty" + tC + "Sex"
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
                        tStr = tStr + "Missing" + tC
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
                '
                '  indexyear = c_index_year_type_code STR
                '
                If IsNull(!c_index_year_type_code) Then
                    tStr = tStr + "Unknown" + tC
                Else
                    tStr = tStr + Trim(!c_index_year_type_code) + tC
                End If
                '
                '  indexyear = c_index_year_type_desc STR
                '
                If IsNull(!c_index_year_type_desc) Then
                    tStr = tStr + "Unknown" + tC
                Else
                    tStr = tStr + Trim(!c_index_year_type_desc) + tC
                End If
                '
                '  indexyear = c_index_year_type_hz STR
                '
                If Not (tCodeStr = "ascii") Then
                    If IsNull(!c_index_year_type_hz) Then
                        tStr = tStr + "Unknown" + tC
                    Else
                        tStr = tStr + Trim(!c_index_year_type_hz) + tC
                    End If
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
    '  now places
    '
    '  There are various fileds from which to address ID: c_index_addr_id, c_entry_addr_id, c_office_addr_id, and c_inst_addr_id
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
        '  first add index addresses for ZZ_SCRATCH_P_TEXT to ZZ_ADDRESSES
        '
        cmdSQL.CommandText = "DELETE * FROM ZZ_ADDRESSES"
        cmdSQL.Execute tRecDeleted
        '
        tQueryStr = "INSERT INTO ZZ_ADDRESSES ( c_addr_id, c_name, c_name_chn, x_coord, y_coord ) " + _
                    "SELECT DISTINCT ZZ_SCRATCH_P_TEXT.c_addr_id, ZZ_SCRATCH_P_TEXT.c_addr_name, ZZ_SCRATCH_P_TEXT.c_addr_chn, " + _
                        "ZZ_SCRATCH_P_TEXT.x_coord, ZZ_SCRATCH_P_TEXT.y_coord " + _
                    "FROM ZZ_SCRATCH_P_TEXT"
        
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecDeleted
        '
        ' now from ZZ_SCRATCH_BIOG_ADDR_DATA:  the index address for the kin and associates are not in this table, and all place
        '    associations for the listed people (more than index addresses) will be in this table
        '
        tQueryStr = "INSERT INTO ZZ_ADDRESSES ( c_addr_id, c_name, c_name_chn, x_coord, y_coord ) " + _
                    "SELECT DISTINCT ZZ_SCRATCH_BIOG_ADDR_DATA.c_addr_id, ZZ_SCRATCH_BIOG_ADDR_DATA.c_addr_name, " + _
                        "ZZ_SCRATCH_BIOG_ADDR_DATA.c_addr_chn, ZZ_SCRATCH_BIOG_ADDR_DATA.x_coord, ZZ_SCRATCH_BIOG_ADDR_DATA.y_coord " + _
                    "FROM ZZ_SCRATCH_BIOG_ADDR_DATA"
                
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecDeleted
        '
        ' now c_entry_addr_id
        '
        tQueryStr = "INSERT INTO ZZ_ADDRESSES ( c_addr_id, c_name, c_name_chn, x_coord, y_coord ) " + _
                    "SELECT DISTINCT ZZ_SCRATCH_ENTRY.c_entry_addr_id, ZZ_SCRATCH_ENTRY.c_entry_addr_name, " + _
                        "ZZ_SCRATCH_ENTRY.c_entry_addr_chn, ZZ_SCRATCH_ENTRY.c_entry_xcoord, ZZ_SCRATCH_ENTRY.c_entry_ycoord " + _
                    "FROM ZZ_SCRATCH_ENTRY " + _
                    "WHERE (((ZZ_SCRATCH_ENTRY.c_entry_addr_id)>0))"
                
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecDeleted
        '
        ' now c_inst_addr_id
        '
        tQueryStr = "INSERT INTO ZZ_ADDRESSES ( c_addr_id, c_name, c_name_chn, x_coord, y_coord ) " + _
                    "SELECT DISTINCT SOCIAL_INSTITUTION_ADDR.c_inst_addr_id, ZZ_SCRATCH_ENTRY.c_inst_name_py, " + _
                        "ZZ_SCRATCH_ENTRY.c_inst_name_hz, SOCIAL_INSTITUTION_ADDR.inst_xcoord, SOCIAL_INSTITUTION_ADDR.inst_ycoord " + _
                    "FROM ZZ_SCRATCH_ENTRY INNER JOIN SOCIAL_INSTITUTION_ADDR ON " + _
                        "(ZZ_SCRATCH_ENTRY.c_inst_name_code = SOCIAL_INSTITUTION_ADDR.c_inst_name_code) AND " + _
                        "(ZZ_SCRATCH_ENTRY.c_inst_code = SOCIAL_INSTITUTION_ADDR.c_inst_code) " + _
                    "WHERE (((SOCIAL_INSTITUTION_ADDR.c_inst_addr_id)>0))"
                
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecDeleted
        '
        ' now c_office_addr_id
        '
        tQueryStr = "INSERT INTO ZZ_ADDRESSES ( c_addr_id, c_name, c_name_chn, x_coord, y_coord ) " + _
                    "SELECT DISTINCT ZZ_SCRATCH_OFFICE.c_office_addr_id, ZZ_SCRATCH_OFFICE.c_office_addr_name, " + _
                        "ZZ_SCRATCH_OFFICE.c_office_addr_chn, ZZ_SCRATCH_OFFICE.office_x_coord, ZZ_SCRATCH_OFFICE.office_y_coord " + _
                    "FROM ZZ_SCRATCH_OFFICE " + _
                    "WHERE (((ZZ_SCRATCH_OFFICE.c_office_addr_id)>0))"
                
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecDeleted
        '
        ' finally, institution addresses from the office table
        '
        tQueryStr = "INSERT INTO ZZ_ADDRESSES ( c_addr_id, c_name, c_name_chn, x_coord, y_coord ) " + _
                    "SELECT DISTINCT SOCIAL_INSTITUTION_ADDR.c_inst_addr_id, ZZ_SCRATCH_OFFICE.c_inst_name_py, " + _
                        "ZZ_SCRATCH_OFFICE.c_inst_name_hz, SOCIAL_INSTITUTION_ADDR.inst_xcoord, SOCIAL_INSTITUTION_ADDR.inst_ycoord " + _
                    "FROM ZZ_SCRATCH_OFFICE INNER JOIN SOCIAL_INSTITUTION_ADDR ON " + _
                        "(ZZ_SCRATCH_OFFICE.c_inst_name_code = SOCIAL_INSTITUTION_ADDR.c_inst_name_code) AND " + _
                        "(ZZ_SCRATCH_OFFICE.c_inst_code = SOCIAL_INSTITUTION_ADDR.c_inst_code) " + _
                    "WHERE (((SOCIAL_INSTITUTION_ADDR.c_inst_addr_id)>0))"
                
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecDeleted
        '
        ' now get the distinct records from ZZ_ADDRESSES
        '
        tQueryStr = "SELECT DISTINCT ZZ_ADDRESSES.c_addr_id, ZZ_ADDRESSES.c_name, ZZ_ADDRESSES.c_name_chn, " + _
                        "ZZ_ADDRESSES.x_coord, ZZ_ADDRESSES.y_coord " + _
                    "FROM ZZ_ADDRESSES"
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
    '  the difficulty here is that we want information on both the people (in ZZ_SCRATCH_BIOG_ADDR_DATA) and their associates who appear in
    '    other tables (listed in ZZ_SCRATCH_P_TEXT)
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
        '  get all the index address relations
        '
        cmdSQL.CommandText = "DELETE * FROM ZZ_PLACE"
        cmdSQL.Execute tRecDeleted
        '
        tQueryStr = "INSERT INTO ZZ_PLACE ( c_personid, c_addr_id, c_rel_code ) " + _
                    "SELECT DISTINCT ZZ_SCRATCH_P_TEXT.c_person_id, BIOG_MAIN.c_index_addr_id, BIOG_MAIN.c_index_addr_type_code " + _
                    "FROM ZZ_SCRATCH_P_TEXT INNER JOIN BIOG_MAIN ON ZZ_SCRATCH_P_TEXT.c_person_id = BIOG_MAIN.c_personid " + _
                    "WHERE (((BIOG_MAIN.c_index_addr_id)>0))"
        
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecDeleted
    '
    ' fill in the outer join fields
    '
    tQueryStr = "UPDATE ZZ_PLACE LEFT JOIN BIOG_ADDR_CODES ON ZZ_PLACE.c_rel_code = BIOG_ADDR_CODES.c_addr_type " + _
                "SET ZZ_PLACE.c_rel_desc = [BIOG_ADDR_CODES].[c_addr_desc], ZZ_PLACE.c_rel_chn = [BIOG_ADDR_CODES].[c_addr_desc_chn]"
        '
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecDeleted
    '
        '  now the BIOG_ADDR data
        '
        tQueryStr = "INSERT INTO ZZ_PLACE ( c_personid, c_addr_id, c_rel_code, c_rel_desc, c_rel_chn ) " + _
                    "SELECT DISTINCT ZZ_SCRATCH_BIOG_ADDR_DATA.c_personid, ZZ_SCRATCH_BIOG_ADDR_DATA.c_addr_id, " + _
                        "ZZ_SCRATCH_BIOG_ADDR_DATA.c_addr_type, ZZ_SCRATCH_BIOG_ADDR_DATA.c_addr_desc, " + _
                        "ZZ_SCRATCH_BIOG_ADDR_DATA.c_addr_desc_chn " + _
                    "FROM ZZ_SCRATCH_BIOG_ADDR_DATA " + _
                    "WHERE (((ZZ_SCRATCH_BIOG_ADDR_DATA.c_addr_id)>0))"
        
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecDeleted
        '
        tQueryStr = "SELECT DISTINCT ZZ_PLACE.c_personid, ZZ_PLACE.c_addr_id, ZZ_PLACE.c_rel_code FROM ZZ_PLACE"

        Set tRstPeoplePlace = CurrentDb.OpenRecordset(tQueryStr)
        
        tStr = "NameID" + tC + "PlaceID" + tC + "PersonPlaceCode"
            
        gStream.WriteText tStr, adWriteLine
            
        With tRstPeoplePlace
            .MoveFirst
            Do While Not .EOF
                If Not IsNull(!c_personid) Then
                    '
                    tStr = Trim(Str(!c_personid)) + tC
                        '
                    tStr = tStr + Trim(Str(!c_addr_id)) + tC
                    '
                    tStr = tStr + Trim(Str(!c_rel_code))
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
    '  now peoplePlaceCode:  use ZZ_PLACE
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
        tQueryStr = "SELECT DISTINCT ZZ_PLACE.c_rel_code, ZZ_PLACE.c_rel_desc, ZZ_PLACE.c_rel_chn FROM ZZ_PLACE"

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
                If Not IsNull(!c_rel_code) Then
                    '
                    tStr = Trim(Str(!c_rel_code)) + tC
                    '
                    tStr = tStr + !c_rel_desc
                    '
                    If Not (tCodeStr = "ascii") Then
                        tStr = tStr + tC + !c_rel_chn
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
    ' the hard work is over: now we just need to write the status data, entry data, office data, and institution data
    '
    ' now the PeopleStatus file
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
        Set tRstPeopleStatus = CurrentDb.OpenRecordset("ZZ_SCRATCH_STATUS", dbOpenDynaset)
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
    '
    ' now the PeopleOffice file
    '
    dlgSaveAs.InitialFileName = "PeopleOffice_" + tCodeStr + ".csv"
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

        tStr = "NameID" + tC + "OfficeCode" + tC + "OfficeAddrID" + tC + "SocialInstID" + tC + "SocialInstID" + tC + _
                    "PostingFirstYear" + tC + "PostingLastYear" + tC + "PostingDynasty"
                    
        gStream.WriteText tStr, adWriteLine
        '
        Set tRstPeopleOffice = CurrentDb.OpenRecordset("ZZ_SCRATCH_OFFICE", dbOpenDynaset)
        With tRstPeopleOffice
            .MoveFirst
            Do While Not .EOF
                '  the ID of the person
                tStr = Trim(Str(!c_personid)) + tC
                '
                '  office code
                '
                If IsNull(!c_office_id) Then
                    tStr = tStr + "0" + tC
                Else
                    tStr = tStr + Trim(Str(!c_office_id)) + tC
                End If
                '
                '  office addr id
                '
                If IsNull(!c_office_addr_id) Then
                    tStr = tStr + "0" + tC
                Else
                    tStr = tStr + Trim(Str(!c_office_addr_id)) + tC
                End If
                '
                '  social inst IDs
                '
                If IsNull(!c_inst_code) Then
                    tStr = tStr + "0" + tC
                Else
                    tStr = tStr + Trim(Str(!c_inst_code)) + tC
                End If
                '
                If IsNull(!c_inst_name_code) Then
                    tStr = tStr + "0" + tC
                Else
                    tStr = tStr + Trim(Str(!c_inst_name_code)) + tC
                End If
                '
                '  posting first year
                '
                If IsNull(!c_firstyear) Then
                    tStr = tStr + "0" + tC
                Else
                    tStr = tStr + Trim(Str(!c_firstyear)) + tC
                End If
                '
                '  posting last year
                '
                If IsNull(!c_lastyear) Then
                    tStr = tStr + "0" + tC
                Else
                    tStr = tStr + Trim(Str(!c_lastyear)) + tC
                End If
                '
                '  posting dynasty
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
    ' get office codes
    '
    dlgSaveAs.InitialFileName = "OfficeCodes_" + tCodeStr + ".csv"
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
            tStr = "OfficeCode" + tC + "OfficeTrans" + tC + "OfficePinyin"
        Else
            tStr = "OfficeCode" + tC + "OfficeTrans" + tC + "OfficePinyin" + tC + "OfficeHZ"
        End If
        gStream.WriteText tStr, adWriteLine
        '
        ' get the codes
        '
        tQueryStr = "SELECT DISTINCT ZZ_SCRATCH_OFFICE.c_office_id, ZZ_SCRATCH_OFFICE.c_office_trans, " + _
                        "ZZ_SCRATCH_OFFICE.c_office_pinyin, ZZ_SCRATCH_OFFICE.c_office_chn " + _
                    "FROM ZZ_SCRATCH_OFFICE"
                    
        Set tRstOfficeCodes = CurrentDb.OpenRecordset(tQueryStr)
        With tRstOfficeCodes
            .MoveFirst
            Do While Not .EOF
                '
                tStr = Trim(Str(!c_office_id)) + tC
                '
                '  office trans
                '
                If IsNull(!c_office_trans) Then
                    tStr = tStr + "Missing" + tC
                Else
                    tStr = tStr + Trim(!c_office_trans) + tC
                End If
                '
                '  office pinyin
                '
                If IsNull(!c_office_pinyin) Then
                    tStr = tStr + "Missing"
                Else
                    tStr = tStr + Trim(!c_office_pinyin)
                End If
                '
                '  office HZ
                '
                If Not (tCodeStr = "ascii") Then
                    tStr = tStr + tC + Trim(!c_office_chn)
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
    ' Are there any records to process
    '
    If DCount("c_entry_code", "ZZ_SCRATCH_ENTRY") > 0 Then
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
                    "AssocPersonID" + tC + "AssocRelCode" + tC + "SocialInstID" + tC + "SocialInstNameID" + tC + "EntryYear" + tC + "EntryDynasty"
            gStream.WriteText tStr, adWriteLine
            '
            Set tRstPeopleEntry = CurrentDb.OpenRecordset("ZZ_SCRATCH_ENTRY", dbOpenDynaset)
            With tRstPeopleEntry
                If Not .EOF Then
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
                        '  social inst IDs
                        '
                        If IsNull(!c_inst_code) Then
                            tStr = tStr + "0" + tC
                        Else
                            tStr = tStr + Trim(Str(!c_inst_code)) + tC
                        End If
                        '
                        If IsNull(!c_inst_name_code) Then
                            tStr = tStr + "0" + tC
                        Else
                            tStr = tStr + Trim(Str(!c_inst_name_code)) + tC
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
                End If
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
            tQueryStr = "SELECT DISTINCT ZZ_SCRATCH_ENTRY.c_entry_code, ZZ_SCRATCH_ENTRY.c_entry_desc, ZZ_SCRATCH_ENTRY.c_entry_chn " + _
                        "FROM ZZ_SCRATCH_ENTRY WHERE (ZZ_SCRATCH_ENTRY.c_entry_code > -1)"
            Set tRstEntryCodes = CurrentDb.OpenRecordset(tQueryStr)
            With tRstEntryCodes
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
    End If
    '
    ' the last step is to collect institution codes from Entry and Office
    ' first, see if there are any
    '
    cmdSQL.CommandText = "DELETE * FROM ZZ_SCRATCH_BIOG_INST_DATA"
    cmdSQL.Execute tRecDeleted
    '
    tCount = 0
    cmdSQL.CommandText = "INSERT INTO ZZ_SCRATCH_BIOG_INST_DATA ( c_inst_code, c_inst_name_code, c_inst_name_hz, c_inst_name_py ) " + _
                         "SELECT DISTINCT ZZ_SCRATCH_ENTRY.c_inst_code, ZZ_SCRATCH_ENTRY.c_inst_name_code, " + _
                            "ZZ_SCRATCH_ENTRY.c_inst_name_hz, ZZ_SCRATCH_ENTRY.c_inst_name_py " + _
                         "FROM ZZ_SCRATCH_ENTRY " + _
                         "WHERE (((ZZ_SCRATCH_ENTRY.c_inst_code)>0))"
    cmdSQL.Execute tRecDeleted
    tCount = tRecDeleted

    cmdSQL.CommandText = "INSERT INTO ZZ_SCRATCH_BIOG_INST_DATA ( c_inst_code, c_inst_name_code, c_inst_name_hz, c_inst_name_py ) " + _
                         "SELECT ZZ_SCRATCH_OFFICE.c_inst_code, ZZ_SCRATCH_OFFICE.c_inst_name_code, " + _
                            "ZZ_SCRATCH_OFFICE.c_inst_name_hz, ZZ_SCRATCH_OFFICE.c_inst_name_py " + _
                         "FROM ZZ_SCRATCH_OFFICE " + _
                         "WHERE (((ZZ_SCRATCH_OFFICE.c_inst_code)>0))"
    cmdSQL.Execute tRecDeleted
    tCount = tCount + tRecDeleted
    '
    If tCount > 0 Then
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
            gStream.Mode = adModeReadWrite
            gStream.Type = adTypeText
            gStream.Open
    
            If tCodeStr = "ascii" Then
                tStr = "InstitutionCode" + tC + "InstitutionNameCode" + tC + "InstitutionName" + tC + "InstitutionAddrID"
            Else
                tStr = "InstitutionCode" + tC + "InstitutionNameCode" + tC + "InstitutionName" + tC + "InstitutionNameHZ" + tC + "InstitutionAddrID"
            End If
            gStream.WriteText tStr, adWriteLine
    
            tQueryStr = "SELECT DISTINCT ZZ_SCRATCH_BIOG_INST_DATA.c_inst_code, ZZ_SCRATCH_BIOG_INST_DATA.c_inst_name_code, " + _
                            "ZZ_SCRATCH_BIOG_INST_DATA.c_inst_name_hz, ZZ_SCRATCH_BIOG_INST_DATA.c_inst_name_py, " + _
                            "SOCIAL_INSTITUTION_ADDR.c_inst_addr_id " + _
                        "FROM ZZ_SCRATCH_BIOG_INST_DATA INNER JOIN SOCIAL_INSTITUTION_ADDR ON " + _
                            "(ZZ_SCRATCH_BIOG_INST_DATA.c_inst_name_code = SOCIAL_INSTITUTION_ADDR.c_inst_name_code) AND " + _
                            "(ZZ_SCRATCH_BIOG_INST_DATA.c_inst_code = SOCIAL_INSTITUTION_ADDR.c_inst_code)"
                            
            Set tRstInstitutionCodes = CurrentDb.OpenRecordset(tQueryStr)
            
            With tRstInstitutionCodes
                .MoveFirst
                Do While Not .EOF
                    '
                    tStr = Trim(Str(!c_inst_code)) + tC
                    '
                    If IsNull(!c_inst_name_code) Then
                        tStr = tStr + "0" + tC
                    Else
                        tStr = tStr + Trim(Str(!c_inst_name_code)) + tC
                    End If
                    '
                    '  Name (pinyin)
                    '
                    If IsNull(!c_inst_name_py) Then
                        tStr = tStr + "Missing" + tC
                    Else
                        tStr = tStr + Trim(!c_inst_name_py) + tC
                    End If
                    '
                    '  Name (HZ)
                    '
                    If Not (tCodeStr = "ascii") Then
                        If IsNull(!c_inst_name_hz) Then
                            tStr = tStr + "Missing" + tC
                        Else
                            tStr = tStr + Trim(!c_inst_name_hz) + tC
                        End If
                    End If
                    '
                    If IsNull(!c_inst_addr_id) Then
                        tStr = tStr + "0"
                    Else
                        tStr = tStr + Trim(Str(!c_inst_addr_id))
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
    End If
    'Set the object variable to Nothing.
    Set dlgSaveAs = Nothing
    MsgBox "Finished saving to Neo4j"
    
Exit_CmdNeo4j_Click:
    Exit Sub

Err_CmdNeo4j_Click:
    MsgBox Err.Description
    Resume Exit_CmdNeo4j_Click

End Sub

Private Sub CmdRun_Click()
On Error GoTo Err_CmdRun_Click
    Dim cmdSQL As ADODB.Command, tRecCount As Long, tQueryStr As String
    '
    Set cmdSQL = New ADODB.Command
    cmdSQL.ActiveConnection = CurrentProject.Connection
    cmdSQL.CommandType = adCmdText

    ' fill in the data for ZZ_SCRATCH_IMPORT_PEOPLE

    tQueryStr = "UPDATE ZZ_SCRATCH_IMPORT_PEOPLE INNER JOIN BIOG_MAIN ON ZZ_SCRATCH_IMPORT_PEOPLE.c_person_id = BIOG_MAIN.c_personid " + _
        "SET ZZ_SCRATCH_IMPORT_PEOPLE.c_name = [BIOG_MAIN].[c_name], ZZ_SCRATCH_IMPORT_PEOPLE.c_name_chn = [BIOG_MAIN].[c_name_chn], " + _
            "ZZ_SCRATCH_IMPORT_PEOPLE.c_female = [BIOG_MAIN].[c_female], ZZ_SCRATCH_IMPORT_PEOPLE.c_sex = IIf([BIOG_MAIN].[c_female], 'F', 'M'), " + _
            "ZZ_SCRATCH_IMPORT_PEOPLE.c_index_year = [BIOG_MAIN].[c_index_year], " + _
            "ZZ_SCRATCH_IMPORT_PEOPLE.c_index_year_type_code = [BIOG_MAIN].[c_index_year_type_code], " + _
            "ZZ_SCRATCH_IMPORT_PEOPLE.c_dy = [BIOG_MAIN].[c_dy], ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_id = [BIOG_MAIN].[c_index_addr_id], " + _
            "ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_type = [BIOG_MAIN].[c_index_addr_type_code];"
    
    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecCount
    '
    ' now get the outer join data
    
    tQueryStr = "UPDATE ( ( ( ZZ_SCRATCH_IMPORT_PEOPLE LEFT JOIN INDEXYEAR_TYPE_CODES ON " + _
                    "ZZ_SCRATCH_IMPORT_PEOPLE.c_index_year_type_code = INDEXYEAR_TYPE_CODES.c_index_year_type_code ) LEFT JOIN DYNASTIES " + _
                    "ON ZZ_SCRATCH_IMPORT_PEOPLE.c_dy = DYNASTIES.c_dy ) LEFT JOIN ADDR_CODES " + _
                    "ON ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_id = ADDR_CODES.c_addr_id ) " + _
                    "LEFT JOIN BIOG_ADDR_CODES ON ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_type = BIOG_ADDR_CODES.c_addr_type " + _
                "SET ZZ_SCRATCH_IMPORT_PEOPLE.c_index_year_type_desc = [INDEXYEAR_TYPE_CODES].[c_index_year_type_desc], " + _
                    "ZZ_SCRATCH_IMPORT_PEOPLE.c_index_year_type_chn = [INDEXYEAR_TYPE_CODES].[c_index_year_type_hz], " + _
                    "ZZ_SCRATCH_IMPORT_PEOPLE.c_dynasty = [DYNASTIES].[c_dynasty], ZZ_SCRATCH_IMPORT_PEOPLE.c_dynasty_chn = [DYNASTIES].[c_dynasty_chn], " + _
                    "ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_name = [ADDR_CODES].[c_name], ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_chn = [ADDR_CODES].[c_name_chn], " + _
                    "ZZ_SCRATCH_IMPORT_PEOPLE.x_coord = [ADDR_CODES].[x_coord], ZZ_SCRATCH_IMPORT_PEOPLE.y_coord = [ADDR_CODES].[y_coord], " + _
                    "ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_desc = [BIOG_ADDR_CODES].[c_addr_desc], ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_desc_chn = [BIOG_ADDR_CODES].[c_addr_desc_chn]"

    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecCount
    '
    If ChkStatus.Value Then
        Call queryStatus
    End If
    
    If ChkOffice.Value Then
        Call queryOffice
    End If
    
    If ChkEntry.Value Then
        Call queryEntry
    End If
    
    If ChkText.Value Then
        Call queryText
    End If
    
    If ChkAddr.Value Then
        Call queryAddr
    End If
    
    If CmdGIS.Enabled Then
        CmdNeo4j.Enabled = True
    Else
        CmdNeo4j.Enabled = False
    End If
    
Exit_CmdRun_Click:
    '
    ' close the tables
    '
    Exit Sub

Err_CmdRun_Click:
    MsgBox Err.Description + tErrorStr
    Resume Exit_CmdRun_Click

End Sub


Private Sub Form_Open(Cancel As Integer)
    Dim tRstDummy As DAO.Recordset
    Dim cmdDel As ADODB.Command, tRecDeleted As Long
    '
    ' Clear the input and output tables
    '
    Set cmdDel = New ADODB.Command
    cmdDel.ActiveConnection = CurrentProject.Connection
    cmdDel.CommandType = adCmdText
    cmdDel.CommandText = "Delete * from ZZ_SCRATCH_IMPORT_PEOPLE"
    cmdDel.Execute tRecDeleted
    '
    ' clear status
    '
    Set Me.STATUS.Form.Recordset = CurrentDb.OpenRecordset("Z_SCRATCH_DUMMY_SC", dbOpenDynaset)
    cmdDel.CommandText = "Delete * from ZZ_SCRATCH_STATUS"
    cmdDel.Execute tRecDeleted
    Set Me.STATUS.Form.Recordset = CurrentDb.OpenRecordset("ZZ_SCRATCH_STATUS", dbOpenDynaset)
    
    cmdDel.CommandText = "DELETE * FROM ZZ_SCRATCH_P_STATUS"
    cmdDel.Execute gStatusRecCount
    '
    ' clear office
    Set Me.OFFICE.Form.Recordset = CurrentDb.OpenRecordset("Z_SCRATCH_DUMMY_OF", dbOpenDynaset)
    cmdDel.CommandText = "Delete * from ZZ_SCRATCH_OFFICE"
    cmdDel.Execute tRecDeleted
    Set Me.OFFICE.Form.Recordset = CurrentDb.OpenRecordset("ZZ_SCRATCH_OFFICE", dbOpenDynaset)
    '
    ' clear entry
    '
    Set Me.ENTRY.Form.Recordset = CurrentDb.OpenRecordset("Z_SCRATCH_DUMMY_ENTRY", dbOpenDynaset)
    cmdDel.CommandText = "Delete * from ZZ_SCRATCH_ENTRY"
    cmdDel.Execute tRecDeleted
    Set Me.ENTRY.Form.Recordset = CurrentDb.OpenRecordset("ZZ_SCRATCH_ENTRY", dbOpenDynaset)
    '
    ' clear text data
    '
    Set Me.TEXT.Form.Recordset = CurrentDb.OpenRecordset("Z_SCRATCH_DUMMY_TR", dbOpenDynaset)
    cmdDel.CommandText = "Delete * from ZZ_SCRATCH_BIOG_TEXT_DATA"
    cmdDel.Execute tRecDeleted
    Set Me.TEXT.Form.Recordset = CurrentDb.OpenRecordset("ZZ_SCRATCH_BIOG_TEXT_DATA", dbOpenDynaset)
    
    cmdDel.CommandText = "DELETE * FROM ZZ_SCRATCH_P_TEXT"
    cmdDel.Execute gStatusRecCount
    '
    ' clear place association
    '
    Set Me.PLACE.Form.Recordset = CurrentDb.OpenRecordset("Z_SCRATCH_DUMMY_BA", dbOpenDynaset)
    cmdDel.CommandText = "Delete * from ZZ_SCRATCH_BIOG_ADDR_DATA"
    cmdDel.Execute tRecDeleted
    Set Me.PLACE.Form.Recordset = CurrentDb.OpenRecordset("ZZ_SCRATCH_BIOG_ADDR_DATA", dbOpenDynaset)
    
    Set cmdDel = Nothing
        '
        'Set tRstDummy = CurrentDb.OpenRecordset("Z_SCRATCH_DUMMY_KIN", dbOpenDynaset)
        'Set Forms!LookAtKinship!frmZZ_SCRATCH_KIN.Form.Recordset = tRstDummy
        'gRstPersonID.Close
        'Set gRstPersonID = CurrentDb.OpenRecordset("ZZ_SCRATCH_KIN", dbOpenDynaset)
        'Set Forms!LookAtKinship!frmZZ_SCRATCH_KIN.Form.Recordset = gRstPersonID
        '
        'Set Forms!LookAtKinship!frmZZ_SCRATCH_KINNET.Form.Recordset = tRstDummy
        'gRstKinList.Close
        'Set gRstKinList = CurrentDb.OpenRecordset("ZZ_SCRATCH_KINNET", dbOpenDynaset)
        'Set Forms!LookAtKinship!frmZZ_SCRATCH_KINNET.Form.Recordset = gRstKinList
        '
        'Set tRstDummy = Nothing
    '
    '  first determine the language
    gLCID = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    If gLCID = 2052 Or gLCID = 4100 Then      ' 2052 = PRC, 4100 = Singapore
        gDisplayLanguage = "S"
    ElseIf gLCID = 3076 Or gLCID = 1028 Then  ' 3076 = Hong Kong, 1028 = Taiwan
        gDisplayLanguage = "T"
        'Call changeDisplayLanguage
    Else
        gDisplayLanguage = "E"
        'Call changeDisplayLanguage
    End If

    If DCount("*", "ZZ_STORE_PERSON_ID") > 0 Then
        CmdRecallID.Enabled = True
    End If
    '
    ' initially all 5 checkboxes are checked
    '
    gChkCount = 5
    gChkGisCount = 0
    gPeopleCount = 0
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
    Dim tLabelLanguage(3, 31) As String, tLang As Integer
    
    Dim tRstLabelList As DAO.Recordset, ti As Integer
    
    Set tRstLabelList = CurrentDb.OpenRecordset("FormLabels", dbOpenTable)
    
    tRstLabelList.Index = "label"
    
    gLabelsOK = False
    With tRstLabelList
        .MoveFirst
        ti = 1
        
        Do While ti < 31 And Not .EOF
            If !c_form = "LAG" Then
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
        Me.CmdImport.Caption = tLabelLanguage(tLang, 1)
        Me.CmdRecallID.Caption = tLabelLanguage(tLang, 2)
        Me.CmdRun.Caption = tLabelLanguage(tLang, 3)
        Me.CmdStoreID.Caption = tLabelLanguage(tLang, 4)
        Me.CmdFanti.Caption = tLabelLanguage(tLang, 5)
        Me.CmdJianti.Caption = tLabelLanguage(tLang, 6)
        Me.CmdHelp.Caption = tLabelLanguage(tLang, 7)
        Me.CmdClose.Caption = tLabelLanguage(tLang, 8)
        Me.CmdGIS.Caption = tLabelLanguage(tLang, 9)
        '
        Me.LblChkStatus.Caption = tLabelLanguage(tLang, 10)
        Me.LblChkOffice.Caption = tLabelLanguage(tLang, 11)
        Me.LblChkEntry.Caption = tLabelLanguage(tLang, 12)
        Me.LblChkText.Caption = tLabelLanguage(tLang, 13)
        Me.LblChkAddr.Caption = tLabelLanguage(tLang, 14)
        Me.LblChkGisStatus.Caption = tLabelLanguage(tLang, 15)
        Me.LblChkGisOfficeOffice.Caption = tLabelLanguage(tLang, 16)
        Me.LblChkGisOfficePeople.Caption = tLabelLanguage(tLang, 17)
        Me.LblChkGisEntry.Caption = tLabelLanguage(tLang, 18)
        Me.LblChkGisText.Caption = tLabelLanguage(tLang, 19)
        Me.LblChkGisAddr.Caption = tLabelLanguage(tLang, 20)
        '
        Me.PageStatus.Caption = tLabelLanguage(tLang, 21)
        Me.PageOffice.Caption = tLabelLanguage(tLang, 22)
        Me.PageEntry.Caption = tLabelLanguage(tLang, 23)
        Me.PageTexts.Caption = tLabelLanguage(tLang, 24)
        Me.PagePlace.Caption = tLabelLanguage(tLang, 25)
        '
        Me.LblOptIndexAddr.Caption = tLabelLanguage(tLang, 26)
        Me.LblOptAllAddr.Caption = tLabelLanguage(tLang, 27)
        Me.LblGIS_PY.Caption = tLabelLanguage(tLang, 28)
        '
        Me.CmdNeo4j.Caption = tLabelLanguage(tLang, 29)
        Me.LblDisplay.Caption = tLabelLanguage(tLang, 30)
    End If
    
End Sub

Private Sub CmdHelp_Click()
On Error GoTo Err_CmdHelp_Click
    Dim tStrPDF As String
    
    tStrPDF = Application.CurrentProject.Path + "\HelpFiles\HelpFile_LookAtKinship.pdf"
    
    'MsgBox tStrPDF
    
    Application.FollowHyperlink tStrPDF, , True
    

Exit_CmdHelp_Click:
    Exit Sub

Err_CmdHelp_Click:
    MsgBox Err.Description
    Resume Exit_CmdHelp_Click
    
End Sub

Private Sub WriteKML_Office()

    Dim tStrKML As String, tPinyin As Boolean
    '
    '  This program will dump the results to a .kml file
    '
    If gOfficeRecCount = 0 Then
        MsgBox "There are no records to save."
        GoTo Exit_WriteKML_Office
    End If
    '
    Dim tStream As ADODB.Stream
    Set tStream = New ADODB.Stream
        
    tPinyin = False
    If GISFrame.Value = 1 Then
        tStream.Charset = "utf-8"
        tCodeStr = "UTF8"
    ElseIf GISFrame.Value = 2 Then
        tStream.Charset = "gb2312"
        tCodeStr = "GB2312"
    Else
        tStream.Charset = "ascii"
        tCodeStr = "ASCII"
        tPinyin = True
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
    Dim tStr As String, tTab As String, ti As Integer
    Dim tFileSystem, tGDF
        
    Set dlgSaveAs = Application.FileDialog(msoFileDialogSaveAs)
    
    dlgSaveAs.InitialFileName = "office_office_gis_" + tCodeStr + ".kml"
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
            GoTo Exit_WriteKML_Office
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
        'SELECT ZZ_SCRATCH_OFFICE.c_name AS Name, ZZ_SCRATCH_OFFICE.c_name_chn AS NameChn,
        'ZZ_SCRATCH_OFFICE.c_index_year AS IndexYear, ZZ_SCRATCH_OFFICE.c_sex AS Sex,
        'ZZ_SCRATCH_OFFICE.c_addr_name AS AddrName, ZZ_SCRATCH_OFFICE.c_addr_chn AS AddrChn,
        'Str(ZZ_SCRATCH_OFFICE.x_coord) AS PersonX, Str(ZZ_SCRATCH_OFFICE.y_coord) AS PersonY,
        'ZZ_SCRATCH_OFFICE.c_office_trans AS Office, ZZ_SCRATCH_OFFICE.c_office_chn AS OfficeChn,
        'ZZ_SCRATCH_OFFICE.c_firstyear AS FirstYear, ZZ_SCRATCH_OFFICE.c_lastyear AS LastYear,
        'ZZ_SCRATCH_OFFICE.c_dy_desc AS Dynasty,
        'ZZ_SCRATCH_OFFICE.c_office_addr_name AS OfficeAddr,
        'ZZ_SCRATCH_OFFICE.c_office_addr_chn AS OfficeAddrChn,
        'Str(ZZ_SCRATCH_OFFICE.office_x_coord) AS X, Str(ZZ_SCRATCH_OFFICE.office_y_coord) AS Y,
        'ZZ_SCRATCH_OFFICE.office_xy_count AS XY_count
        
        '    tStr = "PostingID" (c_posting_id) + "Office" (c_name) + "OfficeChn" (c_name_chn) + _
                "FirstYear" (c_firstyear) + "LastYear" + (c_lastyear) _
                "Dynasty" (c_dy) + "OfficeAddr" (c_office_addr_name) + "OfficeAddrHZ" (c_office_addr_chn) + _
                "X" (office_x_coord) + "Y" (office_y_coord) + "xy_count" (office_xy_count)
        '
        ' process the table
        '
        Set tRstNode = OFFICE.Form.Recordset
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
        tStream.WriteText tC + "<Style id=" + tDQ + "office-balloon-template" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + "<BalloonStyle>", adWriteLine
        tStream.WriteText tC + tC + tC + "<text>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "<![CDATA[", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "ID: $[OfficePosting/PostingID] <br/>", adWriteLine
        If Not tPinyin Then
            tStream.WriteText tC + tC + tC + tC + "Person Chn: $[OfficePosting/PersonNameHZ] <br/>", adWriteLine
        End If
        tStream.WriteText tC + tC + tC + tC + "Begin Year: $[OfficePosting/BeginYear] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "End Year: $[OfficePosting/EndYear] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "Office Desc: $[OfficePosting/OfficeName] <br/>", adWriteLine
        If Not tPinyin Then
            tStream.WriteText tC + tC + tC + tC + "Office Chn: $[OfficePosting/OfficeNameHZ] <br/>", adWriteLine
        End If
        tStream.WriteText tC + tC + tC + tC + "Office Dynasty: $[OfficePosting/OfficeDyn] <br/>", adWriteLine
        If tPinyin Then
            tStream.WriteText tC + tC + tC + tC + "Address: $[OfficePosting/AddrName] <br/>", adWriteLine
        Else
            tStream.WriteText tC + tC + tC + tC + "Address: $[OfficePosting/AddrName] $[OfficePosting/AddrNameHZ] <br/>", adWriteLine
        End If
        tStream.WriteText tC + tC + tC + tC + "XY Count: $[OfficePosting/XYCount] <br/><br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "]]>", adWriteLine
        tStream.WriteText tC + tC + tC + "</text>", adWriteLine
        tStream.WriteText tC + tC + "</BalloonStyle>", adWriteLine
        tStream.WriteText tC + "</Style>", adWriteLine
        tStream.WriteText tC + "<!-- Declare the type " + tDQ + "OfficePosting" + tDQ + " with 6 fields -->", adWriteLine
        tStream.WriteText tC + "<Schema name=" + tDQ + "OfficePosting" + tDQ + " id=" + tDQ + "OfficePostingId" + tDQ + ">", adWriteLine
        If Not tPinyin Then
            tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "PersonNameHZ" + tDQ + ">", adWriteLine
            tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Person Chn]]></displayName>", adWriteLine
            tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        End If
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "AddrName" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Address]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        If Not tPinyin Then
            tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "AddrNameHZ" + tDQ + ">", adWriteLine
            tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Address Chn]]></displayName>", adWriteLine
            tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        End If
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "uint" + tDQ + " name=" + tDQ + "PostingID" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[ID]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "int" + tDQ + " name=" + tDQ + "FirstYear" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Begin Year]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "int" + tDQ + " name=" + tDQ + "LastYear" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[End Year]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "OfficeName" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Office Name]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        If Not tPinyin Then
            tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "OfficeNameHZ" + tDQ + ">", adWriteLine
            tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Office Chn]]></displayName>", adWriteLine
            tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        End If
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "int" + tDQ + " name=" + tDQ + "OfficeDyn" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Office Dyn]]></displayName>", adWriteLine
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
                
                If IsNull(!c_person_name) Then
                    tStr = "[Bad Data] " + Str(!c_posting_id)
                Else
                    tStr = !c_person_name + " " + Str(!c_posting_id)
                End If
                tStream.WriteText tC + tC + "<name>" + tStr + "</name>", adWriteLine
                
                tStream.WriteText tC + tC + "<styleUrl>#office-balloon-template</styleUrl>", adWriteLine
                '
                '  First Year as time stamp
                '
                If IsNull(!c_firstyear) Then
                    tStr = "0"
                Else
                    tStr = Str(!c_firstyear)
                End If
                tStream.WriteText tC + tC + "<TimeStamp>" + tStr + "</TimeStamp>", adWriteLine
                '
                tStream.WriteText tC + tC + "<ExtendedData>", adWriteLine
                tStream.WriteText tC + tC + tC + "<SchemaData schemaUrl=" + tDQ + "#OfficePostingID" + tDQ + ">", adWriteLine
                '
                '  posting ID
                '
                tStr = Str(!c_posting_id)
                tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "PostingID" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                '
                '  Person Name Chn
                '
                If Not tPinyin Then
                    If IsNull(!c_person_name_chn) Then
                        tStr = tStr + "[Bad Data]"
                    Else
                        If Trim(!c_person_name_chn) = "" Then
                            tStr = "[?]"
                        Else
                            tStr = !c_person_name_chn
                        End If
                    End If
                    tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "PersonNameHZ" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                End If
                '
                '  Office Name
                '
                If IsNull(!c_office_trans) Then
                    tStr = tStr + "[Bad Data]"
                Else
                    If Trim(!c_office_trans) = "" Then
                        tStr = "[?]"
                    Else
                        tStr = !c_office_trans
                    End If
                End If
                tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "OfficeName" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                '
                '  Office Chinese Name
                '
                If Not tPinyin Then
                    If IsNull(!c_office_chn) Then
                        tStr = tStr + "[Bad Data]"
                    Else
                        If Trim(!c_office_chn) = "" Then
                            tStr = "[?]"
                        Else
                            tStr = !c_office_chn
                        End If
                    End If
                    tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "OfficeNameHZ" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                End If
                '
                '  First Year
                '
                If IsNull(!c_firstyear) Then
                    tStr = "0"
                Else
                    tStr = Str(!c_firstyear)
                End If
                tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "FirstYear" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                '
                '  Last Year
                '
                If IsNull(!c_lastyear) Then
                    tStr = "0"
                Else
                    tStr = Str(!c_lastyear)
                End If
                tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "LastYear" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                '
                '  Office Dynasty
                '
                If IsNull(!c_dy) Then
                    tStr = "0"
                Else
                    tStr = Str(!c_dy)
                End If
                tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "OfficeDyn" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                '
                '  Address Name
                '
                If IsNull(!c_office_addr_name) Then
                    tStr = "[?]"
                ElseIf Trim(!c_office_addr_name) = "" Then
                    tStr = "[?]"
                Else
                    tStr = !c_office_addr_name
                End If
                tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "AddrName" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                '
                '  Address Name Chinese
                '
                If Not tPinyin Then
                    If IsNull(!c_office_addr_chn) Then
                        tStr = "[?]"
                    ElseIf Trim(!c_office_addr_chn) = "" Then
                        tStr = "[?]"
                    Else
                        tStr = !c_office_addr_chn
                    End If
                    tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "AddrNameHZ" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                End If
                '
                '  XY Count
                '
                If IsNull(!office_xy_count) Then
                    tStr = "0"
                Else
                    tStr = Str(!office_xy_count)
                End If
                tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "XYCount" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                '
                tStream.WriteText tC + tC + tC + "</SchemaData>", adWriteLine
                tStream.WriteText tC + tC + "</ExtendedData>", adWriteLine
                tStream.WriteText tC + tC + "<Point>", adWriteLine
                '
                '  coordinates
                If IsNull(!office_x_coord) Then
                    tStr = "0"
                Else
                    tStr = Str(!office_x_coord)
                End If
                
                If IsNull(!office_y_coord) Then
                    tStr = tStr + ",0"
                Else
                    tStr = tStr + "," + Str(!office_y_coord)
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
    
Exit_WriteKML_Office:
    Exit Sub

Err_WriteKML_Office:
    MsgBox Err.Description
    Resume Exit_WriteKML_Office
    

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

    tStrQuery = "INSERT INTO ZZ_STORE_PERSON_ID ( c_personid ) SELECT DISTINCT ZZ_SCRATCH_IMPORT_PEOPLE.c_person_id FROM ZZ_SCRATCH_IMPORT_PEOPLE"
    
    cmdSQL.CommandText = tStrQuery
    cmdSQL.Execute tRecCount
    MsgBox "Person IDs successfully stored.  Click on 'Recall Person IDs' to reuse these IDs in other forms."

End Sub
Private Sub CmdRecallID_Click()
On Error GoTo Err_CmdRecallID_Click
    Dim tStrSQL As String, tRecCount As Long, tStrQuestion As String, tRst As DAO.Recordset, tID As Long

    Set cmdSQL = New ADODB.Command
    cmdSQL.ActiveConnection = CurrentProject.Connection
    cmdSQL.CommandType = adCmdText
    
    tRecCount = DCount("*", "ZZ_SCRATCH_IMPORT_PEOPLE")
    
    If tRecCount > 0 Then
        If tRecCount = 1 Then
            tStrQuestion = "Do you wish to replace the current person?"
        Else
            tStrQuestion = "Do you wish to replace the current list of IDs?"
        End If
        ' Display message.
        If MsgBox(tStrQuestion, vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
            Exit Sub
        Else
            cmdSQL.CommandText = "Delete * from ZZ_SCRATCH_IMPORT_PEOPLE"
            cmdSQL.Execute tRecCount
        End If
    End If
    '
    ' Clear the error table now that we are ready to go
    '
    cmdSQL.CommandText = "Delete * from InputErrorList"
    cmdSQL.Execute tRecDeleted
    '
    '  copy the IDs
    '
    tStrSQL = "INSERT INTO ZZ_SCRATCH_IMPORT_PEOPLE ( c_person_id ) SELECT DISTINCT c_personid FROM ZZ_STORE_PERSON_ID"

    cmdSQL.CommandText = tStrSQL
    cmdSQL.Execute gPeopleCount
        
    If gPeopleCount = 0 Then
        CmdRun.Enabled = False
        CmdStoreID.Enabled = False
    Else
        If gChkCount > 0 Then
            CmdRun.Enabled = True
        End If
        CmdStoreID.Enabled = True
    End If
        
    Set cmdSQL = Nothing
    
Exit_CmdRecallID_Click:
    Exit Sub

Err_CmdRecallID_Click:
    MsgBox Err.Description
    Resume Exit_CmdRecallID_Click

End Sub
Private Sub queryStatus()
    Dim cmdSQL As ADODB.Command, tStrInsert As String, tStrSelect As String, tStrFrom As String, tRst As DAO.Recordset
    
    Set cmdSQL = New ADODB.Command
    cmdSQL.ActiveConnection = CurrentProject.Connection
    cmdSQL.CommandType = adCmdText
    
    '  first clear ZZ_SCRATCH_STATUS
    
    Set Me.STATUS.Form.Recordset = CurrentDb.OpenRecordset("Z_SCRATCH_DUMMY_SC", dbOpenDynaset)
    
    cmdSQL.CommandText = "DELETE * FROM ZZ_SCRATCH_STATUS"
    cmdSQL.Execute gStatusRecCount

    cmdSQL.CommandText = "DELETE * FROM ZZ_SCRATCH_P_STATUS"
    cmdSQL.Execute gStatusRecCount

    tStrInsert = "INSERT INTO ZZ_SCRATCH_STATUS (c_personid, c_sequence, c_status_code, c_status_desc, c_status_desc_chn, c_firstyear, " + _
                    "c_fy_nh_code, c_fy_nh_year, c_fy_range, c_lastyear, c_ly_nh_code, c_ly_nh_year, c_ly_range, c_source, " + _
                    "c_pages, c_notes, c_name, c_name_chn, c_sex, c_index_year, c_index_year_type_code, c_index_year_type_desc, " + _
                    "c_index_year_type_hz, c_dy, c_dynasty, c_dynasty_chn, c_addr_id, c_addr_name, c_addr_chn, x_coord, y_coord, c_ethnicity_code ) "
    tStrSelect = "SELECT STATUS_DATA.c_personid, STATUS_DATA.c_sequence, STATUS_DATA.c_status_code, STATUS_CODES.c_status_desc, STATUS_CODES.c_status_desc_chn, " + _
                    "STATUS_DATA.c_firstyear, STATUS_DATA.c_fy_nh_code, STATUS_DATA.c_fy_nh_year, STATUS_DATA.c_fy_range, STATUS_DATA.c_lastyear, " + _
                    "STATUS_DATA.c_ly_nh_code, STATUS_DATA.c_ly_nh_year, STATUS_DATA.c_ly_range, STATUS_DATA.c_source, STATUS_DATA.c_pages, " + _
                    "STATUS_DATA.c_notes, ZZ_SCRATCH_IMPORT_PEOPLE.c_name, ZZ_SCRATCH_IMPORT_PEOPLE.c_name_chn, ZZ_SCRATCH_IMPORT_PEOPLE.c_sex, " + _
                    "ZZ_SCRATCH_IMPORT_PEOPLE.c_index_year, ZZ_SCRATCH_IMPORT_PEOPLE.c_index_year_type_code, " + _
                    "ZZ_SCRATCH_IMPORT_PEOPLE.c_index_year_type_desc, ZZ_SCRATCH_IMPORT_PEOPLE.c_index_year_type_chn, " + _
                    "ZZ_SCRATCH_IMPORT_PEOPLE.c_dy, ZZ_SCRATCH_IMPORT_PEOPLE.c_dynasty, ZZ_SCRATCH_IMPORT_PEOPLE.c_dynasty_chn, " + _
                    "ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_id, ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_name, ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_chn, " + _
                    "ZZ_SCRATCH_IMPORT_PEOPLE.x_coord, ZZ_SCRATCH_IMPORT_PEOPLE.y_coord, BIOG_MAIN.c_ethnicity_code "
    tStrFrom = "FROM ( BIOG_MAIN INNER JOIN ( STATUS_CODES INNER JOIN STATUS_DATA ON STATUS_CODES.c_status_code = STATUS_DATA.c_status_code ) " + _
                    "ON BIOG_MAIN.c_personid = STATUS_DATA.c_personid ) INNER JOIN ZZ_SCRATCH_IMPORT_PEOPLE " + _
                    "ON BIOG_MAIN.c_personid = ZZ_SCRATCH_IMPORT_PEOPLE.c_person_id "

     cmdSQL.CommandText = tStrInsert + tStrSelect + tStrFrom
    cmdSQL.Execute gStatusRecCount
    '
    ' there are outer join fields to update
    '
    If gStatusRecCount > 0 Then
        cmdSQL.CommandText = "UPDATE ( ( ( ( ( ZZ_SCRATCH_STATUS LEFT JOIN TEXT_CODES ON ZZ_SCRATCH_STATUS.c_source = TEXT_CODES.c_textid ) " + _
            "LEFT JOIN ETHNICITY_TRIBE_CODES ON ZZ_SCRATCH_STATUS.c_ethnicity_code = ETHNICITY_TRIBE_CODES.c_ethnicity_code ) " + _
            "LEFT JOIN NIAN_HAO ON ZZ_SCRATCH_STATUS.c_fy_nh_code = NIAN_HAO.c_nianhao_id ) " + _
            "LEFT JOIN NIAN_HAO AS NIAN_HAO_1 ON ZZ_SCRATCH_STATUS.c_ly_nh_code = NIAN_HAO_1.c_nianhao_id ) " + _
            "LEFT JOIN YEAR_RANGE_CODES ON ZZ_SCRATCH_STATUS.c_fy_range = YEAR_RANGE_CODES.c_range_code ) " + _
            "LEFT JOIN YEAR_RANGE_CODES AS YEAR_RANGE_CODES_1 ON ZZ_SCRATCH_STATUS.c_ly_range = YEAR_RANGE_CODES_1.c_range_code " + _
        "SET ZZ_SCRATCH_STATUS.c_fy_nh_chn = [NIAN_HAO].[c_nianhao_chn], ZZ_SCRATCH_STATUS.c_fy_nh_py = [NIAN_HAO].[c_nianhao_pin], " + _
            "ZZ_SCRATCH_STATUS.c_fy_range_desc = [YEAR_RANGE_CODES].[c_range], ZZ_SCRATCH_STATUS.c_fy_range_chn = [YEAR_RANGE_CODES].[c_range_chn], " + _
            "ZZ_SCRATCH_STATUS.c_ly_nh_chn = [NIAN_HAO_1].[c_nianhao_chn], ZZ_SCRATCH_STATUS.c_ly_nh_py = [NIAN_HAO_1].[c_nianhao_pin], " + _
            "ZZ_SCRATCH_STATUS.c_ly_range_desc = [YEAR_RANGE_CODES_1].[c_range], ZZ_SCRATCH_STATUS.c_ly_range_chn = [YEAR_RANGE_CODES_1].[c_range_chn], " + _
            "ZZ_SCRATCH_STATUS.c_title_chn = [TEXT_CODES].[c_title_chn], ZZ_SCRATCH_STATUS.c_title = [TEXT_CODES].[c_title], " + _
            "ZZ_SCRATCH_STATUS.c_ethnicity_chn = [ETHNICITY_TRIBE_CODES].[c_name_chn], ZZ_SCRATCH_STATUS.c_ethnicity_rmn = [ETHNICITY_TRIBE_CODES].[c_name]"
        cmdSQL.Execute tsRecCount
    End If
    '
    '  the final step is to insert the people into P_STATUS and calculate the xy_count
    '
    If gStatusRecCount > 0 Then
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
        '  calculate the xy_count
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
            
        Me.ChkGisStatus.Enabled = True
        If Me.ChkGisStatus.Value = False Then
            gChkGisCount = gChkGisCount + 1
            Me.ChkGisStatus.Value = True
        End If
        CmdGIS.Enabled = True
    Else
        If ChkGisStatus.Value Then
            Me.ChkGisStatus.Value = False
            gChkGisCount = gChkGisCount - 1
        End If
        Me.ChkGisStatus.Enabled = False
        
        If gChkGisCount = 0 Then
            CmdGIS.Enabled = False
        End If
    End If
    
    Set Me.STATUS.Form.Recordset = CurrentDb.OpenRecordset("ZZ_SCRATCH_STATUS", dbOpenDynaset)
End Sub
Private Sub queryOffice()
    Dim cmdSQL As ADODB.Command, tStrInsert As String, tStrSelect As String, tStrFrom As String, tRst As DAO.Recordset
    
    Set cmdSQL = New ADODB.Command
    cmdSQL.ActiveConnection = CurrentProject.Connection
    cmdSQL.CommandType = adCmdText
    
    '  first clear ZZ_SCRATCH_OFFICE
    Set Me.OFFICE.Form.Recordset = CurrentDb.OpenRecordset("Z_SCRATCH_DUMMY_OF", dbOpenDynaset)
    
    cmdSQL.CommandText = "DELETE * FROM ZZ_SCRATCH_OFFICE"
    cmdSQL.Execute gOfficeRecCount

    tStrInsert = "INSERT INTO ZZ_SCRATCH_OFFICE (c_personid, c_person_name, c_person_name_chn, c_female, c_sex, c_index_year, c_index_year_type_code, " + _
                    "c_index_year_type_desc, c_index_year_type_hz, c_person_dy, c_person_dynasty, c_person_dy_chn, c_addr_id, " + _
                    "c_addr_name, c_addr_chn, c_addr_type, c_addr_desc, c_addr_desc_chn, x_coord, y_coord, c_office_id, c_office_pinyin, " + _
                    "c_office_chn, c_office_trans, c_sequence, c_firstyear, c_fy_nh_code, c_fy_nh_year, c_fy_range, c_lastyear, " + _
                    "c_ly_nh_code, c_ly_nh_year, c_ly_range, c_appt_code, c_assume_office_code, c_office_category_id, c_inst_code, " + _
                    "c_inst_name_code, c_source, c_pages, c_notes, c_fy_intercalary, c_fy_month, c_ly_intercalary, c_ly_month, " + _
                    "c_fy_day, c_ly_day, c_fy_day_gz, c_ly_day_gz, c_dy, c_office_addr_id, c_office_addr_name, c_office_addr_chn, " + _
                    "office_x_coord, office_y_coord ) "

    tStrSelect = "SELECT POSTED_TO_ADDR_DATA.c_personid, ZZ_SCRATCH_IMPORT_PEOPLE.c_name, ZZ_SCRATCH_IMPORT_PEOPLE.c_name_chn, ZZ_SCRATCH_IMPORT_PEOPLE.c_female, " + _
                    "ZZ_SCRATCH_IMPORT_PEOPLE.c_sex, ZZ_SCRATCH_IMPORT_PEOPLE.c_index_year, ZZ_SCRATCH_IMPORT_PEOPLE.c_index_year_type_code, " + _
                    "ZZ_SCRATCH_IMPORT_PEOPLE.c_index_year_type_desc, ZZ_SCRATCH_IMPORT_PEOPLE.c_index_year_type_chn, ZZ_SCRATCH_IMPORT_PEOPLE.c_dy, " + _
                    "ZZ_SCRATCH_IMPORT_PEOPLE.c_dynasty, ZZ_SCRATCH_IMPORT_PEOPLE.c_dynasty_chn, ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_id, " + _
                    "ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_name, ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_chn, ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_type, " + _
                    "ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_desc, ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_desc_chn, ZZ_SCRATCH_IMPORT_PEOPLE.x_coord, " + _
                    "ZZ_SCRATCH_IMPORT_PEOPLE.y_coord, POSTED_TO_ADDR_DATA.c_office_id, OFFICE_CODES.c_office_pinyin, OFFICE_CODES.c_office_chn, " + _
                    "OFFICE_CODES.c_office_trans, POSTED_TO_OFFICE_DATA.c_sequence, POSTED_TO_OFFICE_DATA.c_firstyear, " + _
                    "POSTED_TO_OFFICE_DATA.c_fy_nh_code, POSTED_TO_OFFICE_DATA.c_fy_nh_year, POSTED_TO_OFFICE_DATA.c_fy_range, " + _
                    "POSTED_TO_OFFICE_DATA.c_lastyear, POSTED_TO_OFFICE_DATA.c_ly_nh_code, POSTED_TO_OFFICE_DATA.c_ly_nh_year, " + _
                    "POSTED_TO_OFFICE_DATA.c_ly_range, POSTED_TO_OFFICE_DATA.c_appt_code, POSTED_TO_OFFICE_DATA.c_assume_office_code, " + _
                    "POSTED_TO_OFFICE_DATA.c_office_category_id, POSTED_TO_OFFICE_DATA.c_inst_code, POSTED_TO_OFFICE_DATA.c_inst_name_code, " + _
                    "POSTED_TO_OFFICE_DATA.c_source, POSTED_TO_OFFICE_DATA.c_pages, POSTED_TO_OFFICE_DATA.c_notes, " + _
                    "POSTED_TO_OFFICE_DATA.c_fy_intercalary, POSTED_TO_OFFICE_DATA.c_fy_month, POSTED_TO_OFFICE_DATA.c_ly_intercalary, " + _
                    "POSTED_TO_OFFICE_DATA.c_ly_month, POSTED_TO_OFFICE_DATA.c_fy_day, POSTED_TO_OFFICE_DATA.c_ly_day, " + _
                    "POSTED_TO_OFFICE_DATA.c_fy_day_gz, POSTED_TO_OFFICE_DATA.c_ly_day_gz, POSTED_TO_OFFICE_DATA.c_dy, " + _
                    "POSTED_TO_ADDR_DATA.c_addr_id, ADDR_CODES.c_name, ADDR_CODES.c_name_chn, ADDR_CODES.x_coord, ADDR_CODES.y_coord "

    tStrFrom = "FROM ADDR_CODES INNER JOIN ( ( POSTED_TO_OFFICE_DATA INNER JOIN ( POSTED_TO_ADDR_DATA INNER JOIN ZZ_SCRATCH_IMPORT_PEOPLE " + _
                    "ON POSTED_TO_ADDR_DATA.c_personid = ZZ_SCRATCH_IMPORT_PEOPLE.c_person_id ) " + _
                    "ON POSTED_TO_OFFICE_DATA.c_posting_id = POSTED_TO_ADDR_DATA.c_posting_id ) " + _
                    "INNER JOIN OFFICE_CODES ON POSTED_TO_ADDR_DATA.c_office_id = OFFICE_CODES.c_office_id ) " + _
                    "ON ADDR_CODES.c_addr_id = POSTED_TO_ADDR_DATA.c_addr_id"

    cmdSQL.CommandText = tStrInsert + tStrSelect + tStrFrom
    cmdSQL.Execute gOfficeRecCount
    '
    ' now fill in the outer join fields
    '
    If gOfficeRecCount > 0 Then
        tStrUpdate = "UPDATE ( ( ( ( ( ( ( ( ( ( ZZ_SCRATCH_OFFICE LEFT JOIN SOCIAL_INSTITUTION_NAME_CODES " + _
                    "ON ZZ_SCRATCH_OFFICE.c_inst_name_code = SOCIAL_INSTITUTION_NAME_CODES.c_inst_name_code ) LEFT JOIN TEXT_CODES " + _
                    "ON ZZ_SCRATCH_OFFICE.c_source = TEXT_CODES.c_textid ) LEFT JOIN NIAN_HAO " + _
                    "ON ZZ_SCRATCH_OFFICE.c_fy_nh_code = NIAN_HAO.c_nianhao_id ) " + _
                    "LEFT JOIN YEAR_RANGE_CODES ON ZZ_SCRATCH_OFFICE.c_fy_range = YEAR_RANGE_CODES.c_range_code ) " + _
                    "LEFT JOIN NIAN_HAO AS NIAN_HAO_1 " + _
                    "ON ZZ_SCRATCH_OFFICE.c_ly_nh_code = NIAN_HAO_1.c_nianhao_id ) LEFT JOIN YEAR_RANGE_CODES AS YEAR_RANGE_CODES_1 " + _
                    "ON ZZ_SCRATCH_OFFICE.c_ly_range = YEAR_RANGE_CODES_1.c_range_code ) LEFT JOIN APPOINTMENT_CODES " + _
                    "ON ZZ_SCRATCH_OFFICE.c_appt_code = APPOINTMENT_CODES.c_appt_code ) LEFT JOIN ASSUME_OFFICE_CODES " + _
                    "ON ZZ_SCRATCH_OFFICE.c_assume_office_code = ASSUME_OFFICE_CODES.c_assume_office_code ) LEFT JOIN OFFICE_CATEGORIES " + _
                    "ON ZZ_SCRATCH_OFFICE.c_office_category_id = OFFICE_CATEGORIES.c_office_category_id ) LEFT JOIN GANZHI_CODES " + _
                    "ON ZZ_SCRATCH_OFFICE.c_fy_day_gz = GANZHI_CODES.c_ganzhi_code ) LEFT JOIN GANZHI_CODES AS GANZHI_CODES_1 " + _
                    "ON ZZ_SCRATCH_OFFICE.c_ly_day_gz = GANZHI_CODES_1.c_ganzhi_code "
        tStrSet = "SET ZZ_SCRATCH_OFFICE.c_fy_nh_chn = [NIAN_HAO].[c_nianhao_chn], ZZ_SCRATCH_OFFICE.c_fy_nh_py = [NIAN_HAO].[c_nianhao_pin], " + _
                    "ZZ_SCRATCH_OFFICE.c_fy_range_desc = [YEAR_RANGE_CODES].[c_range], " + _
                    "ZZ_SCRATCH_OFFICE.c_fy_range_chn = [YEAR_RANGE_CODES].[c_range_chn], " + _
                    "ZZ_SCRATCH_OFFICE.c_ly_nh_chn = [NIAN_HAO_1].[c_nianhao_chn], ZZ_SCRATCH_OFFICE.c_ly_nh_py = [NIAN_HAO_1].[c_nianhao_pin], " + _
                    "ZZ_SCRATCH_OFFICE.c_ly_range_desc = [YEAR_RANGE_CODES_1].[c_range], " + _
                    "ZZ_SCRATCH_OFFICE.c_ly_range_chn = [YEAR_RANGE_CODES_1].[c_range_chn], " + _
                    "ZZ_SCRATCH_OFFICE.c_appt_desc_chn = [APPOINTMENT_CODES].[c_appt_desc_chn], " + _
                    "ZZ_SCRATCH_OFFICE.c_appt_desc = [APPOINTMENT_CODES].[c_appt_desc], ZZ_SCRATCH_OFFICE.c_title_chn = [TEXT_CODES].[c_title_chn], " + _
                    "ZZ_SCRATCH_OFFICE.c_assume_office_desc_chn = [ASSUME_OFFICE_CODES].[c_assume_office_desc_chn], " + _
                    "ZZ_SCRATCH_OFFICE.c_assume_office_desc = [ASSUME_OFFICE_CODES].[c_assume_office_desc], " + _
                    "ZZ_SCRATCH_OFFICE.c_inst_name_hz = [SOCIAL_INSTITUTION_NAME_CODES].[c_inst_name_hz], " + _
                    "ZZ_SCRATCH_OFFICE.c_inst_name_py = [SOCIAL_INSTITUTION_NAME_CODES].[c_inst_name_py], " + _
                    "ZZ_SCRATCH_OFFICE.c_title = [TEXT_CODES].[c_title], ZZ_SCRATCH_OFFICE.c_fy_day_gz_chn = [GANZHI_CODES].[c_ganzhi_chn], " + _
                    "ZZ_SCRATCH_OFFICE.c_fy_day_gz_py = [GANZHI_CODES].[c_ganzhi_py], " + _
                    "ZZ_SCRATCH_OFFICE.c_ly_day_gz_chn = [GANZHI_CODES_1].[c_ganzhi_chn], " + _
                    "ZZ_SCRATCH_OFFICE.c_ly_day_gz_py = [GANZHI_CODES_1].[c_ganzhi_py], " + _
                    "ZZ_SCRATCH_OFFICE.c_category_desc = [OFFICE_CATEGORIES].[c_category_desc], " + _
                    "ZZ_SCRATCH_OFFICE.c_category_desc_chn = [OFFICE_CATEGORIES].[c_category_desc_chn]"

        cmdSQL.CommandText = tStrUpdate + tStrSet
        cmdSQL.Execute tRecCount
    End If
    '
    '  the next step is to calculate the xy_count for office addresses
    '
     If gOfficeRecCount > 0 Then
        '
        '  Now the People
        '
        tStrQuery = "INSERT INTO ZZ_SCRATCH_P_OFFICE ( c_personid, c_name, c_name_chn, c_index_year, " + _
            "c_index_year_type_code, c_index_year_type_desc, c_index_year_type_hz, " + _
            "c_female, c_sex, c_dy, c_dynasty, c_dynasty_chn, c_addr_id, c_addr_name, c_addr_chn, c_addr_type, " + _
            "c_addr_desc, c_addr_desc_chn, x_coord, y_coord ) " + _
            "SELECT DISTINCT ZZ_SCRATCH_OFFICE.c_personid, ZZ_SCRATCH_OFFICE.c_person_name AS c_name, " + _
            "ZZ_SCRATCH_OFFICE.c_person_name_chn AS c_name_chn, ZZ_SCRATCH_OFFICE.c_index_year, " + _
            "ZZ_SCRATCH_OFFICE.c_index_year_type_code, ZZ_SCRATCH_OFFICE.c_index_year_type_desc, ZZ_SCRATCH_OFFICE.c_index_year_type_hz, " + _
            "ZZ_SCRATCH_OFFICE.c_female, ZZ_SCRATCH_OFFICE.c_sex, ZZ_SCRATCH_OFFICE.c_person_dy AS c_dy, ZZ_SCRATCH_OFFICE.c_person_dynasty AS c_dynasty, " + _
            "ZZ_SCRATCH_OFFICE.c_person_dy_chn AS c_dynasty_chn, ZZ_SCRATCH_OFFICE.c_addr_id, " + _
            "ZZ_SCRATCH_OFFICE.c_addr_name, ZZ_SCRATCH_OFFICE.c_addr_chn, ZZ_SCRATCH_OFFICE.c_addr_type, " + _
            "ZZ_SCRATCH_OFFICE.c_addr_desc, ZZ_SCRATCH_OFFICE.c_addr_desc_chn, ZZ_SCRATCH_OFFICE.x_coord, " + _
            "ZZ_SCRATCH_OFFICE.y_coord " + _
            "FROM ZZ_SCRATCH_OFFICE"
        '
        cmdSQL.CommandText = tStrQuery
        cmdSQL.Execute tRecDeleted
        '
        ' first the people xy
        '
        ' use three SQL calls
        '
        cmdSQL.CommandText = "DELETE * FROM tmpXY"
        cmdSQL.Execute tRecDeleted
        '
        tStrQuery = "INSERT INTO tmpXY ( x_coord, y_coord, CountOfx_coord, CountOfy_coord ) " + _
            "SELECT ZZ_SCRATCH_P_OFFICE.x_coord, ZZ_SCRATCH_P_OFFICE.y_coord, Count(ZZ_SCRATCH_P_OFFICE.x_coord) " + _
            "AS CountOfx_coord, Count(ZZ_SCRATCH_P_OFFICE.y_coord) AS CountOfy_coord " + _
            "FROM ZZ_SCRATCH_P_OFFICE " + _
            "GROUP BY ZZ_SCRATCH_P_OFFICE.x_coord, ZZ_SCRATCH_P_OFFICE.y_coord;"

        '
        cmdSQL.CommandText = tStrQuery
        cmdSQL.Execute tRecDeleted
        '
        tStrQuery = "UPDATE tmpXY INNER JOIN ZZ_SCRATCH_P_OFFICE ON (tmpXY.y_coord = " + _
            "ZZ_SCRATCH_P_OFFICE.y_coord) AND (tmpXY.x_coord = ZZ_SCRATCH_P_OFFICE.x_coord) " + _
            "SET ZZ_SCRATCH_P_OFFICE.xy_count = [tmpXY].[CountOfx_coord];"

        cmdSQL.CommandText = tStrQuery
        cmdSQL.Execute tRecDeleted
        '
        ' then the offices
        '
        cmdSQL.CommandText = "DELETE * FROM tmpXY"
        cmdSQL.Execute tRecDeleted
        '
        tStrQuery = "INSERT INTO tmpXY ( x_coord, y_coord, CountOfx_coord, CountOfy_coord ) " + _
            "SELECT ZZ_SCRATCH_OFFICE.office_x_coord, ZZ_SCRATCH_OFFICE.office_y_coord, " + _
            "Count(ZZ_SCRATCH_OFFICE.office_x_coord) AS CountOfx_coord, " + _
            "Count(ZZ_SCRATCH_OFFICE.office_y_coord) AS CountOfy_coord " + _
            "FROM ZZ_SCRATCH_OFFICE " + _
            "GROUP BY ZZ_SCRATCH_OFFICE.office_x_coord, ZZ_SCRATCH_OFFICE.office_y_coord"
        '
        cmdSQL.CommandText = tStrQuery
        cmdSQL.Execute tRecDeleted
            
        tStrQuery = "UPDATE tmpXY INNER JOIN ZZ_SCRATCH_OFFICE ON (tmpXY.y_coord = " + _
            "ZZ_SCRATCH_OFFICE.office_y_coord) AND (tmpXY.x_coord = ZZ_SCRATCH_OFFICE.office_x_coord) " + _
            "SET ZZ_SCRATCH_OFFICE.office_xy_count = [tmpXY].[CountOfx_coord];"

        cmdSQL.CommandText = tStrQuery
        cmdSQL.Execute tRecDeleted
        
        Me.ChkGisOffice.Enabled = True
        Me.ChkGisOfficePeople.Enabled = True
        If Me.ChkGisOffice.Value = False Then
            gChkGisCount = gChkGisCount + 1
            Me.ChkGisOffice.Value = True
        End If
        If Me.ChkGisOfficePeople.Value = False Then
            gChkGisCount = gChkGisCount + 1
            Me.ChkGisOfficePeople.Value = True
        End If
        CmdGIS.Enabled = True
    Else
        If ChkGisOffice.Value Then
            Me.ChkGisOffice.Value = False
            gChkGisCount = gChkGisCount - 1
        End If
        If ChkGisOfficePeople.Value Then
            Me.ChkGisOfficePeople.Value = False
            gChkGisCount = gChkGisCount - 1
        End If
        Me.ChkGisOfficePeople.Enabled = False
        Me.ChkGisOffice.Enabled = False
        
        If gChkGisCount = 0 Then
            CmdGIS.Enabled = False
        End If
    End If
    '
    Set Me.OFFICE.Form.Recordset = CurrentDb.OpenRecordset("ZZ_SCRATCH_OFFICE", dbOpenDynaset)
End Sub
Private Sub queryEntry()
    Dim cmdSQL As ADODB.Command, tStrInsert As String, tStrSelect As String, tStrFrom As String, tRst As DAO.Recordset
    
    Set cmdSQL = New ADODB.Command
    cmdSQL.ActiveConnection = CurrentProject.Connection
    cmdSQL.CommandType = adCmdText
    
    '  first clear ZZ_SCRATCH_ENTRY
    
    Set Me.ENTRY.Form.Recordset = CurrentDb.OpenRecordset("Z_SCRATCH_DUMMY_ENTRY", dbOpenDynaset)
    
    cmdSQL.CommandText = "DELETE * FROM ZZ_SCRATCH_ENTRY"
    cmdSQL.Execute gEntryRecCount
    '
    ' stopped here

    tStrInsert = "INSERT INTO ZZ_SCRATCH_ENTRY ( c_personid, c_name, c_name_chn, c_index_year, c_index_year_type_code, c_index_year_type_desc, c_index_year_type_hz, " + _
                "c_dy, c_dynasty, c_dynasty_chn, c_addr_id, c_addr_name, c_addr_chn, c_addr_type, c_addr_desc, c_addr_desc_chn, x_coord, y_coord, " + _
                "c_entry_code, c_entry_desc, c_entry_chn, c_sequence, c_exam_rank, c_kin_code, c_kin_id, c_assoc_code, c_assoc_id, c_year, c_age, " + _
                "c_inst_code, c_inst_name_code, c_entry_addr_id, c_source, c_parental_status_code ) "
    tStrSelect = "SELECT ZZ_SCRATCH_IMPORT_PEOPLE.c_person_id, ZZ_SCRATCH_IMPORT_PEOPLE.c_name, ZZ_SCRATCH_IMPORT_PEOPLE.c_name_chn, " + _
                "ZZ_SCRATCH_IMPORT_PEOPLE.c_index_year, ZZ_SCRATCH_IMPORT_PEOPLE.c_index_year_type_code, ZZ_SCRATCH_IMPORT_PEOPLE.c_index_year_type_desc, " + _
                "ZZ_SCRATCH_IMPORT_PEOPLE.c_index_year_type_chn, ZZ_SCRATCH_IMPORT_PEOPLE.c_dy, ZZ_SCRATCH_IMPORT_PEOPLE.c_dynasty, " + _
                "ZZ_SCRATCH_IMPORT_PEOPLE.c_dynasty_chn, ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_id, ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_name, " + _
                "ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_chn, ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_type, ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_desc, " + _
                "ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_desc_chn, ZZ_SCRATCH_IMPORT_PEOPLE.x_coord, ZZ_SCRATCH_IMPORT_PEOPLE.y_coord, ENTRY_DATA.c_entry_code, " + _
                "ENTRY_CODES.c_entry_desc, ENTRY_CODES.c_entry_desc_chn, ENTRY_DATA.c_sequence,  ENTRY_DATA.c_exam_rank, ENTRY_DATA.c_kin_code, " + _
                "ENTRY_DATA.c_kin_id, ENTRY_DATA.c_assoc_code, ENTRY_DATA.c_assoc_id, ENTRY_DATA.c_year, ENTRY_DATA.c_age, ENTRY_DATA.c_inst_code, " + _
                "ENTRY_DATA.c_inst_name_code, ENTRY_DATA.c_entry_addr_id, ENTRY_DATA.c_source, ENTRY_DATA.c_parental_status_code "
    tStrFrom = "FROM ZZ_SCRATCH_IMPORT_PEOPLE INNER JOIN ( ENTRY_CODES INNER JOIN ENTRY_DATA ON ENTRY_CODES.c_entry_code = ENTRY_DATA.c_entry_code ) " + _
                "ON ZZ_SCRATCH_IMPORT_PEOPLE.c_person_id = ENTRY_DATA.c_personid"
   
    cmdSQL.CommandText = tStrInsert + tStrSelect + tStrFrom
    cmdSQL.Execute gEntryRecCount
    '
    '  the final step is to calculate the xy_count
    '
    If gEntryRecCount > 0 Then
        '
    ' get the outer join field information
    '
    tQueryStr = "UPDATE ( ( ( ( ( ( ( ZZ_SCRATCH_ENTRY LEFT JOIN BIOG_MAIN ON ZZ_SCRATCH_ENTRY.c_kin_id = BIOG_MAIN.c_personid ) LEFT JOIN ASSOC_CODES " + _
            "ON ZZ_SCRATCH_ENTRY.c_assoc_code = ASSOC_CODES.c_assoc_code ) LEFT JOIN KINSHIP_CODES ON ZZ_SCRATCH_ENTRY.c_kin_code = KINSHIP_CODES.c_kincode ) " + _
            "LEFT JOIN BIOG_MAIN AS BIOG_MAIN_1 ON ZZ_SCRATCH_ENTRY.c_assoc_id = BIOG_MAIN_1.c_personid ) LEFT JOIN SOCIAL_INSTITUTION_NAME_CODES " + _
            "ON ZZ_SCRATCH_ENTRY.c_inst_name_code = SOCIAL_INSTITUTION_NAME_CODES.c_inst_name_code ) LEFT JOIN PARENTAL_STATUS_CODES " + _
            "ON ZZ_SCRATCH_ENTRY.c_parental_status_code = PARENTAL_STATUS_CODES.c_parental_status_code ) LEFT JOIN ADDR_CODES " + _
            "ON ZZ_SCRATCH_ENTRY.c_entry_addr_id = ADDR_CODES.c_addr_id ) LEFT JOIN TEXT_CODES ON ZZ_SCRATCH_ENTRY.c_source = TEXT_CODES.c_textid " + _
        "SET ZZ_SCRATCH_ENTRY.c_kin_desc = [KINSHIP_CODES].[c_kinrel], ZZ_SCRATCH_ENTRY.c_kin_name = [BIOG_MAIN].[c_name], " + _
            "ZZ_SCRATCH_ENTRY.c_kin_chn = [BIOG_MAIN].[c_name_chn], " + _
            "ZZ_SCRATCH_ENTRY.c_assoc_desc = [ASSOC_CODES].[c_assoc_desc], ZZ_SCRATCH_ENTRY.c_assoc_desc_chn = [ASSOC_CODES].[c_assoc_desc_chn], " + _
            "ZZ_SCRATCH_ENTRY.c_assoc_name = [BIOG_MAIN_1].[c_name], ZZ_SCRATCH_ENTRY.c_assoc_name_chn = [BIOG_MAIN_1].[c_name_chn], " + _
            "ZZ_SCRATCH_ENTRY.c_inst_name_hz = [SOCIAL_INSTITUTION_NAME_CODES].[c_inst_name_hz], " + _
            "ZZ_SCRATCH_ENTRY.c_inst_name_py = [SOCIAL_INSTITUTION_NAME_CODES].[c_inst_name_py], " + _
            "ZZ_SCRATCH_ENTRY.c_parental_status_desc = [PARENTAL_STATUS_CODES].[c_parental_status_desc], " + _
            "ZZ_SCRATCH_ENTRY.c_parental_status_desc_chn = [PARENTAL_STATUS_CODES].[c_parental_status_desc_chn], " + _
            "ZZ_SCRATCH_ENTRY.c_entry_addr_name = [ADDR_CODES].[c_name], ZZ_SCRATCH_ENTRY.c_entry_addr_chn = [ADDR_CODES].[c_name_chn], " + _
            "ZZ_SCRATCH_ENTRY.c_entry_xcoord = [ADDR_CODES].[x_coord], ZZ_SCRATCH_ENTRY.c_entry_ycoord = [ADDR_CODES].[y_coord], " + _
            "ZZ_SCRATCH_ENTRY.c_source_text = [TEXT_CODES].[c_title], ZZ_SCRATCH_ENTRY.c_source_text_chn = [TEXT_CODES].[c_title_chn]"
    '
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecCount
        '
        cmdSQL.CommandText = "Delete * from tmpXY"
        cmdSQL.Execute tRecDeleted
        '
        tQueryStr = "INSERT INTO tmpXY ( x_coord, y_coord, CountOfx_coord, CountOfy_coord ) " + _
            "SELECT ZZ_SCRATCH_ENTRY.x_coord, ZZ_SCRATCH_ENTRY.y_coord, Count(ZZ_SCRATCH_ENTRY.x_coord) " + _
            "AS CountOfx_coord, Count(ZZ_SCRATCH_ENTRY.y_coord) AS CountOfy_coord " + _
            "FROM ZZ_SCRATCH_ENTRY " + _
            "GROUP BY ZZ_SCRATCH_ENTRY.x_coord, ZZ_SCRATCH_ENTRY.y_coord"
        '
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecCount
        '
        tQueryStr = "UPDATE tmpXY INNER JOIN ZZ_SCRATCH_ENTRY ON (tmpXY.y_coord = " + _
            "ZZ_SCRATCH_ENTRY.y_coord) AND (tmpXY.x_coord = ZZ_SCRATCH_ENTRY.x_coord) " + _
            "SET ZZ_SCRATCH_ENTRY.xy_count = [tmpXY].[CountOfx_coord]"
        
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecCount
            
        Me.ChkGisEntry.Enabled = True
        If Me.ChkGisEntry.Value = False Then
            gChkGisCount = gChkGisCount + 1
            Me.ChkGisEntry.Value = True
        End If
        CmdGIS.Enabled = True
    Else
        If ChkGisEntry.Value Then
            Me.ChkGisEntry.Value = False
            gChkGisCount = gChkGisCount - 1
        End If
        Me.ChkGisEntry.Enabled = False
        
        If gChkGisCount = 0 Then
            CmdGIS.Enabled = False
        End If
    End If
    
    Set Me.ENTRY.Form.Recordset = CurrentDb.OpenRecordset("ZZ_SCRATCH_ENTRY", dbOpenDynaset)
End Sub

Private Sub queryText()
    Dim cmdSQL As ADODB.Command, tStrInsert As String, tStrSelect As String, tStrFrom As String, tRst As DAO.Recordset
    
    Set cmdSQL = New ADODB.Command
    cmdSQL.ActiveConnection = CurrentProject.Connection
    cmdSQL.CommandType = adCmdText
    
    '  first clear ZZ_SCRATCH_BIOG_TEXT_DATA
    
    Set Me.TEXT.Form.Recordset = CurrentDb.OpenRecordset("Z_SCRATCH_DUMMY_TR", dbOpenDynaset)
    
    cmdSQL.CommandText = "DELETE * FROM ZZ_SCRATCH_BIOG_TEXT_DATA"
    cmdSQL.Execute gTextRecCount
    
    cmdSQL.CommandText = "DELETE * FROM ZZ_SCRATCH_P_TEXT"
    cmdSQL.Execute gTextRecCount

    tStrInsert = "INSERT INTO ZZ_SCRATCH_BIOG_TEXT_DATA ( c_personid, c_name, c_name_chn, c_sex, c_index_year, c_index_year_type_code, c_index_year_type_desc, " + _
                "c_index_year_type_hz, c_dy, c_dynasty, c_dynasty_chn, c_addr_id, c_addr_name, c_addr_chn, x_coord, y_coord, c_textid, " + _
                "c_role_id, c_role_desc, c_role_desc_chn, c_source, c_pages, c_notes, c_title_chn, c_title, c_text_cat_code, c_text_cat_type_id ) "
    tStrSelect = "SELECT BIOG_TEXT_DATA.c_personid, ZZ_SCRATCH_IMPORT_PEOPLE.c_name, ZZ_SCRATCH_IMPORT_PEOPLE.c_name_chn, ZZ_SCRATCH_IMPORT_PEOPLE.c_sex, " + _
                "ZZ_SCRATCH_IMPORT_PEOPLE.c_index_year, ZZ_SCRATCH_IMPORT_PEOPLE.c_index_year_type_code, ZZ_SCRATCH_IMPORT_PEOPLE.c_index_year_type_desc, " + _
                "ZZ_SCRATCH_IMPORT_PEOPLE.c_index_year_type_chn, ZZ_SCRATCH_IMPORT_PEOPLE.c_dy, ZZ_SCRATCH_IMPORT_PEOPLE.c_dynasty, " + _
                "ZZ_SCRATCH_IMPORT_PEOPLE.c_dynasty_chn, ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_id, ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_name, " + _
                "ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_chn, ZZ_SCRATCH_IMPORT_PEOPLE.x_coord, ZZ_SCRATCH_IMPORT_PEOPLE.y_coord, BIOG_TEXT_DATA.c_textid, " + _
                "BIOG_TEXT_DATA.c_role_id, TEXT_ROLE_CODES.c_role_desc, TEXT_ROLE_CODES.c_role_desc_chn, BIOG_TEXT_DATA.c_source, BIOG_TEXT_DATA.c_pages, " + _
                "BIOG_TEXT_DATA.c_notes, TEXT_CODES.c_title_chn, TEXT_CODES.c_title, TEXT_CODES.c_bibl_cat_code, c_text_type_id  "
    tStrFrom = "FROM ( ( ZZ_SCRATCH_IMPORT_PEOPLE INNER JOIN BIOG_TEXT_DATA ON ZZ_SCRATCH_IMPORT_PEOPLE.c_person_id = BIOG_TEXT_DATA.c_personid ) " + _
                "INNER JOIN TEXT_ROLE_CODES ON BIOG_TEXT_DATA.c_role_id = TEXT_ROLE_CODES.c_role_id ) " + _
                "INNER JOIN TEXT_CODES ON BIOG_TEXT_DATA.c_textid = TEXT_CODES.c_textid"

    cmdSQL.CommandText = tStrInsert + tStrSelect + tStrFrom
    cmdSQL.Execute gTextRecCount
    '
    If gTextRecCount > 0 Then
        '
        ' get the outer join fields
        '
        tQueryStr = "UPDATE ( ( ZZ_SCRATCH_BIOG_TEXT_DATA LEFT JOIN TEXT_CODES ON ZZ_SCRATCH_BIOG_TEXT_DATA.c_source = TEXT_CODES.c_textid ) LEFT JOIN TEXT_BIBLCAT_CODES " + _
                "ON ZZ_SCRATCH_BIOG_TEXT_DATA.c_text_cat_code = TEXT_BIBLCAT_CODES.c_text_cat_code ) LEFT JOIN TEXT_BIBLCAT_TYPES " + _
                "ON ZZ_SCRATCH_BIOG_TEXT_DATA.c_text_cat_type_id = TEXT_BIBLCAT_TYPES.c_text_cat_type_id " + _
            "SET ZZ_SCRATCH_BIOG_TEXT_DATA.c_text_cat_desc = [TEXT_BIBLCAT_CODES].[c_text_cat_desc], " + _
                "ZZ_SCRATCH_BIOG_TEXT_DATA.c_text_cat_desc_chn = [TEXT_BIBLCAT_CODES].[c_text_cat_desc_chn], " + _
                "ZZ_SCRATCH_BIOG_TEXT_DATA.c_text_cat_pinyin = [TEXT_BIBLCAT_CODES].[c_text_cat_pinyin], " + _
                "ZZ_SCRATCH_BIOG_TEXT_DATA.c_text_cat_type_desc = [TEXT_BIBLCAT_TYPES].[c_text_cat_type_desc], " + _
                "ZZ_SCRATCH_BIOG_TEXT_DATA.c_text_cat_type_desc_chn = [TEXT_BIBLCAT_TYPES].[c_text_cat_type_desc_chn], " + _
                "ZZ_SCRATCH_BIOG_TEXT_DATA.c_source_chn = [TEXT_CODES].[c_title_chn], ZZ_SCRATCH_BIOG_TEXT_DATA.c_source_title = [TEXT_CODES].[c_title]"

        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecCount
            '
        '  the final step is to calculate the xy_count
        '
        tQueryStr = "INSERT INTO ZZ_SCRATCH_P_TEXT ( c_person_id, c_name_chn, c_name, c_sex, c_dy, c_dynasty, c_dynasty_chn, c_index_year, " + _
                "c_index_year_type_code, c_index_year_type_desc, c_index_year_type_hz, c_addr_id, c_addr_name, c_addr_chn, x_coord, y_coord ) " + _
            "SELECT DISTINCT ZZ_SCRATCH_BIOG_TEXT_DATA.c_personid, ZZ_SCRATCH_BIOG_TEXT_DATA.c_name_chn, ZZ_SCRATCH_BIOG_TEXT_DATA.c_name, " + _
                "ZZ_SCRATCH_BIOG_TEXT_DATA.c_sex, ZZ_SCRATCH_BIOG_TEXT_DATA.c_dy, ZZ_SCRATCH_BIOG_TEXT_DATA.c_dynasty, " + _
                "ZZ_SCRATCH_BIOG_TEXT_DATA.c_dynasty_chn, ZZ_SCRATCH_BIOG_TEXT_DATA.c_index_year, ZZ_SCRATCH_BIOG_TEXT_DATA.c_index_year_type_code, " + _
                "ZZ_SCRATCH_BIOG_TEXT_DATA.c_index_year_type_desc, ZZ_SCRATCH_BIOG_TEXT_DATA.c_index_year_type_hz, " + _
                "ZZ_SCRATCH_BIOG_TEXT_DATA.c_addr_id, ZZ_SCRATCH_BIOG_TEXT_DATA.c_addr_name, ZZ_SCRATCH_BIOG_TEXT_DATA.c_addr_chn, " + _
                "ZZ_SCRATCH_BIOG_TEXT_DATA.x_coord, ZZ_SCRATCH_BIOG_TEXT_DATA.y_coord " + _
            "FROM ZZ_SCRATCH_BIOG_TEXT_DATA"

        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecCount
        
        '
        '  the final step is to calculate the xy_count
        '
        cmdSQL.CommandText = "Delete * from tmpXY"
        cmdSQL.Execute tRecDeleted
        '
        tQueryStr = "INSERT INTO tmpXY ( x_coord, y_coord, CountOfx_coord, CountOfy_coord ) " + _
            "SELECT ZZ_SCRATCH_P_TEXT.x_coord, ZZ_SCRATCH_P_TEXT.y_coord, Count(ZZ_SCRATCH_P_TEXT.x_coord) " + _
            "AS CountOfx_coord, Count(ZZ_SCRATCH_P_TEXT.y_coord) AS CountOfy_coord " + _
            "FROM ZZ_SCRATCH_P_TEXT " + _
            "GROUP BY ZZ_SCRATCH_P_TEXT.x_coord, ZZ_SCRATCH_P_TEXT.y_coord"
        '
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecCount
        '
        tQueryStr = "UPDATE tmpXY INNER JOIN ZZ_SCRATCH_P_TEXT ON (tmpXY.y_coord = " + _
            "ZZ_SCRATCH_P_TEXT.y_coord) AND (tmpXY.x_coord = ZZ_SCRATCH_P_TEXT.x_coord) " + _
            "SET ZZ_SCRATCH_P_TEXT.xy_count = [tmpXY].[CountOfx_coord]"
        
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecCount
            
        Me.ChkGisText.Enabled = True
        If Me.ChkGisText.Value = False Then
            gChkGisCount = gChkGisCount + 1
            Me.ChkGisText.Value = True
        End If
        CmdGIS.Enabled = True
    Else
        If ChkGisText.Value Then
            Me.ChkGisText.Value = False
            gChkGisCount = gChkGisCount - 1
        End If
        Me.ChkGisText.Enabled = False
        
        If gChkGisCount = 0 Then
            CmdGIS.Enabled = False
        End If
    End If
    
    Set Me.TEXT.Form.Recordset = CurrentDb.OpenRecordset("ZZ_SCRATCH_BIOG_TEXT_DATA", dbOpenDynaset)
End Sub

Private Sub queryAddr()
    Dim cmdSQL As ADODB.Command, tStrInsert As String, tStrSelect As String, tStrFrom As String, tRst As DAO.Recordset
    
    Set cmdSQL = New ADODB.Command
    cmdSQL.ActiveConnection = CurrentProject.Connection
    cmdSQL.CommandType = adCmdText
    
    '  first clear ZZ_SCRATCH_STATUS
    
    Set Me.PLACE.Form.Recordset = CurrentDb.OpenRecordset("Z_SCRATCH_DUMMY_BA", dbOpenDynaset)
    
    cmdSQL.CommandText = "DELETE * FROM ZZ_SCRATCH_BIOG_ADDR_DATA"
    cmdSQL.Execute gPlaceRecCount

    cmdSQL.CommandText = "DELETE * FROM ZZ_SCRATCH_PEOPLE"
    cmdSQL.Execute gPlaceRecCount

    tStrInsert = "INSERT INTO ZZ_SCRATCH_BIOG_ADDR_DATA ( c_personid, c_name, c_name_chn, c_index_year, c_index_year_type_code, c_index_year_type_desc, " + _
                "c_index_year_type_hz, c_dy, c_dynasty, c_dynasty_chn, c_addr_id, c_addr_name, c_addr_chn, x_coord, y_coord, c_addr_type, " + _
                "c_addr_desc, c_addr_desc_chn, c_sequence, c_firstyear, c_lastyear, c_source, c_pages, c_notes, c_fy_nh_code, c_ly_nh_code, " + _
                "c_fy_nh_year, c_ly_nh_year, c_fy_range, c_ly_range, c_natal, c_fy_intercalary, c_ly_intercalary, c_fy_month, " + _
                "c_ly_month, c_fy_day, c_ly_day, c_fy_day_gz, c_ly_day_gz ) "
    tStrSelect = "SELECT ZZ_SCRATCH_IMPORT_PEOPLE.c_person_id, ZZ_SCRATCH_IMPORT_PEOPLE.c_name, ZZ_SCRATCH_IMPORT_PEOPLE.c_name_chn, " + _
                "ZZ_SCRATCH_IMPORT_PEOPLE.c_index_year, ZZ_SCRATCH_IMPORT_PEOPLE.c_index_year_type_code, ZZ_SCRATCH_IMPORT_PEOPLE.c_index_year_type_desc, " + _
                "ZZ_SCRATCH_IMPORT_PEOPLE.c_index_year_type_chn, ZZ_SCRATCH_IMPORT_PEOPLE.c_dy, ZZ_SCRATCH_IMPORT_PEOPLE.c_dynasty, " + _
                "ZZ_SCRATCH_IMPORT_PEOPLE.c_dynasty_chn, BIOG_ADDR_DATA.c_addr_id, ADDR_CODES.c_name, ADDR_CODES.c_name_chn, ADDR_CODES.x_coord, " + _
                "ADDR_CODES.y_coord, BIOG_ADDR_DATA.c_addr_type, BIOG_ADDR_CODES.c_addr_desc, BIOG_ADDR_CODES.c_addr_desc_chn, BIOG_ADDR_DATA.c_sequence, " + _
                "BIOG_ADDR_DATA.c_firstyear, BIOG_ADDR_DATA.c_lastyear, BIOG_ADDR_DATA.c_source, BIOG_ADDR_DATA.c_pages, BIOG_ADDR_DATA.c_notes, " + _
                "BIOG_ADDR_DATA.c_fy_nh_code, BIOG_ADDR_DATA.c_ly_nh_code, BIOG_ADDR_DATA.c_fy_nh_year, BIOG_ADDR_DATA.c_ly_nh_year, " + _
                "BIOG_ADDR_DATA.c_fy_range, BIOG_ADDR_DATA.c_ly_range, BIOG_ADDR_DATA.c_natal, BIOG_ADDR_DATA.c_fy_intercalary, " + _
                "BIOG_ADDR_DATA.c_ly_intercalary, BIOG_ADDR_DATA.c_fy_month, BIOG_ADDR_DATA.c_ly_month, BIOG_ADDR_DATA.c_fy_day, BIOG_ADDR_DATA.c_ly_day, " + _
                "BIOG_ADDR_DATA.c_fy_day_gz, BIOG_ADDR_DATA.c_ly_day_gz "
    If Me.FrameQueryAddress.Value = 1 Then
        tStrFrom = "FROM ( ZZ_SCRATCH_IMPORT_PEOPLE INNER JOIN ( BIOG_ADDR_CODES INNER JOIN BIOG_ADDR_DATA " + _
                "ON BIOG_ADDR_CODES.c_addr_type = BIOG_ADDR_DATA.c_addr_type ) ON ZZ_SCRATCH_IMPORT_PEOPLE.c_person_id = BIOG_ADDR_DATA.c_personid ) " + _
                "INNER JOIN ADDR_CODES ON BIOG_ADDR_DATA.c_addr_id = ADDR_CODES.c_addr_id"
    Else
        tStrFrom = "FROM ( ZZ_SCRATCH_IMPORT_PEOPLE INNER JOIN ( BIOG_ADDR_CODES INNER JOIN BIOG_ADDR_DATA " + _
                "ON BIOG_ADDR_CODES.c_addr_type = BIOG_ADDR_DATA.c_addr_type ) " + _
                "ON ZZ_SCRATCH_IMPORT_PEOPLE.c_person_id = BIOG_ADDR_DATA.c_personid  AND ZZ_SCRATCH_IMPORT_PEOPLE.c_addr_id = BIOG_ADDR_DATA.c_addr_id ) " + _
                "INNER JOIN ADDR_CODES ON BIOG_ADDR_DATA.c_addr_id = ADDR_CODES.c_addr_id"
    End If
    cmdSQL.CommandText = tStrInsert + tStrSelect + tStrFrom
    cmdSQL.Execute gPlaceRecCount
    '
    If gPlaceRecCount > 0 Then
        '
        ' get the outer join field information
        '
        tQueryStr = "UPDATE ( ( ( ( ( ( ZZ_SCRATCH_BIOG_ADDR_DATA LEFT JOIN TEXT_CODES ON ZZ_SCRATCH_BIOG_ADDR_DATA.c_source = TEXT_CODES.c_textid ) " + _
                "LEFT JOIN NIAN_HAO ON ZZ_SCRATCH_BIOG_ADDR_DATA.c_fy_nh_code = NIAN_HAO.c_nianhao_id ) LEFT JOIN NIAN_HAO AS NIAN_HAO_1 " + _
                "ON ZZ_SCRATCH_BIOG_ADDR_DATA.c_ly_nh_code = NIAN_HAO_1.c_nianhao_id ) LEFT JOIN YEAR_RANGE_CODES " + _
                "ON ZZ_SCRATCH_BIOG_ADDR_DATA.c_fy_range = YEAR_RANGE_CODES.c_range_code ) LEFT JOIN YEAR_RANGE_CODES AS YEAR_RANGE_CODES_1 " + _
                "ON ZZ_SCRATCH_BIOG_ADDR_DATA.c_ly_range = YEAR_RANGE_CODES_1.c_range_code ) LEFT JOIN GANZHI_CODES " + _
                "ON ZZ_SCRATCH_BIOG_ADDR_DATA.c_fy_day_gz = GANZHI_CODES.c_ganzhi_code ) LEFT JOIN GANZHI_CODES AS GANZHI_CODES_1 " + _
                "ON ZZ_SCRATCH_BIOG_ADDR_DATA.c_ly_day_gz = GANZHI_CODES_1.c_ganzhi_code " + _
            "SET ZZ_SCRATCH_BIOG_ADDR_DATA.c_source_title = [TEXT_CODES].[c_title], ZZ_SCRATCH_BIOG_ADDR_DATA.c_source_chn = [TEXT_CODES].[c_title_chn], " + _
                "ZZ_SCRATCH_BIOG_ADDR_DATA.c_fy_nh_chn = [NIAN_HAO].[c_nianhao_chn], ZZ_SCRATCH_BIOG_ADDR_DATA.c_fy_nh_py = [NIAN_HAO].[c_nianhao_pin], " + _
                "ZZ_SCRATCH_BIOG_ADDR_DATA.c_ly_nh_chn = [NIAN_HAO_1].[c_nianhao_chn], ZZ_SCRATCH_BIOG_ADDR_DATA.c_ly_nh_py = [NIAN_HAO_1].[c_nianhao_pin], " + _
                "ZZ_SCRATCH_BIOG_ADDR_DATA.c_fy_range_desc = [YEAR_RANGE_CODES].[c_range], " + _
                "ZZ_SCRATCH_BIOG_ADDR_DATA.c_fy_range_chn = [YEAR_RANGE_CODES].[c_range_chn], " + _
                "ZZ_SCRATCH_BIOG_ADDR_DATA.c_ly_range_desc = [YEAR_RANGE_CODES_1].[c_range], " + _
                "ZZ_SCRATCH_BIOG_ADDR_DATA.c_ly_range_chn = [YEAR_RANGE_CODES_1].[c_range_chn], " + _
                "ZZ_SCRATCH_BIOG_ADDR_DATA.c_fy_day_gz_chn = [GANZHI_CODES].[c_ganzhi_chn], " + _
                "ZZ_SCRATCH_BIOG_ADDR_DATA.c_fy_day_gz_py = [GANZHI_CODES].[c_ganzhi_py], " + _
                "ZZ_SCRATCH_BIOG_ADDR_DATA.c_ly_day_gz_chn = [GANZHI_CODES_1].[c_ganzhi_chn], " + _
                "ZZ_SCRATCH_BIOG_ADDR_DATA.c_ly_day_gz_py = [GANZHI_CODES_1].[c_ganzhi_py]"
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecDeleted
            '
        '  the final step is to calculate the xy_count
        '
        tQueryStr = "INSERT INTO ZZ_SCRATCH_PEOPLE ( c_person_id, c_name, c_name_chn, c_index_year, c_index_year_type_code, c_index_year_type_desc, " + _
                "c_index_year_type_hz, c_addr_id, c_addr_name, c_addr_chn, x_coord, y_coord, c_addr_type, c_addr_desc, c_addr_desc_chn, c_dy, " + _
                "c_dynasty, c_dynasty_chn ) " + _
            "SELECT DISTINCT ZZ_SCRATCH_BIOG_ADDR_DATA.c_personid, ZZ_SCRATCH_BIOG_ADDR_DATA.c_name, ZZ_SCRATCH_BIOG_ADDR_DATA.c_name_chn, " + _
                "ZZ_SCRATCH_BIOG_ADDR_DATA.c_index_year, ZZ_SCRATCH_BIOG_ADDR_DATA.c_index_year_type_code, ZZ_SCRATCH_BIOG_ADDR_DATA.c_index_year_type_desc, " + _
                "ZZ_SCRATCH_BIOG_ADDR_DATA.c_index_year_type_hz, ZZ_SCRATCH_BIOG_ADDR_DATA.c_addr_id, ZZ_SCRATCH_BIOG_ADDR_DATA.c_addr_name, " + _
                "ZZ_SCRATCH_BIOG_ADDR_DATA.c_addr_chn, ZZ_SCRATCH_BIOG_ADDR_DATA.x_coord, ZZ_SCRATCH_BIOG_ADDR_DATA.y_coord, " + _
                "ZZ_SCRATCH_BIOG_ADDR_DATA.c_addr_type, ZZ_SCRATCH_BIOG_ADDR_DATA.c_addr_desc, ZZ_SCRATCH_BIOG_ADDR_DATA.c_addr_desc_chn, " + _
                "ZZ_SCRATCH_BIOG_ADDR_DATA.c_dy, ZZ_SCRATCH_BIOG_ADDR_DATA.c_dynasty, ZZ_SCRATCH_BIOG_ADDR_DATA.c_dynasty_chn " + _
            "FROM ZZ_SCRATCH_BIOG_ADDR_DATA"
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecDeleted
        '
        cmdSQL.CommandText = "Delete * from tmpXY"
        cmdSQL.Execute tRecDeleted
        '
        tQueryStr = "INSERT INTO tmpXY ( x_coord, y_coord, CountOfx_coord, CountOfy_coord ) " + _
            "SELECT ZZ_SCRATCH_PEOPLE.x_coord, ZZ_SCRATCH_PEOPLE.y_coord, Count(ZZ_SCRATCH_PEOPLE.x_coord) " + _
            "AS CountOfx_coord, Count(ZZ_SCRATCH_PEOPLE.y_coord) AS CountOfy_coord " + _
            "FROM ZZ_SCRATCH_PEOPLE " + _
            "GROUP BY ZZ_SCRATCH_PEOPLE.x_coord, ZZ_SCRATCH_PEOPLE.y_coord"
        '
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecCount
        '
        tQueryStr = "UPDATE tmpXY INNER JOIN ZZ_SCRATCH_PEOPLE ON (tmpXY.y_coord = " + _
            "ZZ_SCRATCH_PEOPLE.y_coord) AND (tmpXY.x_coord = ZZ_SCRATCH_PEOPLE.x_coord) " + _
            "SET ZZ_SCRATCH_PEOPLE.xy_count = [tmpXY].[CountOfx_coord]"
        
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecCount
            
        Me.ChkGisAddr.Enabled = True
        If Me.ChkGisAddr.Value = False Then
            gChkGisCount = gChkGisCount + 1
            Me.ChkGisAddr.Value = True
        End If
        CmdGIS.Enabled = True
    Else
        If ChkGisAddr.Value Then
            Me.ChkGisAddr.Value = False
            gChkGisCount = gChkGisCount - 1
        End If
        Me.ChkGisAddr.Enabled = False
        
        If gChkGisCount = 0 Then
            CmdGIS.Enabled = False
        End If
    End If
    
    Set Me.PLACE.Form.Recordset = CurrentDb.OpenRecordset("ZZ_SCRATCH_BIOG_ADDR_DATA", dbOpenDynaset)
End Sub

Private Sub WriteGIS_OfficeOffice()
On Error GoTo Err_WriteGIS_OfficeOffice
    '
    '  If it is a KML file, call the routine and exit
    '
    If ChkKML.Value Then
        Call WriteKML_Office
        Exit Sub
    End If
    '
    '  This program will dump the results to a .gis file
    '
    If gOfficeRecCount = 0 Then
        MsgBox "There are no records to save."
        GoTo Exit_WriteGIS_OfficeOffice
    End If
    '
    Dim tStream As ADODB.Stream, tPinyin As Boolean
    
    Set tStream = New ADODB.Stream
    
    tPinyin = False
    If Me.GISFrame.Value = 1 Then
        tStream.Charset = "utf-8"
        tCodeStr = "UTF8"
    ElseIf GISFrame.Value = 2 Then
        tStream.Charset = "gb18030"
        tCodeStr = "GB18030"
    Else
        tStream.Charset = "ascii"
        tCodeStr = "ASCII"
        tPinyin = True
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
    Dim tStr As String, tTab As String, ti As Integer
    Dim tFileSystem, tGDF
        
    Set dlgSaveAs = Application.FileDialog(msoFileDialogSaveAs)
    
    dlgSaveAs.InitialFileName = "office_office_gis_" + tCodeStr + ".txt"
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
            GoTo Exit_WriteGIS_OfficeOffice
        Else
            '  make sure the file name has a txt extension
            If Len(tFileName) < 5 Then
                tFileName = tFileName + ".txt"
            ElseIf Not (LCase(Right(tFileName, 4)) = ".txt") Then
                tFileName = tFileName + ".tab"
            End If
        End If
        '
        '  write the file
        '
        'SELECT ZZ_SCRATCH_OFFICE.c_name AS Name, ZZ_SCRATCH_OFFICE.c_name_chn AS NameChn,
        'ZZ_SCRATCH_OFFICE.c_index_year AS IndexYear, ZZ_SCRATCH_OFFICE.c_sex AS Sex,
        'ZZ_SCRATCH_OFFICE.c_addr_name AS AddrName, ZZ_SCRATCH_OFFICE.c_addr_chn AS AddrChn,
        'Str(ZZ_SCRATCH_OFFICE.x_coord) AS PersonX, Str(ZZ_SCRATCH_OFFICE.y_coord) AS PersonY,
        'ZZ_SCRATCH_OFFICE.c_office_trans AS Office, ZZ_SCRATCH_OFFICE.c_office_chn AS OfficeChn,
        'ZZ_SCRATCH_OFFICE.c_firstyear AS FirstYear, ZZ_SCRATCH_OFFICE.c_lastyear AS LastYear,
        'ZZ_SCRATCH_OFFICE.c_dy_desc AS Dynasty,
        'ZZ_SCRATCH_OFFICE.c_office_addr_name AS OfficeAddr,
        'ZZ_SCRATCH_OFFICE.c_office_addr_chn AS OfficeAddrChn,
        'Str(ZZ_SCRATCH_OFFICE.office_x_coord) AS X, Str(ZZ_SCRATCH_OFFICE.office_y_coord) AS Y,
        'ZZ_SCRATCH_OFFICE.office_xy_count AS XY_count
        '
        ' process the table
        '
        Set tRstNode = Me.OFFICE.Form.Recordset
        tTab = Chr(9) ' the tab character
        '
        With tRstNode
            '
            ' write the header
            '
            If tPinyin Then
                tStr = "Office" + tTab + "FirstYear" + tTab + "LastYear" + tTab + _
                    "Dynasty" + tTab + "OfficeAddr" + tTab + _
                    "X" + tTab + "Y" + tTab + "xy_count"
            Else
                tStr = "Office" + tTab + "OfficeChn" + tTab + "FirstYear" + tTab + "LastYear" + tTab + _
                    "Dynasty" + tTab + "OfficeAddr" + tTab + "OfficeAddrChn" + tTab + _
                    "X" + tTab + "Y" + tTab + "xy_count"
            End If
            
            tStream.WriteText tStr, adWriteLine
            
            .MoveFirst
            ' MsgBox "writing file"
            Do While Not .EOF
                ' must guard against NULLs
                '
                If IsNull(!c_office_trans) Then
                    tStr = "No Translation" + tTab
                Else
                    tStr = !c_office_trans + tTab
                End If
                
                If Not tPinyin Then
                    If IsNull(!c_office_chn) Then
                        tStr = tStr + "[?]" + tTab
                    Else
                        tStr = tStr + !c_office_chn + tTab
                    End If
                End If
                
                If IsNull(!c_firstyear) Then
                    tStr = tStr + "[?]" + tTab
                Else
                    tStr = tStr + Str(!c_firstyear) + tTab
                End If
                
                If IsNull(!c_lastyear) Then
                    tStr = tStr + "[?]" + tTab
                Else
                    tStr = tStr + Str(!c_lastyear) + tTab
                End If
                
                If IsNull(!c_dynasty) Then
                    tStr = tStr + "[?]" + tTab
                Else
                    tStr = tStr + !c_dynasty + tTab
                End If
                
                If IsNull(!c_office_addr_name) Then
                    tStr = tStr + "[?]" + tTab
                ElseIf Trim(!c_office_addr_name) = "" Then
                    tStr = tStr + "[?]" + tTab
                Else
                    tStr = tStr + !c_office_addr_name + tTab
                End If
                
                If Not tPinyin Then
                    If IsNull(!c_office_addr_chn) Then
                        tStr = tStr + "[?]" + tTab
                    ElseIf Trim(!c_office_addr_chn) = "" Then
                        tStr = tStr + "[?]" + tTab
                    Else
                        tStr = tStr + !c_office_addr_chn + tTab
                    End If
                End If
                    
                If IsNull(!office_x_coord) Then
                    tStr = tStr + "0" + tTab
                Else
                    tStr = tStr + Str(!office_x_coord) + tTab
                End If
                    
                If IsNull(!office_y_coord) Then
                    tStr = tStr + "0" + tTab
                Else
                    tStr = tStr + Str(!office_y_coord) + tTab
                End If
                    
                If IsNull(!office_xy_count) Then
                    tStr = tStr + "0"
                Else
                    tStr = tStr + Str(!office_xy_count)
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
    
Exit_WriteGIS_OfficeOffice:
    Exit Sub

Err_WriteGIS_OfficeOffice:
    MsgBox Err.Description
    Resume Exit_WriteGIS_OfficeOffice
    
End Sub
Private Sub WriteGIS_Status()
On Error GoTo Err_WriteGIS_Status
    '
    '  If it is a KML file, call the routine and exit
    '
    If ChkKML.Value Then
        Call WriteKML_Status
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
    If gStatusRecCount = 0 Then
        MsgBox "There are no records to save."
        GoTo Exit_WriteGIS_Status
    End If
    '
    Dim tStream As ADODB.Stream
    Set tStream = New ADODB.Stream
    
    tPinyin = False
    If GISFrame.Value = 1 Then
        tStream.Charset = "utf-8"
        tCodeStr = "UTF8"
    ElseIf GISFrame.Value = 3 Then
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

    dlgSaveAs.InitialFileName = "status_gis_" + tCodeStr + ".tab"
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
            GoTo Exit_WriteGIS_Status
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
        Set tRstNode = CurrentDb.OpenRecordset("ZZ_SCRATCH_P_STATUS", dbOpenDynaset)
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
    
Exit_WriteGIS_Status:
    Exit Sub

Err_WriteGIS_Status:
    MsgBox Err.Description
    Resume Exit_WriteGIS_Status
    
End Sub
Private Sub WriteKML_Status()
On Error GoTo Err_WriteKML_Status
    '
    '  This program will dump the results to a .gis file
    '
    If gStatusRecCount = 0 Then
        MsgBox "There are no records to save."
        GoTo Exit_WriteKML_Status
    End If
    '
    Dim tStream As ADODB.Stream
    Set tStream = New ADODB.Stream
    
    tPinyin = False
    If GISFrame.Value = 1 Then
        tStream.Charset = "utf-8"
        tCodeStr = "UTF8"
    ElseIf GISFrame.Value = 3 Then
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

    dlgSaveAs.InitialFileName = "status_gis_" + tCodeStr + ".kml"
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
            GoTo Exit_WriteKML_Status
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
        Set tRstNode = CurrentDb.OpenRecordset("ZZ_SCRATCH_P_STATUS", dbOpenDynaset)
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
        If Not tPinyin Then
            tStream.WriteText tC + tC + tC + tC + "Name Chn: $[StatusGIS/NameHZ] <br/>", adWriteLine
        End If
        tStream.WriteText tC + tC + tC + tC + "Index Year: $[StatusGIS/IndexYear] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "Sex: $[StatusGIS/Sex] <br/>", adWriteLine
        If tPinyin Then
            tStream.WriteText tC + tC + tC + tC + "Address: $[StatusGIS/AddrName] <br/>", adWriteLine
        Else
            tStream.WriteText tC + tC + tC + tC + "Address: $[StatusGIS/AddrName] $[StatusGIS/AddrHZ] <br/>", adWriteLine
        End If
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
        If Not tPinyin Then
            tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "NameHZ" + tDQ + ">", adWriteLine
            tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Name Chn]]></displayName>", adWriteLine
            tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        End If
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "Sex" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Sex]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "AddrName" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Address]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        If Not tPinyin Then
            tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "AddrHZ" + tDQ + ">", adWriteLine
            tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Address Chn]]></displayName>", adWriteLine
            tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        End If
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
                If Not tPinyin Then
                    If IsNull(!c_name_chn) Then
                        tStr = tStr + "[Bad Data]"
                    Else
                        If Trim(!c_name_chn) = "" Then
                            tStr = "[?]"
                        Else
                            tStr = !c_name_chn
                        End If
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
                If Not tPinyin Then
                    If IsNull(!c_addr_chn) Then
                        tStr = "[?]"
                    ElseIf Trim(!c_addr_chn) = "" Then
                        tStr = "[?]"
                    Else
                        tStr = !c_addr_chn
                    End If
                    tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "AddrHZ" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                End If
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

Exit_WriteKML_Status:
    Exit Sub

Err_WriteKML_Status:
    MsgBox Err.Description
    Resume Exit_WriteKML_Status

End Sub
Private Sub WriteGIS_OfficePeople()
On Error GoTo Err_WriteGIS_OfficePeople
    Dim tPinyin As Boolean
    '
    '  If it is a KML file, call the routine and exit
    '
    If ChkKML.Value Then
        Call WriteKML_OfficePeople
        Exit Sub
    End If
    '
    '  This program will dump the results to a .gis file
    '
    If gOfficeRecCount = 0 Then
        MsgBox "There are no records to save."
        GoTo Exit_WriteGIS_OfficePeople
    End If
    '
    tPinyin = False
    '
    Dim tStream As ADODB.Stream
    Set tStream = New ADODB.Stream
        
    If GISFrame.Value = 1 Then
        tStream.Charset = "utf-8"
        tCodeStr = "UTF8"
    ElseIf GISFrame.Value = 2 Then
        tStream.Charset = "gb18030"
        tCodeStr = "GB18030"
    Else
        tStream.Charset = "ascii"
        tCodeStr = "ASCII"
        tPinyin = True
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
    Dim tStr As String, tTab As String, ti As Integer
    Dim tFileSystem, tGDF
        
    Set dlgSaveAs = Application.FileDialog(msoFileDialogSaveAs)
    
    dlgSaveAs.InitialFileName = "office_people_gis_" + tCodeStr + ".tab"
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
            GoTo Exit_WriteGIS_OfficePeople
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
        Set tRstNode = CurrentDb.OpenRecordset("ZZ_SCRATCH_P_OFFICE", dbOpenDynaset)
        tTab = Chr(9) ' the tab character
        '
        With tRstNode
            '
            ' write the header
            '
            If tPinyin Then
                tStr = "Name" + tTab + "Sex" + tTab + "IndexYear" + tTab + _
                    "AddrID" + tTab + "AddrName" + tTab + _
                    "X" + tTab + "Y" + tTab + "xy_count"
            Else
                tStr = "Name" + tTab + "NameChn" + tTab + "Sex" + tTab + "IndexYear" + tTab + _
                    "AddrID" + tTab + "AddrName" + tTab + "AddrChn" + tTab + _
                    "X" + tTab + "Y" + tTab + "xy_count"
            End If
            tStream.WriteText tStr, adWriteLine
            
            .MoveFirst
            Do While Not .EOF
                ' must guard against NULLs
                '
                If Trim(!c_name) = "" Then
                    tStr = "[?]" + tTab
                Else
                    tStr = !c_name + tTab
                End If
                
                If Not tPinyin Then
                    If Trim(!c_name_chn) = "" Then
                        tStr = tStr + "[?]" + tTab
                    Else
                        tStr = tStr + !c_name_chn + tTab
                    End If
                End If
                
                'If IsNull(!c_sex) Then
                '    tStr = tStr + "?" + tTab
                'Else
                '    tStr = tStr + !c_sex + tTab
                'End If
                tStr = tStr + IIf(!c_female, "F", "M") + tTab
                
                If IsNull(!c_index_year) Then
                    tStr = tStr + "-2000" + tTab
                Else
                    tStr = tStr + Str(!c_index_year) + tTab
                End If
                    
                ' here guard against blanks as well
                    
                If IsNull(!c_addr_id) Then
                    tStr = tStr + "0" + tTab
                Else
                    tStr = tStr + Str(!c_addr_id) + tTab
                End If
                
                If IsNull(!c_addr_name) Then
                    tStr = tStr + "[?]" + tTab
                ElseIf Trim(!c_addr_name) = "" Then
                    tStr = tStr + "[?]" + tTab
                Else
                    tStr = tStr + !c_addr_name + tTab
                End If
                    
                If Not tPinyin Then
                    If IsNull(!c_addr_chn) Then
                        tStr = tStr + "[?]" + tTab
                    ElseIf Trim(!c_addr_chn) = "" Then
                        tStr = tStr + "[?]" + tTab
                    Else
                        tStr = tStr + !c_addr_chn + tTab
                    End If
                End If
                    
                If IsNull(!x_coord) Then
                    tStr = tStr + "0" + tTab
                Else
                    tStr = tStr + Str(!x_coord) + tTab
                End If
                    
                If IsNull(!y_coord) Then
                    tStr = tStr + "0" + tTab
                Else
                    tStr = tStr + Str(!y_coord) + tTab
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
    
Exit_WriteGIS_OfficePeople:
    Exit Sub

Err_WriteGIS_OfficePeople:
    MsgBox Err.Description
    Resume Exit_WriteGIS_OfficePeople
    
End Sub

Private Sub WriteKML_OfficePeople()

    Dim tStrKML As String, tPinyin As Boolean
    '
    '  This program will dump the results to a .kml file
    '
    If gOfficeRecCount = 0 Then
        MsgBox "There are no records to save."
        GoTo Exit_WriteKML_OfficePeople
    End If
    '
    Dim tStream As ADODB.Stream
    Set tStream = New ADODB.Stream
        
    tPinyin = False
    If GISFrame.Value = 1 Then
        tStream.Charset = "utf-8"
        tCodeStr = "UTF8"
    ElseIf GISFrame.Value = 2 Then
        tStream.Charset = "gb18030"
        tCodeStr = "GB18030"
    Else
        tStream.Charset = "ascii"
        tCodeStr = "ASCII"
        tPinyin = True
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
    Dim tStr As String, tTab As String, ti As Integer
    Dim tFileSystem, tGDF
        
    Set dlgSaveAs = Application.FileDialog(msoFileDialogSaveAs)
    
    dlgSaveAs.InitialFileName = "office_people_gis_" + tCodeStr + ".kml"
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
            GoTo Exit_WriteKML_OfficePeople
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
        'SELECT ZZ_SCRATCH_OFFICE.c_name AS Name, ZZ_SCRATCH_OFFICE.c_name_chn AS NameChn,
        'ZZ_SCRATCH_OFFICE.c_index_year AS IndexYear, ZZ_SCRATCH_OFFICE.c_sex AS Sex,
        'ZZ_SCRATCH_OFFICE.c_addr_name AS AddrName, ZZ_SCRATCH_OFFICE.c_addr_chn AS AddrChn,
        'Str(ZZ_SCRATCH_OFFICE.x_coord) AS PersonX, Str(ZZ_SCRATCH_OFFICE.y_coord) AS PersonY,
        'ZZ_SCRATCH_OFFICE.c_office_trans AS Office, ZZ_SCRATCH_OFFICE.c_office_chn AS OfficeChn,
        'ZZ_SCRATCH_OFFICE.c_firstyear AS FirstYear, ZZ_SCRATCH_OFFICE.c_lastyear AS LastYear,
        'ZZ_SCRATCH_OFFICE.c_dy_desc AS Dynasty,
        'ZZ_SCRATCH_OFFICE.c_office_addr_name AS OfficeAddr,
        'ZZ_SCRATCH_OFFICE.c_office_addr_chn AS OfficeAddrChn,
        'Str(ZZ_SCRATCH_OFFICE.office_x_coord) AS X, Str(ZZ_SCRATCH_OFFICE.office_y_coord) AS Y,
        'ZZ_SCRATCH_OFFICE.office_xy_count AS XY_count
        
        '    tStr = "PostingID" (c_posting_id) + "Office" (c_name) + "OfficeChn" (c_name_chn) + _
                "FirstYear" (c_firstyear) + "LastYear" + (c_lastyear) _
                "Dynasty" (c_dy) + "OfficeAddr" (c_office_addr_name) + "OfficeAddrHZ" (c_office_addr_chn) + _
                "X" (office_x_coord) + "Y" (office_y_coord) + "xy_count" (office_xy_count)
        '
        ' process the table
        '
        Set tRstNode = CurrentDb.OpenRecordset("ZZ_SCRATCH_P_OFFICE", dbOpenDynaset)
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
        tStream.WriteText tC + "<Style id=" + tDQ + "office-balloon-template" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + "<BalloonStyle>", adWriteLine
        tStream.WriteText tC + tC + tC + "<text>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "<![CDATA[", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "ID: $[OfficePosting/PersonID] <br/>", adWriteLine
        If Not tPinyin Then
            tStream.WriteText tC + tC + tC + tC + "Name Chn: $[OfficePosting/NameHZ] <br/>", adWriteLine
        End If
        tStream.WriteText tC + tC + tC + tC + "Index Year: $[OfficePosting/IndexYear] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "Dynasty: $[OfficePosting/Dyn] <br/>", adWriteLine
        If tPinyin Then
            tStream.WriteText tC + tC + tC + tC + "Address: $[OfficePosting/AddrName] <br/>", adWriteLine
        Else
            tStream.WriteText tC + tC + tC + tC + "Address: $[OfficePosting/AddrName] $[OfficePosting/AddrNameHZ] <br/>", adWriteLine
        End If
        tStream.WriteText tC + tC + tC + tC + "XY Count: $[OfficePosting/XYCount] <br/><br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "]]>", adWriteLine
        tStream.WriteText tC + tC + tC + "</text>", adWriteLine
        tStream.WriteText tC + tC + "</BalloonStyle>", adWriteLine
        tStream.WriteText tC + "</Style>", adWriteLine
        tStream.WriteText tC + "<!-- Declare the type " + tDQ + "OfficePosting" + tDQ + " with 6 fields -->", adWriteLine
        tStream.WriteText tC + "<Schema name=" + tDQ + "OfficePosting" + tDQ + " id=" + tDQ + "OfficePersonID" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "uint" + tDQ + " name=" + tDQ + "PersonID" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[ID]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        If Not tPinyin Then
            tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "NameHZ" + tDQ + ">", adWriteLine
            tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Person Chn]]></displayName>", adWriteLine
            tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        End If
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "AddrName" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Address]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        If Not tPinyin Then
            tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "AddrNameHZ" + tDQ + ">", adWriteLine
            tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Address Chn]]></displayName>", adWriteLine
            tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        End If
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "IndexYear" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Index Year]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "int" + tDQ + " name=" + tDQ + "Dyn" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Dyn]]></displayName>", adWriteLine
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
                    tStr = "[Bad Data] " + Str(!c_personid)
                Else
                    tStr = !c_name + " " + Str(!c_personid)
                End If
                tStream.WriteText tC + tC + "<name>" + tStr + "</name>", adWriteLine
                
                tStream.WriteText tC + tC + "<styleUrl>#office-balloon-template</styleUrl>", adWriteLine
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
                tStream.WriteText tC + tC + tC + "<SchemaData schemaUrl=" + tDQ + "#OfficePersonID" + tDQ + ">", adWriteLine
                '
                '  person ID
                '
                tStr = Str(!c_personid)
                tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "PersonID" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                '
                '  Person Name Chn
                '
                If Not tPinyin Then
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
                End If
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
                '  Dynasty
                '
                If IsNull(!c_dy) Then
                    tStr = "0"
                Else
                    tStr = Str(!c_dy)
                End If
                tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "Dyn" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
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
                If Not tPinyin Then
                    If IsNull(!c_addr_chn) Then
                        tStr = "[?]"
                    ElseIf Trim(!c_addr_chn) = "" Then
                        tStr = "[?]"
                    Else
                        tStr = !c_addr_chn
                    End If
                    tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "AddrNameHZ" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                End If
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
    
Exit_WriteKML_OfficePeople:
    Exit Sub

Err_WriteKML_OfficePeople:
    MsgBox Err.Description
    Resume Exit_WriteKML_OfficePeople

End Sub

Private Sub WriteKML_Entry()
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

    Dim tStrKML As String, tPinyin As Boolean
    '
    '  This program will dump the results to a .gis file
    '
    If gEntryRecCount = 0 Then
        MsgBox "There are no records to save."
        GoTo Exit_WriteKML_Entry
    End If
    '
    Dim tStream As ADODB.Stream
    Set tStream = New ADODB.Stream
    
    tPinyin = False
    If GISFrame.Value = 1 Then
        tStream.Charset = "utf-8"
        tCodeStr = "UTF8"
    ElseIf GISFrame.Value = 2 Then
        tStream.Charset = "gb2312"
        tCodeStr = "GB2312"
    Else
        tStream.Charset = "ascii"
        tCodeStr = "ASCII"
        tPinyin = True
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
            GoTo Exit_WriteKML_Entry
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
        Set tRstNode = CurrentDb.OpenRecordset("ZZ_SCRATCH_ENTRY", dbOpenDynaset)
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
        If Not tPinyin Then
            tStream.WriteText tC + tC + tC + tC + "$[EntryPerson/PersonNameHZ] <br/>", adWriteLine
        End If
        tStream.WriteText tC + tC + tC + tC + "ID: $[EntryPerson/PersonID] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "Index Year: $[EntryPerson/IndexYear] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "Entry Year: $[EntryPerson/EntryYear] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "Entry Desc: $[EntryPerson/EntryDesc] <br/>", adWriteLine
        If Not tPinyin Then
            tStream.WriteText tC + tC + tC + tC + "Entry Chn: $[EntryPerson/EntryDescHZ] <br/>", adWriteLine
        End If
        tStream.WriteText tC + tC + tC + tC + "Entry Rank: $[EntryPerson/EntryRank] <br/>", adWriteLine
        If tPinyin Then
            tStream.WriteText tC + tC + tC + tC + "Address: $[EntryPerson/AddrName] <br/>", adWriteLine
        Else
            tStream.WriteText tC + tC + tC + tC + "Address: $[EntryPerson/AddrName] $[EntryPerson/AddrNameHZ] <br/>", adWriteLine
        End If
        tStream.WriteText tC + tC + tC + tC + "XY Count: $[EntryPerson/XYCount] <br/><br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "]]>", adWriteLine
        tStream.WriteText tC + tC + tC + "</text>", adWriteLine
        tStream.WriteText tC + tC + "</BalloonStyle>", adWriteLine
        tStream.WriteText tC + "</Style>", adWriteLine
        tStream.WriteText tC + "<!-- Declare the type " + tDQ + "EntryPerson" + tDQ + " with 10 fields -->", adWriteLine
        tStream.WriteText tC + "<Schema name=" + tDQ + "EntryPerson" + tDQ + " id=" + tDQ + "EntryPersonId" + tDQ + ">", adWriteLine
        If Not tPinyin Then
            tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "PersonNameHZ" + tDQ + ">", adWriteLine
            tStream.WriteText tC + tC + tC + "<displayName><![CDATA[<b>Person</b>]]></displayName>", adWriteLine
            tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        End If
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "AddrName" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[<b>Person</b>]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        If Not tPinyin Then
            tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "AddrNameHZ" + tDQ + ">", adWriteLine
            tStream.WriteText tC + tC + tC + "<displayName><![CDATA[<b>Person</b>]]></displayName>", adWriteLine
            tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        End If
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
        If Not tPinyin Then
            tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "EntryDescHZ" + tDQ + ">", adWriteLine
            tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Entry Chn]]></displayName>", adWriteLine
            tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        End If
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
                If Not tPinyin Then
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
                End If
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
                If Not tPinyin Then
                    If IsNull(!c_entry_chn) Then
                        tStr = "[Missing Data]"
                    Else
                        tStr = !c_entry_chn
                    End If
                    tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "EntryDescHZ" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                End If
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
                If Not tPinyin Then
                    If IsNull(!c_addr_chn) Then
                        tStr = "[?]"
                    ElseIf Trim(!c_addr_chn) = "" Then
                        tStr = "[?]"
                    Else
                        tStr = !c_addr_chn
                    End If
                    tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "AddrNameHZ" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                End If
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
    
Exit_WriteKML_Entry:
    Exit Sub

Err_WriteKML_Entry:
    MsgBox Err.Description
    Resume Exit_WriteKML_Entry
    
End Sub

Private Sub WriteGIS_Entry()
On Error GoTo Err_WriteGIS_Entry
    Dim tPinyin As Boolean
    '
    '  If it is a KML file, call the routine and exit
    '
    If ChkKML.Value Then
        Call WriteKML_Entry
        Exit Sub
    End If
    '
    '  This program will dump the results to a .gis file
    '
    If gEntryRecCount = 0 Then
        MsgBox "There are no records to save."
        GoTo Exit_WriteGIS_Entry
    End If
    '
    Dim tStream As ADODB.Stream
    Set tStream = New ADODB.Stream
    
    tPinyin = False
    If GISFrame.Value = 1 Then
        tStream.Charset = "utf-8"
        tCodeStr = "UTF8"
    ElseIf GISFrame.Value = 2 Then
        tStream.Charset = "gb2312"
        tCodeStr = "GB2312"
    Else
        tStream.Charset = "ascii"
        tCodeStr = "ASCII"
        tPinyin = True
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
    Dim tStr As String, tTab As String, ti As Integer
    Dim tFileSystem, tGDF
        
    Set dlgSaveAs = Application.FileDialog(msoFileDialogSaveAs)
    
    dlgSaveAs.InitialFileName = "entry_gis_" + tCodeStr + ".tab"
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
            GoTo Exit_WriteGIS_Entry
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
        'Name,NameChn,PersonID,IndexYear,EntryDesc,EntryChn,EntryYear,EntryRank,KinType,KinName,KinChn,
        'AddrName,AddrChn,X,Y,xy_count
        '
        ' process the table
        '
        Set tRstNode = CurrentDb.OpenRecordset("ZZ_SCRATCH_ENTRY", dbOpenDynaset)
        tTab = Chr(9) ' the tab character
        '
        With tRstNode
            '
            ' write the header
            '
            If tPinyin Then
                tStr = "Name" + tTab + "PersonID" + tTab + "IndexYear" + tTab + _
                    "EntryDesc" + tTab + "EntryYear" + tTab + "EntryRank" + tTab + _
                    "KinType" + tTab + "KinName" + tTab + _
                    "AddrName" + tTab + "X" + tTab + "Y" + tTab + "xy_count"
            Else
                tStr = "Name" + tTab + "NameChn" + tTab + "PersonID" + tTab + "IndexYear" + tTab + _
                    "EntryDesc" + tTab + "EntryChn" + tTab + "EntryYear" + tTab + "EntryRank" + tTab + _
                    "KinType" + tTab + "KinName" + tTab + "KinChn" + tTab + _
                    "AddrName" + tTab + "AddrChn" + tTab + "X" + tTab + "Y" + tTab + "xy_count"
            End If
            tStream.WriteText tStr, adWriteLine
            .MoveFirst
            Do While Not .EOF
                ' must guard against NULLs
                '
                If Trim(!c_name) = "" Then
                    tStr = "[?]" + tTab
                Else
                    tStr = !c_name + tTab
                End If
                
                If Not tPinyin Then
                    If Trim(!c_name_chn) = "" Then
                        tStr = tStr + "[?]" + tTab
                    Else
                        tStr = tStr + !c_name_chn + tTab
                    End If
                End If
                
                tStr = tStr + Trim(Str(!c_personid)) + tTab
                    
                If IsNull(!c_index_year) Then
                    tStr = tStr + "-2000" + tTab
                Else
                    tStr = tStr + Str(!c_index_year) + tTab
                End If
                    
                If IsNull(!c_entry_desc) Then
                    tStr = tStr + "[?]" + tTab
                ElseIf Trim(!c_entry_desc) = "" Then
                    tStr = tStr + "[?]" + tTab
                Else
                    tStr = tStr + !c_entry_desc + tTab
                End If
                    
                If Not tPinyin Then
                    If IsNull(!c_entry_chn) Then
                        tStr = tStr + "[?]" + tTab
                    ElseIf Trim(!c_entry_chn) = "" Then
                        tStr = tStr + "[?]" + tTab
                    Else
                        tStr = tStr + !c_entry_chn + tTab
                    End If
                End If
                    
                If IsNull(!c_year) Then
                    tStr = tStr + "[?]" + tTab
                Else
                    tStr = tStr + Trim(Str(!c_year)) + tTab
                End If
                    
                If IsNull(!c_exam_rank) Then
                    tStr = tStr + "[?]" + tTab
                ElseIf Trim(!c_exam_rank) = "" Then
                    tStr = tStr + "[?]" + tTab
                Else
                    tStr = tStr + !c_exam_rank + tTab
                End If
                    
                If IsNull(!c_kin_desc) Then
                    tStr = tStr + "[?]" + tTab
                ElseIf Trim(!c_kin_desc) = "" Then
                    tStr = tStr + "[?]" + tTab
                Else
                    tStr = tStr + !c_kin_desc + tTab
                End If
                    
                If IsNull(!c_kin_name) Then
                    tStr = tStr + "[?]" + tTab
                ElseIf Trim(!c_kin_name) = "" Then
                    tStr = tStr + "[?]" + tTab
                Else
                    tStr = tStr + !c_kin_name + tTab
                End If
                    
                If Not tPinyin Then
                    If IsNull(!c_kin_chn) Then
                        tStr = tStr + "[?]" + tTab
                    ElseIf Trim(!c_kin_chn) = "" Then
                        tStr = tStr + "[?]" + tTab
                    Else
                        tStr = tStr + !c_kin_chn + tTab
                    End If
                End If
                    
                ' here guard against blanks as well
                    
                If IsNull(!c_addr_name) Then
                    tStr = tStr + "[?]" + tTab
                ElseIf Trim(!c_addr_name) = "" Then
                    tStr = tStr + "[?]" + tTab
                Else
                    tStr = tStr + !c_addr_name + tTab
                End If
                    
                If Not tPinyin Then
                    If IsNull(!c_addr_chn) Then
                        tStr = tStr + "[?]" + tTab
                    ElseIf Trim(!c_addr_chn) = "" Then
                        tStr = tStr + "[?]" + tTab
                    Else
                        tStr = tStr + !c_addr_chn + tTab
                    End If
                End If
                    
                If IsNull(!x_coord) Then
                    tStr = tStr + "0" + tTab
                Else
                    tStr = tStr + Str(!x_coord) + tTab
                End If
                    
                If IsNull(!y_coord) Then
                    tStr = tStr + "0" + tTab
                Else
                    tStr = tStr + Str(!y_coord) + tTab
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
    
Exit_WriteGIS_Entry:
    Exit Sub

Err_WriteGIS_Entry:
    MsgBox Err.Description
    Resume Exit_WriteGIS_Entry
End Sub

Private Sub WriteKML_Text()
On Error GoTo Err_WriteKML_Text
    '
    '  This program will dump the results to a .gis file
    '
    If gTextRecCount = 0 Then
        MsgBox "There are no records to save."
        GoTo Exit_WriteKML_Text
    End If
    '
    Dim tStream As ADODB.Stream
    Set tStream = New ADODB.Stream
    
    tPinyin = False
    If GISFrame.Value = 1 Then
        tStream.Charset = "utf-8"
        tCodeStr = "UTF8"
    ElseIf GISFrame.Value = 2 Then
        tStream.Charset = "gb18030"
        tCodeStr = "GB18030"
    Else
        tStream.Charset = "ascii"
        tCodeStr = "ASCII"
        tPinyin = True
    End If
    tStream.Mode = adModeReadWrite
    tStream.Type = adTypeText
    tStream.Open
    '
    '  next get a file
    '
    
    Set dlgSaveAs = Application.FileDialog(msoFileDialogSaveAs)

    dlgSaveAs.InitialFileName = "text_gis_" + tCodeStr + ".kml"
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
            GoTo Exit_WriteKML_Text
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
        Set tRstNode = CurrentDb.OpenRecordset("ZZ_SCRATCH_P_TEXT", dbOpenDynaset)
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
        tStream.WriteText tC + "<Style id=" + tDQ + "TextCat-balloon-template" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + "<BalloonStyle>", adWriteLine
        tStream.WriteText tC + tC + tC + "<text>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "<![CDATA[", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "ID: $[TextCatGIS/PersonID] <br/>", adWriteLine
        If Not tPinyin Then
            tStream.WriteText tC + tC + tC + tC + "Name Chn: $[TextCatGIS/NameHZ] <br/>", adWriteLine
        End If
        tStream.WriteText tC + tC + tC + tC + "Index Year: $[TextCatGIS/IndexYear] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "Sex: $[TextCatGIS/Sex] <br/>", adWriteLine
        If tPinyin Then
            tStream.WriteText tC + tC + tC + tC + "Address: $[TextCatGIS/AddrName] <br/>", adWriteLine
        Else
            tStream.WriteText tC + tC + tC + tC + "Address: $[TextCatGIS/AddrName] $[TextCatGIS/AddrHZ] <br/>", adWriteLine
        End If
        tStream.WriteText tC + tC + tC + tC + "XY Count: $[TextCatGIS/XYCount] <br/><br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "]]>", adWriteLine
        tStream.WriteText tC + tC + tC + "</text>", adWriteLine
        tStream.WriteText tC + tC + "</BalloonStyle>", adWriteLine
        tStream.WriteText tC + "</Style>", adWriteLine
        tStream.WriteText tC + "<!-- Declare the type " + tDQ + "TextCatGIS" + tDQ + " with 6 fields -->", adWriteLine
        tStream.WriteText tC + "<Schema name=" + tDQ + "TextCatGIS" + tDQ + " id=" + tDQ + "TextCatGISId" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "uint" + tDQ + " name=" + tDQ + "PersonID" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Person ID]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        If Not tPinyin Then
            tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "NameHZ" + tDQ + ">", adWriteLine
            tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Name Chn]]></displayName>", adWriteLine
            tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        End If
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "Sex" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Sex]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "AddrName" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Address]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        If Not tPinyin Then
            tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "AddrHZ" + tDQ + ">", adWriteLine
            tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Address Chn]]></displayName>", adWriteLine
            tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        End If
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
                
                tStream.WriteText tC + tC + "<styleUrl>#TextCat-balloon-template</styleUrl>", adWriteLine
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
                tStream.WriteText tC + tC + tC + "<SchemaData schemaUrl=" + tDQ + "#TextCatGISId" + tDQ + ">", adWriteLine
                '
                '  person ID
                '
                tStr = Str(!c_person_id)
                tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "PersonID" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                '
                '  Person Name Chn
                '
                If Not tPinyin Then
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
                End If
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
                If Not tPinyin Then
                    If IsNull(!c_addr_chn) Then
                        tStr = "[?]"
                    ElseIf Trim(!c_addr_chn) = "" Then
                        tStr = "[?]"
                    Else
                        tStr = !c_addr_chn
                    End If
                    tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "AddrHZ" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                End If
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

Exit_WriteKML_Text:
    Exit Sub

Err_WriteKML_Text:
    MsgBox Err.Description
    Resume Exit_WriteKML_Text

End Sub

Private Sub WriteGIS_Text()
On Error GoTo Err_WriteGIS_Text
    '
    '  If it is a KML file, call the routine and exit
    '
    If ChkKML.Value Then
        Call WriteKML_Text
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
    If gTextRecCount = 0 Then
        MsgBox "There are no records to save."
        GoTo Exit_WriteGIS_Text
    End If
    '
    Dim tStream As ADODB.Stream
    Set tStream = New ADODB.Stream
    
    tPinyin = False
    If GISFrame.Value = 1 Then
        tStream.Charset = "utf-8"
        tCodeStr = "UTF8"
    ElseIf GISFrame.Value = 2 Then
        tStream.Charset = "gb18030"
        tCodeStr = "GB18030"
    Else
        tStream.Charset = "ascii"
        tCodeStr = "ASCII"
        tPinyin = True
    End If
    tStream.Mode = adModeReadWrite
    tStream.Type = adTypeText
    tStream.Open
    '
    '  next get a file
    '
    
    Set dlgSaveAs = Application.FileDialog(msoFileDialogSaveAs)

    dlgSaveAs.InitialFileName = "text_gis_" + tCodeStr + ".tab"
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
            GoTo Exit_WriteGIS_Text
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
        Set tRstNode = CurrentDb.OpenRecordset("ZZ_SCRATCH_P_TEXT", dbOpenDynaset)
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
    
Exit_WriteGIS_Text:
    Exit Sub

Err_WriteGIS_Text:
    MsgBox Err.Description
    Resume Exit_WriteGIS_Text
    
End Sub

Private Sub WriteKML_Addr()
On Error GoTo Err_WriteKML_Addr
    '
    '  This program will dump the results to a .gis file
    '
    If gPlaceRecCount = 0 Then
        MsgBox "There are no records to save."
        GoTo Exit_WriteKML_Addr
    End If
    '
    Dim tStream As ADODB.Stream, tPinyin As Boolean
    Set tStream = New ADODB.Stream
    
    tPinyin = False
    If GISFrame.Value = 1 Then
        tStream.Charset = "utf-8"
        tCodeStr = "UTF8"
    ElseIf GISFrame.Value = 3 Then
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

    dlgSaveAs.InitialFileName = "addr_gis_" + tCodeStr + ".kml"
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
            GoTo Exit_WriteKML_Addr
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
        Set tRstNode = CurrentDb.OpenRecordset("ZZ_SCRATCH_PEOPLE", dbOpenDynaset)
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
        tStream.WriteText tC + "<Style id=" + tDQ + "addr-balloon-template" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + "<BalloonStyle>", adWriteLine
        tStream.WriteText tC + tC + tC + "<text>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "<![CDATA[", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "ID: $[AddrGIS/PersonID] <br/>", adWriteLine
        If Not tPinyin Then
            tStream.WriteText tC + tC + tC + tC + "Name Chn: $[AddrGIS/NameHZ] <br/>", adWriteLine
        End If
        tStream.WriteText tC + tC + tC + tC + "Index Year: $[AddrGIS/IndexYear] <br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "Sex: $[AddrGIS/Sex] <br/>", adWriteLine
        If tPinyin Then
            tStream.WriteText tC + tC + tC + tC + "Address: $[AddrGIS/AddrName] <br/>", adWriteLine
        Else
            tStream.WriteText tC + tC + tC + tC + "Address: $[AddrGIS/AddrName] $[AddrGIS/AddrHZ] <br/>", adWriteLine
        End If
        tStream.WriteText tC + tC + tC + tC + "XY Count: $[AddrGIS/XYCount] <br/><br/>", adWriteLine
        tStream.WriteText tC + tC + tC + tC + "]]>", adWriteLine
        tStream.WriteText tC + tC + tC + "</text>", adWriteLine
        tStream.WriteText tC + tC + "</BalloonStyle>", adWriteLine
        tStream.WriteText tC + "</Style>", adWriteLine
        tStream.WriteText tC + "<!-- Declare the type " + tDQ + "AddrGIS" + tDQ + " with 6 fields -->", adWriteLine
        tStream.WriteText tC + "<Schema name=" + tDQ + "AddrGIS" + tDQ + " id=" + tDQ + "AddrGISId" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "uint" + tDQ + " name=" + tDQ + "PersonID" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Person ID]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        If Not tPinyin Then
            tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "NameHZ" + tDQ + ">", adWriteLine
            tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Name Chn]]></displayName>", adWriteLine
            tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        End If
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "Sex" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Sex]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "AddrName" + tDQ + ">", adWriteLine
        tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Address]]></displayName>", adWriteLine
        tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        If Not tPinyin Then
            tStream.WriteText tC + tC + "<SimpleField type=" + tDQ + "string" + tDQ + " name=" + tDQ + "AddrHZ" + tDQ + ">", adWriteLine
            tStream.WriteText tC + tC + tC + "<displayName><![CDATA[Address Chn]]></displayName>", adWriteLine
            tStream.WriteText tC + tC + "</SimpleField>", adWriteLine
        End If
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
                tStream.WriteText tC + tC + tC + "<SchemaData schemaUrl=" + tDQ + "#AddrGISId" + tDQ + ">", adWriteLine
                '
                '  person ID
                '
                tStr = Str(!c_person_id)
                tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "PersonID" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                '
                '  Person Name Chn
                '
                If Not tPinyin Then
                    If IsNull(!c_name_chn) Then
                        tStr = tStr + "[Bad Data]"
                    Else
                        If Trim(!c_name_chn) = "" Then
                            tStr = "[?]"
                        Else
                            tStr = !c_name_chn
                        End If
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
                tStr = IIf(!c_female, "F", "M")
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
                If Not tPinyin Then
                    If IsNull(!c_addr_chn) Then
                        tStr = "[?]"
                    ElseIf Trim(!c_addr_chn) = "" Then
                        tStr = "[?]"
                    Else
                        tStr = !c_addr_chn
                    End If
                    tStream.WriteText tC + tC + tC + tC + "<SimpleData name=" + tDQ + "AddrHZ" + tDQ + ">" + tStr + "</SimpleData>", adWriteLine
                End If
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

Exit_WriteKML_Addr:
    Exit Sub

Err_WriteKML_Addr:
    MsgBox Err.Description
    Resume Exit_WriteKML_Addr

End Sub

Private Sub WriteGIS_Addr()
On Error GoTo Err_WriteGIS_Addr
    '
    '  If it is a KML file, call the routine and exit
    '
    If ChkKML.Value Then
        Call WriteKML_Addr
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
    If gPlaceRecCount = 0 Then
        MsgBox "There are no records to save."
        GoTo Exit_WriteGIS_Addr
    End If
    '
    Dim tStream As ADODB.Stream
    Set tStream = New ADODB.Stream
    
    tPinyin = False
    If GISFrame.Value = 1 Then
        tStream.Charset = "utf-8"
        tCodeStr = "UTF8"
    ElseIf GISFrame.Value = 3 Then
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

    dlgSaveAs.InitialFileName = "place_gis_" + tCodeStr + ".tab"
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
            GoTo Exit_WriteGIS_Addr
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
        Set tRstNode = CurrentDb.OpenRecordset("ZZ_SCRATCH_PEOPLE", dbOpenDynaset)
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
                
                tStr = tStr + IIf(!c_female, "F", "M") + tC
                
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
    
Exit_WriteGIS_Addr:
    Exit Sub

Err_WriteGIS_Addr:
    MsgBox Err.Description
    Resume Exit_WriteGIS_Addr
    
End Sub

