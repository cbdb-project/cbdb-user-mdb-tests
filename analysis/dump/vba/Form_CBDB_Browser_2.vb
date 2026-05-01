Option Compare Database
Public gDisplayLanguage As String, gLabelsOK As Boolean, gPersonID As Long


Private Sub CmdSaveToFile_Click()
On Error GoTo Err_CmdSaveToFile_Click
    Dim tPersonName As String
    '
    ' first, just test if it knows the perdon ID
    '
    gPersonID = Me.BIOG_MAIN_2_Subform.Form.c_personid.Value
    'MsgBox Str(gPersonID)
    tPersonName = Me.frmPeopleLookup2.Form.c_name_chn.Value
    If IsNull(tPersonName) Then
        MsgBox "Name is NULL"
        Exit Sub
    End If
    
    ' the routine tests whether thre isinformation to be written for each category
    '
    ' set the language
    Dim tmli As MsoLanguageID, tLang As String
    ' get the labels
    tmli = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    
    If tmli = msoLanguageIDSimplifiedChinese Then
        tLang = "S"
    ElseIf tmli = msoLanguageIDTraditionalChinese Then
        tLang = "T"
    ElseIf tmli = msoLanguageIDEnglishUS Then
        tLang = "E"
    Else
        tLang = "E"
    End If
    '
    '  The challenge here is that we have 7 tables of results:
    '       BIOG_MAIN
    '       ZZ_SCRATCH_KIN
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
    Dim tRstPeople As DAO.Recordset, tRst As DAO.Recordset
    Dim tStr As String, tC As String
    Dim tQueryStr As String, tPersonID As Long, tCount As Long, tStrPersonID As String
    '
    
    tStrPersonID = Str(gPersonID)
    
    Dim gStream As ADODB.Stream, tCodeStr As String
    '
    ' set up the stream to write to
    
    Set gStream = New ADODB.Stream
    '
    gStream.Charset = "utf-8"
    tCodeStr = "UTF8"
    '
    ' Other options
        'gStream.Charset = "big5"
        'tCodeStr = "BIG5"
        'gStream.Charset = "gb2312"
        'tCodeStr = "GB2312"
        'gStream.Charset = "ascii"
        'tCodeStr = "ascii"
    '
    tC = Chr(44) ' the comma
    '
    Dim cmdSQL As ADODB.Command, tRecCount As Long
    Set cmdSQL = New ADODB.Command
    cmdSQL.ActiveConnection = CurrentProject.Connection
    cmdSQL.CommandType = adCmdText
    
    ' get the basic information
    
    ' Create the file

    Set dlgSaveAs = Application.FileDialog(msoFileDialogSaveAs)

    dlgSaveAs.InitialFileName = tPersonName + tCodeStr + ".htm"
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
            GoTo Exit_CmdSaveToFile_Click
        Else
            '  make sure the file name has a htm extension
            If Len(tFileName) < 5 Then
                tFileName = tFileName + ".htm"
            ElseIf Not (LCase(Right(tFileName, 4)) = ".htm") Then
                tFileName = tFileName + ".htm"
            End If
        End If
        '
        '  we have a file name:  now open the stream for writing
            
        gStream.Mode = adModeReadWrite
        gStream.Type = adTypeText
        gStream.Open
        
    tQueryStr = "SELECT BIOG_MAIN.c_personid, BIOG_MAIN.c_name, BIOG_MAIN.c_name_chn, BIOG_MAIN.c_index_year, BIOG_MAIN.c_index_year_type_code,  " + _
                    "INDEXYEAR_TYPE_CODES.c_index_year_type_desc, INDEXYEAR_TYPE_CODES.c_index_year_type_hz, BIOG_MAIN.c_female, " + _
                    "BIOG_MAIN.c_index_addr_id, ADDR_CODES.c_name AS c_addr_name, ADDR_CODES.c_name_chn AS c_addr_chn, " + _
                    "BIOG_MAIN.c_index_addr_type_code, BIOG_ADDR_CODES.c_addr_desc, BIOG_ADDR_CODES.c_addr_desc_chn, BIOG_MAIN.c_ethnicity_code, " + _
                    "ETHNICITY_TRIBE_CODES.c_name AS c_ethnicity_rmn, ETHNICITY_TRIBE_CODES.c_name_chn AS c_ethnicity_chn, " + _
                    "BIOG_MAIN.c_household_status_code, HOUSEHOLD_STATUS_CODES.c_household_status_desc, " + _
                    "HOUSEHOLD_STATUS_CODES.c_household_status_desc_chn, BIOG_MAIN.c_birthyear, BIOG_MAIN.c_deathyear, BIOG_MAIN.c_death_age, " + _
                    "BIOG_MAIN.c_fl_earliest_year, BIOG_MAIN.c_fl_latest_year, BIOG_MAIN.c_name_proper, BIOG_MAIN.c_name_rm " + _
                "FROM DYNASTIES RIGHT JOIN ( CHORONYM_CODES RIGHT JOIN ( HOUSEHOLD_STATUS_CODES RIGHT JOIN ( ETHNICITY_TRIBE_CODES " + _
                    "RIGHT JOIN ( ( ( BIOG_MAIN LEFT JOIN INDEXYEAR_TYPE_CODES " + _
                    "ON BIOG_MAIN.c_index_year_type_code = INDEXYEAR_TYPE_CODES.c_index_year_type_code ) LEFT JOIN ADDR_CODES " + _
                    "ON BIOG_MAIN.c_index_addr_id = ADDR_CODES.c_addr_id ) LEFT JOIN BIOG_ADDR_CODES " + _
                    "ON BIOG_MAIN.c_index_addr_type_code = BIOG_ADDR_CODES.c_addr_type ) " + _
                    "ON ETHNICITY_TRIBE_CODES.c_ethnicity_code = BIOG_MAIN.c_ethnicity_code ) " + _
                    "ON HOUSEHOLD_STATUS_CODES.c_household_status_code = BIOG_MAIN.c_household_status_code ) " + _
                    "ON CHORONYM_CODES.c_choronym_code = BIOG_MAIN.c_choronym_code ) ON DYNASTIES.c_dy = BIOG_MAIN.c_dy " + _
                "WHERE (((BIOG_MAIN.c_personid)=" + tStrPersonID + "))"
        '
        MsgBox "Writing header"
        gStream.WriteText "<HTML>", adWriteLine
        gStream.WriteText "<BODY>", adWriteLine
        tStr = "<P><B>Basic Information</B></P>"
        gStream.WriteText tStr, adWriteLine
        '
        MsgBox "Getting Person Data"
        Set tRstPeople = CurrentDb.OpenRecordset(tQueryStr)
        MsgBox "Writing Person Basic Data"
        With tRstPeople
            .MoveFirst
            Do While Not .EOF
                '  the ID of the person
                gStream.WriteText "<P><I>CBDB ID</I>", adWriteLine
                gStream.WriteText Trim(Str(!c_personid)) + "</P>", adWriteLine
                '
                '  name
                '
                gStream.WriteText "<P><I>Name</I>", adWriteLine
                If IsNull(!c_name) Then
                    tStr = "[Missing]"
                Else
                    tStr = !c_name
                End If
                '
               If IsNull(!c_name_chn) Then
                    tStr = tStr + " [Missing]"
                Else
                    tStr = tStr + " " + !c_name_chn
                End If
                gStream.WriteText tStr, adWriteLine
                
                ' c_name_proper, c_name_rm
                If Not IsNull(!c_name_proper) Then
                    gStream.WriteText "<BR>" + !c_name_proper, adWriteLine
                End If
                If Not IsNull(!c_name_rm) Then
                    gStream.WriteText "<BR>" + !c_name_rm, adWriteLine
                End If
                gStream.WriteText "</P>", adWriteLine
                        
                gStream.WriteText "<P><I>Sex</I>", adWriteLine
                If IsNull(!c_female) Then
                    tStr = "[Unknown]"
                Else
                    tStr = IIf(!c_female, "F", "M")
                End If
                gStream.WriteText tStr + "</P>", adWriteLine
                '
                '  indexyear = c_index_year INT
                '
                gStream.WriteText "<P><I>Index Year</I>", adWriteLine
                If IsNull(!c_index_year) Then
                    tStr = "[Unknown]"
                Else
                    tStr = Trim(Str(!c_index_year))
                End If
                gStream.WriteText tStr + "<BR>", adWriteLine
                '
                '  indexyear = c_index_year_type_desc STR
                '
                gStream.WriteText "<I>Index Year Type</I>", adWriteLine
                If IsNull(!c_index_year_type_desc) Then
                    tStr = "Unknown "
                Else
                    tStr = Trim(!c_index_year_type_desc) + " "
                End If
                '
                '  indexyear = c_index_year_type_hz STR
                '
                If Not IsNull(!c_index_year_type_hz) Then
                    tStr = tStr + Trim(!c_index_year_type_hz)
                End If
                gStream.WriteText tStr, adWriteLine
                
                ' Additional Year Information
                
                '"c_birthyear, c_deathyear, c_death_age, c_fl_earliest_year, c_fl_latest_year "
                If Not IsNull(!c_birthyear) Then
                    gStream.WriteText "<BR>Birth Year: " + Str(!c_birthyear), adWriteLine
                End If
                If Not IsNull(!c_deathyear) Then
                    gStream.WriteText "<BR>Death Year: " + Str(!c_deathyear), adWriteLine
                End If
                If Not IsNull(!c_death_age) Then
                    gStream.WriteText "<BR>Death Age: " + Str(!c_death_age), adWriteLine
                End If
                If Not IsNull(!c_fl_earliest_year) Then
                    gStream.WriteText "<BR>Earliest Floruit Year: " + Str(!c_fl_earliest_year), adWriteLine
                End If
                If Not IsNull(!c_fl_latest_year) Then
                    gStream.WriteText "<BR>Latest Floruit Year: " + Str(!c_fl_latest_year), adWriteLine
                End If
                gStream.WriteText "</P>", adWriteLine

                '  dynasty information
                '
                tStr = "<P><I>Dynasty</I>: "
                If IsNull(!c_dynasty) Then
                    tStr = tStr + "[Unknown]"
                Else
                    tStr = tStr + !c_dynasty + " " + !c_dynasty_chn
                End If
                gStream.WriteText tStr + "</P>", adWriteLine
                '
                ' Index Address information
                
                'c_index_addr_name, c_index_addr_chn, c_index_addr_type_desc, c_index_addr_type_chn, x_coord, y_coord, "
                
                gStream.WriteText "<P><I>Index Address</I>", adWriteLine
                If IsNull(!c_index_addr_name) Then
                    gStream.WriteText "[Unknown]</P>", adWriteLine
                Else
                    tStr = Trim(!c_index_addr_name) + " " + Trim(!c_index_addr_chn)
                    gStream.WriteText tStr + "<BR>", adWriteLine
                    If Not IsNull(!x_coord) Then
                        gStream.WriteText "Coordinates: " + Str(!x_coord) + tC + Str(!y_coord) + "<BR>", adWriteLine
                    End If
                    '
                    gStream.WriteText "<I>Index Address Type</I>", adWriteLine
                    If IsNull(!c_index_addr_type_desc) Then
                        tStr = "Unknown "
                    Else
                        tStr = Trim(!c_index_addr_type_desc) + " "
                    End If
                    '
                    If Not IsNull(!c_index_addr_type_chn) Then
                        tStr = tStr + Trim(!c_index_addr_type_chn)
                    End If
                    gStream.WriteText tStr + "</P>", adWriteLine
                End If
                
                '", c_ethnicity_chn, c_ethnicity_rmn, "
                If Not IsNull(!c_ethnicity_chn) Then
                    gStream.WriteText "<P>Ethnicity: " + !c_ethnicity_rmn + " " + !c_ethnicity_chn + "</P>", adWriteLine
                End If
                '"c_choronym_desc, c_choronym_chn, "
                If Not IsNull(!c_choronym_desc) Then
                    gStream.WriteText "<P>Choronym: " + !c_choronym_desc + " " + !c_choronym_chn + "</P>", adWriteLine
                End If
                '"c_household_status_desc, c_household_status_desc_chn "
                If Not IsNull(!c_household_status_desc) Then
                    gStream.WriteText "<P>Household Status: " + !c_household_status_desc + " " + !c_household_status_desc_chn + "</P>", adWriteLine
                End If
                .MoveNext
            Loop
        End With
    Else
        'The user pressed Cancel.
        GoTo Exit_CmdSaveToFile_Click
    End If
    '
    ' Alt Names
    '
    MsgBox "Writing alternate names"
    tQueryStr = "SELECT ALTNAME_DATA.c_personid, ALTNAME_DATA.c_alt_name, ALTNAME_DATA.c_alt_name_chn, ALTNAME_DATA.c_alt_name_type_code, " + _
                    "ALTNAME_CODES.c_name_type_desc, ALTNAME_CODES.c_name_type_desc_chn, ALTNAME_DATA.c_sequence, ALTNAME_DATA.c_source, " + _
                    "TEXT_CODES.c_title_chn, TEXT_CODES.c_title, ALTNAME_DATA.c_pages, ALTNAME_DATA.c_notes " + _
                "FROM TEXT_CODES RIGHT JOIN ( ALTNAME_CODES INNER JOIN ALTNAME_DATA ON ALTNAME_CODES.c_name_type_code = ALTNAME_DATA.c_alt_name_type_code ) " + _
                    "ON TEXT_CODES.c_textid = ALTNAME_DATA.c_source " + _
                "WHERE (((ALTNAME_DATA.c_personid)=" + tStrPersonID + "))"
    Set tRst = CurrentDb.OpenRecordset(tQueryStr)
    
    If Not (tRst.EOF) Then
        '
        tRst.MoveLast
        tStr = "<P><B><I>Alternate Names</I></B> "
        If tRst.RecordCount = 1 Then
            tStr = tStr + "(1 record)"
        Else
            tStr = tStr + "(" + Trim(Str(tRst.RecordCount)) + " records)"
        End If
        gStream.WriteText tStr + "</P>", adWriteLine
        '
        ' !c_alt_name, !c_alt_name_chn, " + _
        ' !c_name_type_desc, !c_name_type_desc_chn, !c_sequence, !c_source, " + _
        ' !c_title_chn, !c_title, !c_pages, !c_notes
        '
        With tRst
            .MoveFirst
            Do While Not .EOF
                '
                tStr = "<P><I>Name</I>: " + Trim(!c_alt_name) + " " + Trim(!c_alt_name_chn) + "<BR>"
                gStream.WriteText tStr, adWriteLine
                '
                tStr = "Type: " + Trim(!c_name_type_desc) + " " + Trim(!c_name_type_desc_chn) + "<BR>"
                gStream.WriteText tStr, adWriteLine
                '
                If Not IsNull(!c_sequence) Then
                    tStr = "Sequence: " + Trim(Str(!c_sequence)) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                If Not IsNull(!c_title) Then
                    tStr = "Source: " + Trim(!c_title) + " " + Trim(!c_title_chn)
                    If Not IsNull(!c_pages) Then
                        tStr = tStr + " (" + Trim(!c_pages) + ")"
                    End If
                    gStream.WriteText tStr + "<BR>", adWriteLine
                End If
                '
                If Not IsNull(!c_notes) Then
                    tStr = "Notes: " + Trim(!c_notes) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                gStream.WriteText "</P>", adWriteLine
                '
                .MoveNext
            Loop
        End With
    End If
    '
    '  now Places
    '
    ' replace the tLang condition with the count of addresses
    '
    tQueryStr = "SELECT ADDR_CODES.c_name AS c_addr_name, ADDR_CODES.c_name_chn AS c_addr_chn, BIOG_ADDR_CODES.c_addr_desc, BIOG_ADDR_CODES.c_addr_desc_chn, " + _
            "BIOG_ADDR_DATA.c_firstyear, BIOG_ADDR_DATA.c_lastyear, TEXT_CODES.c_title AS c_source_title, TEXT_CODES.c_title_chn AS c_source_chn, " + _
            "BIOG_ADDR_DATA.c_personid " + _
        "FROM ( ( BIOG_ADDR_CODES INNER JOIN BIOG_ADDR_DATA ON BIOG_ADDR_CODES.c_addr_type = BIOG_ADDR_DATA.c_addr_type ) INNER JOIN ADDR_CODES " + _
            "ON BIOG_ADDR_DATA.c_addr_id = ADDR_CODES.c_addr_id ) LEFT JOIN TEXT_CODES ON BIOG_ADDR_DATA.c_source = TEXT_CODES.c_textid " + _
        "WHERE (((BIOG_ADDR_DATA.c_personid) = " + tStrPersonID + "))"
        
    Set tRst = CurrentDb.OpenRecordset(tQueryStr)
    
    MsgBox "Writing address information"
    If Not (tRst.EOF) Then
        '
        '  now the BIOG_ADDR data
        '
        tRst.MoveLast
        tStr = "<P><B><I>Place Information</I></B> "
        If tRst.RecordCount = 1 Then
            tStr = tStr + "(1 record)"
        Else
            tStr = tStr + "(" + Trim(Str(tRst.RecordCount)) + " records)"
        End If
        gStream.WriteText tStr + "</P>", adWriteLine
        '
        ' c_addr_name, c_addr_chn, x_coord, y_coord, c_addr_desc, c_addr_desc_chn, "
        ' c_firstyear, c_lastyear, "
        ' c_source_title, c_source_chn, c_pages, c_notes "
        
        With tRst
            .MoveFirst
            Do While Not .EOF
                tStr = "<P><I>" + Trim(!c_addr_desc) + " " + Trim(c_addr_desc_chn) + "</I>: " + Trim(!c_addr_name) + " " + Trim(!c_addr_chn)
                gStream.WriteText tStr, adWriteLine
                
                If Not IsNull(!x_coord) Then
                    '
                    tStr = " (" + Trim(Str(!x_coord)) + tC + Trim(Str(!y_coord)) + ")<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                
                If Not IsNull(!c_firstyear) Then
                    tStr = "First year: " + Trim(Str(!c_firstyear)) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                
                If Not IsNull(!c_lastyear) Then
                    tStr = "Last year: " + Trim(Str(!c_lastyear)) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                
                If Not IsNull(!c_source_title) Then
                    tStr = "Source: " + Trim(!c_source_title) + " " + Trim(!c_source_chn)
                    If Not IsNull(!c_pages) Then
                        tStr = tStr + " " + Trim(!c_pages)
                    End If
                    tStr = tStr + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                
                If Not IsNull(!c_notes) Then
                    tStr = "Notes: " + Trim(!c_notes) + ")<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                gStream.WriteText "</P>", adWriteLine
                .MoveNext
            Loop
        End With
        '
    End If
    
    '  now kinship
    '
    tQueryStr = "SELECT ZZ_SCRATCH_KIN.c_kin_id, ZZ_SCRATCH_KIN.c_kin_name, ZZ_SCRATCH_KIN.c_kin_chn, ZZ_SCRATCH_KIN.c_kin_index_year, " + _
            "ZZ_SCRATCH_KIN.c_kin_dynasty, ZZ_SCRATCH_KIN.c_kin_dynasty_chn, ZZ_SCRATCH_KIN.c_kin_sex, ZZ_SCRATCH_KIN.c_kin_rel_total, " + _
            "ZZ_SCRATCH_KIN.c_up, ZZ_SCRATCH_KIN.c_down, ZZ_SCRATCH_KIN.c_marriage, ZZ_SCRATCH_KIN.c_collateral, ZZ_SCRATCH_KIN.c_kin_addr_name, " + _
            "ZZ_SCRATCH_KIN.c_kin_addr_chn, ZZ_SCRATCH_KIN.kin_x_coord, ZZ_SCRATCH_KIN.kin_y_coord, ZZ_SCRATCH_KIN.c_source_text, " + _
            "ZZ_SCRATCH_KIN.c_source_text_chn, ZZ_SCRATCH_KIN.c_pages, ZZ_SCRATCH_KIN.c_notes " + _
        "FROM ZZ_SCRATCH_KIN"
    Set tRst = CurrentDb.OpenRecordset(tQueryStr)
    
    MsgBox "Writing kinship data"
    If Not (tRst.EOF) Then
        '
        tRst.MoveLast
        tStr = "<P><B><I>Kinship</I></B> "
        If tRst.RecordCount = 1 Then
            tStr = tStr + "(1 record)"
        Else
            tStr = tStr + "(" + Trim(Str(tRst.RecordCount)) + " records)"
        End If
        gStream.WriteText tStr + "</P>", adWriteLine
        
        ' c_kin_id, c_kin_name, c_kin_chn, c_kin_sex,
        ' c_kin_rel_total, c_up, c_down, c_marriage, c_collateral,
        ' c_kin_index_year, c_kin_dynasty, c_kin_dynasty_chn,
        ' c_kin_addr_name, c_kin_addr_chn, kin_x_coord, kin_y_coord,
        ' c_source_text, c_source_text_chn, c_pages, c_notes"
        '
        With tRst
            .MoveFirst
            Do While Not .EOF
                tStr = "<P>" + Trim(!c_kin_name) + " " + Trim(!c_kin_chn) + "(" + !c_kin_sex + ") [CBDB ID " + Trim(Str(!c_kin_id)) + "]<BR>"
                gStream.WriteText tStr, adWriteLine
                
                tStr = "Relationship: " + Trim(!c_kin_rel_total) + " (" + Str(!c_up) + "-" + Str(!c_down) + "-" + _
                    Str(!c_marriage) + "-" + Str(!c_collateral) + ")<BR>"
                gStream.WriteText tStr, adWriteLine
                
                If Not IsNull(!c_kin_index_year) Then
                    tStr = "Index Year: " + Trim(Str(!c_kin_index_year)) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                
                If Not IsNull(!c_kin_dynasty) Then
                    tStr = "Dynasty: " + Trim(!c_kin_dynasty) + " " + Trim(!c_kin_dynasty_chn) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                
                If Not IsNull(!c_kin_addr_name) Then
                    tStr = "Index Address: " + Trim(!c_kin_addr_name) + " " + Trim(!c_kin_addr_chn)
                    If Not IsNull(!kin_x_coord) Then
                        tStr = tStr + " (" + Trim(Str(!kin_x_coord)) + tC + Trim(Str(!kin_y_coord)) + ")"
                    End If
                    tStr = tStr + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                
                If Not IsNull(!c_source_text) Then
                    tStr = "Source: " + Trim(!c_source_text) + " " + Trim(!c_source_text_chn)
                    If Not IsNull(!c_pages) Then
                        tStr = tStr + " (" + Trim(!c_pages) + ")"
                    End If
                    tStr = tStr + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                
                If Not IsNull(!c_notes) Then
                    tStr = "Notes: " + Trim(!c_notes) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                gStream.WriteText "</P>", adWriteLine
                .MoveNext
            Loop
        End With
        '
    End If
    '
    '  now Associations: Because of the outer joins, I need to put the data into a scratch table
    '
    ' clear the scratch table
    '
    cmdSQL.CommandText = "DELETE * FROM ZZ_SCRATCH_ASSOC"
    cmdSQL.Execute tRecCount
    '
    tQueryStr = "INSERT INTO ZZ_SCRATCH_ASSOC (c_person_id, c_assoc_id, c_assoc_name, c_assoc_chn, c_assoc_index_year, c_assoc_dy, c_assoc_sex, " + _
            "c_assoc_addr_id, c_assoc_desc, c_assoc_desc_chn, c_assoc_first_year, c_assoc_last_year, c_source, c_pages, c_notes, c_assoc_place_addr_id, " + _
            "c_litgenre_code, c_occasion_code, c_topic_code, c_inst_code, c_inst_name_code, c_text_title, c_assoc_count ) " + _
        "SELECT ASSOC_DATA.c_personid, ASSOC_DATA.c_assoc_id, BIOG_MAIN.c_name, BIOG_MAIN.c_name_chn, BIOG_MAIN.c_index_year, BIOG_MAIN.c_dy, " + _
            "IIf([c_female], 'F', 'M') AS c_assoc_sex, BIOG_MAIN.c_index_addr_id, ASSOC_CODES.c_assoc_desc, ASSOC_CODES.c_assoc_desc_chn, " + _
            "ASSOC_DATA.c_assoc_first_year, ASSOC_DATA.c_assoc_last_year, ASSOC_DATA.c_source, ASSOC_DATA.c_pages, ASSOC_DATA.c_notes, " + _
            "ASSOC_DATA.c_addr_id, ASSOC_DATA.c_litgenre_code, ASSOC_DATA.c_occasion_code, ASSOC_DATA.c_topic_code, ASSOC_DATA.c_inst_code, " + _
            "ASSOC_DATA.c_inst_name_code, ASSOC_DATA.c_text_title, ASSOC_DATA.c_assoc_count " + _
        "FROM ASSOC_CODES INNER JOIN ( ASSOC_DATA INNER JOIN BIOG_MAIN ON ASSOC_DATA.c_assoc_id = BIOG_MAIN.c_personid ) " + _
            "ON ASSOC_CODES.c_assoc_code = ASSOC_DATA.c_assoc_code " + _
        "WHERE (((ASSOC_DATA.c_personid)=" + tStrPersonID + "))"
    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecCount
    '
    '   fill in the outer join data
    '
    tQueryStr = "UPDATE ( ( ( ( ( ( ZZ_SCRATCH_ASSOC LEFT JOIN ADDR_CODES ON ZZ_SCRATCH_ASSOC.c_addr_id = ADDR_CODES.c_addr_id ) LEFT JOIN SOCIAL_INSTITUTION_NAME_CODES " + _
            "ON ZZ_SCRATCH_ASSOC.c_inst_name_code = SOCIAL_INSTITUTION_NAME_CODES.c_inst_name_code ) LEFT JOIN DYNASTIES " + _
            "ON ZZ_SCRATCH_ASSOC.c_assoc_dy = DYNASTIES.c_dy ) LEFT JOIN LITERARYGENRE_CODES " + _
            "ON ZZ_SCRATCH_ASSOC.c_litgenre_code = LITERARYGENRE_CODES.c_lit_genre_code ) LEFT JOIN OCCASION_CODES " + _
            "ON ZZ_SCRATCH_ASSOC.c_occasion_code = OCCASION_CODES.c_occasion_code ) LEFT JOIN SCHOLARLYTOPIC_CODES " + _
            "ON ZZ_SCRATCH_ASSOC.c_topic_code = SCHOLARLYTOPIC_CODES.c_topic_code ) INNER JOIN TEXT_CODES ON ZZ_SCRATCH_ASSOC.c_source = TEXT_CODES.c_textid " + _
        "SET ZZ_SCRATCH_ASSOC.c_assoc_dynasty = [DYNASTIES].[c_dynasty], ZZ_SCRATCH_ASSOC.c_assoc_dynasty_chn = [DYNASTIES].[c_dynasty_chn], " + _
            "ZZ_SCRATCH_ASSOC.c_assoc_addr_name = [ADDR_CODES].[c_name], ZZ_SCRATCH_ASSOC.c_assoc_addr_chn = [ADDR_CODES].[c_name_chn], " + _
            "ZZ_SCRATCH_ASSOC.assoc_xcoord = [ADDR_CODES].[x_coord], ZZ_SCRATCH_ASSOC.assoc_ycoord = [ADDR_CODES].[y_coord], " + _
            "ZZ_SCRATCH_ASSOC.c_litgenre_desc = [LITERARYGENRE_CODES].[c_lit_genre_desc], " + _
            "ZZ_SCRATCH_ASSOC.c_litgenre_desc_chn = [LITERARYGENRE_CODES].[c_lit_genre_desc_chn], " + _
            "ZZ_SCRATCH_ASSOC.c_occasion_desc = [OCCASION_CODES].[c_occasion_desc], " + _
            "ZZ_SCRATCH_ASSOC.c_occasion_desc_chn = [OCCASION_CODES].[c_occasion_desc_chn], " + _
            "ZZ_SCRATCH_ASSOC.c_topic_desc = [SCHOLARLYTOPIC_CODES].[c_topic_desc], " + _
            "ZZ_SCRATCH_ASSOC.c_topic_desc_chn = [SCHOLARLYTOPIC_CODES].[c_topic_desc_chn], " + _
            "ZZ_SCRATCH_ASSOC.c_inst_name_py = [SOCIAL_INSTITUTION_NAME_CODES].[c_inst_name_py], " + _
            "ZZ_SCRATCH_ASSOC.c_inst_name_hz = [SOCIAL_INSTITUTION_NAME_CODES].[c_inst_name_hz], " + _
            "ZZ_SCRATCH_ASSOC.c_source_text = [TEXT_CODES].[c_title], ZZ_SCRATCH_ASSOC.c_source_text_chn = [TEXT_CODES].[c_title_chn]"
    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecCount
    '
    Set tRst = CurrentDb.OpenRecordset("ZZ_SCRATCH_ASSOC")
    
    MsgBox "Writing association data"
    If Not (tRst.EOF) Then
        '
        tRst.MoveLast
        tStr = "<P><B><I>Associations</I></B> "
        If tRst.RecordCount = 1 Then
            tStr = tStr + "(1 record)"
        Else
            tStr = tStr + "(" + Trim(Str(tRst.RecordCount)) + " records)"
        End If
        gStream.WriteText tStr + "</P>", adWriteLine
        '
        With tRst
            .MoveFirst
            Do While Not .EOF
                tStr = "<P>" + !c_assoc_name + " " + !c_assoc_chn + "(" + !c_assoc_sex + ") <BR>"
                gStream.WriteText tStr, adWriteLine
                
                tStr = !c_assoc_desc + " " + !c_assoc_desc_chn
                If Not IsNull(!c_assoc_year) Then
                    tStr = tStr + "(Year: " + Trim(Str(!c_assoc_year)) + ")"
                End If
                If Not IsNull(!c_assoc_count) Then
                    tStr = tStr + "[Count: " + Trim(Str(!c_assoc_count)) + "]"
                End If
                gStream.WriteText tStr + "<BR>", adWriteLine
                
                If Not IsNull(!c_text_title) Then
                    tStr = "Title: " + Trim(!c_text_title) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                
                tStr = "<I>Personal Information</I><BR>"
                gStream.WriteText tStr, adWriteLine
                
                If Not IsNull(!c_assoc_index_year) Then
                    tStr = "Index Year: " + Trim(Str(!c_assoc_index_year)) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                
                If Not IsNull(!c_assoc_dynasty) Then
                    tStr = "Dynasty: " + Trim(!c_assoc_dynasty) + Trim(!c_assoc_dynasty_chn) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                
                If Not IsNull(!c_assoc_addr_name) Then
                    tStr = "Index Address: " + Trim(!c_assoc_addr_name) + Trim(!c_assoc_addr_chn)
                    If Not IsNull(!assoc_xcoord) Then
                        tStr = tStr + " (" + Trim(Str(!assoc_xcoord)) + tC + Trim(Str(!assoc_ycoord)) + ")"
                    End If
                    gStream.WriteText tStr + "<BR>", adWriteLine
                End If
                
                If Not IsNull(!c_lit_genre_desc) Then
                    tStr = "Literary Genre: " + Trim(!c_lit_genre_desc) + Trim(!c_lit_genre_desc_chn) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                
                If Not IsNull(!c_occasion_desc) Then
                    tStr = "Occasion: " + Trim(!c_occasion_desc) + Trim(!c_occasion_desc_chn) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                
                If Not IsNull(!c_topic_desc) Then
                    tStr = "Topic: " + Trim(!c_topic_desc) + Trim(!c_topic_desc_chn) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                
                If Not IsNull(!c_inst_name_py) Then
                    tStr = "Institution: " + Trim(!c_inst_name_py) + Trim(!c_inst_name_hz) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                
                If Not IsNull(!c_source_title) Then
                    tStr = "Source: " + Trim(!c_source_title) + " " + Trim(!c_source_chn)
                    If Not IsNull(!c_pages) Then
                        tStr = tStr + " (" + Trim(!c_pages) + ")"
                    End If
                    tStr = tStr + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                
                If Not IsNull(!c_notes) Then
                    tStr = "Notes: " + Trim(!c_notes) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                .MoveNext
                gStream.WriteText "</P>", adWriteLine
            Loop
        End With
        '
    End If
    '
    ' now Entry
    '
    tQueryStr = "SELECT ENTRY_CODES.c_entry_desc, ENTRY_CODES.c_entry_desc_chn, KINSHIP_CODES.c_kinrel, KINSHIP_CODES.c_kinrel_chn, ASSOC_CODES.c_assoc_desc, " + _
            "ASSOC_CODES.c_assoc_desc_chn, BIOG_MAIN.c_name AS c_kin_name, BIOG_MAIN.c_name_chn AS c_kin_name_chn, BIOG_MAIN_1.c_name AS c_assoc_name, " + _
            "BIOG_MAIN_1.c_name_chn AS c_assoc_name_chn,  PARENTAL_STATUS_CODES.c_parental_status_desc, " + _
            "PARENTAL_STATUS_CODES.c_parental_status_desc_chn, TEXT_CODES.c_title, " + _
            "TEXT_CODES.c_title_chn, ADDR_CODES.c_name AS c_entry_addr_name, ADDR_CODES.c_name_chn AS c_entry_addr_chn, ADDR_CODES.x_coord AS c_entry_xcoord, " + _
            "ADDR_CODES.y_coord AS c_entry_ycoord, ENTRY_DATA.c_year, ENTRY_DATA.c_sequence, ENTRY_DATA.c_age, ENTRY_DATA.c_exam_rank " + _
        "FROM TEXT_CODES RIGHT JOIN ( PARENTAL_STATUS_CODES RIGHT JOIN ( ADDR_CODES RIGHT JOIN ( ( ( ASSOC_CODES INNER JOIN ( ( KINSHIP_CODES " + _
            "INNER JOIN ( ENTRY_CODES INNER JOIN ENTRY_DATA ON ENTRY_CODES.c_entry_code = ENTRY_DATA.c_entry_code ) " + _
            "ON KINSHIP_CODES.c_kincode = ENTRY_DATA.c_kin_code ) INNER JOIN BIOG_MAIN " + _
            "ON ENTRY_DATA.c_kin_id = BIOG_MAIN.c_personid ) ON ASSOC_CODES.c_assoc_code = ENTRY_DATA.c_assoc_code ) INNER JOIN BIOG_MAIN AS BIOG_MAIN_1 " + _
            "ON ENTRY_DATA.c_assoc_id = BIOG_MAIN_1.c_personid ) INNER JOIN SOCIAL_INSTITUTION_NAME_CODES " + _
            "ON ENTRY_DATA.c_inst_name_code = SOCIAL_INSTITUTION_NAME_CODES.c_inst_name_code ) ON ADDR_CODES.c_addr_id = ENTRY_DATA.c_entry_addr_id ) " + _
            "ON PARENTAL_STATUS_CODES.c_parental_status_code = ENTRY_DATA.c_parental_status ) ON TEXT_CODES.c_textid = ENTRY_DATA.c_source " + _
        "WHERE (((ENTRY_DATA.c_personid) = " + tStrPersonID + "))"
    Set tRst = CurrentDb.OpenRecordset(tQueryStr)
    
    MsgBox "Writing entry data"
    If Not (tRst.EOF) Then
        '
        ' !c_entry_desc, !c_entry_desc_chn, !c_year, !c_sequence,!c_age, !c_exam_rank,
        ' !c_parental_status_desc, !c_parental_status_desc_chn,
        ' !c_entry_addr_name, !c_entry_addr_chn, !c_entry_xcoord, !c_entry_ycoord
        ' !c_kinrel_chn, !c_kinrel, !c_kin_name, !c_kin_name_chn,
        ' !c_assoc_desc, !c_assoc_desc_chn, !c_assoc_name, !c_assoc_name_chn,
        ' !c_inst_name_hz, !c_inst_name_py,
        ' !c_title_chn, !c_title, !c_pages, !c_notes,
        '
        tRst.MoveLast
        tStr = "<P><B><I>Entry into Government</I></B> "
        If tRst.RecordCount = 1 Then
            tStr = tStr + "(1 record)"
        Else
            tStr = tStr + "(" + Trim(Str(tRst.RecordCount)) + " records)"
        End If
        gStream.WriteText tStr + "</P>", adWriteLine
        '
        With tRst
            .MoveFirst
            Do While Not .EOF
                '
                tStr = "<P><I>Entry</I>: " + Trim(!c_entry_desc) + " " + Trim(!c_entry_desc_chn) + "<BR>"
                gStream.WriteText tStr, adWriteLine
                '
                If Not IsNull(!c_year) Then
                    tStr = "Year: " + Trim(Str(!c_year)) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                If Not IsNull(!c_sequence) Then
                    tStr = "Sequence: " + Trim(Str(!c_sequence)) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                If Not IsNull(!c_age) Then
                    tStr = "Age: " + Trim(Str(!c_age)) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                If Not IsNull(!c_exam_rank) Then
                    tStr = "Exam Rank: " + Trim(!c_exam_rank) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                If Not IsNull(!c_parental_status_desc) Then
                    tStr = "Parental Status: " + Trim(!c_parental_status_desc) + " " + Trim(!c_parental_status_desc_chn) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                If Not IsNull(!c_entry_addr_name) Then
                    tStr = "Location: " + Trim(!c_entry_addr_name) + " " + Trim(!c_entry_addr_chn)
                    If Not IsNull(!c_entry_xcoord) Then
                        tStr = tStr + " (" + Trim(Str(!c_entry_xcoord)) + tC + Trim(Str(!c_entry_ycoord)) + ")"
                    End If
                    gStream.WriteText tStr + "<BR>", adWriteLine
                End If
                '
                If Not IsNull(!c_kinrel) Then
                    tStr = "Kinship Relation: " + Trim(!c_kinrel) + " " + Trim(!c_kinrel_chn)
                    gStream.WriteText tStr + "<BR>", adWriteLine
                    If Not IsNull(!c_kin_name) Then
                        tStr = "Kin Name: " + Trim(!c_kin_name) + " " + Trim(!c_kin_name_chn)
                        gStream.WriteText tStr + "<BR>", adWriteLine
                    End If
                End If
                '
                If Not IsNull(!c_assoc_desc) Then
                    tStr = "Association: " + Trim(!c_assoc_desc) + " " + Trim(!c_assoc_desc_chn)
                    gStream.WriteText tStr + "<BR>", adWriteLine
                    If Not IsNull(!c_kin_name) Then
                        tStr = "Associate Name: " + Trim(!c_assoc_name) + " " + Trim(!c_assoc_name_chn)
                        gStream.WriteText tStr + "<BR>", adWriteLine
                    End If
                End If
                '
                If Not IsNull(!c_inst_name_py) Then
                    tStr = "Institution: " + Trim(!c_inst_name_py) + " " + Trim(!c_inst_name_hz) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                If Not IsNull(!c_title) Then
                    tStr = "Source: " + Trim(!c_title) + " " + Trim(!c_title_chn)
                    If Not IsNull(!c_pages) Then
                        tStr = tStr + " (" + Trim(!c_pages) + ")"
                    End If
                    gStream.WriteText tStr + "<BR>", adWriteLine
                End If
                '
                If Not IsNull(!c_notes) Then
                    tStr = "Notes: " + Trim(!c_notes) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                gStream.WriteText "</P>", adWriteLine
                '
                .MoveNext
            Loop
        End With
    End If
    '
    ' now Office
    '
    '  because there are outer joins, I need to copy the data to a scratch table then deal with the outer join
    '
    cmdSQL.CommandText = "DELETE * FROM ZZ_SCRATCH_OFFICE"
    cmdSQL.Execute tRecCount
    '
    tQueryStr = "INSERT INTO ZZ_SCRATCH_OFFICE ( c_personid, c_office_id, c_posting_id, c_sequence, c_firstyear, c_lastyear, c_appt_code, " + _
            "c_assume_office_code, c_inst_code, c_inst_name_code, c_source, c_pages, c_dy,  c_appt_desc_chn,  c_appt_desc,  c_assume_office_desc_chn, " + _
            "c_assume_office_desc, c_inst_name_hz, c_inst_name_py, c_dynasty, c_dynasty_chn, c_office_pinyin, " + _
            "c_office_chn, c_office_trans, c_office_addr_id ) " + _
        "SELECT POSTED_TO_OFFICE_DATA.c_personid, POSTED_TO_OFFICE_DATA.c_office_id, POSTED_TO_OFFICE_DATA.c_posting_id, POSTED_TO_OFFICE_DATA.c_sequence," + _
            "POSTED_TO_OFFICE_DATA.c_firstyear, POSTED_TO_OFFICE_DATA.c_lastyear, POSTED_TO_OFFICE_DATA.c_appt_code, POSTED_TO_OFFICE_DATA.c_assume_office_code, " + _
            "POSTED_TO_OFFICE_DATA.c_inst_code, POSTED_TO_OFFICE_DATA.c_inst_name_code, POSTED_TO_OFFICE_DATA.c_source, POSTED_TO_OFFICE_DATA.c_pages, " + _
            "POSTED_TO_OFFICE_DATA.c_dy, APPOINTMENT_CODES.c_appt_desc_chn, APPOINTMENT_CODES.c_appt_desc, ASSUME_OFFICE_CODES.c_assume_office_desc_chn, " + _
            "ASSUME_OFFICE_CODES.c_assume_office_desc, SOCIAL_INSTITUTION_NAME_CODES.c_inst_name_hz, SOCIAL_INSTITUTION_NAME_CODES.c_inst_name_py, " + _
            "DYNASTIES.c_dynasty, DYNASTIES.c_dynasty_chn, OFFICE_CODES.c_office_pinyin, " + _
            "OFFICE_CODES.c_office_chn, OFFICE_CODES.c_office_trans, POSTED_TO_ADDR_DATA.c_addr_id " + _
        "FROM ( TEXT_CODES RIGHT JOIN ( SOCIAL_INSTITUTION_NAME_CODES INNER JOIN ( OFFICE_CODES INNER JOIN ( DYNASTIES " + _
            "RIGHT JOIN ( ASSUME_OFFICE_CODES RIGHT JOIN (APPOINTMENT_CODES RIGHT JOIN POSTED_TO_OFFICE_DATA " + _
            "ON APPOINTMENT_CODES.c_appt_code = POSTED_TO_OFFICE_DATA.c_appt_code ) " + _
            "ON ASSUME_OFFICE_CODES.c_assume_office_code = POSTED_TO_OFFICE_DATA.c_assume_office_code ) " + _
            "ON DYNASTIES.c_dy = POSTED_TO_OFFICE_DATA.c_dy) ON OFFICE_CODES.c_office_id = POSTED_TO_OFFICE_DATA.c_office_id ) " + _
            "ON SOCIAL_INSTITUTION_NAME_CODES.c_inst_name_code = POSTED_TO_OFFICE_DATA.c_inst_name_code ) " + _
            "ON TEXT_CODES.c_textid = POSTED_TO_OFFICE_DATA.c_source ) LEFT JOIN POSTED_TO_ADDR_DATA " + _
            "ON (POSTED_TO_OFFICE_DATA.c_office_id = POSTED_TO_ADDR_DATA.c_office_id ) " + _
            "AND ( POSTED_TO_OFFICE_DATA.c_posting_id = POSTED_TO_ADDR_DATA.c_posting_id ) " + _
        "WHERE (((POSTED_TO_ADDR_DATA.c_personid)=" + tStrPersonID + "))"
        '
    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecCount
    
    MsgBox "Writing postings data"
    If (tRecCount > 0) Then
    '
    '  now get the address data
    '
    tQueryStr = "UPDATE ZZ_SCRATCH_OFFICE LEFT JOIN ADDR_CODES ON ZZ_SCRATCH_OFFICE.c_office_addr_id = ADDR_CODES.c_addr_id " + _
                "SET ZZ_SCRATCH_OFFICE.c_office_addr_name = [ADDR_CODES].[c_name], ZZ_SCRATCH_OFFICE.c_office_addr_chn = [ADDR_CODES].[c_name_chn], " + _
                    "ZZ_SCRATCH_OFFICE.office_x_coord = [ADDR_CODES].[x_coord], ZZ_SCRATCH_OFFICE.office_y_coord = [ADDR_CODES].[y_coord]"
'
    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecCount
    '
    Set tRst = CurrentDb.OpenRecordset("ZZ_SCRATCH_OFFICE")
        tRst.MoveLast
        tStr = "<P><B><I>Office Appointments</I></B> "
        If tRst.RecordCount = 1 Then
            tStr = tStr + "(1 record)"
        Else
            tStr = tStr + "(" + Trim(Str(tRst.RecordCount)) + " records)"
        End If
        gStream.WriteText tStr + "</P>", adWriteLine
        '
        With tRst
            .MoveFirst
            '
            ' !c_personid, !c_office_pinyin, !c_office_chn, !c_office_trans
            ' !c_office_addr_name, !c_office_addr_chn
            ' !office_x_coord, !office_y_coord
            ' !c_sequence, !c_dynasty, !c_dynasty_chn, !c_firstyear, !c_lastyear
            ' !c_category_desc, !c_category_desc_chn, !c_appt_desc_chn, !c_appt_desc
            ' !c_assume_office_desc_chn, !c_assume_office_desc
            ' !c_inst_name_py, !c_inst_name_hz
            ' !c_title, !c_title_chn, !c_pages, !c_notes

            Do While Not .EOF
                '  the ID of the person
                tStr = "<P><I>Appointment</I>: " + Trim(!c_office_pinyin) + " " + Trim(!c_office_chn)
                If Not IsNull(!c_office_trans) Then
                    tStr = tStr + " " + Trim(!c_office_trans)
                End If
                gStream.WriteText tStr + "<BR>", adWriteLine
                '
                If Not IsNull(!c_office_addr_name) Then
                    tStr = "Location: " + Trim(!c_office_addr_name) + " " + Trim(!c_office_addr_chn)
                    If Not IsNull(!office_x_coord) Then
                        tStr = tStr + " (" + Trim(Str(!office_x_coord)) + tC + Trim(Str(!office_y_coord)) + ")"
                    End If
                    gStream.WriteText tStr + "<BR>", adWriteLine
                End If
                '
                If Not IsNull(!c_sequence) Then
                    tStr = "Sequence: " + Trim(Str(!c_sequence)) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                If Not IsNull(!c_dynasty) Then
                    tStr = "Dynasty: " + Trim(!c_dynasty) + " " + Trim(!c_dynasty_chn) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                If Not IsNull(!c_firstyear) Then
                    tStr = "First Year: " + Trim(Str(!c_firstyear)) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                If Not IsNull(!c_lastyear) Then
                    tStr = "Last Year: " + Trim(Str(!c_lastyear)) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                If Not IsNull(!c_category_desc) Then
                    tStr = "Category: " + Trim(!c_category_desc) + " " + Trim(!c_category_desc_chn) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                If Not IsNull(!c_appt_desc) Then
                    tStr = "Appointment Type: " + Trim(!c_appt_desc) + " " + Trim(!c_appt_desc_chn) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                If Not IsNull(!c_assume_office_desc) Then
                    tStr = "Assuming Office: " + Trim(!c_assume_office_desc) + " " + Trim(!c_assume_office_desc_chn) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                If Not IsNull(!c_inst_name_py) Then
                    tStr = "Institution: " + Trim(!c_inst_name_py) + " " + Trim(!c_inst_name_hz) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                If Not IsNull(!c_title) Then
                    tStr = "Source: " + Trim(!c_title) + " " + Trim(!c_title_chn)
                    If Not IsNull(!c_pages) Then
                        tStr = tStr + " (" + Trim(!c_pages) + ")"
                    End If
                    gStream.WriteText tStr + "<BR>", adWriteLine
                End If
                '
                If Not IsNull(!c_notes) Then
                    tStr = "Notes: " + Trim(!c_notes) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                gStream.WriteText "</P>", adWriteLine
                '
                .MoveNext
            Loop
        End With
    End If
    '
    ' now Status
    '
    tQueryStr = "SELECT STATUS_DATA.c_personid, STATUS_DATA.c_sequence, STATUS_DATA.c_status_code, STATUS_CODES.c_status_desc, STATUS_CODES.c_status_desc_chn, " + _
                "STATUS_DATA.c_firstyear, STATUS_DATA.c_lastyear, STATUS_DATA.c_source, TEXT_CODES.c_title_chn, TEXT_CODES.c_title, STATUS_DATA.c_pages, " + _
                "STATUS_DATA.c_notes " + _
            "FROM ( STATUS_CODES INNER JOIN STATUS_DATA ON ( STATUS_CODES.c_status_code = STATUS_DATA.c_status_code ) " + _
                "AND (STATUS_CODES.c_status_code = STATUS_DATA.c_status_code ) ) LEFT JOIN TEXT_CODES ON STATUS_DATA.c_source = TEXT_CODES.c_textid " + _
                "WHERE (((STATUS_DATA.c_personid)=" + tStrPersonID + "))"

    Set tRst = CurrentDb.OpenRecordset(tQueryStr)
    
    MsgBox "Writing status data"
    If Not (tRst.EOF) Then
        '
        tRst.MoveLast
        tStr = "<P><B><I>Social Distinction</I></B> "
        If tRst.RecordCount = 1 Then
            tStr = tStr + "(1 record)"
        Else
            tStr = tStr + "(" + Trim(Str(tRst.RecordCount)) + " records)"
        End If
        gStream.WriteText tStr + "</P>", adWriteLine
        '
        ' !c_personid, !c_status_desc, !c_status_desc_chn,
        ' !c_firstyear, !c_lastyear, !c_source, !c_title_chn, !c_title, !c_pages, !c_notes
        '
        With tRst
            .MoveFirst
            Do While Not .EOF
                '
                tStr = "<P><I>Status</I>: " + Trim(!c_status_desc) + " " + Trim(!c_status_desc_chn) + "<BR>"
                gStream.WriteText tStr, adWriteLine
                '
                If Not IsNull(!c_firstyear) Then
                    tStr = "First Year: " + Trim(Str(!c_firstyear)) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                If Not IsNull(!c_lastyear) Then
                    tStr = "Last Year: " + Trim(Str(!c_lastyear)) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                If Not IsNull(!c_title) Then
                    tStr = "Source: " + Trim(!c_title) + " " + Trim(!c_title_chn)
                    If Not IsNull(!c_pages) Then
                        tStr = tStr + " (" + Trim(!c_pages) + ")"
                    End If
                    gStream.WriteText tStr + "<BR>", adWriteLine
                End If
                '
                If Not IsNull(!c_notes) Then
                    tStr = "Notes: " + Trim(!c_notes) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                gStream.WriteText "</P>", adWriteLine
                '
                .MoveNext
            Loop
        End With
    End If
    '
    '  now Organizations: once again, the outer join for the address requires copying the data to a scratch file
    '
    cmdSQL.CommandText = "DELETE * FROM ZZ_SCRATCH_BIOG_INST_DATA"
    cmdSQL.Execute tRecCount
    '
    tQueryStr = "INSERT INTO ZZ_SCRATCH_BIOG_INST_DATA ( c_personid, c_inst_name_hz, c_inst_name_py, c_bi_role_desc, c_bi_role_chn, c_bi_begin_year, c_bi_end_year, " + _
                "c_source, c_source_chn, c_source_py, c_pages, c_notes, c_inst_addr_type_code, c_inst_addr_id, inst_xcoord, inst_ycoord ) " + _
            "SELECT BIOG_INST_DATA.c_personid, SOCIAL_INSTITUTION_NAME_CODES.c_inst_name_hz, SOCIAL_INSTITUTION_NAME_CODES.c_inst_name_py, " + _
                "BIOG_INST_CODES.c_bi_role_desc, BIOG_INST_CODES.c_bi_role_chn, BIOG_INST_DATA.c_bi_begin_year, BIOG_INST_DATA.c_bi_end_year, " + _
                "BIOG_INST_DATA.c_source, TEXT_CODES.c_title_chn, TEXT_CODES.c_title, BIOG_INST_DATA.c_pages, BIOG_INST_DATA.c_notes, " + _
                "SOCIAL_INSTITUTION_ADDR.c_inst_addr_type_code, SOCIAL_INSTITUTION_ADDR.c_inst_addr_id, " + _
                "SOCIAL_INSTITUTION_ADDR.inst_xcoord, SOCIAL_INSTITUTION_ADDR.inst_ycoord " + _
            "FROM (TEXT_CODES RIGHT JOIN (BIOG_INST_CODES INNER JOIN (BIOG_INST_DATA INNER JOIN SOCIAL_INSTITUTION_NAME_CODES " + _
                "ON BIOG_INST_DATA.c_inst_name_code = SOCIAL_INSTITUTION_NAME_CODES.c_inst_name_code) " + _
                "ON BIOG_INST_CODES.c_bi_role_code = BIOG_INST_DATA.c_bi_role_code ) " + _
                "ON TEXT_CODES.c_textid = BIOG_INST_DATA.c_source ) LEFT JOIN SOCIAL_INSTITUTION_ADDR " + _
                "ON ( BIOG_INST_DATA.c_inst_code = SOCIAL_INSTITUTION_ADDR.c_inst_code ) " + _
                    "AND ( BIOG_INST_DATA.c_inst_name_code = SOCIAL_INSTITUTION_ADDR.c_inst_name_code ) " + _
                "WHERE (((BIOG_INST_DATA.c_personid) = " + tStrPersonID + ")) "
        '
    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecCount

    MsgBox "Writing institution data"
    If (tRecCount > 0) Then
        '
        '  now get the address data
        '
        tQueryStr = "UPDATE ( ZZ_SCRATCH_BIOG_INST_DATA LEFT JOIN SOCIAL_INSTITUTION_ADDR_TYPES " + _
                        "ON ZZ_SCRATCH_BIOG_INST_DATA.c_inst_addr_type_code = SOCIAL_INSTITUTION_ADDR_TYPES.c_inst_addr_type_code ) " + _
                        "LEFT JOIN ADDR_CODES ON ZZ_SCRATCH_BIOG_INST_DATA.c_inst_addr_id = ADDR_CODES.c_addr_id " + _
                    "SET ZZ_SCRATCH_BIOG_INST_DATA.c_inst_addr_type_desc = [SOCIAL_INSTITUTION_ADDR_TYPES].[c_inst_addr_type_desc], " + _
                        "ZZ_SCRATCH_BIOG_INST_DATA.c_inst_addr_type_chn = [SOCIAL_INSTITUTION_ADDR_TYPES].[c_inst_addr_type_chn], " + _
                        "ZZ_SCRATCH_BIOG_INST_DATA.c_inst_addr_pinyin = [ADDR_CODES].[c_name], " + _
                        "ZZ_SCRATCH_BIOG_INST_DATA.c_inst_addr_chn = [ADDR_CODES].[c_name_chn]"
'
        cmdSQL.CommandText = tQueryStr
        cmdSQL.Execute tRecCount
            '
        Set tRst = CurrentDb.OpenRecordset("ZZ_SCRATCH_BIOG_INST_DATA")
            tRst.MoveLast
            tStr = "<P><B><I>Institurions</I></B> "
            If tRst.RecordCount = 1 Then
                tStr = tStr + "(1 record)"
            Else
                tStr = tStr + "(" + Trim(Str(tRst.RecordCount)) + " records)"
            End If
            gStream.WriteText tStr + "</P>", adWriteLine
        '
        ' !c_personid, !c_inst_name_hz, !c_inst_name_py, _
        ' !c_inst_type_hz, !c_inst_addr_pinyin, !c_inst_addr_chn,
        ' !c_inst_addr_type_desc, !c_inst_addr_type_chn, !c_bi_role_desc,
        ' !c_bi_role_chn, !c_bi_begin_year, !c_bi_end_year,
        ' !c_source_chn, !c_source_py, !c_pages, !c_notes, !inst_xcoord, !inst_ycoord
        '
        With tRst
            .MoveFirst
            Do While Not .EOF
                '
                tStr = "<P><I>Institution</I>: " + Trim(!c_inst_name_py) + " " + Trim(!c_inst_name_hz) + "<BR>"
                gStream.WriteText tStr, adWriteLine
                '
                If Not IsNull(!c_inst_type_hz) Then
                    tStr = "Type: " + Trim(!c_inst_type_hz) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                
                If Not IsNull(!c_inst_addr_pinyin) Then
                    tStr = "Location: " + Trim(!c_inst_addr_pinyin) + " " + Trim(!c_inst_addr_chn)
                    If Not IsNull(!inst_xcoord) Then
                        tStr = tStr + " (" + Trim(Str(!inst_xcoord)) + tC + Trim(Str(!inst_ycoord)) + ")"
                    End If
                    If Not IsNull(!c_inst_addr_type_desc) Then
                        gStream.WriteText tStr + "<BR>", adWriteLine
                        tStr = "Type: " + !c_inst_addr_type_desc + " " + !c_inst_addr_type_chn
                    End If
                    gStream.WriteText tStr + "<BR>", adWriteLine
                End If

                tStr = "Role: " + Trim(!c_bi_role_desc) + " " + Trim(!c_bi_role_chn) + "<BR>"
                gStream.WriteText tStr, adWriteLine
                
                If Not IsNull(!c_bi_begin_year) Then
                    tStr = "First Year: " + Trim(Str(!c_bi_begin_year)) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                If Not IsNull(!c_bi_end_year) Then
                    tStr = "Last Year: " + Trim(Str(!c_bi_end_year)) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                If Not IsNull(!c_source_py) Then
                    tStr = "Source: " + Trim(!c_source_py) + " " + Trim(!c_source_chn)
                    If Not IsNull(!c_pages) Then
                        tStr = tStr + " (" + Trim(!c_pages) + ")"
                    End If
                    gStream.WriteText tStr + "<BR>", adWriteLine
                End If
                '
                If Not IsNull(!c_notes) Then
                    tStr = "Notes: " + Trim(!c_notes) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                gStream.WriteText "</P>", adWriteLine
                '
                .MoveNext
            Loop
        End With
    End If
    '
    '  now Texts: once again, because of an outer join, I need to dump the records to a scratch table
    '
    cmdSQL.CommandText = "DELETE * FROM ZZ_SCRATCH_BIOG_TEXT_DATA"
    cmdSQL.Execute tRecCount
    '
    tQueryStr = "INSERT INTO ZZ_SCRATCH_BIOG_TEXT_DATA ( c_personid, c_textid, c_role_id, c_year, c_title_chn, c_title, c_role_desc, c_role_desc_chn, " + _
                "c_text_cat_code, c_source, c_pages, c_notes ) " + _
            "SELECT BIOG_TEXT_DATA.c_personid, BIOG_TEXT_DATA.c_textid, BIOG_TEXT_DATA.c_role_id, BIOG_TEXT_DATA.c_year, TEXT_CODES.c_title_chn, " + _
                "TEXT_CODES.c_title, TEXT_ROLE_CODES.c_role_desc, TEXT_ROLE_CODES.c_role_desc_chn, TEXT_CODES.c_bibl_cat_code, " + _
                "BIOG_TEXT_DATA.c_source, BIOG_TEXT_DATA.c_pages, BIOG_TEXT_DATA.c_notes " + _
            "FROM TEXT_ROLE_CODES INNER JOIN ( TEXT_CODES INNER JOIN BIOG_TEXT_DATA ON TEXT_CODES.c_textid = BIOG_TEXT_DATA.c_textid ) " + _
                "ON TEXT_ROLE_CODES.c_role_id = BIOG_TEXT_DATA.c_role_id " + _
            "WHERE (((BIOG_TEXT_DATA.c_personid) = " + tStrPersonID + ")) "

    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecCount
    '
    
    MsgBox "Writing text data"
    If tRecCount > 0 Then
        '
    ' source data
    '
    tQueryStr = "UPDATE ( ( ZZ_SCRATCH_BIOG_TEXT_DATA LEFT JOIN TEXT_BIBLCAT_CODES ON ZZ_SCRATCH_BIOG_TEXT_DATA.c_text_cat_code = TEXT_BIBLCAT_CODES.c_text_cat_code ) " + _
                    "LEFT JOIN TEXT_CODES ON ZZ_SCRATCH_BIOG_TEXT_DATA.c_source = TEXT_CODES.c_textid ) LEFT JOIN TEXT_BIBLCAT_CODE_TYPE_REL " + _
                    "ON ZZ_SCRATCH_BIOG_TEXT_DATA.c_text_cat_code = TEXT_BIBLCAT_CODE_TYPE_REL.c_text_cat_code " + _
                "SET ZZ_SCRATCH_BIOG_TEXT_DATA.c_text_cat_desc = [TEXT_BIBLCAT_CODES].[c_text_cat_desc], " + _
                    "ZZ_SCRATCH_BIOG_TEXT_DATA.c_text_cat_desc_chn = [TEXT_BIBLCAT_CODES].[c_text_cat_desc_chn], " + _
                    "ZZ_SCRATCH_BIOG_TEXT_DATA.c_text_cat_pinyin = [TEXT_BIBLCAT_CODES].[c_text_cat_pinyin], " + _
                    "ZZ_SCRATCH_BIOG_TEXT_DATA.c_source_chn = [TEXT_CODES].[c_title_chn], " + _
                    "ZZ_SCRATCH_BIOG_TEXT_DATA.c_source_title = [TEXT_CODES].[c_title], " + _
                    "ZZ_SCRATCH_BIOG_TEXT_DATA.c_text_cat_type_id = [TEXT_BIBLCAT_CODE_TYPE_REL].[c_text_cat_type_id]"

    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecCount
    '
    ' now BIBLCAT_TYPE data
    '
    tQueryStr = "UPDATE ZZ_SCRATCH_BIOG_TEXT_DATA LEFT JOIN TEXT_BIBLCAT_TYPES ON ZZ_SCRATCH_BIOG_TEXT_DATA.c_text_cat_type_id = TEXT_BIBLCAT_TYPES.c_text_cat_type_id " + _
                "SET  ZZ_SCRATCH_BIOG_TEXT_DATA.c_text_cat_type_desc = [TEXT_BIBLCAT_TYPES].[c_text_cat_type_desc], " + _
                    "ZZ_SCRATCH_BIOG_TEXT_DATA.c_text_cat_type_desc_chn = [TEXT_BIBLCAT_TYPES].[c_text_cat_type_desc_chn]"

    cmdSQL.CommandText = tQueryStr
    cmdSQL.Execute tRecCount
    '
    Set tRst = CurrentDb.OpenRecordset("ZZ_SCRATCH_BIOG_TEXT_DATA")
      tRst.MoveLast
        tStr = "<P><B><I>Texts</I></B> "
        If tRst.RecordCount = 1 Then
            tStr = tStr + "(1 record)"
        Else
            tStr = tStr + "(" + Trim(Str(tRst.RecordCount)) + " records)"
        End If
        gStream.WriteText tStr + "</P>", adWriteLine
        '
        ' !c_title, !c_title_chn, !c_bibl_cat_desc, !c_bibl_cat_desc_chn,
        ' !c_role_desc, !c_role_desc_chn,
        ' !c_source_title, !c_source_chn, !c_pages, !c_notes,
        
        With tRst
            .MoveFirst
            Do While Not .EOF
                '
                tStr = "<P><I>Text</I>: "
                If Not IsNull(!c_title) Then
                    tStr = tStr + Trim(!c_title) + " "
                End If
                If IsNull(!c_title_chn) Then
                    tStr = tStr + "Title Missing"
                Else
                    tStr = tStr + Trim(!c_title_chn)
                End If
                tStr = tStr + "<BR>"
                gStream.WriteText tStr, adWriteLine
                '
                If Not IsNull(!c_bibl_cat_desc) Then
                    tStr = "Category: " + Trim(!c_bibl_cat_desc) + " " + Trim(!c_bibl_cat_desc_chn) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                If Not IsNull(!c_role_desc) Then
                    tStr = "Role: " + Trim(!c_role_desc) + " " + Trim(!c_role_desc_chn) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                If Not IsNull(!c_source_title) Then
                    tStr = "Source: " + Trim(!c_source_title) + " " + Trim(!c_source_chn)
                    If Not IsNull(!c_pages) Then
                        tStr = tStr + " (" + Trim(!c_pages) + ")"
                    End If
                    gStream.WriteText tStr + "<BR>", adWriteLine
                End If
                '
                If Not IsNull(!c_notes) Then
                    tStr = "Notes: " + Trim(!c_notes) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                gStream.WriteText "</P>", adWriteLine
                '
                .MoveNext
            Loop
        End With
    End If
    '
    '  now Sources
    '
    tQueryStr = "SELECT BIOG_SOURCE_DATA.c_personid, BIOG_SOURCE_DATA.c_textid, BIOG_SOURCE_DATA.c_pages, BIOG_SOURCE_DATA.c_notes, TEXT_CODES.c_title_chn, " + _
            "TEXT_CODES.c_title, IIf(IsNull([TEXT_CODES].[c_url_api]), '', [TEXT_CODES].[c_url_api]) + " + _
            "IIf(IsNull([TEXT_CODES].[c_url_api]),'', [BIOG_SOURCE_DATA].[c_pages]) + " + _
            "IIf( IsNull([TEXT_CODES].[c_url_api_coda]), '', [TEXT_CODES].[c_url_api_coda]) AS c_hyperlink " + _
        "FROM TEXT_CODES INNER JOIN BIOG_SOURCE_DATA ON TEXT_CODES.c_textid = BIOG_SOURCE_DATA.c_textid " + _
        "WHERE  (((BIOG_SOURCE_DATA.c_personid) = " + tStrPersonID + "))"
    Set tRst = CurrentDb.OpenRecordset(tQueryStr)
    
    MsgBox "Writing source data"
    If Not (tRst.EOF) Then
        '
        tRst.MoveLast
        tStr = "<P><B><I>Sources</I></B> "
        If tRst.RecordCount = 1 Then
            tStr = tStr + "(1 record)"
        Else
            tStr = tStr + "(" + Trim(Str(tRst.RecordCount)) + " records)"
        End If
        gStream.WriteText tStr + "</P>", adWriteLine
        '
        ' !c_title_chn, !c_title,
        ' !c_pages, !c_notes, !c_hyperlink
        '
        With tRst
            .MoveFirst
            Do While Not .EOF
                '
                tStr = "<P><I>Source</I>: " + Trim(!c_title) + " " + Trim(!c_title_chn)
                If Not IsNull(!c_pages) Then
                    tStr = tStr + " (" + Trim(!c_pages) + ")"
                End If
                gStream.WriteText tStr + "<BR>", adWriteLine
                '
                If Not IsNull(!c_hyperlink) Then
                    tStr = "Link: " + Trim(!c_hyperlink) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                If Not IsNull(!c_notes) Then
                    tStr = "Notes: " + Trim(!c_notes) + "<BR>"
                    gStream.WriteText tStr, adWriteLine
                End If
                '
                gStream.WriteText "</P>", adWriteLine
                '
                .MoveNext
            Loop
        End With
    End If
    '
    'Set the object variable to Nothing.
    Set dlgSaveAs = Nothing
    gStream.WriteText "</BODY>", adWriteLine
    gStream.WriteText "</HTML>", adWriteLine
    ' now make sure all the data is copied to tStream
    gStream.Flush
    ' and write the stream to the file
    gStream.SaveToFile tFileName, adSaveCreateOverWrite
    '
    gStream.Close
    ' MsgBox "Finished"
    MsgBox "Finished saving to File"
    
Exit_CmdSaveToFile_Click:
    Exit Sub

Err_CmdSaveToFile_Click:
    MsgBox Err.Description
    Resume Exit_CmdSaveToFile_Click
    
End Sub

Private Sub CmdSearchByOffice_Click()
    Dim stDocName As String, stLinkCriteria As String
    Dim tRstSearch As DAO.Recordset

    stDocName = "frmSearchPeopleOffice"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , acDialog
    
    If CurrentProject.AllForms("frmSearchPeopleOffice").IsLoaded Then
        Set tRstSearch = CurrentDb.OpenRecordset("Z_NAME_SEARCH", dbOpenDynaset)
        If tRstSearch.RecordCount > 0 Then
            Me.CmdClearSearch.Enabled = True
            Set Me.frmPeopleLookup2.Form.Recordset = tRstSearch
        End If
                
        DoCmd.Close acForm, stDocName
    End If
            
End Sub

Private Sub CmdStoreID_Click()
    Dim cmdSQL As ADODB.Command, tRecCount As Variant, tRst As DAO.Recordset, tID As Long
    
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
    '
    ' get the value
    '
    Set tRst = Me.frmPeopleLookup2.Form.Recordset
    tID = tRst!c_personid

    cmdSQL.CommandText = "INSERT INTO ZZ_STORE_PERSON_ID ( c_personid ) SELECT " + Str(tID) + " AS c_personid"
    cmdSQL.Execute tRecCount
    MsgBox "Person IDs successfully stored.  Click on 'Recall Person IDs' to reuse these IDs in other forms."
End Sub


Private Sub Form_Open(Cancel As Integer)
    BIOG_MAIN_2_Subform.Form.OrderBy = "c_personid"
    BIOG_MAIN_2_Subform.Form.OrderByOn = True
    frmPeopleLookup2.Form.OrderBy = "c_name"
    frmPeopleLookup2.Form.OrderByOn = True

    ' set the language
    Dim tmli As MsoLanguageID
    ' get the labels
    tmli = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    gLabelsOK = True
    If tmli = msoLanguageIDSimplifiedChinese Then
        gDisplayLanguage = "S"
    ElseIf tmli = msoLanguageIDTraditionalChinese Then
        gDisplayLanguage = "T"
    ElseIf tmli = msoLanguageIDEnglishUS Then
        gDisplayLanguage = "E"
    Else
        gDisplayLanguage = "E"
    End If
    Call changeDisplayLanguage
    
    gPersonID = 0

End Sub
Private Sub CmdSearch_Click()
On Error GoTo Err_CmdSearch_Click
    
    Dim tRstSearch As DAO.Recordset, tStr As String, tQt As String, tQuery As QueryDef, tStrName As String
    Dim cmdSQL As ADODB.Command, tRecNum As Long
    
    tQt = Chr(34)
    
    '  first make sure that the browser recordset is a dummy
    
    Set tRstSearch = CurrentDb.OpenRecordset("Z_SCRATCH_DUMMY_NAME_SEARCH", dbOpenDynaset)
    Set Me.frmPeopleLookup2.Form.Recordset = tRstSearch
    
    '  Now zap ZZ_NAME_SEARCH
    
    Set cmdSQL = New ADODB.Command
    cmdSQL.ActiveConnection = CurrentProject.Connection
    cmdSQL.CommandType = adCmdText
    '
    cmdSQL.CommandText = "Delete * from ZZ_NAME_SEARCH"
    cmdSQL.Execute tRecNum
    
    ' now populate from ZZZ_NAMES
    
    If IsNull(TxtNameChn.Value) Then
        If IsNull(TxtName.Value) Then
            tStr = "Quit"
        Else
            If Me.TxtName.Value = "" Then
                tStr = "Quit"
            Else
                tStrName = TxtName.Value
                If Left(tStrName, 1) = "!" Then
                    tStrName = Mid(TxtName.Value, 2)
                    tStr = " Left(c_name," + Str(Len(tStrName)) + ") = " + tQt + Trim(tStrName) + tQt
                ElseIf UCase(Left(tStrName, 1)) = Left(tStrName, 1) Then
                    tStr = " Left(c_name," + Str(Len(tStrName)) + ") = " + tQt + Trim(tStrName) + tQt + _
                        " OR c_name LIKE " + tQt + "%" + " " + Trim(tStrName) + "%" + tQt
                Else
                    tStr = " c_name LIKE " + tQt + "%" + Trim(tStrName) + "%" + tQt
                End If
            End If
        End If
    Else
        If Me.TxtNameChn.Value = "" Then
            If IsNull(TxtName.Value) Then
                tStr = "Quit"
            Else
                If Me.TxtName.Value = "" Then
                    tStr = "Quit"
                Else
                    tStrName = TxtName.Value
                    If Left(tStrName, 1) = "!" Then
                        tStrName = Mid(TxtName.Value, 2)
                        tStr = " Left(c_name," + Str(Len(tStrName)) + ") = " + tQt + Trim(tStrName) + tQt
                    ElseIf UCase(Left(tStrName, 1)) = Left(tStrName, 1) Then
                        tStr = " Left(c_name," + Str(Len(tStrName)) + ") = " + tQt + Trim(tStrName) + tQt + _
                            " OR c_name LIKE " + tQt + "%" + " " + Trim(tStrName) + "%" + tQt
                    Else
                        tStr = " c_name LIKE " + tQt + "%" + Trim(tStrName) + "%" + tQt
                    End If
                End If
            End If
        Else
            tStr = " c_name_chn LIKE " + tQt + "%" + Trim(TxtNameChn.Value) + "%" + tQt
        End If
    End If
        
    If Not (tStr = "Quit") Then
        tStr = "INSERT INTO ZZ_NAME_SEARCH SELECT c_personid, c_name, c_name_chn " + _
            "FROM ZZZ_NAMES WHERE" + tStr
        
        cmdSQL.CommandText = tStr
        cmdSQL.Execute tRecNum
    End If
    
    Set tRstSearch = CurrentDb.OpenRecordset("ZZ_NAME_SEARCH", dbOpenDynaset)
    If tRstSearch.RecordCount = 0 Then
        Set tRstSearch = CurrentDb.OpenRecordset("BIOG_MAIN", dbOpenDynaset)
        Me.CmdClearSearch.Enabled = False
    Else
        Me.CmdClearSearch.Enabled = True
    End If
    ' tRstSearch.Index = "c_name"
    Set Me.frmPeopleLookup2.Form.Recordset = tRstSearch
    

Exit_CmdSearch_Click:
    Exit Sub

Err_CmdSearch_Click:
    MsgBox Err.Description
    Resume Exit_CmdSearch_Click

    
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
Private Sub changeDisplayLanguage()
    Dim tLabelLanguage(3, 8) As String, tLang As Integer
    
    Dim tRstLabelList As DAO.Recordset, ti As Integer
    
    Set tRstLabelList = CurrentDb.OpenRecordset("FormLabels", dbOpenTable)
    
    tRstLabelList.Index = "label"
    
    gLabelsOK = False
    With tRstLabelList
        .MoveFirst
        ti = 1
        
        Do While ti < 8 And Not .EOF
            If !c_form = "BROWSE" Then
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
        Me.CmdSearch.Caption = tLabelLanguage(tLang, 1)
        Me.CmdFanti.Caption = tLabelLanguage(tLang, 2)
        Me.CmdJianti.Caption = tLabelLanguage(tLang, 3)
        Me.CmdClearSearch.Caption = tLabelLanguage(tLang, 5)
        Me.CmdSaveToFile.Caption = tLabelLanguage(tLang, 7)
        
        ' now do the subform
        BIOG_MAIN_2_Subform.Form.gDisplayLanguage = gDisplayLanguage
        BIOG_MAIN_2_Subform.Form.changeDisplayLanguage
    End If
    
End Sub

Private Sub TxtNameChn_LostFocus()
    '
    ' clear the pinyin name
    '
    TxtName.Value = ""
    
End Sub

Private Sub TxtName_LostFocus()
    '
    ' clear the Chinese name
    '
    TxtNameChn.Value = ""
End Sub
Private Sub CmdClearSearch_Click()
On Error GoTo Err_CmdClearSearch_Click

    Dim tRst As DAO.Recordset, tQuery As QueryDef, tID As Long, tQueryStr As String
    
    Set tRst = Me.frmPeopleLookup2.Form.Recordset
    
    tID = tRst!c_personid
    
    tQueryStr = "SELECT BIOG_MAIN.c_personid, BIOG_MAIN.c_name, BIOG_MAIN.c_name_chn" + _
        " FROM BIOG_MAIN ORDER BY BIOG_MAIN.c_name"
    
    'Set tQuery = CurrentDb.QueryDefs("Selected PERSON_NAME_DATA Query")
    'Set tRst = tQuery.OpenRecordset(dbOpenDynaset)
    'Set tRst = CurrentDb.OpenRecordset("BIOG_MAIN", dbOpenDynaset)
    Set tRst = CurrentDb.OpenRecordset(tQueryStr)
    'tRst.Index = "c_name"
    tRst.FindFirst "c_personid = " + Trim(Str(tID))
    
    Set Me.frmPeopleLookup2.Form.Recordset = tRst

    Me.CmdClearSearch.Enabled = False
    Me.TxtName.Value = ""
    Me.TxtNameChn.Value = ""

Exit_CmdClearSearch_Click:
    Exit Sub

Err_CmdClearSearch_Click:
    MsgBox Err.Description
    Resume Exit_CmdClearSearch_Click
    
End Sub

