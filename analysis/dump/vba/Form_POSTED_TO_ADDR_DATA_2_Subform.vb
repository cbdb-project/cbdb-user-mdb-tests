Option Compare Database
Public gDisplayLanguage As String, gLabelsOK As Boolean, gTest As Integer

Public Sub changeDisplayLanguage()
    Dim tLabelLanguage(3, 23) As String, tLang As Integer
    
    Dim tRstLabelList As DAO.Recordset, ti As Integer
    
    Set tRstLabelList = CurrentDb.OpenRecordset("FormLabels", dbOpenTable)
    
    tRstLabelList.Index = "label"
    
    gLabelsOK = False
    With tRstLabelList
        .MoveFirst
        ti = 24
        
        'Do While tI < 23 And Not .EOF
        '    If !c_form = "SF_POFF" Then
        '        gLabelsOK = True
        '        If tI <> !c_label_id Then
        '            MsgBox "Uh oh:  mismatched label table"
        '            gLabelsOK = False
        '            Exit Do
        '        End If
        '        tLabelLanguage(1, tI) = !c_english
        '        tLabelLanguage(2, tI) = !c_fanti
        '        tLabelLanguage(3, tI) = !c_jianti
        '        tI = tI + 1
        '    End If
        '    .MoveNext
        'Loop
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
        'Me.LblSequence.Caption = tLabelLanguage(tLang, 1)
        'Me.LblOfficeCat.Caption = tLabelLanguage(tLang, 2)
        'Me.LblPostCat.Caption = tLabelLanguage(tLang, 3)
        'Me.LblFirstyear.Caption = tLabelLanguage(tLang, 4)
        'Me.LblFYRange.Caption = tLabelLanguage(tLang, 5)
        'Me.LblAssume.Caption = tLabelLanguage(tLang, 6)
        'Me.LblFYNHYear.Caption = tLabelLanguage(tLang, 7)
        'Me.LblFYIintercalary.Caption = tLabelLanguage(tLang, 8)
        'Me.LblFYGZ.Caption = tLabelLanguage(tLang, 9)
        'Me.LblLastyear.Caption = tLabelLanguage(tLang, 10)
        'Me.LblLYRange.Caption = tLabelLanguage(tLang, 11)
        'Me.LblLYNHYear.Caption = tLabelLanguage(tLang, 12)
        'Me.LblLYIntercalary.Caption = tLabelLanguage(tLang, 13)
        'Me.LblLYGZ.Caption = tLabelLanguage(tLang, 14)
        'Me.LblPages.Caption = tLabelLanguage(tLang, 15)
        'Me.LblNotes.Caption = tLabelLanguage(tLang, 16)
        'Me.CmdPickOffice.Caption = tLabelLanguage(tLang, 17)
        'Me.CmdPickFY_NH.Caption = tLabelLanguage(tLang, 18)
        'Me.CmdPickLY_NH.Caption = tLabelLanguage(tLang, 19)
        'Me.CmdPickSource.Caption = tLabelLanguage(tLang, 20)
        'Me.CmdAddNew.Caption = tLabelLanguage(tLang, 21)
        'Me.CmdDelete.Caption = tLabelLanguage(tLang, 22)
        
    End If
    
End Sub
Private Sub CmdPlace_Click()
On Error GoTo Err_CmdPlace_Click


    If Me.Dirty Then Me.Dirty = False
    DoCmd.Close

Exit_CmdPlace_Click:
    Exit Sub

Err_CmdPlace_Click:
    MsgBox Err.Description
    Resume Exit_CmdPlace_Click
    
End Sub